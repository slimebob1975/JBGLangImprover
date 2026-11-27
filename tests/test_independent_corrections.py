"""Bounded recovery of model-proposed, independently safe corrections."""
import json
import logging
from unittest.mock import patch
from zipfile import ZipFile

import pytest
from docx import Document
from lxml import etree

from app.src.JBGIndependentCorrections import independent_corrections
from app.src.JBGDocumentStructureExtractor import DocumentStructureExtractor
from app.src.JBGChangePlanner import ChangePlanner
from app.src.JBGTokenDiffEngine import TokenDiffEngine
from app.src.JBGDocxPackage import DocxPackage, W_NS
from app.src.JBGTrackedChangesRenderer import TrackedChangesRenderer
from app.src.JBGCommentsRenderer import CommentsRenderer
from tests.test_quality_guards import suggestor, element, proposal, ASSESSMENT
from tests.test_hard_hyphen import visible

LOG = logging.getLogger(__name__)
W = '{' + W_NS + '}'
ASSESS = dict(ASSESSMENT, priority='critical', category='correctness')
OLD = 'Det är viktigt att vi följer upp arbetet. A-kassorna är har det yttersta ansvaret.'
NEW = 'Vi behöver följa upp arbetet. A-kassorna har det yttersta ansvaret.'
OLD2 = 'Det kan möjliggöra bättre arbete. IAF menar att rollfördelningen inom sådana samarbeten är tydliga.'
NEW2 = 'Det kan bidra till bättre arbete. IAF menar att rollfördelningen inom sådana samarbeten är tydlig.'


def run(old=OLD, new=NEW, text=None, assessment=ASSESS, enabled=True, extra=()):
    ai = suggestor([element(old if text is None else text)], enabled)
    raw = [proposal(old, new, assessment=assessment), *extra]
    val = ai._postprocess_suggestions(raw)
    return ai._filter_with_quality_guards(val)


@pytest.mark.parametrize('old,new,expected_old,expected_new,parent_guard', [
    (OLD, NEW, 'A-kassorna är har det', 'A-kassorna har det', 'guard_modality'),
    (OLD2, NEW2, 'rollfördelningen inom sådana samarbeten är tydliga',
     'rollfördelningen inom sådana samarbeten är tydlig', None),
])
def test_recover_one_repair_and_preserve_parent_diagnostics(old, new, expected_old, expected_new, parent_guard):
    accepted, rows = run(old, new)
    assert len(rows) == 2 and len(accepted) == 1
    assert not rows[0].accepted and rows[1].accepted
    assert (accepted[0].old, accepted[0].new) == (expected_old, expected_new)
    assert accepted[0].model_assessment is None
    assert accepted[0].assessment_status == 'missing'
    assert accepted[0].motivation != rows[0].suggestion.motivation
    assert 'split_parent_replaced' in {i.code for i in rows[0].issues}
    assert 'split_recovered' in {i.code for i in rows[1].issues}
    if parent_guard:
        assert parent_guard in {i.code for i in rows[0].issues}


@pytest.mark.parametrize('assessment', [None, {}, dict(ASSESS, meaning_risk='medium'),
    dict(ASSESS, meaning_risk='high'), dict(ASSESS, confidence='medium'),
    dict(ASSESS, category='clarity'), dict(ASSESS, requires_rewrite=True)])
def test_no_recovery_without_qualified_parent(assessment):
    _, rows = run(assessment=assessment)
    assert not any(i.code.startswith('split_') for row in rows for i in row.issues)


def test_no_recovery_for_legacy():
    _, rows = run(enabled=False)
    assert len(rows) == 1 and not rows[0].accepted


@pytest.mark.parametrize('source', [OLD + ' ' + OLD, 'TODO: ' + OLD,
    'Källa: ' + OLD, '”' + OLD + '”'])
def test_no_bypass_of_ambiguous_or_protected_source(source):
    accepted, rows = run(text=source)
    assert not accepted
    assert not any(i.code == 'split_recovered' for row in rows for i in row.issues)


def test_no_recovery_with_competing_proposals():
    _, rows = run(extra=[proposal('yttersta ansvaret', 'fulla ansvaret')])
    assert not any(i.code.startswith('split_') for row in rows for i in row.issues)


@pytest.mark.parametrize('old,new,source', [
    ('A-kassorna är har det yttersta ansvaret.', 'A-kassorna har det yttersta ansvaret.', None),
    (OLD, NEW.replace('A-kassorna har det', 'A-kassorna har ett'), None),
    (OLD, NEW.replace('A-kassorna har det', 'A-kassorna har detta'), None),
    (OLD, NEW.replace('A-kassorna har det yttersta ansvaret.', 'De har det yttersta ansvaret.'), None),
    (OLD, NEW + ' En ny mening.', None),
    (OLD, NEW, OLD + ' A-kassorna är har det.'),
    ('Vi skriver en text. Vi granskar texten.', 'Vi formulerar en text. Vi kontrollerar texten.', None),
    (OLD2, NEW2.replace('tydlig.', 'tydligare.'), None),
    ('Det är viktigt att rollfördelningen är tydliga och arbetet ska följas upp.',
     'Rollfördelningen är tydlig och arbetet följs upp.', None),
])
def test_rule_does_not_guess_or_split_dependent_edits(old, new, source):
    assert independent_corrections(old, new, old if source is None else source) == []


@pytest.mark.parametrize('method', ['suggest_changes', 'suggest_changes_token_aware_batching'])
def test_pipeline_persists_parent_and_repair(tmp_path, method):
    ai = suggestor([element(OLD)])
    response = json.dumps([proposal(OLD, NEW, assessment=ASSESS)])
    with patch('app.src.JBGLangImprovSuggestorAI.create_openai_client', return_value=object()), \
         patch.object(ai, '_call_model', return_value=response), \
         patch.object(ai, '_call_model_for_chunks', return_value=[response]):
        getattr(ai, method)()
    assert len(ai.json_suggestions) == len(ai.validated_suggestions) == 1
    assert len(ai.filtered_review) == 2
    path = tmp_path / 'filter.json'
    ai.save_filter_report(path)
    rows = json.loads(path.read_text(encoding='utf-8'))
    assert rows[0]['accepted'] is False and rows[1]['accepted'] is True
    assert rows[1]['suggestion']['assessment_status'] == 'missing'
    assert 'model_assessment' not in rows[1]['suggestion']
    assert rows[1]['suggestion']['old'] == 'A-kassorna är har det'


@pytest.mark.parametrize('old,new,corrected', [(OLD, NEW, OLD.replace('är har', 'har')),
    (OLD2, NEW2, OLD2.replace('är tydliga', 'är tydlig'))])
def test_real_word_roundtrip_without_crossing_protected_hyphen(tmp_path, old, new, corrected):
    source, dest = tmp_path / 'input.docx', tmp_path / 'output.docx'
    doc = Document()
    p = doc.add_paragraph()
    # Keep an immutable hyphen in the same paragraph, outside the actual repair.
    p.add_run('Bakgrund om a')
    etree.SubElement(p.add_run()._r, W + 'noBreakHyphen')
    p.add_run('kassorna. ' + old)
    doc.save(source)
    structure = DocumentStructureExtractor(str(source), LOG).extract()
    e = next(e for e in structure['elements'] if e['type'] == 'paragraph')
    ai = suggestor(structure['elements'])
    val = ai._postprocess_suggestions([proposal(old, new, e['element_id'], assessment=ASSESS)])
    accepted, rows = ai._filter_with_quality_guards(val)
    assert len(accepted) == 1 and len(rows) == 2
    plans = ChangePlanner(structure, TokenDiffEngine(), LOG).build_plans(accepted)
    with DocxPackage(str(source), LOG) as package:
        results = TrackedChangesRenderer(package, LOG, word_level_changes=True).apply_plans(plans)
        assert len(results) == 1 and results[0].applied
        comments = CommentsRenderer(package, LOG).apply_comments_for_results(results)
        assert len(comments) == 1 and comments[0].applied
        package.save(str(dest))
    with ZipFile(dest) as z:
        root = etree.fromstring(z.read('word/document.xml'))
    p = root.find('.//' + W + 'body/' + W + 'p')
    assert visible(p, False) == 'Bakgrund om a-kassorna. ' + old
    assert visible(p, True) == 'Bakgrund om a-kassorna. ' + corrected
    assert len(list(p.iter(W + 'noBreakHyphen'))) == 1
    assert not any(n.tag in (W + 'ins', W + 'del') for h in p.iter(W + 'noBreakHyphen') for n in h.iterancestors())
