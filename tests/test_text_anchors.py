"""Regression tests for position-preserving minimization and fail-closed anchors."""
import logging
from dataclasses import replace
from zipfile import ZipFile

import pytest
from docx import Document
from lxml import etree

from app.src.JBGChangePlanner import ChangePlanner
from app.src.JBGDocxPackage import DocxPackage, W_NS
from app.src.JBGDocumentStructureExtractor import DocumentStructureExtractor
from app.src.JBGLangImprovSuggestorAI import SuggestedChange
from app.src.JBGReadabilityMetrics import apply_suggestions_to_text
from app.src.JBGSimpleMarkupRenderer import SimpleMarkupRenderer
from app.src.JBGTextAnchors import locate_span
from app.src.JBGTokenDiffEngine import TokenDiffEngine
from app.src.JBGTrackedChangesRenderer import TrackedChangesRenderer
from tests.test_hard_hyphen import visible
from tests.test_quality_guards import element, proposal, suggestor

LOG = logging.getLogger(__name__)
W = '{' + W_NS + '}'
NS = {'w': W_NS}

# Frozen, correctly contextualized model inputs, based on the failures in T.
CASES = [
    ('Kommunalarbetarnas a-kassa använder AI-systemet. Systemet förklarar reglerna och hur de ska tillämpas.',
     'hur de ska tillämpas', 'hur reglerna ska tillämpas'),
    ('Arbetsförmedlingen har redovisat både regelbaserade automationer och AI-verktyg. '
     'Systeminventeringen visar att systemen används i kärnprocesser och i flera fall påverkar beslut.',
     'och i flera fall påverkar', 'och att de i flera fall påverkar'),
    ('Systeminventering: en strukturerad kartläggning av tekniska system, till exempel automationer, '
     'och vad de används till', 'vad de används till', 'vad de används till.'),
    ('De riskbaserade kontrollerna har genomförts i arbetslöshetskassorna.',
     'i arbetslöshetskassorna', 'av arbetslöshetskassorna'),
    ('Vi använder verktyget och de kontrollerar reglerna.', 'de kontrollerar', 'handläggarna kontrollerar'),
    ('Vi har genomfört en  granskning av ärendet.', 'genomfört en granskning av', 'granskat'),
]


def prepare(text, old, new, ident='p1', kind='paragraph', assessment=None):
    ai = suggestor([element(text, ident, kind)], assess_changes=False)
    raw = proposal(old, new, ident, kind, assessment=assessment)
    return ai, ai._validate_docx_suggestion(raw), raw


@pytest.mark.parametrize('text,old,new', CASES)
def test_minimization_preserves_full_result_and_planner_position(text, old, new):
    ai, original, _ = prepare(text, old, new)
    stabilized = ai._minimize_and_filter_suggestion(original)
    assert stabilized is not None and stabilized.safe_to_apply, stabilized
    original_span = locate_span(old, text)
    assert original_span is not None
    a, b = original_span
    expected = text[:a] + new + text[b:]
    plans = ChangePlanner(ai.json_structured_document, TokenDiffEngine(), LOG).build_plans([stabilized])
    assert len(plans) == 1
    plan = plans[0]
    assert a <= plan.anchor.start < plan.anchor.end <= b
    assert text[:plan.anchor.start] + plan.new_text + text[plan.anchor.end:] == expected
    assert apply_suggestions_to_text(text, [dict(old=stabilized.old, new=stabilized.new)]) == (expected, 1, 0, 0)


@pytest.mark.parametrize('text,old,new', [
    ('Vi använder verktyget.', 'de', 'reglerna'),
    ('Vi använder verktyget och de kontrollerar reglerna.', 'de', 'reglerna'),
    ('Vi granskar och analyserar och redovisar.', 'och', 'samt'),
    ('till exempel vad de används till', 'till', 'till.'),
    ('ord ord ord', 'ord ord', 'ord'),
    ('ord  ord  ord', 'ord ord', 'ord'),
    ('en  granskning och en\u00a0granskning', 'en granskning', 'kontroll'),
    ('Antalet ökade t.ex. kraftigt.', 'ex.', 'exempel'),
    ('a-kassorna granskar.', 'kassorna', 'myndigheterna'),
    ('a\u2011kassorna granskar.', 'kassorna', 'myndigheterna'),
    ('De riskbaserade kontrollerna fungerar.', 'i', 'av'),
])
def test_ambiguous_and_partial_raw_anchors_fail_closed(text, old, new):
    ai = suggestor([element(text)])
    raw = proposal(old, new)
    assert not ai._postprocess_suggestions([raw])
    unchecked = SuggestedChange('paragraph', 'p1', None, old, new, None, 'exact')
    assert not ChangePlanner(ai.json_structured_document, TokenDiffEngine(), LOG).build_plans([unchecked])
    assert apply_suggestions_to_text(text, [dict(old=old, new=new)]) == (text, 0, 1, 0)
    assert not ai._filter_with_quality_guards([unchecked])[0]


def test_minimized_candidate_must_remain_inside_original_span(monkeypatch):
    ai, original, _ = prepare('Första ordet är svårt. Det andra uttrycket är oklart.',
                               'Det andra uttrycket är oklart', 'Det andra uttrycket är tydligt')
    monkeypatch.setattr(ai, '_minimize_change_span', lambda s: replace(s, old='svårt', new='tydligt'))
    result = ai._minimize_and_filter_suggestion(original)
    assert result.safe_to_apply
    assert (result.old, result.new) == (original.old, original.new)


def test_minimized_candidate_must_produce_identical_text(monkeypatch):
    ai, original, _ = prepare('Det andra uttrycket är oklart.', 'uttrycket är oklart', 'uttrycket är tydligt')
    monkeypatch.setattr(ai, '_minimize_change_span', lambda s: replace(s, old='oklart', new='svårt'))
    result = ai._minimize_and_filter_suggestion(original)
    assert result.safe_to_apply
    assert (result.old, result.new) == (original.old, original.new)


def test_exact_match_cannot_fall_back_to_normalization():
    assert locate_span('ord', 'ord ord') is None
    assert locate_span('ord ord', 'ord ord ord') is None
    assert locate_span('de', 'använder') is None
    assert locate_span('ord  ord', 'ord\u00a0ord') == (0, 7)
    assert locate_span('ord ', 'ord\u00a0\u00a0slut') == (0, 5)


def test_normalized_only_match_disabled():
    ai, original, raw = prepare('Vi har genomfört en  granskning.', 'en granskning', 'en kontroll')
    ai.allow_normalized_matches = False
    assert not ai._postprocess_suggestions([raw])
    result = ai._minimize_and_filter_suggestion(original)
    assert not result.safe_to_apply


@pytest.mark.parametrize('text,old,new', CASES)
@pytest.mark.parametrize('mode', ['word', 'span', 'simple'])
def test_full_pipeline_in_word(tmp_path, text, old, new, mode):
    source, output = tmp_path / 'source.docx', tmp_path / 'result.docx'
    doc = Document()
    # Include the punctuation failure in a table, not just a body paragraph.
    kind = 'table_cell' if text.startswith('Systeminventering:') else 'paragraph'
    if kind == 'table_cell':
        cell = doc.add_table(rows=1, cols=1).cell(0, 0)
        cell.paragraphs[0].text = 'Detta stycke ska stå kvar.'
        cell.add_paragraph(text)
    else:
        doc.add_paragraph(text)
    doc.save(source)
    structure = DocumentStructureExtractor(str(source), LOG).extract()
    target = next(e for e in structure['elements'] if e.get('text') == text)
    ai = suggestor(structure['elements'], assess_changes=False)
    validated = ai._postprocess_suggestions([proposal(old, new, target['element_id'], kind, assessment=None)])
    accepted, _ = ai._filter_with_quality_guards(validated)
    assert len(accepted) == 1, validated
    plans = ChangePlanner(structure, TokenDiffEngine(), LOG).build_plans(accepted)
    with DocxPackage(str(source), LOG) as package:
        renderer = SimpleMarkupRenderer(package, LOG) if mode == 'simple' else TrackedChangesRenderer(package, LOG, word_level_changes=mode == 'word')
        results = renderer.apply_plans(plans)
        assert len(results) == 1 and results[0].applied, results
        package.save(str(output))
    with ZipFile(output) as archive:
        tree = etree.fromstring(archive.read('word/document.xml'))
    paragraph = tree.xpath('//w:body/w:p' if kind == 'paragraph' else '//w:tbl/w:tr/w:tc/w:p[2]', namespaces=NS)[0]
    a, b = locate_span(old, text)
    expected = text[:a] + new + text[b:]
    if mode == 'simple':
        # Simple markup keeps old text in struck-through runs.
        # The accepted view skips only those runs.
        plain = ''.join(n.text or '' for n in paragraph.iter(W + 't')
                        if not n.getparent().xpath('./w:rPr/w:strike[@w:val!="0"] | ./w:rPr/w:strike[not(@w:val)]', namespaces=NS))
        # The existing simple-markup mode deliberately inserts neutral
        # visual separators between deleted and inserted runs. Verify the
        # lexical result and exact marked spans rather than those separators.
        assert " ".join(plain.split()) == " ".join(expected.split())
        marked_old = ''.join(n.text or '' for n in paragraph.iter(W + 't')
                             if n.getparent().xpath('./w:rPr/w:strike[@w:val="true"]', namespaces=NS))
        marked_new = ''.join(n.text or '' for n in paragraph.iter(W + 't')
                             if n.getparent().xpath('./w:rPr/w:color[@w:val="0000FF"]', namespaces=NS))
        assert marked_old == plans[0].old_text
        assert marked_new == plans[0].new_text
    else:
        assert visible(paragraph, False) == text
        assert visible(paragraph, True) == expected


@pytest.mark.parametrize('mode', ['word', 'span', 'simple'])
def test_table_render_failure_never_relocates_to_other_paragraph(tmp_path, mode):
    path = tmp_path / 'source.docx'
    doc = Document()
    cell = doc.add_table(rows=1, cols=1).cell(0, 0)
    cell.paragraphs[0].text = 'Ordet finns i första stycket.'
    cell.add_paragraph('Ordet finns i andra stycket.')
    doc.save(path)
    # The extracted structure is intentionally stale for the targeted paragraph.
    stale = {'type': 'docx', 'elements': [element('Detta ord finns inte här.', 'table_1_cell_1_1_p2', 'table_cell')]}
    unchecked = SuggestedChange('table_cell', 'table_1_cell_1_1_p2', None, 'finns', 'förekommer', None, 'exact')
    plans = ChangePlanner(stale, TokenDiffEngine(), LOG).build_plans([unchecked])
    assert len(plans) == 1
    with DocxPackage(str(path), LOG) as package:
        before = package.read_document_tree().getroot().xpath('//w:tbl', namespaces=NS)[0]
        renderer = SimpleMarkupRenderer(package, LOG) if mode == 'simple' else TrackedChangesRenderer(package, LOG, word_level_changes=mode == 'word')
        result = renderer.apply_plans(plans)[0]
        assert not result.applied
        after = package.read_document_tree().getroot().xpath('//w:tbl', namespaces=NS)[0]
        assert etree.tostring(before) == etree.tostring(after)


@pytest.mark.parametrize('mode', ['word', 'span', 'simple'])
def test_footnote_anchor_uses_extractor_position_not_first_occurrence(tmp_path, mode):
    source, enriched = tmp_path / 'source.docx', tmp_path / 'footnote.docx'
    doc = Document()
    p = doc.add_paragraph('En fotnot.')
    etree.SubElement(p.add_run()._r, W + 'footnoteReference').set(W + 'id', '7')
    doc.save(source)
    with DocxPackage(str(source), LOG) as package:
        tree = package.read_footnotes_tree(create_if_missing=True)
        footnote = etree.SubElement(tree.getroot(), W + 'footnote')
        footnote.set(W + 'id', '7')
        first = etree.SubElement(footnote, W + 'p')
        run = etree.SubElement(first, W + 'r')
        etree.SubElement(run, W + 'footnoteRef')
        etree.SubElement(run, W + 'tab')
        etree.SubElement(etree.SubElement(first, W + 'r'), W + 't').text = 'Vi använder systemet.'
        second = etree.SubElement(footnote, W + 'p')
        etree.SubElement(etree.SubElement(second, W + 'r'), W + 't').text = ' Vi förklarar hur de ska tillämpas.'
        package.write_footnotes_tree(tree)
        package.ensure_office_relationship('footnotes', 'footnotes.xml')
        package.ensure_content_type_override('/word/footnotes.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml')
        package.save(str(enriched))
    structure = DocumentStructureExtractor(str(enriched), LOG).extract()
    e = next(e for e in structure['elements'] if e['type'] == 'footnote')
    ai = suggestor(structure['elements'], assess_changes=False)
    raw = proposal('hur de ska tillämpas', 'hur reglerna ska tillämpas', e['element_id'], 'footnote', assessment=None)
    raw['footnote_id'] = e['footnote_id']
    val = ai._postprocess_suggestions([raw])
    accepted, _ = ai._filter_with_quality_guards(val)
    assert len(accepted) == 1
    plans = ChangePlanner(structure, TokenDiffEngine(), LOG).build_plans(accepted)
    with DocxPackage(str(enriched), LOG) as package:
        renderer = SimpleMarkupRenderer(package, LOG) if mode == 'simple' else TrackedChangesRenderer(package, LOG, word_level_changes=mode == 'word')
        results = renderer.apply_plans(plans)
        assert len(results) == 1 and results[0].applied, results
        tree = package.read_footnotes_tree()
        first, second = tree.xpath('//w:footnote[@w:id="7"]/w:p', namespaces=NS)
        assert visible(first, True) == 'Vi använder systemet.'
        assert len(list(first.iter(W + 'footnoteRef'))) == 1
        if mode == 'simple':
            inserted = second.xpath('.//w:r[w:rPr/w:color[@w:val="0000FF"]]/w:t/text()', namespaces=NS)
            assert ''.join(inserted) == plans[0].new_text
        else:
            assert visible(second, True) == ' Vi förklarar hur reglerna ska tillämpas.'
            assert visible(second, False) == ' Vi förklarar hur de ska tillämpas.'


@pytest.mark.parametrize('mode', ['word', 'span', 'simple'])
def test_prefix_check_allows_multiple_nonoverlapping_edits(tmp_path, mode):
    source = tmp_path / 'source.docx'
    doc = Document()
    text = 'Vi granskar kontrolerna och skriver sammanfatningen.'
    doc.add_paragraph(text)
    doc.save(source)
    structure = DocumentStructureExtractor(str(source), LOG).extract()
    target = next(e for e in structure['elements'] if e.get('text') == text)
    suggestions = [SuggestedChange('paragraph', target['element_id'], None, a, b, None, 'exact')
                   for a, b in [('kontrolerna', 'kontrollerna'), ('sammanfatningen', 'sammanfattningen')]]
    plans = ChangePlanner(structure, TokenDiffEngine(), LOG).build_plans(suggestions)
    with DocxPackage(str(source), LOG) as package:
        renderer = SimpleMarkupRenderer(package, LOG) if mode == 'simple' else TrackedChangesRenderer(package, LOG, word_level_changes=mode == 'word')
        results = renderer.apply_plans(plans)
        assert len(results) == 2 and all(r.applied for r in results), results
        paragraph = package.read_document_tree().getroot().xpath('//w:body/w:p', namespaces=NS)[0]
        if mode != 'simple':
            assert visible(paragraph, True) == 'Vi granskar kontrollerna och skriver sammanfattningen.'
            assert visible(paragraph, False) == text


def test_footnote_cross_paragraph_anchor_rejected():
    from app.src.JBGFootnotesPartAdapter import FootnotesPartAdapter
    fn = etree.Element(W + 'footnote')
    for text in ('Det första stycket slutar med ', 'det andra stycket.'):
        etree.SubElement(etree.SubElement(etree.SubElement(fn, W + 'p'), W + 'r'), W + 't').text = text
    text = 'Det första stycket slutar med det andra stycket.'
    structure = {'type': 'docx', 'elements': [element(text, 'f1', 'footnote', footnote_id='7')]}
    proposal_ = SuggestedChange('footnote', 'f1', '7', 'med det', 'före det', None, 'exact')
    plan = ChangePlanner(structure, TokenDiffEngine(), LOG).build_plans([proposal_])[0]
    adapter = object.__new__(FootnotesPartAdapter)
    model = adapter._build_footnote_model(fn, '7')
    renderer = object.__new__(TrackedChangesRenderer)
    with pytest.raises(ValueError, match='crosses unsupported content or paragraphs'):
        adapter.locate_editable_anchor(plan, model, renderer._build_footnote_paragraph_models(model))
