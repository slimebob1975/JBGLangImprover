"""Hard hyphens are visible in the text model, but immutable in OOXML."""
import logging
from zipfile import ZipFile

import pytest
from docx import Document
from lxml import etree

from app.src.JBGChangePlanner import ChangePlanner
from app.src.JBGCommentsRenderer import CommentsRenderer
from app.src.JBGDocumentPartAdapter import DocumentPartAdapter
from app.src.JBGDocumentStructureExtractor import DocumentStructureExtractor
from app.src.JBGDocxPackage import DocxPackage, W_NS
from app.src.JBGFootnotesPartAdapter import FootnotesPartAdapter
from app.src.JBGLangImprovSuggestorAI import SuggestedChange
from app.src.JBGRunEditing import run_visible_text
from app.src.JBGSimpleMarkupRenderer import SimpleMarkupRenderer
from app.src.JBGTokenDiffEngine import TokenDiffEngine
from app.src.JBGTrackedChangesRenderer import TrackedChangesRenderer

W = f'{{{W_NS}}}'
LOG = logging.getLogger(__name__)
LEFT = 'Vi granskar omfattande underlag hos a'
RIGHT = 'kassorna i denna rapport och skriver en sammanfattning.'
OLD = LEFT + '-' + RIGHT
NEW = 'Vi granskar relevanta underlag hos a-kassorna i denna rapport och skriver en kort sammanfattning.'


def add_hyphen(paragraph, mixed=False):
    paragraph.add_run(LEFT)
    run = paragraph.add_run()
    etree.SubElement(run._r, W + 'noBreakHyphen')
    if mixed:
        run.add_text('kassorna')
        paragraph.add_run(RIGHT[len('kassorna'):])
    else:
        paragraph.add_run(RIGHT)
    return run._r


def visible(paragraph, accepted):
    text = []
    for node in paragraph.iter():
        if any(a.tag == W + ('del' if accepted else 'ins') for a in node.iterancestors()):
            continue
        if node.tag in (W + 't', W + 'delText'):
            text.append(node.text or '')
        elif node.tag == W + 'noBreakHyphen':
            text.append('-')
    return ''.join(text)


def run_case(tmp_path, old, new, mode, mixed=False):
    source = tmp_path / 'input.docx'
    dest = tmp_path / 'output.docx'
    doc = Document()
    add_hyphen(doc.add_paragraph(), mixed)
    doc.save(source)
    structure = DocumentStructureExtractor(str(source), LOG).extract()
    element = next(e for e in structure['elements'] if e['type'] == 'paragraph')
    assert element['text'] == OLD
    suggestion = SuggestedChange(element_type='paragraph', element_id=element['element_id'],
                                 footnote_id=None, old=old, new=new, motivation='Test', match_status='exact')
    plans = ChangePlanner(structure, TokenDiffEngine(), LOG).build_plans([suggestion])
    assert len(plans) == 1
    with DocxPackage(str(source), LOG) as package:
        renderer = (SimpleMarkupRenderer(package, LOG) if mode == 'simple'
                    else TrackedChangesRenderer(package, LOG, word_level_changes=mode == 'word'))
        results = renderer.apply_plans(plans)
        if mode != "simple" and results[0].applied:
            comments = CommentsRenderer(package, LOG).apply_comments_for_results(results)
            assert len(comments) == 1 and comments[0].applied
        package.save(str(dest))
    with ZipFile(dest) as z:
        root = etree.fromstring(z.read('word/document.xml'))
    paragraph = root.find('.//' + W + 'body/' + W + 'p')
    hyphens = list(paragraph.iter(W + 'noBreakHyphen'))
    assert len(hyphens) == 1
    assert hyphens[0].getparent().getparent() is paragraph
    return results[0], paragraph


@pytest.mark.parametrize('mode', ['word', 'span', 'simple'])
@pytest.mark.parametrize('mixed', [False, True])
def test_preserve_unchanged_hyphen_run(tmp_path, mode, mixed):
    result, paragraph = run_case(tmp_path, OLD, NEW, mode, mixed)
    assert result.applied, result.message
    if mode != 'simple':
        assert visible(paragraph, False) == OLD
        assert visible(paragraph, True) == NEW


@pytest.mark.parametrize('mode', ['word', 'span', 'simple'])
def test_reject_edit_of_hyphen(tmp_path, mode):
    result, paragraph = run_case(tmp_path, OLD, OLD.replace('a-kassorna', 'akassorna'), mode)
    assert not result.applied
    assert 'unsafe_span_run_content:noBreakHyphen' in result.message
    assert visible(paragraph, True) == OLD
    assert not list(paragraph.iter(W + 'ins'))
    assert not list(paragraph.iter(W + 'del'))


@pytest.mark.parametrize('mode', ['word', 'span', 'simple'])
def test_offset_after_hyphen(tmp_path, mode):
    result, paragraph = run_case(tmp_path, 'sammanfattning', 'rapport', mode)
    assert result.applied, result.message
    if mode != 'simple':
        assert visible(paragraph, True) == OLD.replace('sammanfattning', 'rapport')
        assert visible(paragraph, False) == OLD


def test_extraction_all_main_story_paths(tmp_path):
    doc = Document()
    for p in (doc.add_paragraph(), doc.add_table(1, 1).cell(0, 0).paragraphs[0],
              doc.sections[0].header.paragraphs[0], doc.sections[0].footer.paragraphs[0]):
        add_hyphen(p)
    source = tmp_path / 'all.docx'
    doc.save(source)
    structure = DocumentStructureExtractor(str(source), LOG).extract()
    for kind in ('paragraph', 'table_cell', 'header', 'footer'):
        assert any(e['type'] == kind and e['text'] == OLD for e in structure['elements'])
    p = doc.paragraphs[0]._element
    model = object.__new__(DocumentPartAdapter)._build_paragraph_model(p)
    assert model.visible_text == OLD
    node = next(n for n in model.nodes if n.element.tag == W + 'noBreakHyphen')
    assert (node.start, node.end, node.text) == (len(LEFT), len(LEFT) + 1, '-')
    assert run_visible_text(node.run_element) == '-'
    assert DocumentStructureExtractor._visible_text_from_xml_paragraph(p) == OLD
    footnote = etree.Element(W + 'footnote')
    footnote.append(p)
    assert object.__new__(FootnotesPartAdapter)._build_footnote_model(footnote, '7').visible_text == OLD


@pytest.mark.parametrize('mode', ['word', 'span', 'simple'])
def test_mixed_run_with_changed_text_is_not_rebuilt(tmp_path, mode):
    result, paragraph = run_case(tmp_path, OLD, OLD.replace('a-kassorna', 'a-kassan'), mode, mixed=True)
    assert not result.applied
    assert 'unsafe_span_run_content:noBreakHyphen' in result.message
    assert visible(paragraph, True) == OLD


def test_hyphen_with_other_opaque_child_refuses_whole_span():
    from app.src.JBGRunEditing import SpanSegment, split_at_unchanged_hard_hyphens
    p = Document().add_paragraph()
    protected = add_hyphen(p)
    etree.SubElement(protected, W + 'footnoteReference').set(W + 'id', '7')
    runs = [r._r for r in p.runs]
    segment = SpanSegment(container=p._p, runs=runs, old=OLD, new=NEW,
                          first_run=runs[0], last_run=runs[-1])
    before = etree.tostring(p._p)
    assert split_at_unchanged_hard_hyphens(segment) is None
    assert etree.tostring(p._p) == before
