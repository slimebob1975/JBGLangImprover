"""Safely preserve a standalone footnote run at a proven unchanged boundary.

Frozen input and output text; no AI calls and no relocation of footnote IDs.
"""
import logging
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from zipfile import ZipFile

from docx import Document
from lxml import etree

from app.src.JBGChangePlanner import ChangePlanner
from app.src.JBGCommentsRenderer import CommentsRenderer
from app.src.JBGDocxPackage import DocxPackage, W_NS
from app.src.JBGDocumentStructureExtractor import DocumentStructureExtractor
from app.src.JBGLangImprovSuggestorAI import SuggestedChange
from app.src.JBGTokenDiffEngine import TokenDiffEngine
from app.src.JBGTrackedChangesRenderer import TrackedChangesRenderer
from app.src.JBGRunEditing import SpanSegment

W = f'{{{W_NS}}}'
NS = {'w': W_NS}
BEFORE = 'Vi granskar omfattande underlag för beslut.'
AFTER = ' Därefter skriver vi en sammanfattning.'
OLD = BEFORE + AFTER
NEW = 'Vi granskar relevanta underlag för beslut. Därefter skriver vi en kort sammanfattning.'


def selected_text(paragraph, accept):
    items = []
    for element in paragraph:
        if element.tag == W + 'r' or element.tag == W + ('ins' if accept else 'del'):
            for child in element.iter():
                if child.tag in (W + 't', W + 'delText'):
                    items.append(child.text or '')
    return ''.join(items)


def footnote_positions(paragraph):
    """Count visible characters preceding each standalone reference."""
    position = 0
    locations = []
    for el in paragraph:
        if el.tag == W + 'r' and el.find(W + 'footnoteReference') is not None:
            locations.append((position, el.find(W + 'footnoteReference').get(W + 'id')))
        if el.tag in (W + 'r', W + 'ins', W + 'del'):
            for child in el.iter():
                if child.tag == W + 't':
                    position += len(child.text or '')
    return locations


class FootnoteBoundaryTests(unittest.TestCase):
    def test_safe_boundary_keeps_reference_and_roundtrips(self):
        logger = logging.getLogger('safe-footnote-boundary')
        with tempfile.TemporaryDirectory() as td:
            source = Path(td) / 'source.docx'
            d = Document()
            p = d.add_paragraph()
            p.add_run(BEFORE)
            ref = p.add_run()
            etree.SubElement(ref._r, W + 'footnoteReference').set(W + 'id', '9')
            p.add_run(AFTER)
            d.save(source)
            structure = DocumentStructureExtractor(str(source), logger).extract()
            element = next(e for e in structure['elements'] if e.get('type') == 'paragraph' and e.get('text') == OLD)
            suggestion = SuggestedChange(element_type='paragraph', element_id=element['element_id'],
                                         footnote_id=None, old=OLD, new=NEW,
                                         motivation='Test av bevarad fotnot.', match_status='exact')
            plans = ChangePlanner(structure, TokenDiffEngine(), logger).build_plans([suggestion])
            self.assertEqual(len(plans), 1)
            for enabled in (False, True):
                with self.subTest(word_level=enabled):
                    out = Path(td) / ('word.docx' if enabled else 'span.docx')
                    with DocxPackage(str(source), logger) as package:
                        results = TrackedChangesRenderer(package, logger, word_level_changes=enabled).apply_plans(plans)
                        self.assertEqual(len(results), 1)
                        self.assertTrue(results[0].applied, results[0].message)
                        if enabled:
                            self.assertEqual(results[0].word_level_status, 'span_fallback')
                        comments = CommentsRenderer(package, logger).apply_comments_for_results(results)
                        self.assertEqual(len(comments), 1)
                        self.assertTrue(comments[0].applied, comments[0].message)
                        package.save(str(out))
                    with ZipFile(out) as z:
                        root = etree.fromstring(z.read('word/document.xml'))
                    paragraph = root.xpath('.//w:body/w:p', namespaces=NS)[0]
                    references = paragraph.xpath('.//w:footnoteReference', namespaces=NS)
                    self.assertEqual(len(references), 1)
                    self.assertEqual(references[0].get(W + 'id'), '9')
                    self.assertEqual(references[0].getparent().getparent(), paragraph)
                    self.assertEqual(selected_text(paragraph, False), OLD)
                    self.assertEqual(selected_text(paragraph, True), NEW)
                    # The reference must keep its *character offset* when
                    # revisions are rejected, not just its ID and element.
                    original_offset = len(BEFORE)
                    def position_in_view(accept):
                        seen = 0
                        for outer in paragraph:
                            if outer.tag == W + ('del' if accept else 'ins'):
                                continue
                            for child in outer.iter():
                                if child.tag in (W + 't', W + 'delText'):
                                    seen += len(child.text or '')
                                if child.tag == W + 'footnoteReference':
                                    return seen
                        return None
                    self.assertEqual(position_in_view(False), original_offset)
                    self.assertEqual(position_in_view(True), len('Vi granskar relevanta underlag för beslut.'))
                    self.assertEqual(len(paragraph.xpath('./w:ins', namespaces=NS)), 2)
                    self.assertEqual(len(paragraph.xpath('./w:del', namespaces=NS)), 2)
                    for marker in ('commentRangeStart', 'commentRangeEnd', 'commentReference'):
                        self.assertEqual(len(paragraph.xpath('.//w:' + marker, namespaces=NS)), 1)
                    # Reference is never inside a revision wrapper.
                    self.assertFalse(paragraph.xpath('.//w:ins//w:footnoteReference | .//w:del//w:footnoteReference', namespaces=NS))

    def test_split_never_consumes_text_after_reference_with_prefix(self):
        # v6 used the whole first run for the offset although segment.old
        # excludes its untouched prefix. A stable context after the reference
        # could then make the split succeed at the WRONG position.
        prefix = 'PREFIX-UNCHANGED '
        before = 'The first long sentence has a reference after these words'
        after = ' and then this other sentence continues with ordinary content for a long time.'
        paragraph = etree.Element(W + 'p')
        first = etree.SubElement(paragraph, W + 'r')
        etree.SubElement(first, W + 't').text = prefix + before
        ref = etree.SubElement(paragraph, W + 'r')
        etree.SubElement(ref, W + 'footnoteReference').set(W + 'id', '20')
        last = etree.SubElement(paragraph, W + 'r')
        etree.SubElement(last, W + 't').text = after
        segment = SpanSegment(container=paragraph, runs=[first, ref, last],
                              old=before + after,
                              new='The changed long sentence has a reference after these words' + after,
                              before=prefix, after='', first_run=first, last_run=last)
        split = TrackedChangesRenderer._split_at_stable_footnote_reference(segment)
        self.assertIsNotNone(split)
        self.assertEqual(split[0].old, before)
        self.assertEqual(split[1].old, after)
        self.assertEqual(split[0].before, prefix)
        self.assertEqual(split[0].old + split[1].old, segment.old)

    def test_real_report_reference_offsets_with_partial_first_run(self):
        # Deterministic reproductions of the two displaced references in the
        # v6 test report: 20 after "andra metoder" and 23 before "Sammansättningen".
        # The editable span begins *inside* the first run, after a fixed prefix.
        fixtures = (
            ('20', 'De sammanräknade resultaten i avsnitt 4.2.1 och 4.2.2 ger ett högre mörkertal än ',
             'tidigare uppskattningar gjorda av IAF med andra metoder',
             '. Sammantaget har arbetslöshetskassorna enligt uppskattningarna i den här rapporten upptäckt fel.',
             'IAF:s tidigare uppskattningar med andra metoder',
             '. Sammantaget visar rapporten att a-kassorna har upptäckt fel.'),
            ('23', 'IAF:s tidigare granskningar visar också att ',
             'andelen ersättningstagare som fått återkrav ökade mellan 2017 och 2021.',
             ' Sammansättningen av ersättningstagarna som fått återkrav har också förändrats.',
             'andelen ersättningstagare som fått återkrav ökade mellan 2017 och 2021.',
             ' Sammansättningen av de ersättningstagare som fått återkrav har också förändrats.'),
        )
        for footnote_id, prefix, before, after, revised_before, revised_after in fixtures:
            with self.subTest(footnote_id=footnote_id):
                p = etree.Element(W + 'p')
                left = etree.SubElement(p, W + 'r')
                left_text = etree.SubElement(left, W + 't')
                left_text.text = prefix + before
                ref = etree.SubElement(p, W + 'r')
                etree.SubElement(ref, W + 'footnoteReference').set(W + 'id', footnote_id)
                right = etree.SubElement(p, W + 'r')
                right_text = etree.SubElement(right, W + 't')
                right_text.text = after
                old, new = before + after, revised_before + revised_after
                start = len(prefix)
                first_node = SimpleNamespace(run_element=left, element=left_text,
                                             start=0, end=len(prefix + before))
                last_node = SimpleNamespace(run_element=right, element=right_text,
                                            start=len(prefix + before), end=len(prefix + old))
                renderer = object.__new__(TrackedChangesRenderer)
                renderer.word_level_changes = False
                renderer.change_id_counter = 1
                renderer.author = 'test'
                renderer.logger = logging.getLogger('footnote-offset-regression')
                renderer._rewrite_span(p, first_node, last_node, start, len(prefix + old), old, new)

                def view_with_reference(accept):
                    sequence = []
                    for top in p:
                        if top.tag == W + ('del' if accept else 'ins'):
                            continue
                        for child in top.iter():
                            if child.tag in (W + 't', W + 'delText'):
                                sequence.append(child.text or '')
                            elif child.tag == W + 'footnoteReference':
                                sequence.append('<REF:' + child.get(W + 'id') + '>')
                    return ''.join(sequence)

                self.assertEqual(view_with_reference(False),
                                 prefix + before + '<REF:' + footnote_id + '>' + after)
                self.assertEqual(view_with_reference(True),
                                 prefix + revised_before + '<REF:' + footnote_id + '>' + revised_after)

    def test_rewrite_next_to_reference_is_rejected_unchanged(self):
        paragraph = etree.Element(W + 'p')
        first = etree.SubElement(paragraph, W + 'r')
        first_t = etree.SubElement(first, W + 't')
        first_t.text = BEFORE
        middle = etree.SubElement(paragraph, W + 'r')
        etree.SubElement(middle, W + 'footnoteReference').set(W + 'id', '9')
        last = etree.SubElement(paragraph, W + 'r')
        last_t = etree.SubElement(last, W + 't')
        last_t.text = AFTER
        first_node = SimpleNamespace(run_element=first, element=first_t, start=0, end=len(BEFORE))
        last_node = SimpleNamespace(run_element=last, element=last_t, start=len(BEFORE), end=len(OLD))
        # The AI changed words directly adjacent to the footnote.
        rewrite = 'Vi granskar relevanta underlag före beslutet. Sedan skriver vi en kort sammanfattning.'
        renderer = object.__new__(TrackedChangesRenderer)
        renderer.word_level_changes = True
        before_xml = etree.tostring(paragraph)
        with self.assertRaisesRegex(ValueError, 'unsafe_span_run_content:footnoteReference'):
            renderer._rewrite_span(paragraph, first_node, last_node, 0, len(OLD), OLD, rewrite)
        self.assertEqual(etree.tostring(paragraph), before_xml)

    def test_multiple_refs_and_field_stay_blocked(self):
        paragraph = etree.Element(W + 'p')
        r1 = etree.SubElement(paragraph, W + 'r')
        etree.SubElement(r1, W + 't').text = BEFORE
        ref1 = etree.SubElement(paragraph, W + 'r')
        etree.SubElement(ref1, W + 'footnoteReference').set(W + 'id', '9')
        ref2 = etree.SubElement(paragraph, W + 'r')
        etree.SubElement(ref2, W + 'footnoteReference').set(W + 'id', '10')
        r2 = etree.SubElement(paragraph, W + 'r')
        etree.SubElement(r2, W + 't').text = AFTER
        seg = SpanSegment(container=paragraph, runs=[r1, ref1, ref2, r2], old=OLD, new=NEW,
                          before='', after='', first_run=r1, last_run=r2)
        self.assertIsNone(TrackedChangesRenderer._split_at_stable_footnote_reference(seg))
        paragraph.remove(ref2)
        etree.SubElement(ref1, W + 'fldChar')
        seg.runs = [r1, ref1, r2]
        self.assertIsNone(TrackedChangesRenderer._split_at_stable_footnote_reference(seg))

if __name__ == '__main__':
    unittest.main()
