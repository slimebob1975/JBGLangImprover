"""Frozen proposal-level comparison and lossless formatting boundary tests.

No AI calls. A changed formatting boundary must remain the legacy fallback.
"""
import logging
import tempfile
import unittest
from pathlib import Path
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

W = f'{{{W_NS}}}'
NS = {'w': W_NS}
OLD = 'Vi gör en omfattande granskning och därefter fattar vi beslut.'
NEW = 'Vi gör en grundlig granskning och sedan fattar vi beslut.'


def selected_text(paragraph, accept):
    parts = []
    for element in paragraph:
        if element.tag not in (W + 'r', W + ('ins' if accept else 'del')):
            continue
        for child in element.iter():
            if child.tag in (W + 't', W + 'delText'):
                parts.append(child.text or '')
    return ''.join(parts)


class FormattedBoundaryTests(unittest.TestCase):
    def test_safe_boundary_uses_word_level_without_losing_styles(self):
        logger = logging.getLogger('formatted-boundary')
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / 'original.docx'
            dst = Path(tmp) / 'result.docx'
            doc = Document()
            paragraph = doc.add_paragraph()
            boundary = OLD.index(' och därefter')
            paragraph.add_run(OLD[:boundary])
            paragraph.add_run(OLD[boundary:]).bold = True
            doc.save(src)
            structure = DocumentStructureExtractor(str(src), logger).extract()
            element = next(e for e in structure['elements'] if e.get('text') == OLD)
            suggestions = [SuggestedChange(element_type='paragraph', element_id=element['element_id'],
                                           footnote_id=None, old=OLD, new=NEW,
                                           motivation='Fryst jämförelse.', match_status='exact')]
            plans = ChangePlanner(structure, TokenDiffEngine(), logger).build_plans(suggestions)
            with DocxPackage(str(src), logger) as package:
                results = TrackedChangesRenderer(package, logger, word_level_changes=True).apply_plans(plans)
                self.assertEqual(len(results), 1)
                self.assertTrue(results[0].applied, results[0].message)
                self.assertEqual(results[0].word_level_status, 'word_level')
                comments = CommentsRenderer(package, logger).apply_comments_for_results(results)
                self.assertEqual(len(comments), 1)
                self.assertTrue(comments[0].applied, comments[0].message)
                package.save(str(dst))
            with ZipFile(dst) as z:
                root = etree.fromstring(z.read('word/document.xml'))
                comments = etree.fromstring(z.read('word/comments.xml'))
            p = root.xpath('.//w:body/w:p', namespaces=NS)[0]
            self.assertEqual(selected_text(p, False), OLD)
            self.assertEqual(selected_text(p, True), NEW)
            self.assertEqual(len(p.xpath('./w:del', namespaces=NS)), 2)
            self.assertEqual(len(p.xpath('./w:ins', namespaces=NS)), 2)
            # The first edit is unformatted, and the second edit stays bold.
            for tag in ('ins', 'del'):
                items = p.xpath('./w:' + tag, namespaces=NS)
                self.assertEqual(len(items[0].xpath('.//w:b', namespaces=NS)), 0)
                self.assertEqual(len(items[1].xpath('.//w:b', namespaces=NS)), 1)
            ids = [n.get(W + 'id') for n in p.xpath('./w:ins|./w:del', namespaces=NS)]
            self.assertEqual(len(ids), len(set(ids)))
            self.assertEqual(len(comments.xpath('./w:comment', namespaces=NS)), 1)
            for marker in ('commentRangeStart', 'commentRangeEnd', 'commentReference'):
                self.assertEqual(len(p.xpath('.//w:' + marker, namespaces=NS)), 1)

    def test_partial_outer_runs_keep_unmodified_prefix_and_suffix(self):
        logger = logging.getLogger('partial-formatted-boundary')
        with tempfile.TemporaryDirectory() as tmp:
            src, dst = Path(tmp) / 'original.docx', Path(tmp) / 'result.docx'
            original = 'Före: ' + OLD + ' Efter.'
            doc = Document()
            p = doc.add_paragraph()
            boundary = OLD.index(' och därefter')
            p.add_run('Före: ' + OLD[:boundary])
            p.add_run(OLD[boundary:] + ' Efter.').bold = True
            doc.save(src)
            structure = DocumentStructureExtractor(str(src), logger).extract()
            element = next(e for e in structure['elements'] if e.get('text') == original)
            suggestions = [SuggestedChange(element_type='paragraph', element_id=element['element_id'],
                                           footnote_id=None, old=OLD, new=NEW,
                                           motivation='Fryst jämförelse.', match_status='exact')]
            plans = ChangePlanner(structure, TokenDiffEngine(), logger).build_plans(suggestions)
            with DocxPackage(str(src), logger) as package:
                results = TrackedChangesRenderer(package, logger, word_level_changes=True).apply_plans(plans)
                self.assertEqual(len(results), 1)
                self.assertTrue(results[0].applied, results[0].message)
                self.assertEqual(results[0].word_level_status, 'word_level')
                comments = CommentsRenderer(package, logger).apply_comments_for_results(results)
                self.assertEqual(len(comments), 1)
                self.assertTrue(comments[0].applied, comments[0].message)
                package.save(str(dst))
            with ZipFile(dst) as z:
                root = etree.fromstring(z.read('word/document.xml'))
            p = root.xpath('.//w:body/w:p', namespaces=NS)[0]
            self.assertEqual(selected_text(p, False), original)
            self.assertEqual(selected_text(p, True), 'Före: ' + NEW + ' Efter.')
            for tag in ('ins', 'del'):
                parts = p.xpath('./w:' + tag, namespaces=NS)
                self.assertEqual(len(parts), 2)
                self.assertEqual(len(parts[0].xpath('.//w:b', namespaces=NS)), 0)
                self.assertEqual(len(parts[1].xpath('.//w:b', namespaces=NS)), 1)

    def test_partial_runs_reject_boundary_with_modified_context(self):
        from app.src.JBGRunEditing import SpanSegment
        left_text = 'Före: Vi gör en omfattande'
        right_text = ' granskning och därefter fattar vi beslut. Efter.'
        p = etree.Element(W + 'p')
        left = etree.SubElement(p, W + 'r')
        etree.SubElement(left, W + 't').text = left_text
        right = etree.SubElement(p, W + 'r')
        etree.SubElement(etree.SubElement(right, W + 'rPr'), W + 'b')
        etree.SubElement(right, W + 't').text = right_text
        segment = SpanSegment(container=p, runs=[left, right], old=OLD,
                              new='Vi gör en grundlig bedömning och sedan fattar vi beslut.',
                              before='Före: ', after=' Efter.', first_run=left, last_run=right)
        self.assertIsNone(TrackedChangesRenderer._stable_formatted_run_pieces(segment))

    def test_changed_or_ambiguous_boundary_is_not_split(self):
        from app.src.JBGRunEditing import SpanSegment
        first = 'Vi gör en omfattande'
        second = ' granskning och därefter fattar vi beslut.'
        p = etree.Element(W + 'p')
        left = etree.SubElement(p, W + 'r')
        etree.SubElement(left, W + 't').text = first
        right = etree.SubElement(p, W + 'r')
        etree.SubElement(etree.SubElement(right, W + 'rPr'), W + 'b')
        etree.SubElement(right, W + 't').text = second
        segment = SpanSegment(container=p, runs=[left, right], old=first + second,
                              new=NEW, before='', after='', first_run=left, last_run=right)
        self.assertIsNone(TrackedChangesRenderer._stable_formatted_run_pieces(segment))
        # An invented prefix must not be accepted unless it occurs in the
        # actual run text, in the expected position.
        segment.before = 'prefix '
        self.assertIsNone(TrackedChangesRenderer._stable_formatted_run_pieces(segment))


if __name__ == '__main__':
    unittest.main()
