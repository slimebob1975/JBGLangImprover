"""Safely keep heading bookmarks while making precise tracked word changes.

Real-world shape: text inside a bookmark, a trailing space outside it.
No AI, external data, or private captured paragraphs are needed.
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
from app.src.JBGRunEditing import SpanSegment

W = '{' + W_NS + '}'
NS = {'w': W_NS}


def visible_view(paragraph, accept):
    result = []
    for node in paragraph:
        if node.tag == W + ('del' if accept else 'ins'):
            continue
        if node.tag not in (W + 'r', W + ('ins' if accept else 'del')):
            continue
        for child in node.iter():
            if child.tag in (W + 't', W + 'delText'):
                result.append(child.text or '')
    return ''.join(result)


def bookmark_offsets(paragraph, accept):
    """Bookmark character offsets in a selected Word revision view."""
    result = []
    offset = 0
    for node in paragraph:
        if node.tag == W + ('del' if accept else 'ins'):
            continue
        if node.tag in (W + 'bookmarkStart', W + 'bookmarkEnd'):
            result.append((etree.QName(node).localname, node.get(W + 'id'), offset))
        elif node.tag in (W + 'r', W + ('ins' if accept else 'del')):
            offset += sum(len(child.text or '') for child in node.iter()
                          if child.tag in (W + 't', W + 'delText'))
    return result


class BookmarkTrailingSpaceTests(unittest.TestCase):
    def _exercise(self, before_heading, after_heading, runs):
        logger = logging.getLogger('word-level-bookmark-boundary')
        with tempfile.TemporaryDirectory() as td:
            source = Path(td) / 'source.docx'
            document = Document()
            paragraph = document.add_paragraph()
            etree.SubElement(paragraph._p, W + 'bookmarkStart', {W + 'id': '17', W + 'name': 'Section_17'})
            for text in runs:
                paragraph.add_run(text)
            etree.SubElement(paragraph._p, W + 'bookmarkEnd', {W + 'id': '17'})
            paragraph.add_run(' ')
            document.save(source)
            old, new = before_heading + ' ', after_heading
            structure = DocumentStructureExtractor(str(source), logger).extract()
            element = next(item for item in structure['elements']
                           if item.get('type') == 'paragraph' and item.get('text') == old)
            suggestion = SuggestedChange(element_type='paragraph', element_id=element['element_id'],
                                         footnote_id=None, old=old, new=new,
                                         motivation='Frozen bookmark test', match_status='exact')
            plans = ChangePlanner(structure, TokenDiffEngine(), logger).build_plans([suggestion])
            self.assertEqual(len(plans), 1)
            metrics = {}
            for enabled in (False, True):
                with self.subTest(enabled=enabled):
                    output = Path(td) / ('word.docx' if enabled else 'span.docx')
                    with DocxPackage(str(source), logger) as pkg:
                        results = TrackedChangesRenderer(pkg, logger, word_level_changes=enabled).apply_plans(plans)
                        self.assertEqual(len(results), 1)
                        self.assertTrue(results[0].applied, results[0].message)
                        if enabled:
                            self.assertEqual(results[0].word_level_status, 'word_level')
                        comments = CommentsRenderer(pkg, logger).apply_comments_for_results(results)
                        self.assertEqual(len(comments), 1)
                        self.assertTrue(comments[0].applied, comments[0].message)
                        pkg.save(str(output))
                    with ZipFile(output) as z:
                        root = etree.fromstring(z.read('word/document.xml'))
                    p = root.xpath('.//w:body/w:p', namespaces=NS)[0]
                    self.assertEqual(visible_view(p, False), old)
                    self.assertEqual(visible_view(p, True), new)
                    self.assertEqual(bookmark_offsets(p, False),
                                     [('bookmarkStart', '17', 0), ('bookmarkEnd', '17', len(before_heading))])
                    self.assertEqual(bookmark_offsets(p, True),
                                     [('bookmarkStart', '17', 0), ('bookmarkEnd', '17', len(after_heading))])
                    for marker in ('commentRangeStart', 'commentRangeEnd', 'commentReference'):
                        self.assertEqual(len(p.xpath('.//w:' + marker, namespaces=NS)), 1)
                    self.assertFalse(p.xpath('.//w:bookmarkStart/ancestor::w:ins | .//w:bookmarkEnd/ancestor::w:del', namespaces=NS))
                    metrics[enabled] = sum(len(ch.text or '') for rev in p.xpath('./w:ins | ./w:del', namespaces=NS)
                                           for ch in rev.iter() if ch.tag in (W + 't', W + 'delText'))
            self.assertLess(metrics[True], metrics[False])

    def test_multi_run_heading_bookmark(self):
        self._exercise('Besluten följer inte förvaltningslagens krav på innehåll',
                       'Besluten uppfyller inte förvaltningslagens krav på innehåll',
                       ['Besluten följer inte ', 'f', 'örvaltningslagens krav på innehåll'])

    def test_one_run_heading_bookmark(self):
        self._exercise('Återkallandeorsakerna är vagt beskrivna',
                       'Återkallandeorsakerna är otydligt beskrivna',
                       ['Återkallandeorsakerna är vagt beskrivna'])

    def test_unpaired_bookmark_and_changed_suffix_are_not_promoted(self):
        p = etree.Element(W + 'p')
        first = etree.SubElement(p, W + 'r')
        etree.SubElement(first, W + 't').text = 'Vi gör en bedömning'
        bookmark = etree.SubElement(p, W + 'bookmarkEnd', {W + 'id': '17'})
        second = etree.SubElement(p, W + 'r')
        etree.SubElement(second, W + 't').text = ' '
        main = SpanSegment(p, [first], 'Vi gör en bedömning', 'Vi gör en granskning',
                           first_run=first, last_run=first)
        suffix = SpanSegment(p, [second], ' ', '', first_run=second, last_run=second)
        self.assertIsNone(TrackedChangesRenderer._prepare_bookmark_trailing_space_word_edits(p, [main, suffix]))
        etree.SubElement(p, W + 'bookmarkStart', {W + 'id': '17', W + 'name': 'Late'})
        self.assertIsNone(TrackedChangesRenderer._prepare_bookmark_trailing_space_word_edits(p, [main, suffix]))
        p.remove(p[-1])
        p.insert(0, etree.Element(W + 'bookmarkStart', {W + 'id': '17', W + 'name': 'Good'}))
        suffix.new = 'X'
        self.assertIsNone(TrackedChangesRenderer._prepare_bookmark_trailing_space_word_edits(p, [main, suffix]))


if __name__ == '__main__':
    unittest.main()
