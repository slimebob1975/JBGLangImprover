"""Run-local word edits preserve formatting and a following footnote.

Synthetic text only: actual report paragraphs remain in private uploads fixtures.
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
from app.src.JBGRunEditing import SpanSegment
from app.src.JBGTokenDiffEngine import TokenDiffEngine
from app.src.JBGTrackedChangesRenderer import TrackedChangesRenderer
from scripts.frozen_wordlevel_compare import reference_locations, view_text

W = '{' + W_NS + '}'
NS = {'w': W_NS}


class IsolatedRunWordEditsTests(unittest.TestCase):
    def test_formatted_run_local_changes_keep_following_footnote(self):
        logger = logging.getLogger('isolated-run-word-edits')
        with tempfile.TemporaryDirectory() as td:
            source = Path(td) / 'original.docx'
            output = Path(td) / 'result.docx'
            document = Document()
            p = document.add_paragraph()
            p.add_run('Rapporten innehåller en stor analys ').bold = True
            p.add_run('och ')
            p.add_run('en kort slutsats.').italic = True
            p.add_run(' ')
            marker = p.add_run('')
            etree.SubElement(marker._r, W + 'footnoteReference', {W + 'id': '42'})
            old = 'Rapporten innehåller en stor analys och en kort slutsats. '
            new = 'Rapporten innehåller en tydlig analys och en enkel slutsats.'
            document.save(source)
            structure = DocumentStructureExtractor(str(source), logger).extract()
            target = next(e for e in structure['elements'] if e.get('text') == old)
            change = SuggestedChange(element_type='paragraph', element_id=target['element_id'],
                                     footnote_id=None, old=old, new=new,
                                     motivation='Synthetic local-run test', match_status='exact')
            plans = ChangePlanner(structure, TokenDiffEngine(), logger).build_plans([change])
            self.assertEqual(len(plans), 1)
            with DocxPackage(str(source), logger) as package:
                result = TrackedChangesRenderer(package, logger).apply_plans(plans)[0]
                self.assertTrue(result.applied, result.message)
                self.assertEqual(result.word_level_status, 'word_level')
                comments = CommentsRenderer(package, logger).apply_comments_for_results([result])
                self.assertTrue(comments[0].applied, comments[0].message)
                package.save(str(output))
            with ZipFile(output) as zip_file:
                root = etree.fromstring(zip_file.read('word/document.xml'))
            paragraph = root.xpath('.//w:body/w:p', namespaces=NS)[0]
            self.assertEqual(view_text(paragraph, False), old)
            self.assertEqual(view_text(paragraph, True), new)
            self.assertEqual(reference_locations(paragraph, False), [('42', len(old))])
            self.assertEqual(reference_locations(paragraph, True), [('42', len(new))])
            self.assertEqual(len(paragraph.xpath('.//w:footnoteReference', namespaces=NS)), 1)
            self.assertFalse(paragraph.xpath('.//w:ins//w:footnoteReference | .//w:del//w:footnoteReference', namespaces=NS))
            for element in paragraph.xpath('./w:ins | ./w:del', namespaces=NS):
                text = ''.join(x.text or '' for x in element.iter() if x.tag in (W + 't', W + 'delText'))
                if text in ('stor', 'tydlig'):
                    self.assertTrue(element.xpath('.//w:b', namespaces=NS))
                if text in ('kort', 'enkel'):
                    self.assertTrue(element.xpath('.//w:i', namespaces=NS))
            for tag in ('commentRangeStart', 'commentRangeEnd', 'commentReference'):
                self.assertEqual(len(paragraph.xpath('.//w:' + tag, namespaces=NS)), 1)
            ids = [x.get(W + 'id') for x in paragraph.xpath('./w:ins | ./w:del', namespaces=NS)]
            self.assertEqual(len(ids), len(set(ids)))

    def _segment(self, left, right, old, new, extra=None, container='plain'):
        paragraph = etree.Element(W + 'p')
        first = etree.SubElement(paragraph, W + 'r')
        etree.SubElement(first, W + 't').text = left
        parent = etree.SubElement(paragraph, W + 'hyperlink') if container == 'hyperlink' else paragraph
        second = etree.SubElement(parent, W + 'r')
        if extra:
            etree.SubElement(second, W + extra)
        etree.SubElement(second, W + 't').text = right
        return SpanSegment(container=paragraph, runs=[first, second], old=old, new=new,
                           first_run=first, last_run=second)

    def test_refuses_edits_crossing_run_boundaries(self):
        segment = self._segment('Vi läser analy', 'sen noga.', 'Vi läser analysen noga.',
                                'Vi läser rapporten noga.')
        self.assertIsNone(TrackedChangesRenderer._isolated_run_word_pieces(segment))

    def test_refuses_hyperlink_and_opaque_word_elements(self):
        hyperlink = self._segment('Vi läser ', 'en rapport.', 'Vi läser en rapport.',
                                  'Vi läser en text.', container='hyperlink')
        self.assertIsNone(TrackedChangesRenderer._isolated_run_word_pieces(hyperlink))
        field = self._segment('Vi läser ', 'en rapport.', 'Vi läser en rapport.',
                              'Vi läser en text.', extra='fldChar')
        self.assertIsNone(TrackedChangesRenderer._isolated_run_word_pieces(field))

    def test_refuses_insertion_at_ambiguous_run_boundary(self):
        segment = self._segment('Vi läser ', 'böcker.', 'Vi läser böcker.',
                                'Vi läser gärna böcker.')
        self.assertIsNone(TrackedChangesRenderer._isolated_run_word_pieces(segment))


if __name__ == '__main__':
    unittest.main()
