"""Regression tests for optional word-level tracked change rendering."""
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

NS = {"w": W_NS}


class WordLevelChangesTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.logger = logging.getLogger(f"word-level-{id(self)}")
        self.logger.addHandler(logging.NullHandler())

    def tearDown(self):
        self.tmp.cleanup()

    def render(self, old, new, enabled, multi_run=False, same_style_runs=False):
        source = self.root / 'source.docx'
        destination = self.root / ('on.docx' if enabled else 'off.docx')
        document = Document()
        paragraph = document.add_paragraph()
        if multi_run or same_style_runs:
            cut = old.index(' granskning')
            paragraph.add_run(old[:cut])
            trailing = paragraph.add_run(old[cut:])
            if multi_run:
                trailing.bold = True
        else:
            paragraph.add_run(old)
        document.save(source)
        structure = DocumentStructureExtractor(str(source), self.logger).extract()
        element = next(e for e in structure['elements'] if e.get('text') == old and e.get('type') == 'paragraph')
        suggestion = SuggestedChange(element_type='paragraph', element_id=element['element_id'],
                                     footnote_id=None, old=old, new=new,
                                     motivation='Använd enklare språk.', match_status='exact')
        plans = ChangePlanner(structure, TokenDiffEngine(), self.logger).build_plans([suggestion])
        self.assertEqual(len(plans), 1)
        with DocxPackage(str(source), self.logger) as package:
            render_results = TrackedChangesRenderer(package, self.logger, word_level_changes=enabled).apply_plans(plans)
            self.assertTrue(render_results[0].applied, render_results[0].message)
            comments = CommentsRenderer(package, self.logger).apply_comments_for_results(render_results)
            self.assertTrue(comments[0].applied, comments[0].message)
            package.save(str(destination))
        with ZipFile(destination) as archive:
            document_tree = etree.fromstring(archive.read('word/document.xml'))
            comments_tree = etree.fromstring(archive.read('word/comments.xml'))
        return document_tree, comments_tree

    @staticmethod
    def text_from_revision(root, accept):
        paragraph = root.xpath('.//w:body/w:p', namespaces=NS)[0]
        result = []
        for child in paragraph:
            tag = etree.QName(child).localname
            if tag == 'r':
                result.extend(child.xpath('./w:t/text()', namespaces=NS))
            elif tag == ('ins' if accept else 'del'):
                result.extend(child.xpath('.//w:t/text() | .//w:delText/text()', namespaces=NS))
        return ''.join(result)

    def test_word_level_preserves_unchanged_text_and_comment_spans_all_revisions(self):
        old = 'Myndigheten har gjort en omfattande granskning och därefter fattat beslut.'
        new = 'Myndigheten har gjort en grundlig granskning och sedan fattat beslut.'
        root, comments = self.render(old, new, True)
        self.assertEqual(self.text_from_revision(root, accept=False), old)
        self.assertEqual(self.text_from_revision(root, accept=True), new)
        self.assertEqual(len(root.xpath('.//w:del', namespaces=NS)), 2)
        self.assertEqual(len(root.xpath('.//w:ins', namespaces=NS)), 2)
        self.assertEqual(len(comments.xpath('.//w:comment', namespaces=NS)), 1)
        paragraph = root.xpath('.//w:body/w:p', namespaces=NS)[0]
        tags = [etree.QName(e).localname for e in paragraph]
        self.assertLess(tags.index('commentRangeStart'), tags.index('del'))
        self.assertGreater(tags.index('commentRangeEnd'), tags.index('ins'))
        self.assertEqual(len(root.xpath('.//w:commentReference', namespaces=NS)), 1)

    def test_disabled_keeps_span_behavior(self):
        old = 'Myndigheten har gjort en omfattande granskning och därefter fattat beslut.'
        new = 'Myndigheten har gjort en grundlig granskning och sedan fattat beslut.'
        root, _ = self.render(old, new, False)
        self.assertEqual(self.text_from_revision(root, False), old)
        self.assertEqual(self.text_from_revision(root, True), new)
        self.assertEqual(len(root.xpath('.//w:del', namespaces=NS)), 1)
        self.assertEqual(len(root.xpath('.//w:ins', namespaces=NS)), 1)

    def test_multirun_falls_back_to_span(self):
        old = 'Vi gör en omfattande granskning av handlingarna.'
        new = 'Vi gör en grundlig granskning av handlingarna.'
        root, _ = self.render(old, new, True, multi_run=True)
        self.assertEqual(self.text_from_revision(root, False), old)
        self.assertEqual(self.text_from_revision(root, True), new)
        self.assertEqual(len(root.xpath('.//w:del', namespaces=NS)), 1)

    def test_same_format_multirun_uses_word_level_and_keeps_comment(self):
        old = 'Vi gör en omfattande granskning och därefter fattar vi beslut.'
        new = 'Vi gör en grundlig granskning och sedan fattar vi beslut.'
        root, comments = self.render(old, new, True, same_style_runs=True)
        self.assertEqual(self.text_from_revision(root, False), old)
        self.assertEqual(self.text_from_revision(root, True), new)
        self.assertEqual(len(root.xpath('.//w:del', namespaces=NS)), 2)
        self.assertEqual(len(root.xpath('.//w:ins', namespaces=NS)), 2)
        self.assertEqual(len(comments.xpath('.//w:comment', namespaces=NS)), 1)
        paragraph = root.xpath('.//w:body/w:p', namespaces=NS)[0]
        tags = [etree.QName(e).localname for e in paragraph]
        self.assertLess(tags.index('commentRangeStart'), tags.index('del'))
        self.assertGreater(tags.index('commentRangeEnd'), tags.index('ins'))

    def test_nbsp_does_not_normalize_source(self):
        old = 'Beslutet gäller den 1\u00a0januari 2027.'
        new = 'Beslutet gäller från den 1\u00a0januari 2027.'
        root, _ = self.render(old, new, True)
        self.assertEqual(self.text_from_revision(root, False), old)
        self.assertEqual(self.text_from_revision(root, True), new)
        self.assertEqual(len(root.xpath('.//w:del', namespaces=NS)), 0)
        self.assertEqual(len(root.xpath('.//w:ins', namespaces=NS)), 1)

    def test_exact_diff_handles_nbsp_and_multiple_replacements(self):
        old = 'Myndigheten har en omfattande granskning\u00a0och sedan ett beslut.'
        new = 'Myndigheten har en grundlig granskning\u00a0och därefter ett beslut.'
        root, comments = self.render(old, new, True, same_style_runs=True)
        self.assertEqual(self.text_from_revision(root, False), old)
        self.assertEqual(self.text_from_revision(root, True), new)
        self.assertGreaterEqual(len(root.xpath('.//w:del', namespaces=NS)), 2)
        self.assertEqual(len(comments.xpath('.//w:comment', namespaces=NS)), 1)

    def test_exact_diff_preserves_other_unicode_spaces(self):
        old = 'Uppgifterna gäller 12\u202f345 personer i dag.'
        new = 'Uppgifterna gäller 12\u202f345 personer just i dag.'
        root, _ = self.render(old, new, True)
        self.assertEqual(self.text_from_revision(root, False), old)
        self.assertEqual(self.text_from_revision(root, True), new)
        self.assertEqual(len(root.xpath('.//w:del', namespaces=NS)), 0)
        self.assertEqual(len(root.xpath('.//w:ins', namespaces=NS)), 1)

    def test_no_common_tokens_still_falls_back_safely(self):
        old, new = 'Hej.', 'Adjö!'
        root, _ = self.render(old, new, True)
        self.assertEqual(self.text_from_revision(root, False), old)
        self.assertEqual(self.text_from_revision(root, True), new)
        self.assertEqual(len(root.xpath('.//w:del', namespaces=NS)), 1)
        self.assertEqual(len(root.xpath('.//w:ins', namespaces=NS)), 1)

class MultiRunEligibilityTests(unittest.TestCase):
    """Conservative eligibility checks independent of the OOXML editor."""

    def setUp(self):
        from types import SimpleNamespace
        self.ns = SimpleNamespace

    def paragraph_and_segment(self, chunks, bold_index=None):
        root = etree.Element(f'{{{W_NS}}}p')
        runs = []
        for i, chunk in enumerate(chunks):
            run = etree.SubElement(root, f'{{{W_NS}}}r')
            if i == bold_index:
                etree.SubElement(etree.SubElement(run, f'{{{W_NS}}}rPr'), f'{{{W_NS}}}b')
            etree.SubElement(run, f'{{{W_NS}}}t').text = chunk
            runs.append(run)
        old = ''.join(chunks)
        segment = self.ns(first_run=runs[0], last_run=runs[-1], before='', old=old,
                          new=old.replace('omfattande', 'grundlig'), after='')
        return root, segment, old

    def test_equally_formatted_runs_are_eligible(self):
        paragraph, segment, old = self.paragraph_and_segment(
            ['Vi gör en ', 'omfattande ', 'granskning.'])
        self.assertIsNone(TrackedChangesRenderer._word_run_safety_reason(
            paragraph, segment, old, segment.new))

    def test_mixed_formatting_is_rejected(self):
        paragraph, segment, old = self.paragraph_and_segment(
            ['Vi gör en ', 'omfattande ', 'granskning.'], bold_index=1)
        self.assertEqual(TrackedChangesRenderer._word_run_safety_reason(
            paragraph, segment, old, segment.new), 'mixed_run_formatting')

    def test_interleaved_comment_markers_are_rejected(self):
        paragraph, segment, old = self.paragraph_and_segment(
            ['Vi gör en ', 'omfattande ', 'granskning.'])
        paragraph.insert(1, etree.Element(f'{{{W_NS}}}commentRangeStart'))
        self.assertEqual(TrackedChangesRenderer._word_run_safety_reason(
            paragraph, segment, old, segment.new), 'interleaved_markup')

    def test_hidden_text_or_other_run_children_are_rejected(self):
        paragraph, segment, old = self.paragraph_and_segment(['Vi gör ', 'granskning.'])
        etree.SubElement(paragraph[0], f'{{{W_NS}}}tab')
        self.assertEqual(TrackedChangesRenderer._word_run_safety_reason(
            paragraph, segment, old, segment.new), 'complex_run_tabs_or_breaks')
