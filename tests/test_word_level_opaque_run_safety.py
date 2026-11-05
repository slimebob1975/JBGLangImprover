"""A lossy span fallback must never remove footnote/field/object XML."""
import unittest
from types import SimpleNamespace
from lxml import etree

from app.src.JBGDocxPackage import W_NS
from app.src.JBGTrackedChangesRenderer import TrackedChangesRenderer

W = f"{{{W_NS}}}"


class OpaqueRunSafetyTests(unittest.TestCase):
    def make_span(self, opaque_tag, text='Vi granskar ett ärende.'):
        paragraph = etree.Element(f'{W}p')
        run = etree.SubElement(paragraph, f'{W}r')
        t = etree.SubElement(run, f'{W}t')
        t.text = text
        opaque = etree.SubElement(run, f'{W}{opaque_tag}')
        if opaque_tag == 'footnoteReference':
            opaque.set(f'{W}id', '7')
        node = SimpleNamespace(run_element=run, element=t, start=0, end=len(text))
        return paragraph, run, node, text

    def test_nontext_elements_are_rejected_without_editing(self):
        for tag in ('footnoteReference', 'fldChar', 'instrText', 'drawing', 'pict', 'object'):
            with self.subTest(tag=tag):
                paragraph, run, node, old = self.make_span(tag)
                before = etree.tostring(paragraph)
                renderer = object.__new__(TrackedChangesRenderer)
                renderer.word_level_changes = True
                with self.assertRaisesRegex(ValueError, f'unsafe_span_run_content:{tag}'):
                    renderer._rewrite_span(paragraph, node, node, 0, len(old), old,
                                           'Vi granskar två ärenden.')
                self.assertEqual(etree.tostring(paragraph), before)

    def test_same_protection_when_word_level_off(self):
        paragraph, run, node, old = self.make_span('footnoteReference')
        renderer = object.__new__(TrackedChangesRenderer)
        renderer.word_level_changes = False
        before = etree.tostring(paragraph)
        with self.assertRaisesRegex(ValueError, 'unsafe_span_run_content:footnoteReference'):
            renderer._rewrite_span(paragraph, node, node, 0, len(old), old,
                                   'Vi granskar två ärenden.')
        self.assertEqual(etree.tostring(paragraph), before)

    def test_plain_tab_and_break_are_reproducible(self):
        paragraph = etree.Element(f'{W}p')
        run = etree.SubElement(paragraph, f'{W}r')
        etree.SubElement(run, f'{W}t').text = 'A'
        etree.SubElement(run, f'{W}tab')
        etree.SubElement(run, f'{W}br')
        segment = SimpleNamespace(runs=[run])
        self.assertEqual(TrackedChangesRenderer._unsafe_span_run_elements([segment]), ())
        # Break type/clear attributes cannot safely be regenerated.
        run[-1].set(f'{W}type', 'page')
        self.assertEqual(TrackedChangesRenderer._unsafe_span_run_elements([segment]), ('br',))

    def test_multisegment_is_checked_before_first_rewrite(self):
        paragraph = etree.Element(f'{W}p')
        r1 = etree.SubElement(paragraph, f'{W}r')
        etree.SubElement(r1, f'{W}t').text = 'A'
        r2 = etree.SubElement(paragraph, f'{W}r')
        etree.SubElement(r2, f'{W}t').text = 'B'
        etree.SubElement(r2, f'{W}footnoteReference').set(f'{W}id', '7')
        segments = [SimpleNamespace(runs=[r1]), SimpleNamespace(runs=[r2])]
        self.assertEqual(TrackedChangesRenderer._unsafe_span_run_elements(segments), ('footnoteReference',))


if __name__ == '__main__':
    unittest.main()

class OpaqueRunIntegrationTests(unittest.TestCase):
    """The public renderer must fail safely without rewriting the package."""

    def test_footnote_reference_survives_failed_proposal(self):
        import logging
        import tempfile
        from pathlib import Path
        from zipfile import ZipFile
        from docx import Document
        from app.src.JBGDocxPackage import DocxPackage
        from app.src.JBGDocumentStructureExtractor import DocumentStructureExtractor
        from app.src.JBGChangePlanner import ChangePlanner
        from app.src.JBGTokenDiffEngine import TokenDiffEngine
        from app.src.JBGLangImprovSuggestorAI import SuggestedChange

        logger = logging.getLogger('footnote-loss-guard')
        old = 'Myndigheten granskar ett ärende.'
        new = 'Myndigheten granskar ett särskilt ärende.'
        with tempfile.TemporaryDirectory() as td:
            source = Path(td) / 'source.docx'
            destination = Path(td) / 'result.docx'
            document = Document()
            run = document.add_paragraph().add_run(old)
            etree.SubElement(run._r, f'{W}footnoteReference').set(f'{W}id', '7')
            document.save(source)
            structure = DocumentStructureExtractor(str(source), logger).extract()
            element = next(e for e in structure['elements']
                           if e.get('type') == 'paragraph' and e.get('text') == old)
            suggestion = SuggestedChange(element_type='paragraph', element_id=element['element_id'],
                                         footnote_id=None, old=old, new=new,
                                         motivation='Klarspråk.', match_status='exact')
            plans = ChangePlanner(structure, TokenDiffEngine(), logger).build_plans([suggestion])
            self.assertEqual(len(plans), 1)
            for enabled in (False, True):
                with self.subTest(word_level=enabled):
                    with DocxPackage(str(source), logger) as pkg:
                        results = TrackedChangesRenderer(pkg, logger, word_level_changes=enabled).apply_plans(plans)
                        self.assertEqual(len(results), 1)
                        self.assertFalse(results[0].applied)
                        self.assertIn('unsafe_span_run_content:footnoteReference', results[0].message)
                        pkg.save(str(destination))
                    with ZipFile(destination) as z:
                        root = etree.fromstring(z.read('word/document.xml'))
                    self.assertEqual(len(root.xpath('.//w:footnoteReference', namespaces={'w': W_NS})), 1)
                    self.assertEqual(len(root.xpath('.//w:del | .//w:ins', namespaces={'w': W_NS})), 0)
