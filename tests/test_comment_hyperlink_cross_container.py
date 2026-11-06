"""Keep comments anchored when a tracked edit crosses a hyperlink boundary.

The fixture is synthetic; no private report text or captured OOXML is shipped.
"""
import logging
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from lxml import etree

from scripts.frozen_wordlevel_compare import W, W_NS, run_case, view_text

NS = {"w": W_NS}


class CommentHyperlinkBoundaryTests(unittest.TestCase):
    def test_comment_survives_cross_container_span_fallback(self):
        paragraph = Document().add_paragraph()
        paragraph.add_run("Vi granskar viktiga resultat. ")
        link = etree.SubElement(paragraph._p, W + "hyperlink", {W + "anchor": "Bilaga"})
        link_run = etree.SubElement(link, W + "r")
        etree.SubElement(link_run, W + "t").text = "Se bilagan för mer information."

        source = "Vi granskar viktiga resultat. Se bilagan för mer information."
        old = "viktiga resultat. Se bilagan för mer"
        new = "centrala resultat. Läs bilagan för mer"
        start = source.index(old)
        case = {
            "id": "synthetic_hyperlink_boundary",
            "source_text": source,
            "old": old,
            "new": new,
            "anchor_start": start,
            "anchor_end": start + len(old),
            "paragraph_xml": etree.tostring(paragraph._p, encoding="unicode"),
        }
        with tempfile.TemporaryDirectory() as directory:
            result = run_case(case, Path(directory), logging.getLogger(__name__))
            self.assertTrue(result["invariants_passed"])
            self.assertEqual(result["status"], "span_fallback")
            self.assertEqual(result["fallback_reason"], "multiple_edit_segments")
            # Count revisions both in paragraph runs and inside the hyperlink.
            self.assertEqual(result["revisions"]["ins"], 2)
            self.assertEqual(result["revisions"]["del"], 2)
            with ZipFile(Path(directory) / "result.docx") as archive:
                root = etree.fromstring(archive.read("word/document.xml"))
                comments = etree.fromstring(archive.read("word/comments.xml"))
            p = root.xpath(".//w:body/w:p", namespaces=NS)[0]
            self.assertEqual(view_text(p, False), source)
            self.assertEqual(view_text(p, True), source[:start] + new + source[start + len(old):])
            hyperlinks = p.xpath("./w:hyperlink", namespaces=NS)
            self.assertEqual(len(hyperlinks), 1)
            self.assertEqual(hyperlinks[0].get(W + "anchor"), "Bilaga")
            self.assertEqual(len(p.xpath("./w:commentRangeStart", namespaces=NS)), 1)
            self.assertEqual(len(p.xpath("./w:commentRangeEnd", namespaces=NS)), 1)
            self.assertEqual(len(p.xpath("./w:r/w:commentReference", namespaces=NS)), 1)
            self.assertEqual(len(comments.xpath("./w:comment", namespaces=NS)), 1)
            children = list(p)
            self.assertLess(children.index(p.xpath("./w:commentRangeStart", namespaces=NS)[0]),
                            children.index(hyperlinks[0]))
            self.assertGreater(children.index(p.xpath("./w:commentRangeEnd", namespaces=NS)[0]),
                               children.index(hyperlinks[0]))

    def test_revision_views_inside_hyperlink(self):
        paragraph = etree.Element(W + "p")
        run = etree.SubElement(paragraph, W + "r")
        etree.SubElement(run, W + "t").text = "Se "
        hyperlink = etree.SubElement(paragraph, W + "hyperlink")
        original = etree.SubElement(hyperlink, W + "del")
        original_run = etree.SubElement(original, W + "r")
        etree.SubElement(original_run, W + "delText").text = "bilagan"
        replacement = etree.SubElement(hyperlink, W + "ins")
        replacement_run = etree.SubElement(replacement, W + "r")
        etree.SubElement(replacement_run, W + "t").text = "rapporten"
        self.assertEqual(view_text(paragraph, False), "Se bilagan")
        self.assertEqual(view_text(paragraph, True), "Se rapporten")


if __name__ == "__main__":
    unittest.main()
