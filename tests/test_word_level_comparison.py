"""Deterministic A/B comparison of the same frozen suggestions in both renderers.

No language-model calls, no random text and no document from a prior run.
The test uses the same original DOCX and ChangePlan for each renderer mode.
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

NS = {"w": W_NS}

# These examples are deliberately fixed: differences between runs reflect the
# renderer, not nondeterministic suggestions from a language model.
CASES = (
    ("two_replacements", "Myndigheten har gjort en omfattande granskning och därefter fattat beslut.",
     "Myndigheten har gjort en grundlig granskning och sedan fattat beslut.", "single", "word_level", None),
    ("same_format_runs", "Vi gör en omfattande granskning och därefter fattar vi beslut.",
     "Vi gör en grundlig granskning och sedan fattar vi beslut.", "same", "word_level", None),
    ("mixed_format_runs", "Vi gör en omfattande granskning och därefter fattar vi beslut.",
     "Vi gör en grundlig granskning och sedan fattar vi beslut.", "mixed", "span_fallback", "mixed_run_formatting"),
    ("mixed_format_safe_boundary", "Vi gör en omfattande granskning och därefter fattar vi beslut.",
     "Vi gör en grundlig granskning och sedan fattar vi beslut.", "mixed_safe", "word_level", None),
    ("mixed_format_boundary_changed", "Vi gör en omfattande granskning och därefter fattar vi beslut.",
     "Vi gör en grundlig utredning och sedan fattar vi beslut.", "mixed", "span_fallback", "mixed_run_formatting"),
    ("tab_in_run", "Vi gör en omfattande granskning\tinnan beslut.",
     "Vi gör en grundlig granskning\tinnan beslut.", "single", "span_fallback", "complex_run_tabs_or_breaks"),
    ("nb_space", "Vi granskar 1\u00a0234 ärenden och fattar beslut.",
     "Vi granskar 1\u00a0234 ansökningar och fattar beslut.", "single", "word_level", None),
    ("narrow_nb_space", "Vi granskar 1\u202f234 ärenden och fattar beslut.",
     "Vi granskar 1\u202f234 ansökningar och fattar beslut.", "single", "word_level", None),
    ("no_shared_tokens", "Hej.", "Adjö!", "single", "span_fallback", "diff_no_common_text"),
)


def accepted_or_rejected_text(paragraph, accept):
    """Simulate Word's accept/reject text selection, including tabs and breaks."""
    parts = []
    for child in paragraph:
        tag = etree.QName(child).localname
        if tag != "r" and tag != ("ins" if accept else "del"):
            continue
        for item in child.iter():
            localname = etree.QName(item).localname
            if localname in {"t", "delText"}:
                parts.append(item.text or "")
            elif localname == "tab":
                parts.append("\t")
            elif localname in {"br", "cr"}:
                parts.append("\n")
    return "".join(parts)


class DeterministicWordLevelComparisonTests(unittest.TestCase):
    def test_same_frozen_plans_for_span_and_word_level(self):
        logger = logging.getLogger("word-level-frozen-comparison")
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            counts = {False: {"ins": 0, "del": 0}, True: {"ins": 0, "del": 0}}
            proposal_counts = {"word_level": 0, "span_fallback": 0}
            for name, old, new, structure_kind, expected_status, expected_reason in CASES:
                with self.subTest(case=name):
                    source = tmp / f"{name}.docx"
                    document = Document()
                    p = document.add_paragraph()
                    if structure_kind in {"same", "mixed", "mixed_safe"}:
                        cut = old.index(" och därefter") if structure_kind == "mixed_safe" else old.index(" granskning")
                        p.add_run(old[:cut])
                        run = p.add_run(old[cut:])
                        if structure_kind in {"mixed", "mixed_safe"}:
                            run.bold = True
                    else:
                        p.add_run(old)
                    document.save(source)
                    structure = DocumentStructureExtractor(str(source), logger).extract()
                    element = next(e for e in structure["elements"] if e.get("type") == "paragraph" and e.get("text") == old)
                    suggestion = SuggestedChange(element_type="paragraph", element_id=element["element_id"],
                                                 footnote_id=None, old=old, new=new,
                                                 motivation="Fast testförslag.", match_status="exact")
                    plans = ChangePlanner(structure, TokenDiffEngine(), logger).build_plans([suggestion])
                    self.assertEqual(len(plans), 1)
                    # Both renderers receive this exact same plan, not two AI results.
                    for enabled in (False, True):
                        output = tmp / f"{name}-{'word' if enabled else 'span'}.docx"
                        with DocxPackage(str(source), logger) as package:
                            results = TrackedChangesRenderer(package, logger, word_level_changes=enabled).apply_plans(plans)
                            self.assertEqual(len(results), 1)
                            self.assertTrue(results[0].applied, results[0].message)
                            comments = CommentsRenderer(package, logger).apply_comments_for_results(results)
                            self.assertEqual(len(comments), 1)
                            self.assertTrue(comments[0].applied, comments[0].message)
                            package.save(str(output))
                        with ZipFile(output) as z:
                            root = etree.fromstring(z.read("word/document.xml"))
                            com = etree.fromstring(z.read("word/comments.xml"))
                        paragraph = root.xpath(".//w:body/w:p", namespaces=NS)[0]
                        self.assertEqual(accepted_or_rejected_text(paragraph, False), old)
                        self.assertEqual(accepted_or_rejected_text(paragraph, True), new)
                        self.assertEqual(len(com.xpath(".//w:comment", namespaces=NS)), 1)
                        for marker in ("commentRangeStart", "commentRangeEnd", "commentReference"):
                            self.assertEqual(len(paragraph.xpath(f".//w:{marker}", namespaces=NS)), 1)
                        for kind in ("ins", "del"):
                            elements = paragraph.xpath(f"./w:{kind}", namespaces=NS)
                            counts[enabled][kind] += len(elements)
                            ids = [el.get(f"{{{W_NS}}}id") for el in elements]
                            self.assertEqual(len(ids), len(set(ids)))
                        self.assertEqual(results[0].word_level_status, expected_status if enabled else None)
                        if enabled:
                            self.assertEqual(results[0].word_level_fallback_reason, expected_reason)
                            proposal_counts[results[0].word_level_status] += 1
            # Proposal-level metric: exact same frozen suggestions for both renderers.
            # We aim for 80% word-level among APPLIED proposals over time, not
            # 80% of words or characters. This corpus is a guardrail, not an
            # estimate of performance on real reports.
            self.assertEqual(proposal_counts, {"word_level": 5, "span_fallback": 4})
            self.assertGreater(counts[True]["ins"] + counts[True]["del"],
                               counts[False]["ins"] + counts[False]["del"])


if __name__ == "__main__":
    unittest.main()

class ComplexRunDiagnosticTests(unittest.TestCase):
    """Detailed reasons must remain conservative: no new editable elements."""

    def test_specific_complex_run_reasons(self):
        from types import SimpleNamespace
        cases = (
            ("tab", "complex_run_tabs_or_breaks"),
            ("br", "complex_run_tabs_or_breaks"),
            ("fldChar", "complex_run_fields_or_objects"),
            ("drawing", "complex_run_fields_or_objects"),
            ("bookmarkStart", "complex_run_other_elements"),
        )
        for child_tag, expected in cases:
            with self.subTest(child=child_tag):
                paragraph = etree.Element(f"{{{W_NS}}}p")
                run = etree.SubElement(paragraph, f"{{{W_NS}}}r")
                etree.SubElement(run, f"{{{W_NS}}}t").text = "Vi granskar ett ärende."
                etree.SubElement(run, f"{{{W_NS}}}{child_tag}")
                segment = SimpleNamespace(first_run=run, last_run=run, before="",
                                          old="Vi granskar ett ärende.",
                                          new="Vi granskar två ärenden.", after="")
                self.assertEqual(TrackedChangesRenderer._word_run_safety_reason(
                    paragraph, segment, segment.old, segment.new), expected)

    def test_run_without_text_uses_fallback(self):
        from types import SimpleNamespace
        paragraph = etree.Element(f"{{{W_NS}}}p")
        run = etree.SubElement(paragraph, f"{{{W_NS}}}r")
        etree.SubElement(run, f"{{{W_NS}}}rPr")
        segment = SimpleNamespace(first_run=run, last_run=run, before="", old="A", new="B", after="")
        self.assertEqual(TrackedChangesRenderer._word_run_safety_reason(
            paragraph, segment, "A", "B"), "complex_run_without_text")


if __name__ == "__main__":
    unittest.main()
