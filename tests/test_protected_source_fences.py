"""Hard safety gates for reference lists, footnotes and template paragraphs."""
import logging
import unittest

from app.src.JBGLangImprovSuggestorAI import JBGLangImprovSuggestorAI, SuggestedChange


def el(ident, text, order=None, level=None, type_="paragraph"):
    return {"element_id": ident, "type": type_, "text": text,
            "doc_order": order, "heading_level": level}


class ProtectedSourceFencesTests(unittest.TestCase):
    def setUp(self):
        self.ai = JBGLangImprovSuggestorAI(
            api_key="unused", model="unused", prompt_policy="", temperature=0,
            logger=logging.getLogger("source-fences"),
        )
        self.elements = [
            el("p1", "Slutsatser", 1, 1),
            el("p2", "Arbetsförmedlingen har genomfört en granskning.", 2),
            el("p3", "Som ett led i IAF:s uppföljning ska Arbetsförmedlingen redogöra för [precisera vilka brister som avses].", 3),
            el("p4", "Källor", 4, 1),
            el("p5", "Författningar", 5, 2),
            el("p6", "Förvaltningslag (2017:900).", 6),
            el("p7", "Övrigt", 7, 2),
            el("p8", "JO beslut, Arbetsförmedlingen får allvarlig kritik bl.a. för handläggningen, dnr 5533-2025.", 8),
            el("p9", "Bilaga 1: Granskningsmetod", 9, 1),
            el("p10", "IAF bedömer att utredningen är avslutad.", 10),
            el("p11", "Katrineholm den DD månad ÅÅÅÅ", 11),
            el("p12", "Börja skriva här", 12),
            el("p13", "Här är vår referens till lag (2017:900), men inte hela paragrafen.", 13),
            el("f1", "Arbetsförmedlingens handläggarstöd Jobbgarantin för ungdomar, version 27.", type_="footnote"),
            el("f2", "Undantag gäller för personer som får ersättning under perioden.", type_="footnote"),
            el("f3", "23 § Förvaltningslag (2017:900).", type_="footnote"),
        ]
        self.ai.json_structured_document = {"type": "docx", "elements": self.elements}

    def check(self, id_, old, new, reason):
        typ = "footnote" if id_.startswith("f") else "paragraph"
        s = SuggestedChange(typ, id_, None, old, new, None, "exact")
        self.assertEqual(self.ai._protected_source_reason(s), reason)
        self.assertEqual(self.ai._assess_suggestion_safety(s)["reject"], bool(reason))

    def test_source_section_allows_no_internal_edits_but_stops_at_next_heading(self):
        self.check("p6", "Förvaltningslag", "Förvaltningslagen", "protected_bibliography_section")
        self.check("p8", "bl.a.", "bland annat", "protected_bibliography_section")
        self.check("p2", "genomfört", "gjort", None)
        self.check("p10", "bedömer", "anser", None)

    def test_whole_template_paragraph_protected_even_for_tiny_rewrite(self):
        self.check("p3", "Arbetsförmedlingen", "myndigheten", "protected_template_placeholder")
        self.check("p11", "Katrineholm", "I Katrineholm", "protected_template_placeholder")
        self.check("p12", "Börja", "Starta", "protected_template_text")

    def test_reference_footnotes_protected_explanatory_footnotes_editable(self):
        self.check("f1", "version", "utgåva", "protected_reference_footnote")
        self.check("f3", "Förvaltningslag", "Förvaltningslagen", "protected_reference_footnote")
        self.check("f2", "personer", "deltagare", None)

    def test_ordinary_prose_is_not_frozen_by_authority_name_or_citation(self):
        self.check("p13", "vår referens", "hänvisningen", None)

    def test_missing_structural_metadata_does_not_guess_bibliography(self):
        self.ai.json_structured_document = {"type": "docx", "elements": [
            el("single", "Arbetsförmedlingen har granskat ärenden.")
        ]}
        self.check("single", "har granskat", "granskade", None)


class ProtectedSourcePipelineTests(unittest.TestCase):
    setUp = ProtectedSourceFencesTests.setUp

    def test_reference_and_template_changes_are_filtered_before_planner(self):
        proposals = [
            {"type": "paragraph", "element_id": "p8", "old": "bl.a.", "new": "bland annat"},
            {"type": "paragraph", "element_id": "p3", "old": "Arbetsförmedlingen", "new": "myndigheten"},
            {"type": "paragraph", "element_id": "p2", "old": "genomfört", "new": "gjort"},
        ]
        retained = self.ai._postprocess_suggestions(proposals)
        self.assertEqual([(s.element_id, s.old, s.new) for s in retained], [
            ("p2", "genomfört", "gjort")
        ])

if __name__ == "__main__":
    unittest.main()
