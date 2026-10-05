import json
import logging
import tempfile
import unittest
from pathlib import Path

from app.src.JBGLangImprovSuggestorAI import JBGLangImprovSuggestorAI


class EquivalentRewriteTests(unittest.TestCase):
    """
    Likhetsreglerna jämför tecken. Två slag av korrekta omskrivningar ser olika
    ut men är säkra: samma datum i ett annat skrivsätt, och korta ordbyten utan
    ändrad negation. Fallen kommer från testkörningarna 93dbc428 och 2991c745.
    """

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        logger = logging.getLogger(f"equivalent-{id(self)}")
        logger.handlers.clear()
        logger.addHandler(logging.NullHandler())
        logger.propagate = False
        self.logger = logger

    def tearDown(self):
        self.temp_dir.cleanup()

    def accepted(self, text, old, new, element_type="paragraph"):
        extra = {"footnote_id": "11"} if element_type == "footnote" else {}
        element = {"type": element_type, "element_id": f"{element_type}_1", "text": text, **extra}
        path = Path(self.temp_dir.name) / "s.json"
        path.write_text(json.dumps({"type": "docx", "elements": [element]}, ensure_ascii=False), encoding="utf-8")
        suggestor = JBGLangImprovSuggestorAI(api_key="k", model="m", prompt_policy="p", temperature=1, logger=self.logger)
        suggestor.load_structure(str(path))
        suggestion = {"type": element_type, "element_id": f"{element_type}_1", "old": old, "new": new, **extra}
        return bool(suggestor._postprocess_suggestions([suggestion]))

    # ---------------- Korrekta förslag som tidigare stoppades ----------------

    def test_same_date_written_out(self):
        text = "Beslutet fattades 2024-09-24 av kassan."
        self.assertTrue(self.accepted(text, "2024-09-24", "den 24 september 2024"))
        self.assertTrue(self.accepted(text, "2024-09-24", "24 september 2024"))

    def test_same_date_written_out_in_a_footnote(self):
        self.assertTrue(self.accepted("Avstämning 2025-01-31.", "2025-01-31", "den 31 januari 2025", "footnote"))

    def test_short_synonym_swaps(self):
        self.assertTrue(self.accepted("Beslutet saknar rättslig grund i lagen.", "rättslig grund", "lagstöd"))
        # I en fotnot gäller dessutom den striktare regeln för känsliga element
        self.assertTrue(self.accepted("Den sista avstämningen skedde i januari.", "skedde", "gjordes", "footnote"))

    # ---------------- Fortfarande stoppade ----------------

    def test_a_different_date_is_still_rejected(self):
        self.assertFalse(self.accepted("Beslutet fattades 2024-09-24 av kassan.", "2024-09-24", "den 25 september 2024"))
        self.assertFalse(self.accepted("Beslutet fattades 2024-09-24 av kassan.", "2024-09-24", "i september 2024"))

    def test_adding_or_removing_a_negation_does_not_count_as_a_safe_swap(self):
        # Undantaget gäller inte när en negation läggs till eller tas bort; då
        # avgör den vanliga likhetsregeln. (Ord som liknar varandra, som "alltid"
        # och "aldrig", har aldrig stoppats av likhetsregeln.)
        self.assertFalse(self.accepted("Kassan har ofta kontrollerat detta.", "ofta", "aldrig"))
        self.assertFalse(self.accepted("Kassan har inte kontrollerat detta.", "inte", "nu"))

    def test_digits_and_symbols_are_not_word_swaps(self):
        self.assertFalse(self.accepted("Andelen var 3 − 5 procent.", "−", " ", "footnote"))
        # Testkörningen 2991c745: förvanskad lagrumshänvisning i fotnot 7
        self.assertFalse(self.accepted("Enligt 6−c § gäller detta.", "−c §", " c ", "footnote"))

    def test_a_long_word_replaced_by_a_short_one_is_not_a_safe_swap(self):
        self.assertFalse(self.accepted("Arbetslöshetsförsäkringen gäller.", "Arbetslöshetsförsäkringen", "Den"))


if __name__ == "__main__":
    unittest.main()
