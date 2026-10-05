import json
import logging
import tempfile
import unittest
from pathlib import Path

from app.src.JBGLangImprovSuggestorAI import JBGLangImprovSuggestorAI, SuggestedChange


class SpanMinimizationTests(unittest.TestCase):
    """
    Trimning av förslag till minsta ändring får aldrig dela ett ord, och ett
    fragment avgörs av var spannet ligger i texten, inte av ordets längd.
    Fallen kommer från testkörningarna ba144f3d och 2cfb505b.
    """

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        logger = logging.getLogger(f"span-{id(self)}")
        logger.handlers.clear()
        logger.addHandler(logging.NullHandler())
        logger.propagate = False
        self.logger = logger

    def tearDown(self):
        self.temp_dir.cleanup()

    def minimize(self, text, old, new):
        path = Path(self.temp_dir.name) / "s.json"
        path.write_text(json.dumps({"type": "docx", "elements": [
            {"type": "paragraph", "element_id": "paragraph_1", "text": text}]}, ensure_ascii=False), encoding="utf-8")
        suggestor = JBGLangImprovSuggestorAI(api_key="k", model="m", prompt_policy="p", temperature=1, logger=self.logger)
        suggestor.load_structure(str(path))
        return suggestor._minimize_and_filter_suggestion(SuggestedChange(
            element_type="paragraph", element_id="paragraph_1", footnote_id=None,
            old=old, new=new, motivation=None, match_status="exact",
        ))

    def assertAccepted(self, text, old, new, expected_old, expected_new):
        result = self.minimize(text, old, new)
        self.assertIsNotNone(result)
        self.assertTrue(result.safe_to_apply, result.safety_reason)
        self.assertEqual((result.old, result.new), (expected_old, expected_new))

    def assertRejected(self, text, old, new):
        result = self.minimize(text, old, new)
        self.assertIsNotNone(result)
        self.assertFalse(result.safe_to_apply, (result.old, result.new))

    # ---------------- Förslag som tidigare gick förlorade ----------------

    def test_abbreviations_are_kept_whole(self):
        self.assertAccepted("Detta gäller bl.a. ärenden.", "bl.a.", "bland annat", "bl.a.", "bland annat")
        self.assertAccepted("Antalet ökade t.ex. kraftigt.", "t.ex.", "till exempel", "t.ex.", "till exempel")

    def test_shared_start_of_a_word_is_not_cut_on_the_new_side(self):
        self.assertAccepted("Andel 40–63 % ha diagranm%", "40–63 % ha diagranm%", "40–63 % har diagram",
                            "ha diagranm%", "har diagram")

    def test_short_words_and_acronyms_are_whole_words(self):
        self.assertAccepted("IAF ansvarar för tillsynen.", "IAF",
                            "Inspektionen för arbetslöshetsförsäkringen (IAF)",
                            "IAF", "Inspektionen för arbetslöshetsförsäkringen (IAF)")
        self.assertAccepted("Det är dock inte klart.", "dock inte", "dock, inte", "dock", "dock,")
        self.assertAccepted("Vi läste skrift efter skrift.", "skrift efter skrift", "dokument efter dokument",
                            "skrift efter skrift", "dokument efter dokument")

    def test_whitespace_only_addition_is_ignored(self):
        self.assertIsNone(self.minimize("Enligt 6 d § lagen gäller detta.", "6 d §", "6 d § "))

    # ---------------- Oförändrat beteende ----------------

    def test_ordinary_cases_are_unchanged(self):
        self.assertAccepted("Myndigheten har genomfört en granskning av ärendet.",
                            "Myndigheten har genomfört en granskning av ärendet.",
                            "Myndigheten har granskat ärendet.",
                            "genomfört en granskning av", "granskat")
        self.assertAccepted("Vi granskade kontrolerna noga.", "kontrolerna", "kontrollerna",
                            "kontrolerna", "kontrollerna")
        self.assertAccepted("Beslutet fattades 2024-09-24 av kassan.", "2024-09-24", "den 24 september 2024",
                            "2024-09-24", "den 24 september 2024")

    def test_a_change_inside_a_word_is_widened_to_the_whole_word(self):
        self.assertAccepted("Vi granskade kontrolerna noga.", "Vi granskade kontrolerna noga.",
                            "Vi granskade kontrollerna noga.", "kontrolerna", "kontrollerna")

    # ---------------- Riktiga fragment stoppas fortfarande ----------------

    def test_fragment_from_the_model_inside_an_abbreviation_is_rejected(self):
        self.assertRejected("Detta gäller bl.a. ärenden.", ".a.", "and annat")

    def test_fragment_after_the_first_dot_of_an_abbreviation_is_rejected(self):
        # Utan regeln för punkter i förkortningar skulle "t.ex." bli "t.exempel"
        self.assertRejected("Antalet ökade t.ex. kraftigt.", "ex.", "exempel")

    def test_new_text_glued_onto_a_neighbouring_word_is_rejected(self):
        self.assertRejected("Andel 40–63 % ha diagranm%", " diagranm%", "r diagram")


if __name__ == "__main__":
    unittest.main()
