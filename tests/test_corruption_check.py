import json
import logging
import tempfile
import unittest
from pathlib import Path

from app.src.JBGLangImprovSuggestorAI import JBGLangImprovSuggestorAI


class CorruptionAndSimilarityCheckTests(unittest.TestCase):
    """
    Kontrollen av trasig text och likhetskontrollen får inte stoppa korrekt
    svenska som "IAF:s", "t.ex." eller ett borttaget kommatecken.
    Fallen kommer från testkörningen 2991c745.
    """

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        logger = logging.getLogger(f"corruption-{id(self)}")
        logger.handlers.clear()
        logger.addHandler(logging.NullHandler())
        logger.propagate = False
        self.logger = logger

    def tearDown(self):
        self.temp_dir.cleanup()

    def suggestor(self, elements):
        path = Path(self.temp_dir.name) / "s.json"
        path.write_text(json.dumps({"type": "docx", "elements": elements}, ensure_ascii=False), encoding="utf-8")
        suggestor = JBGLangImprovSuggestorAI(api_key="k", model="m", prompt_policy="p", temperature=1, logger=self.logger)
        suggestor.load_structure(str(path))
        return suggestor

    def accepted(self, text, old, new, element_type="paragraph", **extra):
        element = {"type": element_type, "element_id": f"{element_type}_1", "text": text, **extra}
        suggestion = {"type": element_type, "element_id": f"{element_type}_1", "old": old, "new": new, **extra}
        return bool(self.suggestor([element])._postprocess_suggestions([suggestion]))

    # ---------------- Korrekta förslag som tidigare stoppades ----------------

    def test_genitive_after_abbreviation_in_new_text(self):
        text = ("De sammanräknade resultaten ger ett högre mörkertal än tidigare uppskattningar "
                "gjorda av IAF med andra metoder.")
        self.assertTrue(self.accepted(
            text, "tidigare uppskattningar gjorda av IAF med andra metoder",
            "IAF:s tidigare uppskattningar med andra metoder"))

    def test_genitive_after_abbreviation_in_a_footnote(self):
        # Den verkliga fotnoten 11 från testkörningen
        old = ("Arbetslöshetskassornas kontroller avslutades 31 oktober 2024 och IAF:s sista avstämning av hur "
               "många misstänkta tidrapporter som genererat återkrav i efterhand skedde 31 januari 2025. I de fall "
               "en misstänkt tidrapport blivit ett återkrav redovisas dessa")
        new = ("A-kassornas kontroller avslutades den 31 oktober 2024 och IAF:s sista avstämning av hur många "
               "misstänkta tidrapporter som hade lett till återkrav i efterhand gjordes den 31 januari 2025. Om en "
               "misstänkt tidrapport ledde till ett återkrav redovisas den")
        self.assertTrue(self.accepted(old + " som återkrav.", old, new, element_type="footnote", footnote_id="11"))

    def test_comma_removal(self):
        self.assertTrue(self.accepted("Det är dock, inte helt klart.", "dock, inte", "dock inte"))

    def test_valid_punctuated_forms_are_not_garbled(self):
        suggestor = self.suggestor([])
        for text in ("IAF:s rapport", "ST:s beslut", "t.ex. kraftigt", "bl.a. ärenden", "fr.o.m. 2025",
                     "se iaf.se", "skriv till info@iaf.se", "skrift", "strängt", "avsnitt 4.2.1"):
            with self.subTest(text=text):
                self.assertFalse(suggestor._looks_like_corrupted_text(text))

    # ---------------- Trasig text och tveksamma ändringar stoppas fortfarande ----------------

    def test_garbled_text_is_still_detected(self):
        suggestor = self.suggestor([])
        for text in ("kontrollernaVisar att", "slut.Nästa mening", "ut.Nästa", "re,sultat", "rap;port", "!!!!"):
            with self.subTest(text=text):
                self.assertTrue(suggestor._looks_like_corrupted_text(text))

    def test_other_one_character_replacements_are_still_rejected(self):
        # Testkörningen 2991c745: minustecken ersatt med blanksteg i en fotnot
        self.assertFalse(self.accepted("Andelen var 3 − 5 procent.", "−", " ",
                                       element_type="footnote", footnote_id="7"))


if __name__ == "__main__":
    unittest.main()
