import logging
import unittest

from app.src.JBGLangImprovSuggestorAI import JBGLangImprovSuggestorAI


class SpellingDegradationTests(unittest.TestCase):
    def setUp(self):
        logger = logging.getLogger(f"spelling-safety-test-{id(self)}")
        logger.handlers.clear()
        logger.addHandler(logging.NullHandler())
        self.suggestor = JBGLangImprovSuggestorAI(
            api_key="unused",
            model="unused",
            prompt_policy="",
            temperature=0,
            logger=logger,
        )

    def test_allows_legitimate_plain_language_rewrites(self):
        rewrites = (
            (
                "arbetslöshetskassor har fått möjlighet att faktagranska kapitel 1",
                "a-kassor har fått möjlighet att faktagranska kapitel 1",
            ),
            (
                "IAF har gjort uppskattningen med hjälp av AI-baserade",
                "Vi har gjort uppskattningen med hjälp av AI-baserade",
            ),
            ("av de felaktiga tidrapporterna", "för felaktiga tidrapporter"),
            ("Arbetslöshetskassornas", "A-kassornas"),
            ("kontrollerna", "kontroller"),
        )

        for old, new in rewrites:
            with self.subTest(old=old, new=new):
                self.assertFalse(self.suggestor._looks_like_spelling_degradation(old, new))

    def test_rejects_probable_internal_character_loss(self):
        degradations = (
            ("rapporten", "raporten"),
            ("kontrollerna", "kontrolerna"),
            ("organisation", "organsation"),
        )

        for old, new in degradations:
            with self.subTest(old=old, new=new):
                self.assertTrue(self.suggestor._looks_like_spelling_degradation(old, new))

    def test_allows_corrections_and_non_deletion_edits(self):
        changes = (
            ("raporten", "rapporten"),
            ("organisation", "organisering"),
            ("kontroll", "granskning"),
        )

        for old, new in changes:
            with self.subTest(old=old, new=new):
                self.assertFalse(self.suggestor._looks_like_spelling_degradation(old, new))

    def test_long_plain_language_rewrite_is_not_distorted_by_autojunk(self):
        old = (
            "Arbetslöshetskassorna granskar ärenden och gör uppföljning "
            "av arbetslöshetsförsäkringen. "
        ) * 8
        new = (
            "A-kassorna granskar ärenden och följer upp "
            "arbetslöshetsförsäkringen. "
        ) * 8

        self.assertGreater(self.suggestor._similarity_ratio(old, new), 0.80)
        self.assertFalse(self.suggestor._too_low_overlap(old, new, "table_cell"))


if __name__ == "__main__":
    unittest.main()
