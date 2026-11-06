"""Deterministic proper-name preservation in the pre-render safety gate."""
import logging
import unittest

from app.src.JBGLangImprovSuggestorAI import JBGLangImprovSuggestorAI, SuggestedChange


class ProperNameFencesTests(unittest.TestCase):
    def setUp(self):
        self.ai = JBGLangImprovSuggestorAI(
            api_key="unused", model="unused", prompt_policy="", temperature=0,
            logger=logging.getLogger("proper-name-fences"),
        )
        self.ai.json_structured_document = {
            "type": "docx", "elements": [
                {"element_id": "p1", "type": "paragraph", "text":
                 "Lägst andel hade Akademikernas och högst hade Sveriges Arbetares med 3,7 procent."},
                {"element_id": "p2", "type": "paragraph", "text":
                 "Den nya rapporten kommer från Högsta Förvaltningsdomstolen och granskas nu."},
                {"element_id": "p3", "type": "paragraph", "text":
                 "Arbetsförmedlingen har genomfört en granskning av olika ärenden."},
                {"element_id": "p4", "type": "paragraph", "text":
                 "Det innebär att rapporten blir begripligare."},
            ]
        }

    def check(self, ident, old, new, blocked):
        suggestion = SuggestedChange("paragraph", ident, None, old, new, None, "exact")
        reason = self.ai._assess_suggestion_safety(suggestion)
        self.assertEqual(reason["reason"] == "protected_proper_name_changed", blocked)
        return suggestion

    def test_lowercase_name_inside_sentence_rejected(self):
        self.check("p1", "Arbetares", "arbetares", True)

    def test_name_changed_within_large_rewrite_rejected(self):
        self.check("p1", "högst hade Sveriges Arbetares med 3,7 procent",
                   "störst andel hade Sveriges arbetares med 3,7 procent", True)

    def test_name_not_erased_by_rewrite(self):
        self.check("p2", "Högsta Förvaltningsdomstolen", "domstolen", True)

    def test_normal_prose_rewrite_still_allowed(self):
        self.check("p3", "har genomfört", "gjorde", False)

    def test_single_capitalized_authority_not_blanket_locked(self):
        self.check("p3", "Arbetsförmedlingen", "Myndigheten", False)

    def test_original_name_preserved_when_other_word_changed(self):
        self.check("p1", "hade Sveriges Arbetares med", "uppmättes hos Sveriges Arbetares med", False)

    def test_regular_sentence_initial_words_not_treated_as_two_part_name(self):
        self.check("p4", "Det innebär", "Det betyder", False)

    def test_second_occurrence_is_also_protected(self):
        self.ai.json_structured_document["elements"].append({
            "element_id": "p5", "type": "paragraph",
            "text": "Sveriges Arbetares och Sveriges Arbetares nämns i samma stycke."
        })
        self.check("p5", "Sveriges Arbetares nämns", "Sveriges arbetares nämns", True)

    def test_pipeline_filters_name_edit_after_minimization(self):
        result = self.ai._postprocess_suggestions([
            {"type": "paragraph", "element_id": "p1", "old": "Sveriges Arbetares",
             "new": "Sveriges arbetares"},
            {"type": "paragraph", "element_id": "p3", "old": "genomfört",
             "new": "gjort"},
        ])
        self.assertEqual([(x.element_id, x.old) for x in result], [("p3", "genomfört")])


if __name__ == "__main__":
    unittest.main()
