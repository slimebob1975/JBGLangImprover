import json
import tempfile
import unittest
from pathlib import Path

from app.src.JBGReadabilityMetrics import (
    apply_suggestions_to_text,
    compute_document_readability,
    compute_document_readability_from_files,
    lix_band,
    text_stats,
)


class TextStatsTests(unittest.TestCase):
    def assertStats(self, text, words, sentences, long_words, lix):
        stats = text_stats(text)
        self.assertEqual(
            (stats.words, stats.sentences, stats.long_words),
            (words, sentences, long_words),
            text,
        )
        self.assertAlmostEqual(stats.lix, lix, places=2)

    def test_simple_two_sentence_text(self):
        # 9 ord / 2 meningar + 100 * 2 långa (skickar, beslutet) / 9
        self.assertStats("Vi skickar beslutet i dag. Du får svar snart.", 9, 2, 2, 4.5 + 200 / 9)

    def test_abbreviation_is_one_word_and_no_sentence_break(self):
        # Det, gäller, t.ex., arbetslöshetsförsäkringen: 4 ord, 1 mening, 1 långt
        self.assertStats("Det gäller t.ex. arbetslöshetsförsäkringen.", 4, 1, 1, 29.0)

    def test_decimal_number_does_not_split_sentence(self):
        # 10 ord / 2 meningar + 100 * 2 (Kostnaden, miljoner) / 10
        self.assertStats(
            "Kostnaden var 3.5 miljoner kronor. Det är mer än väntat.", 10, 2, 2, 25.0
        )

    def test_element_without_final_punctuation_counts_as_sentence(self):
        self.assertStats("Sammanfattning av rapporten", 3, 1, 2, 3 + 200 / 3)

    def test_line_breaks_are_sentence_boundaries(self):
        stats = text_stats("Första raden\nAndra raden")
        self.assertEqual((stats.words, stats.sentences), (4, 2))

    def test_six_letter_word_is_not_long_but_seven_is(self):
        self.assertEqual(text_stats("kronor").long_words, 0)
        self.assertEqual(text_stats("skickar").long_words, 1)
        self.assertEqual(text_stats("a-kassor").long_words, 1)  # 7 bokstäver

    def test_empty_text(self):
        stats = text_stats("   ")
        self.assertEqual((stats.words, stats.sentences), (0, 0))
        self.assertIsNone(stats.lix)

    def test_bands(self):
        self.assertEqual(lix_band(25), "Mycket lättläst")
        self.assertEqual(lix_band(35), "Lättläst")
        self.assertEqual(lix_band(45), "Medelsvår")
        self.assertEqual(lix_band(55), "Svår")
        self.assertEqual(lix_band(65), "Mycket svår")
        self.assertIsNone(lix_band(None))


class ApplySuggestionsTests(unittest.TestCase):
    def test_exact_and_whitespace_normalized_matches(self):
        text = "Vi har genomfört en  granskning av ärendet."
        result, applied, unanchored, overlapping = apply_suggestions_to_text(text, [
            {"old": "genomfört en granskning av", "new": "granskat"},
            {"old": "ärendet", "new": "fallet"},
        ])
        self.assertEqual(result, "Vi har granskat fallet.")
        self.assertEqual((applied, unanchored, overlapping), (2, 0, 0))

    def test_unanchored_and_overlapping_are_skipped(self):
        text = "Myndigheten fattar beslut."
        result, applied, unanchored, overlapping = apply_suggestions_to_text(text, [
            {"old": "fattar beslut", "new": "beslutar"},
            {"old": "beslut", "new": "avgörande"},
            {"old": "finns inte", "new": "saknas"},
        ])
        self.assertEqual(result, "Myndigheten beslutar.")
        self.assertEqual((applied, unanchored, overlapping), (1, 1, 1))


class DocumentReadabilityTests(unittest.TestCase):
    def setUp(self):
        self.structure = {
            "type": "docx",
            "elements": [
                {"type": "paragraph", "element_id": "paragraph_1", "text": "Inledning",
                 "heading_level": 1, "doc_order": 1},
                {"type": "paragraph", "element_id": "paragraph_2",
                 "text": "Vi har genomfört en granskning av ärendet.",
                 "heading_level": None, "doc_order": 2},
                {"type": "paragraph", "element_id": "paragraph_3", "text": "Resultat",
                 "heading_level": 1, "doc_order": 3},
                {"type": "table_cell", "element_id": "table_1_cell_1_1_p1",
                 "text": "Kostnaden var hög.", "heading_level": None, "doc_order": 4},
                {"type": "header", "element_id": "header_1", "text": "Rapport 2026:1",
                 "heading_level": None, "doc_order": None},
                {"type": "footnote", "element_id": "footnote_1", "footnote_id": "2",
                 "text": "Källa saknas.", "heading_level": None, "doc_order": None},
            ],
        }
        self.suggestions = [
            {"type": "paragraph", "element_id": "paragraph_2",
             "old": "genomfört en granskning av", "new": "granskat"},
            {"type": "footnote", "element_id": "footnote_1", "footnote_id": "2",
             "old": "Källa saknas", "new": "Uppgiften saknar källa"},
        ]

    def test_before_after_excludes_headings_headers_and_footnotes(self):
        report = compute_document_readability(self.structure, self.suggestions)
        data = report.to_dict()

        # Före: 7 + 3 ord, 2 meningar, långa: genomfört, granskning, ärendet, Kostnaden
        self.assertEqual(data["before"]["words"], 10)
        self.assertEqual(data["before"]["sentences"], 2)
        self.assertEqual(data["before"]["long_words"], 4)
        # Efter: "Vi har granskat ärendet." -> 4 + 3 ord, långa: granskat, ärendet, Kostnaden
        self.assertEqual(data["after"]["words"], 7)
        self.assertEqual(data["after"]["long_words"], 3)
        self.assertEqual(data["suggestions"], {
            "total": 2, "in_scope": 1, "applied_in_memory": 1,
            "unanchored": 0, "overlapping": 0,
        })
        self.assertEqual(report.elements_counted, 2)
        # 10/2 + 100*4/10 = 45.0 -> 7/2 + 100*3/7 = 46.4. En bättre formulering
        # kan alltså ge högre LIX när korta ord försvinner; värdet är en
        # indikator, inte ett kvalitetsmått.
        self.assertEqual(data["before"]["lix"], 45.0)
        self.assertEqual(data["after"]["lix"], 46.4)
        self.assertEqual(data["lix_delta"], 1.4)

    def test_sections_follow_top_level_headings(self):
        report = compute_document_readability(self.structure, self.suggestions)
        sections = report.to_dict()["sections"]
        self.assertEqual([s["heading"] for s in sections], ["Inledning", "Resultat"])
        self.assertEqual(sections[0]["before"]["words"], 7)
        self.assertEqual(sections[0]["after"]["words"], 4)
        self.assertEqual(sections[1]["before"]["words"], 3)

    def test_legacy_structure_without_metadata_still_works(self):
        legacy = {"type": "docx", "elements": [
            {"type": "paragraph", "element_id": "paragraph_1", "text": "Kort text."},
        ]}
        report = compute_document_readability(legacy, [])
        self.assertEqual(report.before.words, 2)
        self.assertEqual(report.sections, [])

    def test_from_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            structure_path = Path(tmp) / "doc_structure.json"
            suggestions_path = Path(tmp) / "doc_suggestions.json"
            structure_path.write_text(json.dumps(self.structure, ensure_ascii=False), encoding="utf-8")
            suggestions_path.write_text(json.dumps(self.suggestions, ensure_ascii=False), encoding="utf-8")

            from_files = compute_document_readability_from_files(
                str(structure_path), str(suggestions_path)
            ).to_dict()
        in_memory = compute_document_readability(self.structure, self.suggestions).to_dict()
        self.assertEqual(from_files, in_memory)


if __name__ == "__main__":
    unittest.main()
