import json
import tempfile
import unittest
from pathlib import Path

from app.src.JBGReadabilityMetrics import (
    apply_suggestions_to_text,
    compute_document_readability,
    compute_document_readability_from_files,
    is_caption,
    is_section_heading,
    lix_band,
    rounded_lix_delta,
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


class RegressionFromRealReportTests(unittest.TestCase):
    """Fall från testkörningen 93dbc428 (Rapportutkast för test av klarspråksgranskning)."""

    def test_layout_line_breaks_inside_a_sentence_are_not_boundaries(self):
        # table_13_cell_1_1_p8: en mening med två manuella radbrytningar
        text = (
            "IAF riktar allvarlig kritik mot den aktör vi har granskat när \n"
            "bristen är av större omfattning eller avser allvarligare avsteg från gällande \n"
            "regelverk, eller av sådan art att den riskerar att skada "
            "arbetslöshetsförsäkringens legitimitet. "
        )
        self.assertEqual(text_stats(text).sentences, 1)

    def test_removing_layout_breaks_does_not_change_sentence_count(self):
        before = "IAF påpekar en brist, när bristen inte har fått några eller endast \nsmå konsekvenser."
        after, applied, _, _ = apply_suggestions_to_text(before, [{
            "old": ", när bristen inte har fått några eller endast \n",
            "new": " när bristen inte har fått några eller bara ",
        }])
        self.assertEqual(applied, 1)
        self.assertEqual(text_stats(before).sentences, 1)
        self.assertEqual(text_stats(after).sentences, 1)

    def test_line_break_before_new_line_item_is_a_boundary(self):
        self.assertEqual(text_stats("Telefon 010-123 45 67\nE-post info@iaf.se").sentences, 2)
        self.assertEqual(text_stats("Första punkten\n• Andra punkten").sentences, 2)

    def test_line_break_after_comma_is_not_a_boundary(self):
        self.assertEqual(text_stats("Granskningen omfattar Arbetsförmedlingen,\nIAF och a-kassorna.").sentences, 1)

    def test_rounded_delta_matches_displayed_values(self):
        # Orundat 51.549 -> 51.591 gav tidigare 0.0 trots visningen 51.5 -> 51.6
        self.assertEqual(rounded_lix_delta(51.549, 51.591), 0.1)
        self.assertIsNone(rounded_lix_delta(None, 50.0))

    def test_captions_are_not_section_headings(self):
        caption = {"type": "paragraph", "heading_level": 1, "style_id": "IAFTabellrubrik",
                   "style_name": "IAF Tabellrubrik",
                   "text": "Tabell 1: Antal återkallanden totalt och antal återkallanden per 100 programdeltagare 2025."}
        self.assertTrue(is_caption(caption))
        self.assertFalse(is_section_heading(caption))
        self.assertTrue(is_caption({"type": "paragraph", "style_name": "Normal", "text": "Figur 3 Andel ärenden"}))
        self.assertTrue(is_caption({"type": "paragraph", "style_id": "Caption", "text": "Källa: IAF"}))
        self.assertFalse(is_caption({"type": "paragraph", "style_name": "Heading 1", "text": "Tabeller och figurer"}))

    def test_sections_ignore_table_headings_captions_and_empty_chapters(self):
        def el(eid, etype, text, level=None, style="Normal", order=None):
            return {"type": etype, "element_id": eid, "text": text, "heading_level": level,
                    "style_id": style, "style_name": style, "doc_order": order}

        structure = {"type": "docx", "elements": [
            el("paragraph_1", "paragraph", "Sammanfattning", 1, "IAFSammanfattning", 1),
            el("paragraph_2", "paragraph", "Myndigheten brister.", None, "Normal", 2),
            # Faktaruta byggd som tabell med rubrikformaterad första rad
            el("table_1_cell_1_1_p1", "table_cell", "IAF riktar en anmärkning till Arbetsförmedlingen för att",
               1, "IAFRubriktilltextruta2", 3),
            el("table_1_cell_1_1_p2", "table_cell", "Beslut saknar rättslig grund.", None, "Normal", 4),
            el("paragraph_3", "paragraph", "Utvecklingen", 1, "IAFRubrik1numrerad-Kapitel", 5),
            el("paragraph_4", "paragraph", "Antalet ökar.", None, "Normal", 6),
            el("paragraph_5", "paragraph", "Tabell 1: Antal återkallanden", 1, "IAFTabellrubrik", 7),
            el("table_2_cell_1_1_p1", "table_cell", "Totalt 1 200 ärenden.", None, "Normal", 8),
            el("paragraph_6", "paragraph", "Bilaga 2:", 1, "IAFBilagerubrik1", 9),
            # Tom rubrik med outline level ska inte öppna ett avsnitt
            el("paragraph_7", "paragraph", "\n", 1, "IAFRubrik1numrerad-Kapitel", 10),
        ]}
        sections = compute_document_readability(structure, []).to_dict()["sections"]

        self.assertEqual([s["heading"] for s in sections], ["Sammanfattning", "Utvecklingen"])
        # Faktarutan räknas in i Sammanfattning, tabellen efter bildtexten i Utvecklingen
        self.assertEqual(sections[0]["before"]["words"], 2 + 4)
        self.assertEqual(sections[1]["before"]["words"], 2 + 4)


if __name__ == "__main__":
    unittest.main()
