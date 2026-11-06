"""Regression tests for optional word-level diagnostics in the about section."""
import unittest
from app.src.JBGAboutSectionRenderer import build_about_section_blocks


class WordLevelDiagnosticsTests(unittest.TestCase):
    def test_diagnostics_shown_when_enabled(self):
        blocks = build_about_section_blocks({
            'docx_mode': 'tracked',
            'word_level_diagnostics': {
                'enabled': True, 'word_level': 3, 'span_fallback': 2,
                'fallback_reasons': {'mixed_run_formatting': 2},
            },
        }, tracked=True)
        self.assertTrue(any('Ordnivåändringar med säker spann-fallback' in b.text for b in blocks))
        rows = [row for block in blocks for row in block.rows]
        self.assertIn(['Ordnivå', '3'], rows)
        self.assertIn(['Fallback till spann', '2'], rows)
        self.assertIn(['Fallback: Blandad formatering (mixed_run_formatting)', '2'], rows)

    def test_specific_diff_fallback_labels(self):
        blocks = build_about_section_blocks({
            'docx_mode': 'tracked',
            'word_level_diagnostics': {
                'enabled': True, 'word_level': 0, 'span_fallback': 2,
                'fallback_reasons': {'diff_no_common_text': 1, 'exact_diff_empty': 1},
            },
        }, tracked=True)
        rows = [row for block in blocks for row in block.rows]
        self.assertIn(['Fallback: Inga oförändrade textdelar (diff_no_common_text)', '1'], rows)
        self.assertIn(['Fallback: Texttrogen diff kunde inte byggas (exact_diff_empty)', '1'], rows)

    def test_equivalent_span_is_not_relabelled_word_level(self):
        blocks = build_about_section_blocks({
            'docx_mode': 'tracked',
            'word_level_diagnostics': {
                'enabled': True, 'word_level': 0, 'span_fallback': 1,
                'fallback_reasons': {'diff_no_common_text_equivalent_span': 1},
            },
        }, tracked=True)
        rows = [row for block in blocks for row in block.rows]
        self.assertIn(['Ordnivå', '0'], rows)
        self.assertIn(['Fallback till spann', '1'], rows)
        self.assertIn(['Fallback: Enskilt ordbyte – redan lika precist i spannläge (diff_no_common_text_equivalent_span)', '1'], rows)

    def test_complex_run_reasons_get_readable_labels(self):
        blocks = build_about_section_blocks({
            'docx_mode': 'tracked',
            'word_level_diagnostics': {
                'enabled': True, 'word_level': 0, 'span_fallback': 2,
                'fallback_reasons': {
                    'complex_run_tabs_or_breaks': 1,
                    'complex_run_fields_or_objects': 1,
                },
            },
        }, tracked=True)
        rows = [row for block in blocks for row in block.rows]
        self.assertIn(['Fallback: Tabbar eller radbrytningar i textruns (complex_run_tabs_or_breaks)', '1'], rows)
        self.assertIn(['Fallback: Fält, referenser eller objekt i textruns (complex_run_fields_or_objects)', '1'], rows)

    def test_no_diagnostics_in_span_mode(self):
        blocks = build_about_section_blocks({'docx_mode': 'tracked'}, tracked=True)
        self.assertFalse(any('Ordnivåändringar med säker spann-fallback' in b.text for b in blocks))
