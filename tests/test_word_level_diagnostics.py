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
        self.assertTrue(any('Experimentell ordnivårendering' in b.text for b in blocks))
        rows = [row for block in blocks for row in block.rows]
        self.assertIn(['Ordnivå', '3'], rows)
        self.assertIn(['Fallback till spann', '2'], rows)
        self.assertIn(['Fallback: mixed_run_formatting', '2'], rows)

    def test_no_diagnostics_in_span_mode(self):
        blocks = build_about_section_blocks({'docx_mode': 'tracked'}, tracked=True)
        self.assertFalse(any('Experimentell ordnivårendering' in b.text for b in blocks))
