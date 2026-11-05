"""No-network tests of removable, isolated split renderability dry-run."""
import logging
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from docx import Document

from app.src.JBGDocumentStructureExtractor import DocumentStructureExtractor
from app.src.JBGLangImprovSuggestorAI import SuggestedChange
from app.src.JBGLanguageImprover import JBGLanguageImprover
from app.src.JBGSuggestionSplitDryRun import SuggestionSplitDryRun


OLD = 'Vi granskar ansökningarna noggrant och skriver sedan ett tydligt beslut.'
NEW = 'Vi kontrollerar ansökningarna noggrant och formulerar sedan ett tydligt beslut.'


class SplitDryRunTests(unittest.TestCase):
    def setUp(self):
        self.logger = logging.getLogger('split-dry-run')

    def _fixture(self, root):
        path = Path(root) / 'original.docx'
        doc = Document()
        doc.add_paragraph(OLD)
        doc.save(path)
        structure = DocumentStructureExtractor(str(path), self.logger).extract()
        element = next(e for e in structure['elements'] if e.get('text') == OLD)
        original = SuggestedChange('paragraph', element['element_id'], None,
                                   OLD, NEW, 'One shared motivation.', 'exact')
        return path, structure, original

    def test_real_renderer_probe_isolated_and_parent_count(self):
        with tempfile.TemporaryDirectory() as temp:
            path, structure, original = self._fixture(temp)
            before = path.read_bytes()
            report = SuggestionSplitDryRun().analyze(
                [original], structure, str(path), self.logger,
            )
            self.assertEqual(report['mode'], 'isolated_renderability_dry_run')
            self.assertEqual(report['original_proposals'], 1)
            self.assertEqual(report['probed_proposals'], 1)
            self.assertEqual(len(report['proposals']), 1)
            self.assertEqual(sum(report['candidate_outcomes'].values()), 1)
            self.assertEqual(original.motivation, 'One shared motivation.')
            self.assertEqual(before, path.read_bytes())
            self.assertEqual(sorted(p.name for p in Path(temp).iterdir()), ['original.docx'])

    def test_split_status_requires_every_piece_word_level(self):
        runner = SuggestionSplitDryRun()
        from types import SimpleNamespace
        good = SimpleNamespace(applied=True, word_level_status='word_level',
                               word_level_fallback_reason=None, message='ok')
        fallback = SimpleNamespace(applied=True, word_level_status='span_fallback',
                                   word_level_fallback_reason='mixed_run_formatting', message='ok')
        self.assertEqual(runner._outcome([good, good], 2), 'word_level')
        self.assertEqual(runner._outcome([good, fallback], 2), 'span_fallback')
        self.assertEqual(runner._outcome([good], 2), 'unrenderable')
        self.assertEqual(runner._result_reasons([good, fallback]), ['mixed_run_formatting'])

    def test_no_candidate_means_no_docx_probe(self):
        runner = SuggestionSplitDryRun()
        original = SuggestedChange('paragraph', 'p1', None, 'Hej', 'Hallå', 'Motivering', 'exact')
        structure = {'elements': [{'element_id': 'p1', 'text': 'Hej'}]}
        with patch.object(runner, '_probe', side_effect=AssertionError('must not probe')):
            report = runner.analyze([original], structure, 'missing.docx', self.logger)
        self.assertEqual(report['probed_proposals'], 0)
        self.assertEqual(report['original_proposals'], 1)

    def test_mode_disabled_leaves_existing_plans_untouched(self):
        improver = JBGLanguageImprover('missing.docx', '', 'unused', '', 0, True,
                                        self.logger, save_intermediate_json=False,
                                        docx_mode='tracked', word_level_changes=True)
        improver.structure = {'elements': [{'element_id': 'p1', 'text': OLD}]}
        original = SuggestedChange('paragraph', 'p1', None, OLD, NEW, 'Motivering', 'exact')
        improver.validated_suggestions = [original]
        with patch.dict('os.environ', {'JBG_SUGGESTION_SPLIT_DRY_RUN': '0',
                                       'JBG_SUGGESTION_SPLIT_SHADOW': '0'}), \
             patch('app.src.JBGSuggestionSplitDryRun.SuggestionSplitDryRun.analyze',
                   side_effect=AssertionError('must not run')):
            plans = improver._build_change_plans()
        self.assertEqual(len(plans), 1)
        self.assertEqual(plans[0].new_text, NEW)

    def test_enabled_mode_passes_original_parent_not_parts(self):
        improver = JBGLanguageImprover('missing.docx', '', 'unused', '', 0, True,
                                        self.logger, save_intermediate_json=False,
                                        docx_mode='tracked', word_level_changes=True)
        improver.structure = {'elements': [{'element_id': 'p1', 'text': OLD}]}
        original = SuggestedChange('paragraph', 'p1', None, OLD, NEW, 'Motivering', 'exact')
        improver.validated_suggestions = [original]
        with patch.dict('os.environ', {'JBG_SUGGESTION_SPLIT_DRY_RUN': '1',
                                       'JBG_SUGGESTION_SPLIT_SHADOW': '0'}), \
             patch('app.src.JBGSuggestionSplitDryRun.SuggestionSplitDryRun.analyze',
                   return_value={'probed_proposals': 1, 'candidate_outcomes':
                                 {'promoted_to_word_level': 0}}):
            plans = improver._build_change_plans()
        self.assertEqual(len(plans), 1)
        self.assertEqual(plans[0].old_text, OLD)
        self.assertEqual(plans[0].new_text, NEW)
        self.assertEqual(plans[0].motivation, original.motivation)


if __name__ == '__main__':
    unittest.main()
