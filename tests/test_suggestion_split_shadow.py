"""Frozen, no-network checks of the removable shadow middleware."""
import logging
import unittest
from unittest.mock import patch

from app.src.JBGSuggestionSplitShadow import SuggestionSplitShadow
from app.src.JBGLangImprovSuggestorAI import SuggestedChange
from app.src.JBGLanguageImprover import JBGLanguageImprover


OLD = 'Vi granskar ansökningarna noggrant och skriver sedan ett tydligt beslut.'
NEW = 'Vi kontrollerar ansökningarna noggrant och formulerar sedan ett tydligt beslut.'


def suggestion(old=OLD, new=NEW, element_id='p1'):
    return SuggestedChange('paragraph', element_id, None, old, new,
                           'Gemensam motivering.', 'exact')


class SuggestionSplitShadowTests(unittest.TestCase):
    def setUp(self):
        self.shadow = SuggestionSplitShadow()
        self.structure = {'elements': [{'element_id': 'p1', 'text': OLD}]}

    def test_lossless_two_edits_and_original_motivation_intact(self):
        item = suggestion()
        report = self.shadow.analyze([item], self.structure)
        self.assertEqual(report['original_proposals'], 1)
        self.assertEqual(report['split_candidates'], 1)
        pieces = report['proposals'][0]['pieces']
        self.assertEqual(len(pieces), 2)
        self.assertEqual(''.join(p['old'] for p in pieces), OLD)
        self.assertEqual(''.join(p['new'] for p in pieces), NEW)
        self.assertEqual(item.motivation, 'Gemensam motivering.')
        self.assertEqual(item.old, OLD)
        self.assertEqual(item.new, NEW)

    def test_ambiguous_source_and_short_anchor_fail_closed(self):
        duplicated = {'elements': [{'element_id': 'p1', 'text': OLD + ' ' + OLD}]}
        self.assertEqual(self.shadow.analyze([suggestion()], duplicated)['split_candidates'], 0)
        self.assertEqual(self.shadow.analyze([suggestion('Hej', 'Adjö')],
                                             {'elements': [{'element_id': 'p1', 'text': 'Hej'}]})['split_candidates'], 0)

    def test_count_is_original_suggestions_not_pieces(self):
        report = self.shadow.analyze([suggestion(), suggestion('Hej', 'Adjö')], self.structure)
        self.assertEqual(report['original_proposals'], 2)
        self.assertEqual(report['split_candidates'] + report['unchanged_proposals'], 2)

    def test_shadow_enabled_keeps_original_change_plan(self):
        improver = JBGLanguageImprover('test.docx', '', 'unused', '', 0, True,
                                        logging.getLogger('shadow-test-on'),
                                        save_intermediate_json=False)
        improver.structure = self.structure
        original = suggestion()
        improver.validated_suggestions = [original]
        with patch.dict('os.environ', {'JBG_SUGGESTION_SPLIT_SHADOW': '1'}):
            plans = improver._build_change_plans()
        self.assertEqual(len(plans), 1)
        self.assertEqual(plans[0].old_text, OLD)
        self.assertEqual(plans[0].new_text, NEW)
        self.assertEqual(plans[0].motivation, original.motivation)

    def test_shadow_disabled_never_imports_or_runs(self):
        # An integration test that enforces the off switch in the orchestrator.
        improver = JBGLanguageImprover('test.docx', '', 'unused', '', 0, True,
                                        logging.getLogger('shadow-test'),
                                        save_intermediate_json=False)
        improver.structure = self.structure
        improver.validated_suggestions = [suggestion()]
        with patch.dict('os.environ', {'JBG_SUGGESTION_SPLIT_SHADOW': '0'}):
            plans = improver._build_change_plans()
        self.assertEqual(len(plans), 1)
        self.assertEqual(plans[0].old_text, OLD)
        self.assertEqual(plans[0].new_text, NEW)


if __name__ == '__main__':
    unittest.main()
