"""Opt-in, isolated Word renderability probe for suggestion split candidates.

Every probe opens its own temporary DOCX package and NEVER saves it.
This is diagnostic only: it does not replace full-document OOXML verification,
comment rendering or conflict analysis. A parent is counted only once.
"""
from collections import Counter
from dataclasses import replace

from app.src.JBGChangePlanner import ChangePlanner
from app.src.JBGDocxPackage import DocxPackage
from app.src.JBGSuggestionSplitShadow import SuggestionSplitShadow
from app.src.JBGTokenDiffEngine import TokenDiffEngine
from app.src.JBGTrackedChangesRenderer import TrackedChangesRenderer


class SuggestionSplitDryRun:
    """Compare original and split proposals using identical renderer settings."""

    def __init__(self, *, min_anchor_chars=12, max_pieces=8):
        self.splitter = SuggestionSplitShadow(
            min_anchor_chars=min_anchor_chars, max_pieces=max_pieces,
        )

    @staticmethod
    def _outcome(results, expected):
        if len(results) != expected or not all(r.applied for r in results):
            return 'unrenderable'
        if all(r.word_level_status == 'word_level' for r in results):
            return 'word_level'
        return 'span_fallback'

    @staticmethod
    def _result_reasons(results):
        return sorted({
            r.word_level_fallback_reason or r.message
            for r in results
            if not r.applied or r.word_level_status == 'span_fallback'
        })

    def _probe(self, docx_path, structure, suggestions, logger):
        """Real planner and OOXML renderer, isolated from the output package."""
        engine = TokenDiffEngine(logger=logger)
        planner = ChangePlanner(structure=structure, diff_engine=engine, logger=logger)
        plans = planner.build_plans(suggestions)
        if len(plans) != len(suggestions):
            return {'status': 'unrenderable', 'reasons': ['planning_failed']}
        if any('overlapping_change_conflict' in p.notes for p in plans):
            return {'status': 'unrenderable', 'reasons': ['overlapping_change_conflict']}
        with DocxPackage(docx_path, logger) as pkg:
            renderer = TrackedChangesRenderer(pkg, logger, word_level_changes=True)
            results = renderer.apply_plans(plans)
            return {
                'status': self._outcome(results, len(suggestions)),
                'reasons': self._result_reasons(results),
            }

    def analyze(self, suggestions, structure, docx_path, logger):
        shadow = self.splitter.analyze(suggestions, structure)
        results = {
            'mode': 'isolated_renderability_dry_run',
            'original_proposals': len(suggestions),
            'split_candidates': shadow['split_candidates'],
            'candidate_outcomes': {},
            'proposals': [],
            'limitations': (
                'Probes are isolated per original proposal. They do not test '
                'inter-proposal conflicts, comment grouping, accepted/rejected '
                'roundtrips or final document integrity. Nothing is saved or applied.'
            ),
        }
        counts = Counter()
        # Scope tightly: only candidates are probed. Unchanged proposals are
        # explicitly not re-counted as improvements or failures.
        for row in shadow['proposals']:
            if row['status'] != 'candidate':
                continue
            index = row['original_index']
            original = suggestions[index]
            # These are *independent* probes with the same parent metadata,
            # not new user-visible suggestions. Motivation is never split.
            children = [replace(original, old=p['old'], new=p['new'],
                                match_status='exact') for p in row['pieces']]
            try:
                parent = self._probe(docx_path, structure, [original], logger)
                split = self._probe(docx_path, structure, children, logger)
            except Exception as exc:
                # Fail this candidate closed, not the entire diagnostic report.
                logger.warning('Isolated split probe failed for %s: %s',
                               original.element_id, exc)
                parent = {'status': 'unrenderable', 'reasons': ['probe_exception']}
                split = {'status': 'unrenderable', 'reasons': ['probe_exception']}
            if parent['status'] == 'word_level':
                category = 'already_word_level'
            elif split['status'] == 'word_level':
                category = 'promoted_to_word_level'
            elif split['status'] == 'unrenderable':
                category = 'unsafe_or_unrenderable'
            else:
                category = 'still_span_fallback'
            counts[category] += 1
            results['proposals'].append({
                'original_index': index,
                'element_id': original.element_id,
                'piece_count': len(children),
                'baseline': parent,
                'split': split,
                'outcome': category,
            })
        results['candidate_outcomes'] = {
            key: counts[key] for key in (
                'already_word_level', 'promoted_to_word_level',
                'still_span_fallback', 'unsafe_or_unrenderable'
            )
        }
        results['probed_proposals'] = len(results['proposals'])
        return results
