"""Opt-in, read-only analysis of lossless proposal splitting.

No SuggestedChange, ChangePlan or OOXML data is modified. This module can be
removed along with its single orchestrator hook without changing rendering.
"""
from dataclasses import dataclass
from difflib import SequenceMatcher


@dataclass(frozen=True)
class SplitPiece:
    old: str
    new: str
    old_start: int
    old_end: int


class SuggestionSplitShadow:
    """Inspect one original proposal at a time; never redefine the denominator."""

    def __init__(self, *, min_anchor_chars=12, max_pieces=8):
        self.min_anchor_chars = min_anchor_chars
        self.max_pieces = max_pieces

    @staticmethod
    def _unique(text, substring):
        return bool(substring) and text.count(substring) == 1

    def analyze(self, suggestions, structure):
        elements = {e.get('element_id'): e for e in structure.get('elements', [])}
        report = {'mode': 'shadow_only', 'original_proposals': len(suggestions),
                  'split_candidates': 0, 'unchanged_proposals': 0,
                  'reasons': {}, 'proposals': []}
        for i, suggestion in enumerate(suggestions):
            element = elements.get(suggestion.element_id)
            source = (element or {}).get('text') or ''
            pieces, reason = self._split(suggestion.old, suggestion.new, source)
            row = {'original_index': i, 'element_id': suggestion.element_id,
                   'status': 'candidate' if pieces else 'unchanged',
                   'reason': reason, 'piece_count': len(pieces) if pieces else 1}
            if pieces:
                # Deliberately retain one parent motivation, not duplicated texts.
                row['pieces'] = [{'old': p.old, 'new': p.new,
                                  'old_start': p.old_start, 'old_end': p.old_end}
                                 for p in pieces]
                report['split_candidates'] += 1
            else:
                report['unchanged_proposals'] += 1
            report['reasons'][reason] = report['reasons'].get(reason, 0) + 1
            report['proposals'].append(row)
        return report

    def _split(self, old, new, source):
        if not old or not new or old == new:
            return None, 'empty_or_unchanged'
        # Explicit exact and unique parent anchor. We do not guess positions.
        if not self._unique(source, old):
            return None, 'parent_anchor_not_unique'
        matcher = SequenceMatcher(a=old, b=new, autojunk=False)
        opcodes = matcher.get_opcodes()
        boundaries = []
        for tag, a, a_end, b, b_end in opcodes:
            if tag != 'equal' or a_end - a < 2 * self.min_anchor_chars:
                continue
            # A long unchanged window provides an unambiguous split point.
            anchor = old[a:a_end]
            if not self._unique(old, anchor) or not self._unique(new, anchor):
                continue
            # Only split BETWEEN two edits. End-of-proposal equal tails
            # must remain part of the previous piece.
            if not any(t != 'equal' and end <= a for t, start, end, _, _ in opcodes):
                continue
            if not any(t != 'equal' and start >= a_end for t, start, end, _, _ in opcodes):
                continue
            mid = (a + a_end) // 2
            # A word boundary is preferable to cutting within a word.
            candidate_cuts = [k for k in range(a + self.min_anchor_chars,
                                              a_end - self.min_anchor_chars + 1)
                              if old[k - 1].isspace() or old[k].isspace()]
            if not candidate_cuts:
                continue
            cut = min(candidate_cuts, key=lambda k: abs(k - mid))
            midpoint = cut - a
            left = old[a:cut]
            right = old[cut:a_end]
            if (len(left) < self.min_anchor_chars or len(right) < self.min_anchor_chars
                    or not self._unique(old, left) or not self._unique(new, left)
                    or not self._unique(old, right) or not self._unique(new, right)):
                continue
            boundaries.append((a + midpoint, b + midpoint))
        if not boundaries:
            return None, 'no_unique_stable_boundary'
        if len(boundaries) + 1 > self.max_pieces:
            return None, 'too_many_pieces'
        cut_positions = [(0, 0), *boundaries, (len(old), len(new))]
        pieces = []
        for (a1, b1), (a2, b2) in zip(cut_positions, cut_positions[1:]):
            left, right = old[a1:a2], new[b1:b2]
            if not left or not right or left == right:
                return None, 'empty_or_unchanged_piece'
            # Each proposed piece must be independently and uniquely locatable.
            if not self._unique(source, left):
                return None, 'piece_anchor_not_unique'
            pieces.append(SplitPiece(left, right, a1, a2))
        if len(pieces) < 2:
            return None, 'not_multiple_edits'
        if ''.join(p.old for p in pieces) != old or ''.join(p.new for p in pieces) != new:
            return None, 'roundtrip_failed'
        return tuple(pieces), 'lossless_unique_boundaries'
