"""Small, explicit recovery rules, not a general grammatical splitter.

Only recover a correction already proposed by the model, with unchanged local
context on both sides. All resulting suggestions must be validated again.
"""
import re
from dataclasses import dataclass


@dataclass(frozen=True)
class CorrectionRule:
    pattern: re.Pattern
    replacement: str
    motivation: str


# Deliberately limited to demonstrably independent repairs. The surrounding
# words are retained in the candidate for unique and safe anchoring.
_RULES = (
    CorrectionRule(re.compile(r"\b([\w-]+) är har (\w+)\b"),
                   r"\1 har \2", "Tar bort det felaktiga extra verbet 'är' i 'är har'."),
    CorrectionRule(re.compile(r"\b((?:[Rr]ollfördelningen|[Aa]nsvarsfördelningen)(?: inom (?:sådana samarbeten|samarbetet))?) är tydliga\b"),
                   r"\1 är tydlig", "Rättar kongruensen: subjektet står i singular och kräver 'tydlig'."),
)


_SENTENCE_BOUNDARY = re.compile(r"(?<=[.!?])\s+(?=[A-ZÅÄÖ])")


def _occurrences(text, phrase):
    return list(re.finditer(r"(?<![\w-])" + re.escape(phrase) + r"(?![\w-])", text))


def independent_corrections(old, new, source):
    """Return one exact, model-supported repair from a composite proposal.

    No fuzzy alignment, no new linguistic guesses, no splitting of arbitrary
    sentences. Require a unique parent anchor and a unique local phrase in
    corresponding sentence positions. A standalone repair is left unchanged.
    """
    if not old or not new or source.count(old) != 1:
        return []
    # Reconstruct the full source: ordinary span minimization may have
    # removed the unchanged words adjacent to the actual correction.
    start = source.index(old)
    revised = source[:start] + new + source[start + len(old):]
    old_sentences = _SENTENCE_BOUNDARY.split(source.strip())
    new_sentences = _SENTENCE_BOUNDARY.split(revised.strip())
    if len(old_sentences) != len(new_sentences) or len(old_sentences) < 2:
        return []
    candidates = []
    for index, (old_sentence, new_sentence) in enumerate(zip(old_sentences, new_sentences)):
        # At least one other sentence must contain an actual change.
        if not any(a.strip() != b.strip() for i, (a, b) in enumerate(zip(old_sentences, new_sentences)) if i != index):
            continue
        for rule in _RULES:
            for match in rule.pattern.finditer(old_sentence):
                before = match.group(0)
                after = match.expand(rule.replacement)
                if (len(_occurrences(source, before)) != 1
                        or len(_occurrences(revised, after)) != 1
                        or not _occurrences(new_sentence, after)
                        or _occurrences(revised, before)):
                    continue
                candidates.append((before, after, rule.motivation))
    return candidates if len(candidates) == 1 else []
