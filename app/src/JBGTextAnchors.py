"""Conservative, shared text anchoring; never choose a first occurrence."""
import re
import unicodedata


def word_char(ch):
    return bool(ch) and (ch.isalnum() or ch in "_-\u2011" or unicodedata.category(ch).startswith("M"))


def splits_word(text, index):
    if not 0 < index < len(text):
        return False
    left, right = text[index - 1], text[index]
    return (word_char(left) and word_char(right)
            or right in ".'\u2019" and word_char(left) and index + 1 < len(text) and word_char(text[index + 1])
            or left in ".'\u2019" and word_char(right) and index >= 2 and word_char(text[index - 2]))


def unique_exact_span(old, text):
    if not old or not old.strip():
        return None
    start = text.find(old)
    if start < 0 or text.find(old, start + 1) >= 0:
        return None
    end = start + len(old)
    if splits_word(text, start) or splits_word(text, end):
        return None
    return start, end


def unique_normalized_span(old, text):
    # Whitespace-only normalization. Keep actual source offsets; lookahead
    # also detects overlapping occurrences. Do not consume surrounding space.
    if not old or not old.strip():
        return None
    pattern = "".join(r"\s+" if part.isspace() else re.escape(part)
                      for part in re.split(r"(\s+)", old) if part)
    if old[0].isspace():
        pattern = r"(?<!\s)" + pattern
    if old[-1].isspace():
        pattern += r"(?!\s)"
    pattern = r"(?=(" + pattern + r"))"
    matches = re.finditer(pattern, text)
    match = next(matches, None)
    if match is None or next(matches, None) is not None:
        return None
    start, end = match.span(1)
    if splits_word(text, start) or splits_word(text, end):
        return None
    return start, end


def locate_span(old, text):
    # An ambiguous or partial exact match must not be rescued by normalization.
    if old in text:
        return unique_exact_span(old, text)
    return unique_normalized_span(old, text)


def verify_plan_prefix(plan, current_text):
    """Reject a changed coordinate system; edits to the right are harmless."""
    start, end = plan.anchor.start, plan.anchor.end
    if (not 0 <= start < end <= len(plan.source_text)
            or plan.source_text[start:end] != plan.old_text
            or current_text[:end] != plan.source_text[:end]):
        raise ValueError("Source anchor prefix mismatch; refusing to relocate")
