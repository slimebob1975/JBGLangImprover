import re
from difflib import SequenceMatcher
from typing import List

try:
    from app.src.JBGChangePlanner import DiffSegment
except ModuleNotFoundError:
    from JBGChangePlanner import DiffSegment


class TokenDiffEngine:
    """
    Bygger en renderingsneutral diff mellan old_text och new_text.

    Princip:
    - tokenisera till ord, whitespace och skiljetecken
    - använd SequenceMatcher på tokennivå
    - slå ihop intilliggande segment av samma typ
    """

    TOKEN_PATTERN = re.compile(
        r"""
        \r\n|[\n\r\t]           # explicita rad-/tabbtecken
        |[ ]+                  # vanliga spaces
        |\d+(?:[.,]\d+)*%?     # tal, ev. decimal/procent
        |[A-Za-zÀ-ÖØ-öø-ÿ_]+   # ord med latinska tecken/underscore
        |[^\w\s]               # skiljetecken/övriga symboler
        """,
        re.VERBOSE | re.UNICODE,
    )

    def __init__(self, logger=None, normalize_nbsp=True):
        self.logger = logger
        self.normalize_nbsp = normalize_nbsp

    def build_diff(self, old_text: str, new_text: str) -> List[DiffSegment]:
        old_tokens = self.tokenize(old_text)
        new_tokens = self.tokenize(new_text)

        matcher = SequenceMatcher(a=old_tokens, b=new_tokens)
        segments: List[DiffSegment] = []

        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == "equal":
                text = self._join_tokens(old_tokens[i1:i2])
                if text:
                    segments.append(DiffSegment(kind="equal", text=text))

            elif tag == "delete":
                text = self._join_tokens(old_tokens[i1:i2])
                if text:
                    segments.append(DiffSegment(kind="delete", text=text))

            elif tag == "insert":
                text = self._join_tokens(new_tokens[j1:j2])
                if text:
                    segments.append(DiffSegment(kind="insert", text=text))

            elif tag == "replace":
                old_part = self._join_tokens(old_tokens[i1:i2])
                new_part = self._join_tokens(new_tokens[j1:j2])

                if old_part:
                    segments.append(DiffSegment(kind="delete", text=old_part))
                if new_part:
                    segments.append(DiffSegment(kind="insert", text=new_part))

        return self._merge_adjacent_segments(segments)

    def tokenize(self, text: str) -> List[str]:
        if text is None:
            return []

        if self.normalize_nbsp:
            text = text.replace("\u00a0", " ")

        tokens = self.TOKEN_PATTERN.findall(text)

        # fallback om regex mot förmodan missar allt men text finns
        if not tokens and text:
            return [text]

        return tokens

    def _join_tokens(self, tokens: List[str]) -> str:
        return "".join(tokens)

    def _merge_adjacent_segments(self, segments: List[DiffSegment]) -> List[DiffSegment]:
        if not segments:
            return []

        merged: List[DiffSegment] = [segments[0]]

        for seg in segments[1:]:
            last = merged[-1]
            if seg.kind == last.kind:
                last.text += seg.text
            else:
                merged.append(seg)

        return merged

    def debug_diff(self, old_text: str, new_text: str) -> List[dict]:
        """
        Hjälpmetod för test/debug.
        """
        return [
            {"kind": seg.kind, "text": seg.text}
            for seg in self.build_diff(old_text, new_text)
        ]
    
if __name__ == "__main__":
    engine = TokenDiffEngine()

    tests = [
        (
            "För varje aktivitetsrapport i under tidsperioden Y för sökande X:",
            "För varje aktivitetsrapport under tidsperioden Y för sökande X:"
        ),
        (
            "Sätt k_i = 0, där k_i är antal sökta jobb som ligger utanför sökandens normala yrkesområde.",
            "Sätt k_i = 0, där k_i är antalet sökta jobb som ligger utanför sökandens normala yrkesområde."
        ),
        (
            "När en potentiell utvidgning har identifierats (dvs. (D) över tröskeln) görs en validering i två steg:",
            "När en potentiell utvidgning har identifierats (det vill säga när D ligger över tröskelvärdet) görs en validering i två steg:"
        ),
    ]

    for i, (old_text, new_text) in enumerate(tests, start=1):
        print(f"\nTEST {i}")
        for row in engine.debug_diff(old_text, new_text):
            print(row)