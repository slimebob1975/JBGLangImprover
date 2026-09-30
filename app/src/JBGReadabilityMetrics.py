"""
Läsbarhetsmått (LIX) före och efter föreslagna ändringar.

Modulen är deterministisk och arbetar enbart mot JSON-data:
- strukturfilen från DocumentStructureExtractor (<fil>_structure.json)
- förslagsfilen från JBGLangImprovSuggestorAI (<fil>_suggestions.json)

LIX = ord/meningar + 100 * långa ord/ord, där ett långt ord har fler än
sex bokstäver.

"Efter"-värdet förutsätter att samtliga förslag i förslagsfilen accepteras.
Förslag som inte går att förankra i texten, eller som överlappar ett tidigare
förslag i samma element, hoppas över och räknas i rapporten.
"""

import json
import re
import sys
from dataclasses import dataclass, field, asdict
from typing import Any, Iterable, Optional


# Standardomfång: löptext i huvuddokumentet. Rubriker räknas inte med
# eftersom korta rubriker utan punkt skulle ge många "meningar" och sänka LIX.
DEFAULT_SCOPE_TYPES = ("paragraph", "table_cell", "textbox")

# Vanlig indelning av LIX-värden.
LIX_BANDS = (
    (30.0, "Mycket lättläst"),
    (40.0, "Lättläst"),
    (50.0, "Medelsvår"),
    (60.0, "Svår"),
    (float("inf"), "Mycket svår"),
)

LONG_WORD_MIN_LETTERS = 7

# Ett ord: en följd av bokstäver/siffror, ev. sammanbunden med bindestreck,
# apostrof eller punkt utan blanksteg (a-kassor, 2020-talet, t.ex., 3.5).
_WORD_RE = re.compile(r"[^\W_]+(?:[-'’.][^\W_]+)*", re.UNICODE)

# Manuell radbrytning som troligen börjar en ny mening/listrad: nästa rad
# börjar med versal, siffra eller punkttecken och föregående rad slutar inte
# med kommatecken, semikolon eller tankstreck. Övriga radbrytningar är
# layoutbrytningar mitt i en mening och behandlas som blanksteg.
_LINE_BREAK_BOUNDARY_RE = re.compile(
    r"(?<![,;–—-])[ \t]*[\r\n]+[ \t]*(?=[\"'”“(\[]?[A-ZÅÄÖÉÜ0-9•▪◦·*])"
)

# Meningsslut: . ! ? eller : följt av blanksteg och versal/siffra/citattecken,
# eller av textens slut. Förkortningar som "t.ex. att" och decimaltal som
# "3.5" ger därför ingen meningsgräns.
_SENTENCE_END_RE = re.compile(
    r"[.!?:]+[\"'”’)\]]*(?=\s+[\"'”“(\[]?[A-ZÅÄÖÉÜ0-9]|\s*$)"
)


# ============================================================================
# Datamodeller
# ============================================================================

@dataclass
class ReadabilityStats:
    words: int = 0
    sentences: int = 0
    long_words: int = 0

    @property
    def lix(self) -> Optional[float]:
        if self.words == 0 or self.sentences == 0:
            return None
        return self.words / self.sentences + 100.0 * self.long_words / self.words

    def add(self, other: "ReadabilityStats") -> None:
        self.words += other.words
        self.sentences += other.sentences
        self.long_words += other.long_words

    def to_dict(self) -> dict:
        lix = self.lix
        return {
            "words": self.words,
            "sentences": self.sentences,
            "long_words": self.long_words,
            "lix": round(lix, 1) if lix is not None else None,
            "band": lix_band(lix),
        }


@dataclass
class SectionReadability:
    heading: Optional[str]
    heading_element_id: Optional[str]
    before: ReadabilityStats = field(default_factory=ReadabilityStats)
    after: ReadabilityStats = field(default_factory=ReadabilityStats)

    def to_dict(self) -> dict:
        return {
            "heading": self.heading,
            "heading_element_id": self.heading_element_id,
            "before": self.before.to_dict(),
            "after": self.after.to_dict(),
        }


@dataclass
class DocumentReadabilityReport:
    scope_types: list[str]
    elements_counted: int = 0
    before: ReadabilityStats = field(default_factory=ReadabilityStats)
    after: ReadabilityStats = field(default_factory=ReadabilityStats)
    suggestions_total: int = 0
    suggestions_in_scope: int = 0
    suggestions_applied: int = 0
    suggestions_unanchored: int = 0
    suggestions_overlapping: int = 0
    sections: list[SectionReadability] = field(default_factory=list)

    def to_dict(self) -> dict:
        delta = rounded_lix_delta(self.before.lix, self.after.lix)
        return {
            "scope_types": list(self.scope_types),
            "headings_excluded": True,
            "assumes_all_suggestions_accepted": True,
            "elements_counted": self.elements_counted,
            "before": self.before.to_dict(),
            "after": self.after.to_dict(),
            "lix_delta": delta,
            "suggestions": {
                "total": self.suggestions_total,
                "in_scope": self.suggestions_in_scope,
                "applied_in_memory": self.suggestions_applied,
                "unanchored": self.suggestions_unanchored,
                "overlapping": self.suggestions_overlapping,
            },
            "sections": [s.to_dict() for s in self.sections],
        }


# ============================================================================
# Grundmått
# ============================================================================

def lix_band(lix: Optional[float]) -> Optional[str]:
    if lix is None:
        return None
    for upper, label in LIX_BANDS:
        if lix < upper:
            return label
    return LIX_BANDS[-1][1]


def rounded_lix_delta(before: Optional[float], after: Optional[float]) -> Optional[float]:
    """
    Skillnaden beräknas från de avrundade värdena, så att den alltid stämmer
    med de värden som visas (51.5 -> 51.6 ger 0.1, inte 0.0).
    """
    if before is None or after is None:
        return None
    return round(round(after, 1) - round(before, 1), 1)


def text_stats(text: str) -> ReadabilityStats:
    """
    Räknar ord, meningar och långa ord i en text (ett element).

    Varje element med minst ett ord räknas som minst en mening, även utan
    avslutande skiljetecken (punktlistor, tabellceller). En manuell
    radbrytning räknas som meningsgräns bara när nästa rad ser ut att börja
    en ny mening eller listrad (se _LINE_BREAK_BOUNDARY_RE); layoutbrytningar
    mitt i en mening räknas som blanksteg.
    """
    stats = ReadabilityStats()
    if not text or not text.strip():
        return stats

    for line in _LINE_BREAK_BOUNDARY_RE.split(text):
        line = re.sub(r"\s+", " ", line)
        line_words = _WORD_RE.findall(line)
        if not line_words:
            continue

        stats.words += len(line_words)
        stats.long_words += sum(
            1 for word in line_words
            if sum(ch.isalpha() for ch in word) >= LONG_WORD_MIN_LETTERS
        )

        sentences = 0
        start = 0
        for match in _SENTENCE_END_RE.finditer(line):
            if _WORD_RE.search(line[start:match.end()]):
                sentences += 1
            start = match.end()
        if _WORD_RE.search(line[start:]):
            sentences += 1  # sista meningen saknar avslutande tecken
        stats.sentences += max(1, sentences)

    return stats


def stats_for_texts(texts: Iterable[str]) -> ReadabilityStats:
    total = ReadabilityStats()
    for text in texts:
        total.add(text_stats(text))
    return total


# ============================================================================
# Tillämpning av förslag i minnet
# ============================================================================

def _locate(old: str, text: str) -> Optional[tuple[int, int]]:
    """Exakt träff först, därefter träff med normaliserade blanksteg."""
    if not old:
        return None
    idx = text.find(old)
    if idx >= 0:
        return idx, idx + len(old)

    tokens = old.split()
    if not tokens:
        return None
    pattern = r"\s+".join(re.escape(tok) for tok in tokens)
    match = re.search(pattern, text)
    if match:
        return match.start(), match.end()
    return None


def apply_suggestions_to_text(text: str, suggestions: list[dict]) -> tuple[str, int, int, int]:
    """
    Applicerar förslag ({"old", "new"}) på en elementtext.

    Returnerar (ny_text, applicerade, ej_förankrade, överlappande).
    Förslagen ankras mot originaltexten, så de påverkar inte varandras
    positioner; överlappande spann hoppas över (första förslaget vinner).
    """
    spans: list[tuple[int, int, str]] = []
    unanchored = overlapping = 0

    for suggestion in suggestions:
        old = suggestion.get("old")
        new = suggestion.get("new")
        if not isinstance(old, str) or not isinstance(new, str):
            unanchored += 1
            continue
        span = _locate(old, text)
        if span is None:
            unanchored += 1
            continue
        start, end = span
        if any(start < s_end and s_start < end for s_start, s_end, _ in spans):
            overlapping += 1
            continue
        spans.append((start, end, new))

    result = text
    for start, end, new in sorted(spans, key=lambda s: s[0], reverse=True):
        result = result[:start] + new + result[end:]

    return result, len(spans), unanchored, overlapping


# ============================================================================
# Dokumentrapport
# ============================================================================

def _is_heading(element: dict) -> bool:
    return element.get("heading_level") is not None


# Bildtexter och tabellrubriker har ofta outline level i mallen (för att synas
# i navigeringsfönstret) men är inte avsnittsrubriker.
_CAPTION_STYLE_TOKENS = ("caption", "beskrivning", "tabellrubrik", "figurrubrik", "diagramrubrik")
_CAPTION_TEXT_RE = re.compile(
    r"^\s*(?:tabell|figur|diagram|bild|karta|table|figure)\s+\d+", re.IGNORECASE
)


def is_caption(element: dict) -> bool:
    style = f"{element.get('style_id') or ''} {element.get('style_name') or ''}".lower()
    if any(token in style for token in _CAPTION_STYLE_TOKENS):
        return True
    return bool(_CAPTION_TEXT_RE.match(element.get("text") or ""))


def is_section_heading(element: dict) -> bool:
    """
    En rubrik som öppnar ett avsnitt i dokumentets disposition: ett stycke i
    brödtexten med rubriknivå >= 1 och synlig text. Rubriker i tabeller och
    textrutor (faktarutor, skalor) samt bildtexter räknas inte.
    """
    return (
        element.get("type") == "paragraph"
        and (element.get("heading_level") or 0) >= 1
        and bool((element.get("text") or "").strip())
        and not is_caption(element)
    )


def compute_document_readability(
    structure: dict,
    suggestions: Optional[list[dict]] = None,
    scope_types: Iterable[str] = DEFAULT_SCOPE_TYPES,
) -> DocumentReadabilityReport:
    scope = tuple(scope_types)
    suggestions = suggestions or []
    elements = structure.get("elements", []) if isinstance(structure, dict) else []

    report = DocumentReadabilityReport(scope_types=list(scope))
    report.suggestions_total = len(suggestions)

    by_element: dict[str, list[dict]] = {}
    for suggestion in suggestions:
        if isinstance(suggestion, dict) and suggestion.get("element_id"):
            by_element.setdefault(suggestion["element_id"], []).append(suggestion)

    in_scope_ids = {
        e.get("element_id") for e in elements
        if e.get("type") in scope and not _is_heading(e)
    }
    report.suggestions_in_scope = sum(
        len(items) for element_id, items in by_element.items() if element_id in in_scope_ids
    )

    # Elementtexter före/efter.
    per_element: dict[str, tuple[ReadabilityStats, ReadabilityStats]] = {}
    for element in elements:
        element_id = element.get("element_id")
        if element_id not in in_scope_ids:
            continue
        text = element.get("text") or ""
        before = text_stats(text)
        after_text, applied, unanchored, overlapping = apply_suggestions_to_text(
            text, by_element.get(element_id, [])
        )
        report.suggestions_applied += applied
        report.suggestions_unanchored += unanchored
        report.suggestions_overlapping += overlapping

        after = text_stats(after_text)
        if before.words == 0 and after.words == 0:
            continue
        report.elements_counted += 1
        report.before.add(before)
        report.after.add(after)
        per_element[element_id] = (before, after)

    report.sections = _sections(elements, per_element)
    return report


def _sections(
    elements: list[dict],
    per_element: dict[str, tuple[ReadabilityStats, ReadabilityStats]],
) -> list[SectionReadability]:
    """
    Grupperar löptexten per avsnitt på översta rubriknivån (lägsta
    heading_level bland avsnittsrubrikerna, se is_section_heading).
    Avsnitt utan löptext utelämnas. Kräver doc_order; saknas det
    returneras en tom lista.
    """
    ordered = [e for e in elements if e.get("doc_order") is not None]
    if not ordered:
        return []
    ordered.sort(key=lambda e: e["doc_order"])

    levels = [e["heading_level"] for e in ordered if is_section_heading(e)]
    if not levels:
        return []
    top_level = min(levels)

    sections: list[SectionReadability] = []
    current = SectionReadability(heading=None, heading_element_id=None)
    for element in ordered:
        if is_section_heading(element) and element["heading_level"] == top_level:
            if current.before.words or current.after.words:
                sections.append(current)
            current = SectionReadability(
                heading=(element.get("text") or "").strip(),
                heading_element_id=element.get("element_id"),
            )
            continue
        stats = per_element.get(element.get("element_id"))
        if stats:
            current.before.add(stats[0])
            current.after.add(stats[1])

    if current.before.words or current.after.words:
        sections.append(current)
    return sections


def compute_document_readability_from_files(
    structure_path: str,
    suggestions_path: Optional[str] = None,
    scope_types: Iterable[str] = DEFAULT_SCOPE_TYPES,
) -> DocumentReadabilityReport:
    with open(structure_path, encoding="utf-8") as f:
        structure = json.load(f)
    suggestions: list[dict] = []
    if suggestions_path:
        with open(suggestions_path, encoding="utf-8") as f:
            loaded = json.load(f)
        if isinstance(loaded, list):
            suggestions = [s for s in loaded if isinstance(s, dict)]
    return compute_document_readability(structure, suggestions, scope_types)


def main():
    if len(sys.argv) not in (2, 3):
        print("Usage: python -m app.src.JBGReadabilityMetrics <structure.json> [suggestions.json]")
        sys.exit(1)
    report = compute_document_readability_from_files(
        sys.argv[1], sys.argv[2] if len(sys.argv) == 3 else None
    )
    print(json.dumps(report.to_dict(), indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
