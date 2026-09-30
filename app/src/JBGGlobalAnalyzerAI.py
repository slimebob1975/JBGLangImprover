"""
Global granskning på dokumentnivå (Fas 3).

Flöde:
1. build_outline(): kompakt disposition av brödtexten i läsordning
   (id, rubriknivå, text). Sidhuvuden, sidfötter och fotnoter ingår inte.
2. En modellanrop per del. Normalt är hela dokumentet en del; mycket långa
   dokument delas vid rubriker på översta nivån.
3. Varje iakttagelse valideras mot strukturen: kända id, citatet måste finnas
   i texten, kategorin måste stödjas. Underkända iakttagelser sparas med skäl.

Iakttagelserna ändrar inte texten. De renderas som Word-kommentarer av
GlobalFindingsRenderer.
"""

import json
import os
import re
import time
from dataclasses import dataclass, field, asdict
from typing import Any, Optional

import openai

try:
    from app.src.JBGReadabilityMetrics import is_caption, is_section_heading
except ModuleNotFoundError:
    from JBGReadabilityMetrics import is_caption, is_section_heading


# Kategorier som koden kan hantera, med den etikett som inleder kommentaren
# i Word. En kategori läggs till här när policyn och renderingen stödjer den,
# t.ex. "inconsistency": "Inkonsekvent påstående".
GLOBAL_CATEGORIES = {
    "repetition": "Onödig upprepning",
    "inconsistency": "Inkonsekvent påstående",
    "error": "Troligt fel",
    "disposition": "Förslag om disposition",
    "heading": "Förslag om rubrik",
}

# Kategorier vars kommentar alltid sitter på en rubrik.
HEADING_ANCHOR_CATEGORIES = {"disposition", "heading"}

# Högsta antal godkända iakttagelser per kategori. Förslag om disposition och
# rubriker hålls få, så att kommentarerna inte uppfattas som ett omdöme.
MAX_PER_CATEGORY = {
    "repetition": 10,
    "inconsistency": 10,
    "error": 10,
    "disposition": 5,
    "heading": 5,
}

# Minsta antal ställen per kategori. Upprepningar och motsägelser kräver
# två ställen; ett troligt fel kan finnas på ett enda.
MIN_LOCATIONS = {
    "repetition": 2,
    "inconsistency": 2,
    "error": 1,
    "disposition": 1,
    "heading": 1,
}

# Kategorier där det motstridiga stället måste citeras och kontrolleras.
CATEGORIES_REQUIRING_RELATED_QUOTE = {"inconsistency"}

OUTLINE_TYPES = ("paragraph", "table_cell", "textbox")
MAX_OUTLINE_CHARS_PER_PART = 240_000     # ca 60 000 tokens
MAX_FINDINGS = 20
MAX_QUOTE_CHARS = 300                    # policyn ber om 150; viss marginal
MIN_QUOTE_FRAGMENT_CHARS = 15            # minst en del av ett citat med "…" så här lång
_ELLIPSIS_RE = re.compile(r"\s*(?:\.{3,}|…)\s*")

DEFAULT_POLICY_PATH = os.path.join(
    os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
    "policy", "global_prompt_policy.md",
)

_MARKERS = ("<!-- START_LOCKED -->", "<!-- END_LOCKED -->",
            "<!-- START_EDITABLE -->", "<!-- END_EDITABLE -->")


# ============================================================================
# Datamodeller
# ============================================================================

@dataclass
class GlobalFinding:
    category: str
    element_ids: list[str]
    quote: str
    description: str
    proposal: str = ""
    part: int = 1
    related_quote: str = ""   # citat ur element_ids[1], om modellen angav ett
    proposed_order: list[str] = field(default_factory=list)   # rubriktexter (disposition)

    @property
    def label(self) -> str:
        return GLOBAL_CATEGORIES.get(self.category, self.category)


@dataclass
class RejectedFinding:
    reason: str
    raw: dict


@dataclass
class GlobalReviewResult:
    parts: int = 0
    raw_count: int = 0
    findings: list[GlobalFinding] = field(default_factory=list)
    rejected: list[RejectedFinding] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    def by_category(self) -> dict[str, int]:
        counts: dict[str, int] = {}
        for finding in self.findings:
            counts[finding.category] = counts.get(finding.category, 0) + 1
        return counts

    def summary(self) -> dict:
        return {
            "parts": self.parts,
            "raw": self.raw_count,
            "accepted": len(self.findings),
            "rejected": len(self.rejected),
            "by_category": self.by_category(),
            "errors": list(self.errors),
        }

    def to_dict(self) -> dict:
        return {
            "summary": self.summary(),
            "findings": [asdict(f) for f in self.findings],
            "rejected": [asdict(r) for r in self.rejected],
        }

    def save(self, path: str) -> str:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, indent=2, ensure_ascii=False)
        return path


# ============================================================================
# Policy
# ============================================================================

def load_global_policy(path: str = DEFAULT_POLICY_PATH) -> str:
    with open(path, encoding="utf-8") as f:
        content = f.read()
    for marker in _MARKERS:
        content = content.replace(marker, "")
    return content.strip()


# ============================================================================
# Disposition
# ============================================================================

def build_outline(structure: dict) -> list[dict]:
    """Brödtextens element i läsordning: {"id", "h"?, "t"}."""
    elements = [
        e for e in structure.get("elements", [])
        if e.get("type") in OUTLINE_TYPES and (e.get("text") or "").strip()
    ]
    # Element utan doc_order (äldre strukturfiler) behåller sin ordning sist.
    elements.sort(key=lambda e: (e.get("doc_order") is None, e.get("doc_order") or 0))

    outline = []
    for element in elements:
        row = {"id": element["element_id"]}
        level = element.get("heading_level")
        if level is not None and level >= 1:
            row["h"] = level
        row["t"] = " ".join((element.get("text") or "").split())
        outline.append(row)
    return outline


def split_outline(structure: dict, outline: list[dict], max_chars: int = MAX_OUTLINE_CHARS_PER_PART) -> list[list[dict]]:
    """
    Delar dispositionen vid avsnittsrubriker på översta nivån så att varje del
    ryms inom max_chars. Ett avsnitt som ensamt är för stort delas vid element.
    """
    def size(rows):
        return sum(len(json.dumps(r, ensure_ascii=False)) + 1 for r in rows)

    if size(outline) <= max_chars:
        return [outline]

    by_id = {e["element_id"]: e for e in structure.get("elements", [])}
    levels = [by_id[r["id"]]["heading_level"] for r in outline
              if r["id"] in by_id and is_section_heading(by_id[r["id"]])]
    top = min(levels) if levels else None

    sections: list[list[dict]] = [[]]
    for row in outline:
        element = by_id.get(row["id"], {})
        if top is not None and is_section_heading(element) and element.get("heading_level") == top and sections[-1]:
            sections.append([])
        sections[-1].append(row)

    parts: list[list[dict]] = [[]]
    for section in sections:
        if parts[-1] and size(parts[-1]) + size(section) > max_chars:
            parts.append([])
        if size(section) > max_chars:
            for row in section:
                if parts[-1] and size(parts[-1]) + size([row]) > max_chars:
                    parts.append([])
                parts[-1].append(row)
        else:
            parts[-1].extend(section)
    return [p for p in parts if p]


# ============================================================================
# Analysator
# ============================================================================

class JBGGlobalAnalyzerAI:
    def __init__(
        self,
        api_key: str,
        model: str,
        temperature: float,
        logger,
        policy: Optional[str] = None,
        usage_tracker=None,
        progress_callback=None,
        max_chars_per_part: int = MAX_OUTLINE_CHARS_PER_PART,
        pause_seconds: float = 0.0,
    ):
        self.api_key = api_key
        self.model = model
        self.temperature = temperature
        self.logger = logger
        self.policy = policy if policy is not None else load_global_policy()
        self.usage_tracker = usage_tracker
        self.progress_callback = progress_callback
        self.max_chars_per_part = max_chars_per_part
        self.pause_seconds = pause_seconds

    # ------------------------------------------------------------------
    # Publikt API
    # ------------------------------------------------------------------

    def analyze(self, structure: dict) -> GlobalReviewResult:
        result = GlobalReviewResult()
        outline = build_outline(structure)
        if not outline:
            self.logger.info("Global review: no body text to analyze")
            return result

        parts = split_outline(structure, outline, self.max_chars_per_part)
        result.parts = len(parts)
        if len(parts) > 1:
            self.logger.warning(
                f"Global review: document split into {len(parts)} parts; "
                "repetitions across parts may be missed"
            )

        client = openai.OpenAI(api_key=self.api_key)
        raw_findings: list[tuple[int, Any]] = []
        for index, part in enumerate(parts, start=1):
            if index > 1 and self.pause_seconds:
                time.sleep(self.pause_seconds)
            # Orkestreraren rapporterar redan starten; här bara delar.
            if len(parts) > 1:
                self._report(f"Granskar dokumentet som helhet (del {index} av {len(parts)})...")
            try:
                raw_text = self._call_model(client, part)
                items = self.parse_response(raw_text)
                raw_findings.extend((index, item) for item in items)
            except Exception as ex:
                message = f"Global review part {index} failed: {ex}"
                self.logger.error(message)
                result.errors.append(message)

        result.raw_count = len(raw_findings)
        self.validate(structure, raw_findings, result)
        self.logger.info(
            f"Global review: {len(result.findings)} findings accepted, "
            f"{len(result.rejected)} rejected, parts={result.parts}"
        )
        return result

    # ------------------------------------------------------------------
    # Modellanrop och tolkning
    # ------------------------------------------------------------------

    def _call_model(self, client, part: list[dict]) -> str:
        lines = "\n".join(json.dumps(row, ensure_ascii=False) for row in part)
        messages = [
            {"role": "system", "content": self.policy},
            {"role": "user", "content": f"Här är dokumentet som ska granskas:\n{lines}"},
        ]
        try:
            response = client.chat.completions.create(
                model=self.model, messages=messages, temperature=self.temperature,
            )
        except Exception:
            if self.usage_tracker is not None:
                self.usage_tracker.record_failure("global", self.model)
            raise
        if self.usage_tracker is not None:
            rec = self.usage_tracker.record("global", self.model, getattr(response, "usage", None))
            self.logger.info(
                f"Token usage (global): prompt={rec.prompt_tokens}, completion={rec.completion_tokens}"
            )
        return response.choices[0].message.content or ""

    @staticmethod
    def parse_response(raw_text: str) -> list:
        text = (raw_text or "").strip()
        if text.startswith("```"):
            text = re.sub(r"^```(?:json)?", "", text).strip()
            text = re.sub(r"```$", "", text).strip()
        if not text:
            return []
        try:
            parsed = json.loads(text)
        except json.JSONDecodeError:
            # Sista utväg: första JSON-objektet eller -listan i svaret
            match = re.search(r"(\{.*\}|\[.*\])", text, re.DOTALL)
            if not match:
                raise ValueError("Global review response was not JSON")
            parsed = json.loads(match.group(1))
        if isinstance(parsed, dict):
            parsed = parsed.get("findings", [])
        if not isinstance(parsed, list):
            raise ValueError("Global review response must contain a list of findings")
        return parsed

    # ------------------------------------------------------------------
    # Validering
    # ------------------------------------------------------------------

    def validate(self, structure: dict, raw_findings: list[tuple[int, Any]], result: GlobalReviewResult) -> None:
        texts = {
            e["element_id"]: e.get("text") or ""
            for e in structure.get("elements", [])
            if e.get("type") in OUTLINE_TYPES
        }
        elements = {e["element_id"]: e for e in structure.get("elements", [])}
        heading_texts = {
            self._normalize_heading(e.get("text") or ""): " ".join((e.get("text") or "").split())
            for e in structure.get("elements", [])
            if e.get("type") in OUTLINE_TYPES and self._is_heading(e)
        }
        seen: set[tuple[str, frozenset]] = set()
        per_category: dict[str, int] = {}

        for part, item in raw_findings:
            if not isinstance(item, dict):
                result.rejected.append(RejectedFinding("not_an_object", {"value": repr(item)}))
                continue

            category = str(item.get("category") or "").strip()
            if category not in GLOBAL_CATEGORIES:
                result.rejected.append(RejectedFinding("unsupported_category", item))
                continue

            description = str(item.get("description") or "").strip()
            if not description:
                result.rejected.append(RejectedFinding("missing_description", item))
                continue

            raw_ids = item.get("element_ids") or []
            if isinstance(raw_ids, str):
                raw_ids = [raw_ids]
            ids = []
            for element_id in raw_ids:
                element_id = str(element_id)
                if element_id in texts and element_id not in ids:
                    ids.append(element_id)
            if not ids:
                result.rejected.append(RejectedFinding("unknown_element_ids", item))
                continue

            quote = " ".join(str(item.get("quote") or "").split())[:MAX_QUOTE_CHARS]
            anchor = next((i for i in ids if quote and self._contains(texts[i], quote)), None)
            if anchor is None:
                result.rejected.append(RejectedFinding("quote_not_found", item))
                continue
            # Elementet som innehåller citatet blir ankare (först i listan).
            ids.remove(anchor)
            ids.insert(0, anchor)

            if category in HEADING_ANCHOR_CATEGORIES and not self._is_heading(elements.get(anchor, {})):
                result.rejected.append(RejectedFinding("anchor_not_a_heading", item))
                continue

            if len(ids) < MIN_LOCATIONS.get(category, 2):
                result.rejected.append(RejectedFinding(f"{category}_needs_two_locations", item))
                continue

            related_quote = " ".join(str(item.get("related_quote") or "").split())[:MAX_QUOTE_CHARS]
            if category in CATEGORIES_REQUIRING_RELATED_QUOTE and not related_quote:
                result.rejected.append(RejectedFinding("missing_related_quote", item))
                continue
            if related_quote:
                related = next((i for i in ids[1:] if self._contains(texts[i], related_quote)), None)
                if related is None:
                    if category in CATEGORIES_REQUIRING_RELATED_QUOTE:
                        result.rejected.append(RejectedFinding("related_quote_not_found", item))
                        continue
                    related_quote = ""   # frivilligt citat som inte stämmer tas bort
                else:
                    # Elementet med det motstridiga citatet blir andra i listan.
                    ids.remove(related)
                    ids.insert(1, related)

            key = (category, frozenset(ids))
            if key in seen:
                result.rejected.append(RejectedFinding("duplicate", item))
                continue
            seen.add(key)

            if (
                len(result.findings) >= MAX_FINDINGS
                or per_category.get(category, 0) >= MAX_PER_CATEGORY.get(category, MAX_FINDINGS)
            ):
                result.rejected.append(RejectedFinding("over_limit", item))
                continue
            per_category[category] = per_category.get(category, 0) + 1

            # Föreslagen ordning godtas bara om alla rubriker finns i dokumentet;
            # annars tas den bort men iakttagelsen behålls.
            proposed_order: list[str] = []
            raw_order = item.get("proposed_order") if category == "disposition" else None
            if isinstance(raw_order, list) and raw_order:
                matched = [heading_texts.get(self._normalize_heading(str(h))) for h in raw_order]
                if all(matched) and len(matched) >= 2:
                    proposed_order = matched

            result.findings.append(GlobalFinding(
                category=category,
                element_ids=ids,
                quote=quote,
                description=description,
                proposal=str(item.get("proposal") or "").strip(),
                part=part,
                related_quote=related_quote,
                proposed_order=proposed_order,
            ))

    @staticmethod
    def _is_heading(element: dict) -> bool:
        return (
            (element.get("heading_level") or 0) >= 1
            and bool((element.get("text") or "").strip())
            and not is_caption(element)
        )

    @staticmethod
    def _normalize_heading(text: str) -> str:
        return " ".join(text.split()).strip(" .:").casefold()

    @staticmethod
    def _contains(text: str, quote: str) -> bool:
        """
        Citatet finns i texten (blanksteg normaliserade). Ett citat med
        utelämningstecken ("…" eller "...") godtas om alla delar finns i
        ordning och minst en del är tillräckligt lång för att inte vara slump.
        """
        normalized = " ".join(text.split())
        if quote in normalized:
            return True
        fragments = [f for f in _ELLIPSIS_RE.split(quote) if f.strip()]
        if len(fragments) < 2 and not _ELLIPSIS_RE.search(quote):
            return False
        if not fragments or max(len(f) for f in fragments) < MIN_QUOTE_FRAGMENT_CHARS:
            return False
        position = 0
        for fragment in fragments:
            index = normalized.find(fragment, position)
            if index < 0:
                return False
            position = index + len(fragment)
        return True

    def _report(self, message: str) -> None:
        self.logger.info(message)
        if self.progress_callback is not None:
            try:
                self.progress_callback(message)
            except Exception as ex:
                self.logger.warning(f"Progress callback failed in global analyzer: {ex}")
