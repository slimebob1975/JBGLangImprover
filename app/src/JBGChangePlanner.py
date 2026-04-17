from dataclasses import dataclass, field
from typing import Optional, Literal

try:
    from app.src.JBGLangImprovSuggestorAI import SuggestedChange
except ModuleNotFoundError:
    from JBGLangImprovSuggestorAI import SuggestedChange


# ============================================================================
# Datamodeller
# ============================================================================

@dataclass
class TextAnchor:
    start: int
    end: int
    matched_text: str
    match_strategy: str  # exact, normalized, fallback


@dataclass
class DiffSegment:
    kind: Literal["equal", "delete", "insert"]
    text: str


@dataclass
class ChangeTarget:
    part_name: str
    element_type: Optional[str]
    element_id: Optional[str]
    footnote_id: Optional[str]


@dataclass
class ChangePlan:
    target: ChangeTarget
    source_text: str
    old_text: str
    new_text: str
    motivation: Optional[str]
    anchor: TextAnchor
    diff_segments: list[DiffSegment] = field(default_factory=list)
    contains_special_tokens: bool = False
    notes: list[str] = field(default_factory=list)


# ============================================================================
# Planner
# ============================================================================

class ChangePlanner:
    """
    Översätter SuggestedChange -> ChangePlan.

    Ansvar:
    - hitta target i dokumentstrukturen
    - mappa till rätt part_name
    - hitta textankare
    - bygga diffsegment via TokenDiffEngine
    - markera risk/specialfall
    - upptäcka överlappande ändringar inom samma target
    """

    DOCX_PART_MAP = {
        "paragraph": "word/document.xml",
        "table_cell": "word/document.xml",
        "textbox": "word/document.xml",
        "footnote": "word/footnotes.xml",
        "header": "word/header1.xml",
        "footer": "word/footer1.xml",
    }

    SENSITIVE_ELEMENT_TYPES = {"footnote", "textbox", "header", "footer", "table_cell"}

    def __init__(self, structure: dict, diff_engine, logger):
        self.structure = structure
        self.diff_engine = diff_engine
        self.logger = logger

    def build_plans(self, suggestions: list[SuggestedChange]) -> list[ChangePlan]:
        plans: list[ChangePlan] = []

        for suggestion in suggestions:
            try:
                plan = self._build_single_plan(suggestion)
            except ValueError as ex:
                self.logger.warning(
                    f"Skipping suggestion during planning: {self._label(suggestion)} - {ex}"
                )
                continue

            plans.append(plan)

        self._annotate_conflicts(plans)
        return plans

    def _build_single_plan(self, suggestion: SuggestedChange) -> ChangePlan:
        return self._build_docx_plan(suggestion)

    # ------------------------------------------------------------------
    # DOCX
    # ------------------------------------------------------------------

    def _build_docx_plan(self, suggestion: SuggestedChange) -> ChangePlan:
        element = self._get_docx_element(suggestion.element_id)
        if element is None:
            raise ValueError(f"Unknown element_id: {suggestion.element_id}")

        source_text = element.get("text", "") or ""

        anchor = self._locate_anchor(
            source_text=source_text,
            old_text=suggestion.old,
            match_status=suggestion.match_status,
        )

        part_name = self._resolve_docx_part_name(
            element_type=suggestion.element_type,
            element=element,
        )

        diff_segments = self.diff_engine.build_diff(
            old_text=suggestion.old,
            new_text=suggestion.new,
        )

        contains_special_tokens = self._detect_special_tokens(
            element_type=suggestion.element_type,
            source_text=source_text,
            diff_segments=diff_segments,
        )

        target = ChangeTarget(
            part_name=part_name,
            element_type=suggestion.element_type,
            element_id=suggestion.element_id,
            footnote_id=suggestion.footnote_id,
        )

        notes = self._build_notes(
            suggestion=suggestion,
            source_text=source_text,
            anchor=anchor,
            diff_segments=diff_segments,
            contains_special_tokens=contains_special_tokens,
        )

        return ChangePlan(
            target=target,
            source_text=source_text,
            old_text=suggestion.old,
            new_text=suggestion.new,
            motivation=suggestion.motivation,
            anchor=anchor,
            diff_segments=diff_segments,
            contains_special_tokens=contains_special_tokens,
            notes=notes,
        )

    # ------------------------------------------------------------------
    # Locatorer
    # ------------------------------------------------------------------

    def _locate_anchor(self, source_text: str, old_text: str, match_status: str) -> TextAnchor:
        exact_index = source_text.find(old_text)
        if exact_index != -1:
            return TextAnchor(
                start=exact_index,
                end=exact_index + len(old_text),
                matched_text=old_text,
                match_strategy="exact",
            )

        normalized_span = self._find_normalized_span(source_text, old_text)
        if normalized_span is not None:
            start, end, matched = normalized_span
            return TextAnchor(
                start=start,
                end=end,
                matched_text=matched,
                match_strategy="normalized",
            )

        raise ValueError("Could not locate old_text within source_text")

    def _find_normalized_span(self, source_text: str, old_text: str) -> Optional[tuple[int, int, str]]:
        normalized_old = self._normalize_text(old_text)
        if not normalized_old:
            return None

        source_spans = self._normalized_char_map(source_text)
        normalized_source = "".join(char for _, char in source_spans)

        idx = normalized_source.find(normalized_old)
        if idx == -1:
            return None

        start_original = source_spans[idx][0]
        end_original = source_spans[idx + len(normalized_old) - 1][0] + 1
        matched_text = source_text[start_original:end_original]
        return start_original, end_original, matched_text

    def _normalized_char_map(self, text: str) -> list[tuple[int, str]]:
        result: list[tuple[int, str]] = []
        previous_was_space = False

        for i, ch in enumerate(text.replace("\u00a0", " ")):
            if ch.isspace():
                if not previous_was_space:
                    result.append((i, " "))
                previous_was_space = True
            else:
                result.append((i, ch))
                previous_was_space = False

        while result and result[0][1] == " ":
            result.pop(0)
        while result and result[-1][1] == " ":
            result.pop()

        return result

    @staticmethod
    def _normalize_text(text: str) -> str:
        import re
        text = text.replace("\u00a0", " ")
        text = re.sub(r"\s+", " ", text)
        return text.strip()

    # ------------------------------------------------------------------
    # Metadata / specialfall
    # ------------------------------------------------------------------

    def _resolve_docx_part_name(self, element_type: Optional[str], element: dict) -> str:
        explicit_part = element.get("part_name")
        if explicit_part:
            return explicit_part

        if element_type not in self.DOCX_PART_MAP:
            return "word/document.xml"

        return self.DOCX_PART_MAP[element_type]

    def _detect_special_tokens(
        self,
        element_type: Optional[str],
        source_text: str,
        diff_segments: list[DiffSegment],
    ) -> bool:
        if element_type in self.SENSITIVE_ELEMENT_TYPES:
            return True

        if "\t" in source_text or "\n" in source_text:
            return True

        if any(seg.text in {"\n", "\r", "\r\n", "\t"} for seg in diff_segments):
            return True

        return False

    def _build_notes(
        self,
        suggestion: SuggestedChange,
        source_text: str,
        anchor: TextAnchor,
        diff_segments: list[DiffSegment],
        contains_special_tokens: bool,
    ) -> list[str]:
        notes: list[str] = []

        if suggestion.match_status == "normalized":
            notes.append("anchor_resolved_via_normalized_match")

        if suggestion.element_type == "footnote":
            notes.append("requires_footnote_preservation")

        if suggestion.element_type == "textbox":
            notes.append("requires_textbox_xml_handling")

        if suggestion.element_type in {"header", "footer"}:
            notes.append("comment_support_may_be_limited")

        if anchor.start == 0 and anchor.end == len(source_text):
            notes.append("replacement_covers_entire_element_text")

        if contains_special_tokens:
            notes.append("contains_special_tokens")

        non_equal_segments = [seg for seg in diff_segments if seg.kind != "equal"]
        if len(non_equal_segments) > 4:
            notes.append("complex_multi_segment_rewrite")

        if self._is_whitespace_sensitive(diff_segments):
            notes.append("whitespace_sensitive_change")

        if self._is_punctuation_sensitive(diff_segments):
            notes.append("punctuation_sensitive_change")

        if len(suggestion.old) > 300:
            notes.append("large_anchor_region")

        return notes

    def _is_whitespace_sensitive(self, diff_segments: list[DiffSegment]) -> bool:
        changed_text = "".join(seg.text for seg in diff_segments if seg.kind != "equal")
        return any(ch.isspace() for ch in changed_text)

    def _is_punctuation_sensitive(self, diff_segments: list[DiffSegment]) -> bool:
        changed_text = "".join(seg.text for seg in diff_segments if seg.kind != "equal")
        punctuation_chars = set(",.;:!?()[]{}\"'”’“‘-–—/%")
        return any(ch in punctuation_chars for ch in changed_text)

    # ------------------------------------------------------------------
    # Konfliktkontroll
    # ------------------------------------------------------------------

    def _annotate_conflicts(self, plans: list[ChangePlan]) -> None:
        """
        Markerar överlappande ändringar inom samma target.
        Första version: vi flaggar konflikter, men filtrerar inte bort dem här.
        """
        grouped: dict[tuple, list[ChangePlan]] = {}

        for plan in plans:
            key = self._target_conflict_key(plan)
            grouped.setdefault(key, []).append(plan)

        for _, group in grouped.items():
            if len(group) < 2:
                continue

            group.sort(key=lambda p: (p.anchor.start, p.anchor.end))

            for i in range(len(group)):
                current = group[i]
                for j in range(i + 1, len(group)):
                    other = group[j]

                    if other.anchor.start >= current.anchor.end:
                        break

                    if self._anchors_overlap(current.anchor, other.anchor):
                        if "overlapping_change_conflict" not in current.notes:
                            current.notes.append("overlapping_change_conflict")
                        if "overlapping_change_conflict" not in other.notes:
                            other.notes.append("overlapping_change_conflict")

    def _target_conflict_key(self, plan: ChangePlan) -> tuple:
        t = plan.target
        return (
            t.part_name,
            t.element_type,
            t.element_id,
            t.footnote_id,
        )

    def _anchors_overlap(self, a: TextAnchor, b: TextAnchor) -> bool:
        return max(a.start, b.start) < min(a.end, b.end)

    # ------------------------------------------------------------------
    # Strukturhjälpare
    # ------------------------------------------------------------------

    def _get_docx_element(self, element_id: Optional[str]) -> Optional[dict]:
        if not element_id:
            return None

        for element in self.structure.get("elements", []):
            if element.get("element_id") == element_id:
                return element
        return None

    def _get_pdf_line(self, page: Optional[int], line: Optional[int]) -> Optional[dict]:
        for page_obj in self.structure.get("pages", []):
            if page_obj.get("page") != page:
                continue
            for line_obj in page_obj.get("lines", []):
                if line_obj.get("line") == line:
                    return line_obj
        return None

    def _label(self, suggestion: SuggestedChange) -> str:
        return f"{suggestion.element_type}:{suggestion.element_id}"