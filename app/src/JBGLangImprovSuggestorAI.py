import json
import openai
import sys
import os
import re
from collections import defaultdict
from difflib import SequenceMatcher
import time
import logging
from dataclasses import dataclass, field
from typing import Optional, Literal, Any

MAX_TOKEN_PER_CALL = 8000


# ============================================================================
# Datamodeller
# ============================================================================

@dataclass
class SuggestedChange:
    element_type: str
    element_id: str
    footnote_id: Optional[str]
    old: str
    new: str
    motivation: Optional[str]
    match_status: Literal["exact", "normalized", "rejected"]
    safe_to_apply: bool = True
    safety_reason: Optional[str] = None


@dataclass
class SuggestionIssue:
    severity: Literal["warning", "error"]
    code: str
    message: str


@dataclass
class FilteredSuggestion:
    suggestion: SuggestedChange
    accepted: bool
    issues: list[SuggestionIssue] = field(default_factory=list)


# ============================================================================
# Kvalitetsfilter
# ============================================================================

class SuggestionQualityFilter:
    TODO_PATTERNS = [
        r"\bTODO\b",
        r"\bFIXME\b",
        r"\bTK\b",
        r"\bXXX\b",
    ]

    def __init__(
        self,
        logger,
        max_relative_growth=2.5,
        max_absolute_growth=220,
        max_old_length_for_locality_warning=350,
        max_sentence_count_warning=3,
        reject_todo_transformations=True,
        reject_large_functional_rewrites=True,
    ):
        self.logger = logger
        self.max_relative_growth = max_relative_growth
        self.max_absolute_growth = max_absolute_growth
        self.max_old_length_for_locality_warning = max_old_length_for_locality_warning
        self.max_sentence_count_warning = max_sentence_count_warning
        self.reject_todo_transformations = reject_todo_transformations
        self.reject_large_functional_rewrites = reject_large_functional_rewrites

    def filter(
        self,
        suggestions: list[SuggestedChange]
    ) -> tuple[list[SuggestedChange], list[FilteredSuggestion]]:
        accepted: list[SuggestedChange] = []
        reviewed: list[FilteredSuggestion] = []

        for suggestion in suggestions:
            result = self._review_single(suggestion)
            reviewed.append(result)

            if result.accepted:
                accepted.append(suggestion)
            else:
                self.logger.warning(
                    f"Rejected suggestion [{self._target_label(suggestion)}]: "
                    f"{'; '.join(issue.code for issue in result.issues if issue.severity == 'error')}"
                )

        return accepted, reviewed

    def _review_single(self, suggestion: SuggestedChange) -> FilteredSuggestion:
        issues: list[SuggestionIssue] = []

        self._check_internal_note_patterns(suggestion, issues)
        self._check_length_growth(suggestion, issues)
        self._check_locality(suggestion, issues)
        self._check_sentence_expansion(suggestion, issues)
        self._check_semantic_shift_markers(suggestion, issues)
        self._check_typography_consistency(suggestion, issues)
        self._check_element_specific_risk(suggestion, issues)

        accepted = not any(issue.severity == "error" for issue in issues)
        return FilteredSuggestion(
            suggestion=suggestion,
            accepted=accepted,
            issues=issues,
        )

    def _check_internal_note_patterns(self, s: SuggestedChange, issues: list[SuggestionIssue]):
        old = s.old or ""
        new = s.new

        for pattern in self.TODO_PATTERNS:
            if re.search(pattern, old, flags=re.IGNORECASE):
                issues.append(SuggestionIssue(
                    severity="error" if self.reject_todo_transformations else "warning",
                    code="internal_note_rewrite",
                    message="Förslaget verkar skriva om en intern arbetsnotering eller TODO-markör."
                ))
                break

        if not old:
            issues.append(SuggestionIssue(
                severity="error",
                code="empty_old",
                message="Old är tom."
            ))

        if new is None:
            issues.append(SuggestionIssue(
                severity="error",
                code="missing_new",
                message="New saknas."
            ))

    def _check_length_growth(self, s: SuggestedChange, issues: list[SuggestionIssue]):
        old_len = len(s.old.strip())
        new_len = len((s.new or "").strip())

        if old_len == 0:
            issues.append(SuggestionIssue(
                severity="error",
                code="zero_old_length",
                message="Old är tom."
            ))
            return

        relative_growth = new_len / max(old_len, 1)
        absolute_growth = new_len - old_len

        if relative_growth > self.max_relative_growth and absolute_growth > self.max_absolute_growth:
            issues.append(SuggestionIssue(
                severity="error" if self.reject_large_functional_rewrites else "warning",
                code="oversized_rewrite",
                message=(
                    f"Omskrivningen är mycket större än originalet "
                    f"(old={old_len}, new={new_len}, ratio={relative_growth:.2f})."
                )
            ))

    def _check_locality(self, s: SuggestedChange, issues: list[SuggestionIssue]):
        old_len = len(s.old.strip())
        if old_len > self.max_old_length_for_locality_warning:
            issues.append(SuggestionIssue(
                severity="warning",
                code="weak_locality",
                message=f"Old är långt ({old_len} tecken), vilket tyder på svag lokalitet."
            ))

    def _check_sentence_expansion(self, s: SuggestedChange, issues: list[SuggestionIssue]):
        old_sentences = self._sentence_count(s.old)
        new_sentences = self._sentence_count(s.new or "")

        if old_sentences >= self.max_sentence_count_warning or new_sentences >= self.max_sentence_count_warning:
            issues.append(SuggestionIssue(
                severity="warning",
                code="multi_sentence_rewrite",
                message=(
                    f"Förslaget omfattar flera meningar "
                    f"(old={old_sentences}, new={new_sentences})."
                )
            ))

    def _check_semantic_shift_markers(self, s: SuggestedChange, issues: list[SuggestionIssue]):
        old = s.old.lower()
        new = (s.new or "").lower()

        functional_shift_patterns = [
            (r"\btodo\b", r"\bta fram\b"),
            (r"\btodo\b", r"\bmed hjälp av\b"),
            (r"\butkast\b", r"\bfärdig\b"),
            (r"\barbetsanteckning\b", r"\bformell\b"),
        ]

        for old_pat, new_pat in functional_shift_patterns:
            if re.search(old_pat, old) and re.search(new_pat, new):
                issues.append(SuggestionIssue(
                    severity="error" if self.reject_large_functional_rewrites else "warning",
                    code="functional_shift",
                    message="Förslaget verkar ändra textens funktion, inte bara språket."
                ))
                break

        style_shift_pairs = [
            ("kvantifiera", "mäta"),
            ("kan antas", "bedöms"),
            ("siffran", "talet"),
        ]

        for a, b in style_shift_pairs:
            if a in old and b in new:
                issues.append(SuggestionIssue(
                    severity="warning",
                    code="stylistic_rewrite",
                    message=f"Stilistiskt byte ({a} → {b}), inte bara ren språkfelrättning."
                ))
                break

    def _check_typography_consistency(self, s: SuggestedChange, issues: list[SuggestionIssue]):
        old_has_typographic_quotes = any(ch in s.old for ch in ["”", "“", "’", "‘"])
        new_has_ascii_quotes = '"' in (s.new or "") or "'" in (s.new or "")

        if old_has_typographic_quotes and new_has_ascii_quotes:
            issues.append(SuggestionIssue(
                severity="warning",
                code="quote_style_shift",
                message="Förslaget byter från typografiska till raka citattecken."
            ))

    def _check_element_specific_risk(self, s: SuggestedChange, issues: list[SuggestionIssue]):
        element_type = s.element_type or ""
        old_len = len(s.old.strip())
        new_len = len((s.new or "").strip())

        if element_type in {"footnote", "table_cell", "textbox", "header", "footer"}:
            if new_len - old_len > 120:
                issues.append(SuggestionIssue(
                    severity="warning",
                    code="sensitive_element_growth",
                    message=f"Texten blir märkbart längre i känsligt element ({element_type})."
                ))

    def _sentence_count(self, text: str) -> int:
        text = text.strip()
        if not text:
            return 0
        parts = re.split(r"(?<=[.!?])\s+|\n+", text)
        parts = [p for p in parts if p.strip()]
        return len(parts)

    def _target_label(self, s: SuggestedChange) -> str:
        return f"{s.element_type}:{s.element_id}"


# ============================================================================
# Suggestor
# ============================================================================

class JBGLangImprovSuggestorAI:
    """
    Version med:
    - robust JSON-parse
    - schema-validering
    - matchningskontroll mot källstrukturen
    - kvalitetsfilter
    - filterrapport
    - minimering av ändringsspann
    - generell säkerhetsrensning av suggestions före planner
    """

    WORD_CHAR_RE = re.compile(r"[\wÅÄÖåäöÀ-ÖØ-öø-ÿ-]", re.UNICODE)

    def __init__(
        self,
        api_key,
        model,
        prompt_policy,
        temperature,
        logger,
        progress_callback=None,
        strict_validation=True,
        allow_normalized_matches=True,
        quality_filter=None,
    ):
        self.api_key = api_key
        self.model = model
        self.policy_prompt = prompt_policy
        self.temperature = temperature
        self.logger = logger
        self.progress_callback = progress_callback

        self.strict_validation = strict_validation
        self.allow_normalized_matches = allow_normalized_matches

        self.file_path = None
        self.json_structured_document = None
        self.json_suggestions = None

        self.validated_suggestions: list[SuggestedChange] = []
        self.filtered_review: list[FilteredSuggestion] = []

        self.quality_filter = quality_filter or SuggestionQualityFilter(logger=self.logger)

    def _report(self, message: str):
        self.logger.info(message)
        if self.progress_callback is not None:
            try:
                self.progress_callback(message)
            except Exception as ex:
                self.logger.warning(
                    f"Progress callback failed in JBGLangImprovSuggestorAI: {ex}"
                )

    def load_structure(self, filepath):
        self.file_path = filepath
        with open(filepath, "r", encoding="utf-8") as f:
            self.json_structured_document = json.load(f)

        if not self.json_structured_document:
            raise ValueError(f"Could not load JSON document from {filepath}")

    def save_as_json(self, output_path=None, use_validated=True):
        if use_validated:
            if not self.validated_suggestions:
                self.suggest_changes_token_aware_batching()
            payload = [self._suggested_change_to_output_dict(s) for s in self.validated_suggestions]
        else:
            if self.json_suggestions is None:
                self.suggest_changes_token_aware_batching()
            payload = self.json_suggestions

        if not output_path:
            output_path = self.file_path + "_suggestions.json"

        try:
            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(payload, f, indent=2, ensure_ascii=False)
            return output_path
        except Exception as e:
            self.logger.error(f"Error saving suggestions JSON: {str(e)}")
            return None

    def save_filter_report(self, output_path=None):
        if output_path is None:
            output_path = self.file_path + "_suggestion_filter_report.json"

        payload = []
        for row in self.filtered_review:
            payload.append({
                "accepted": row.accepted,
                "suggestion": self._suggested_change_to_output_dict(row.suggestion),
                "issues": [
                    {
                        "severity": issue.severity,
                        "code": issue.code,
                        "message": issue.message,
                    }
                    for issue in row.issues
                ],
            })

        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2, ensure_ascii=False)

        return output_path

    def suggest_changes(self):
        self._ensure_structure_loaded()
        self.logger.info(f"The used prompt policy:\n{str(self.policy_prompt)}\n")

        client = openai.OpenAI(api_key=self.api_key)
        messages = [
            {"role": "system", "content": self.policy_prompt},
            {
                "role": "user",
                "content": f"Här är det dokument som ska granskas: {self.json_structured_document}."
            },
        ]

        try:
            raw_text = self._call_model(client, messages)
            parsed = self._parse_model_response(raw_text)
            validated = self._postprocess_suggestions(parsed)

            accepted, reviewed = self.quality_filter.filter(validated)

            self.json_suggestions = parsed
            self.validated_suggestions = accepted
            self.filtered_review = reviewed

        except Exception as e:
            self.logger.error(f"Error during OpenAI suggestion generation: {e}")
            self.json_suggestions = None
            self.validated_suggestions = []
            self.filtered_review = []

    def suggest_changes_token_aware_batching(self, max_tokens_per_call=MAX_TOKEN_PER_CALL):
        self._ensure_structure_loaded()
        self._report("Promptpolicyn laddad. Förbereder API-anrop...")
        self.logger.info(f"The used prompt policy:\n{str(self.policy_prompt)}\n")

        client = openai.OpenAI(api_key=self.api_key)
        system_msg = {"role": "system", "content": self.policy_prompt}
        structure = self.json_structured_document

        if structure["type"] != "docx":
            raise ValueError("Unsupported document type. Only docx is supported.")

        elements = structure["elements"]
        chunks = self._chunk_elements(elements, max_tokens_per_call=max_tokens_per_call)
        num_chunks = len(chunks)

        self._report(f"Dokumentet är stort. Skickar {num_chunks} separata API-anrop.")

        all_raw_suggestions: list[dict[str, Any]] = []
        all_validated: list[SuggestedChange] = []

        for i, chunk in enumerate(chunks, start=1):
            if i > 1:
                time.sleep(5)

            self._report(f"Gör API-anrop {i} av {num_chunks}.")
            messages = [system_msg, self._build_user_message_for_chunk(chunk)]

            try:
                raw_text = self._call_model(client, messages)
                parsed = self._parse_model_response(raw_text)
                validated = self._postprocess_suggestions(parsed)

                all_raw_suggestions.extend(parsed)
                all_validated.extend(validated)

                self._report(f"Klar med API-anrop {i} av {num_chunks}.")
            except Exception as e:
                msg = f"Fel i API-anrop {i} av {num_chunks}: {e}"
                self.logger.error(msg)
                self._report(msg)

        self.json_suggestions = self._deduplicate_raw_suggestions(all_raw_suggestions)
        all_validated = self._deduplicate_validated_suggestions(all_validated)

        accepted, reviewed = self.quality_filter.filter(all_validated)
        self.validated_suggestions = accepted
        self.filtered_review = reviewed

        self._report("Alla AI-förslag är genererade, validerade och filtrerade.")

    # -------------------------------------------------------------------------
    # Modellanrop / parsing
    # -------------------------------------------------------------------------

    def _ensure_structure_loaded(self):
        if not self.json_structured_document:
            raise ValueError("No structure loaded. Call load_structure() first.")

    def _call_model(self, client, messages):
        response = client.chat.completions.create(
            model=self.model,
            messages=messages,
            temperature=self.temperature,
        )
        return response.choices[0].message.content or "[]"

    def _build_user_message_for_chunk(self, chunk):
        return {
            "role": "user",
            "content": f"Här är en del av dokumentet som ska granskas: {json.dumps(chunk, ensure_ascii=False)}",
        }

    def _parse_model_response(self, raw_text: str) -> list[dict[str, Any]]:
        cleaned = self._clean_json_response(raw_text)

        try:
            parsed = json.loads(cleaned)
        except json.JSONDecodeError as ex:
            raise ValueError(f"Model response was not valid JSON: {ex}") from ex

        if not isinstance(parsed, list):
            raise ValueError("Model response must be a JSON list.")

        normalized_items: list[dict[str, Any]] = []
        for idx, item in enumerate(parsed):
            if not isinstance(item, dict):
                self.logger.warning(f"Skipping non-object suggestion at index {idx}: {item!r}")
                continue
            normalized_items.append(item)

        return normalized_items

    def _clean_json_response(self, raw_text):
        if raw_text.strip().startswith("```"):
            cleaned = re.sub(r"^```(?:json)?", "", raw_text.strip())
            cleaned = re.sub(r"```$", "", cleaned.strip())
            return cleaned.strip()
        return raw_text.strip()

    # -------------------------------------------------------------------------
    # Chunkning
    # -------------------------------------------------------------------------

    def _chunk_elements(self, elements, max_tokens_per_call=MAX_TOKEN_PER_CALL):
        chunks = []
        current_chunk = []
        current_char_count = len(self.policy_prompt)

        for elem in elements:
            elem_text = json.dumps(elem, ensure_ascii=False)
            if current_chunk and current_char_count + len(elem_text) > max_tokens_per_call * 4:
                chunks.append(current_chunk)
                current_chunk = [elem]
                current_char_count = len(self.policy_prompt) + len(elem_text)
            else:
                current_chunk.append(elem)
                current_char_count += len(elem_text)

        if current_chunk:
            chunks.append(current_chunk)

        return chunks

    # -------------------------------------------------------------------------
    # Validering / normalisering
    # -------------------------------------------------------------------------

    def _postprocess_suggestions(self, suggestions: list[dict[str, Any]]) -> list[SuggestedChange]:
        structure_type = self.json_structured_document.get("type")
        validated: list[SuggestedChange] = []

        for idx, item in enumerate(suggestions):
            try:
                suggested = self._validate_single_suggestion(item, structure_type=structure_type)
            except ValueError as ex:
                self.logger.warning(f"Rejected suggestion at index {idx}: {ex}")
                continue

            if suggested.match_status == "rejected":
                self.logger.warning(
                    f"Rejected suggestion for unmatchable old-text: {self._safe_preview(item)}"
                )
                continue

            if suggested.match_status == "normalized" and not self.allow_normalized_matches:
                self.logger.warning(
                    f"Rejected normalized-only suggestion: {self._safe_preview(item)}"
                )
                continue

            minimized = self._minimize_and_filter_suggestion(suggested)
            if minimized is None:
                continue

            if not minimized.safe_to_apply:
                self.logger.warning(
                    f"Rejected unsafe suggestion: "
                    f"{minimized.element_type}:{minimized.element_id} "
                    f"reason={minimized.safety_reason}"
                )
                continue

            validated.append(minimized)

        validated = self._deduplicate_validated_suggestions(validated)
        validated = self._validate_and_clean_suggestions(validated)

        self.logger.info(f"Validated suggestions after safety-cleaning: {len(validated)}")
        return validated

    def _validate_single_suggestion(self, item: dict[str, Any], structure_type: str) -> SuggestedChange:
        if structure_type != "docx":
            raise ValueError(f"Unsupported structure type: {structure_type}")
        return self._validate_docx_suggestion(item)

    def _validate_docx_suggestion(self, item: dict[str, Any]) -> SuggestedChange:
        element_type = self._require_str(item, "type")
        element_id = self._require_str(item, "element_id")
        old = self._require_nonempty_str(item, "old")
        new = self._require_str(item, "new")
        motivation = self._optional_str(item.get("motivation"))
        footnote_id = self._optional_str(item.get("footnote_id"))

        if old == new:
            raise ValueError(f"Suggestion has identical old/new for element {element_id}")

        element = self._get_docx_element_by_id(element_id)
        if element is None:
            raise ValueError(f"Unknown element_id: {element_id}")

        if element.get("type") != element_type:
            raise ValueError(
                f"type mismatch for {element_id}: expected {element.get('type')}, got {element_type}"
            )

        if element_type == "footnote":
            expected_footnote_id = self._optional_str(element.get("footnote_id"))
            if expected_footnote_id != footnote_id:
                raise ValueError(
                    f"footnote_id mismatch for {element_id}: expected {expected_footnote_id}, got {footnote_id}"
                )

        element_text = element.get("text", "") or ""
        match_status = self._match_old_against_text(old, element_text)

        if self.strict_validation and match_status == "rejected":
            raise ValueError(f"'old' text not found in source element {element_id}")

        return SuggestedChange(
            element_type=element_type,
            element_id=element_id,
            footnote_id=footnote_id,
            old=old,
            new=new,
            motivation=motivation,
            match_status=match_status,
        )

    # -------------------------------------------------------------------------
    # Minimering av ändringsspann
    # -------------------------------------------------------------------------

    def _minimize_and_filter_suggestion(self, suggestion: SuggestedChange) -> Optional[SuggestedChange]:
        minimized = self._minimize_change_span(suggestion)
        if minimized is None:
            return None

        if minimized.old.strip() == "" and minimized.new.strip() == "":
            self.logger.info(
                f"Ignoring whitespace-only suggestion for "
                f"{minimized.element_type}:{minimized.element_id}"
            )
            return None

        element_text = self._get_element_text_for_suggestion(minimized)
        stabilized = self._stabilize_span_with_element_context(minimized, element_text)

        if not stabilized.safe_to_apply:
            self.logger.warning(
                f"Marked suggestion unsafe [{stabilized.element_type}:{stabilized.element_id}] "
                f"{stabilized.safety_reason}: old={stabilized.old!r}, new={stabilized.new!r}"
            )

        return stabilized

    def _minimize_change_span(self, suggestion: SuggestedChange) -> Optional[SuggestedChange]:
        """
        Reducerar old/new till minsta faktiska diff.
        Konservativ hållning:
        - rena trailing-whitespace-ändringar ignoreras
        - om minimeringen skulle ge old == "" expanderas spannet till ett
          närliggande ord-/frasspann i stället för att skapa ren insertion
        """
        old = suggestion.old
        new = suggestion.new

        if old == new:
            return None

        # 1. Ignorera rena trailing-whitespace-fall helt
        if old.rstrip() == new.rstrip() and old != new:
            trailing_old = old[len(old.rstrip()):]
            trailing_new = new[len(new.rstrip()):]

            if trailing_old and trailing_old.isspace() and trailing_new == "":
                self.logger.info(
                    f"Ignoring trailing-whitespace-only suggestion for "
                    f"{suggestion.element_type}:{suggestion.element_id}"
                )
                return None

        # 2. Trimma gemensam prefix
        prefix_len = 0
        max_prefix = min(len(old), len(new))
        while prefix_len < max_prefix and old[prefix_len] == new[prefix_len]:
            prefix_len += 1

        old_rem = old[prefix_len:]
        new_rem = new[prefix_len:]

        # 3. Trimma gemensam suffix
        suffix_len = 0
        max_suffix = min(len(old_rem), len(new_rem))
        while (
            suffix_len < max_suffix
            and old_rem[-(suffix_len + 1)] == new_rem[-(suffix_len + 1)]
        ):
            suffix_len += 1

        if suffix_len > 0:
            old_rem = old_rem[:-suffix_len]
            new_rem = new_rem[:-suffix_len]

        if old_rem == new_rem:
            return None

        # 4. Konservativ fallback: om old blir tomt, expandera till ord-/frasspann
        if old_rem == "":
            expanded = self._expand_empty_old_minimization(old, new)
            if expanded is None:
                return suggestion
            old_rem, new_rem = expanded

        # 5. Säkerhetsnät: old måste finnas
        if old_rem == "":
            return suggestion

        return SuggestedChange(
            element_type=suggestion.element_type,
            element_id=suggestion.element_id,
            footnote_id=suggestion.footnote_id,
            old=old_rem,
            new=new_rem,
            motivation=suggestion.motivation,
            match_status=suggestion.match_status,
        )

    def _expand_empty_old_minimization(self, old: str, new: str) -> Optional[tuple[str, str]]:
        """
        Om teckenmässig minimering skulle ge old == "", expandera till ett
        stabilt ord-/frasspann. Exempel:
            old="... senare"
            new="... senare i processen"
        -> ("senare", "senare i processen")
        """
        start = 0
        max_len = min(len(old), len(new))
        while start < max_len and old[start] == new[start]:
            start += 1

        left = self._scan_left_to_word_boundary(old, start)

        old_tail = old[left:]
        new_tail = new[left:]

        suffix_len = 0
        max_suffix = min(len(old_tail), len(new_tail))
        while (
            suffix_len < max_suffix
            and old_tail[-(suffix_len + 1)] == new_tail[-(suffix_len + 1)]
        ):
            suffix_len += 1

        if suffix_len > 0:
            old_tail = old_tail[:-suffix_len]
            new_tail = new_tail[:-suffix_len]

        if old_tail == "":
            return None

        return old_tail, new_tail

    def _scan_left_to_word_boundary(self, text: str, idx: int) -> int:
        """
        Flyttar vänster till början av ordet som föregår skillnaden.
        Om vi står mitt i eller precis efter ett ord används ordstart som ankare.
        """
        if idx <= 0:
            return 0

        i = idx

        while i > 0 and not self._is_word_char(text[i - 1]):
            i -= 1

        while i > 0 and self._is_word_char(text[i - 1]):
            i -= 1

        return i

    def _find_exact_span_in_element_text(self, old_text: str, element_text: str) -> Optional[tuple[int, int]]:
        idx = element_text.find(old_text)
        if idx == -1:
            return None
        return idx, idx + len(old_text)

    def _is_word_char(self, ch: str) -> bool:
        return bool(self.WORD_CHAR_RE.match(ch))

    def _expand_span_to_word_boundaries(
        self,
        element_text: str,
        start: int,
        end: int,
    ) -> tuple[int, int]:
        """
        Expanderar ett spann till närmaste ordgränser.
        Konservativt: bara bokstavs-/siffergränser, inte fraser.
        """
        new_start = start
        new_end = end

        while new_start > 0 and self._is_word_char(element_text[new_start - 1]):
            new_start -= 1

        while new_end < len(element_text) and self._is_word_char(element_text[new_end]):
            new_end += 1

        return new_start, new_end

    def _looks_like_fragment_boundary(
        self,
        element_text: str,
        start: int,
        end: int,
    ) -> bool:
        """
        Returnerar True om spannet sannolikt börjar/slutar mitt i ord.
        """
        left_midword = (
            start > 0
            and start < len(element_text)
            and self._is_word_char(element_text[start - 1])
            and self._is_word_char(element_text[start])
        )
        right_midword = (
            end > 0
            and end < len(element_text)
            and self._is_word_char(element_text[end - 1])
            and self._is_word_char(element_text[end])
        )

        return left_midword or right_midword

    def _looks_like_truncated_text(self, text: str) -> bool:
        stripped = text.strip()
        if not stripped:
            return False

        if not any(ch.isalpha() for ch in stripped):
            return False

        common_short_words = {
            "de", "dem", "om", "av", "en", "ett", "är", "vi", "i", "på",
            "att", "för", "med", "och", "det", "den", "som", "har"
        }
        if len(stripped) <= 3 and stripped.lower() not in common_short_words:
            return True

        if stripped.isalpha():
            vowels = set("aeiouyåäöAEIOUYÅÄÖ")
            vowel_count = sum(1 for ch in stripped if ch in vowels)

            if len(stripped) >= 5 and vowel_count <= 1:
                return True

            if len(stripped) <= 4 and stripped.lower() not in common_short_words:
                return True

        if re.search(r"[A-Za-zÅÄÖåäö][^\w\s-]+[A-Za-zÅÄÖåäö]", stripped):
            return True

        return False

    def _causes_obvious_duplication(
        self,
        element_text: str,
        start: int,
        end: int,
        replacement: str,
    ) -> bool:
        before = element_text[:start]
        after = element_text[end:]
        candidate = before + replacement + after

        words = candidate.split()

        for i in range(len(words) - 1):
            if words[i] == words[i + 1] and len(words[i]) > 2:
                return True

        left_char = before[-1] if before else ""
        first_new = replacement[0] if replacement else ""
        last_new = replacement[-1] if replacement else ""
        right_char = after[0] if after else ""

        if replacement:
            if left_char and first_new and self._is_word_char(left_char) and self._is_word_char(first_new):
                return True
            if last_new and right_char and self._is_word_char(last_new) and self._is_word_char(right_char):
                return True

        return False

    def _stabilize_span_with_element_context(
        self,
        suggestion: SuggestedChange,
        element_text: str,
    ) -> SuggestedChange:
        """
        Försök reparera reducerade spann genom att:
        1. hitta exakt old i element_text
        2. expandera till ordgränser om old ser ut att ligga mitt i ord
        3. markera unsafe om spannet fortfarande ser trasigt ut
        """
        span = self._find_exact_span_in_element_text(suggestion.old, element_text)
        if span is None:
            return SuggestedChange(
                element_type=suggestion.element_type,
                element_id=suggestion.element_id,
                footnote_id=suggestion.footnote_id,
                old=suggestion.old,
                new=suggestion.new,
                motivation=suggestion.motivation,
                match_status=suggestion.match_status,
                safe_to_apply=False,
                safety_reason="old_not_found_exactly_in_element_text",
            )

        start, end = span
        old = suggestion.old
        new = suggestion.new

        if self._looks_like_fragment_boundary(element_text, start, end):
            expanded_start, expanded_end = self._expand_span_to_word_boundaries(element_text, start, end)
            expanded_old = element_text[expanded_start:expanded_end]

            prefix_len = start - expanded_start
            suffix_len = expanded_end - end

            expanded_new = new
            if prefix_len > 0:
                expanded_new = expanded_old[:prefix_len] + expanded_new
            if suffix_len > 0:
                expanded_new = expanded_new + expanded_old[len(expanded_old) - suffix_len:]

            old = expanded_old
            new = expanded_new
            start, end = expanded_start, expanded_end

        safe = True
        reason = None

        if self._looks_like_truncated_text(old):
            safe = False
            reason = "old_looks_truncated"
        elif new.strip() and self._looks_like_truncated_text(new):
            safe = False
            reason = "new_looks_truncated"
        elif self._causes_obvious_duplication(element_text, start, end, new):
            safe = False
            reason = "replacement_causes_obvious_duplication"

        return SuggestedChange(
            element_type=suggestion.element_type,
            element_id=suggestion.element_id,
            footnote_id=suggestion.footnote_id,
            old=old,
            new=new,
            motivation=suggestion.motivation,
            match_status=suggestion.match_status,
            safe_to_apply=safe,
            safety_reason=reason,
        )

    # -------------------------------------------------------------------------
    # Generell säkerhetsrensning före planner
    # -------------------------------------------------------------------------

    def _validate_and_clean_suggestions(
        self,
        suggestions: list[SuggestedChange],
    ) -> list[SuggestedChange]:
        suggestions = self._deduplicate_conflicting_same_old(suggestions)

        cleaned: list[SuggestedChange] = []

        for s in suggestions:
            verdict = self._assess_suggestion_safety(s)

            if verdict["reject"]:
                self.logger.warning(
                    f"Rejected unsafe suggestion "
                    f"[{s.element_type}:{s.element_id}] "
                    f"reason={verdict['reason']} "
                    f"old={s.old!r} new={s.new!r}"
                )
                continue

            cleaned.append(
                SuggestedChange(
                    element_type=s.element_type,
                    element_id=s.element_id,
                    footnote_id=s.footnote_id,
                    old=s.old,
                    new=s.new,
                    motivation=s.motivation,
                    match_status=s.match_status,
                    safe_to_apply=True,
                    safety_reason=None,
                )
            )

        return cleaned

    def _deduplicate_conflicting_same_old(
        self,
        suggestions: list[SuggestedChange],
    ) -> list[SuggestedChange]:
        grouped: dict[tuple, list[SuggestedChange]] = defaultdict(list)

        for s in suggestions:
            key = (s.element_type, s.element_id, s.footnote_id, s.old)
            grouped[key].append(s)

        resolved: list[SuggestedChange] = []

        for _, group in grouped.items():
            if len(group) == 1:
                resolved.append(group[0])
                continue

            best = self._choose_best_conflicting_suggestion(group)

            self.logger.warning(
                f"Resolved conflicting suggestions for "
                f"{best.element_type}:{best.element_id} old={best.old!r} "
                f"by keeping new={best.new!r} and dropping {len(group) - 1} alternatives"
            )

            resolved.append(best)

        return resolved

    def _choose_best_conflicting_suggestion(
        self,
        group: list[SuggestedChange],
    ) -> SuggestedChange:
        scored = []

        for s in group:
            score = 0.0

            if not self._looks_like_corrupted_text(s.new):
                score += 3.0

            score += self._similarity_ratio(s.old, s.new)

            old_nonword = len(re.findall(r"[^\w\s]", s.old))
            new_nonword = len(re.findall(r"[^\w\s]", s.new))
            if new_nonword <= old_nonword + 1:
                score += 0.5

            length_delta = abs(len(s.new.strip()) - len(s.old.strip()))
            score -= min(length_delta / 50.0, 1.0)

            scored.append((score, s))

        scored.sort(key=lambda x: x[0], reverse=True)
        return scored[0][1]

    def _assess_suggestion_safety(self, s: SuggestedChange) -> dict:
        old = (s.old or "").strip()
        new = (s.new or "").strip()

        if not old:
            return {"reject": True, "reason": "empty_old"}

        if old == new:
            return {"reject": True, "reason": "old_equals_new"}

        if old.strip() == new.strip() and old != new:
            return {"reject": True, "reason": "whitespace_only_change"}

        if self._looks_like_corrupted_text(new):
            return {"reject": True, "reason": "corrupted_new_text"}

        if self._looks_like_corrupted_text(old):
            return {"reject": True, "reason": "corrupted_old_span"}

        if self._too_low_overlap(old, new, s.element_type):
            return {"reject": True, "reason": "low_similarity"}

        if self._self_overlap_or_merge_risk(old, new):
            return {"reject": True, "reason": "self_overlap_or_merge_risk"}

        if self._looks_like_bad_casing_change(old, new):
            return {"reject": True, "reason": "bad_casing_change"}

        if self._looks_like_spelling_degradation(old, new):
            return {"reject": True, "reason": "spelling_degradation"}

        if s.element_type in {"textbox", "footnote"} and self._similarity_ratio(old, new) < 0.45:
            return {"reject": True, "reason": "low_similarity_sensitive_element"}

        return {"reject": False, "reason": None}

    def _similarity_ratio(self, old: str, new: str) -> float:
        return SequenceMatcher(None, old, new).ratio()

    def _too_low_overlap(self, old: str, new: str, element_type: str) -> bool:
        ratio = self._similarity_ratio(old, new)
        short_local = max(len(old), len(new)) <= 25

        if short_local and ratio < 0.35:
            return True

        if element_type in {"table_cell", "textbox", "footnote"} and ratio < 0.30:
            return True

        return False

    def _looks_like_corrupted_text(self, text: str) -> bool:
        t = (text or "").strip()
        if not t:
            return False

        if re.search(r"[a-zåäö]{2,}[A-ZÅÄÖ][a-zåäö]+", t):
            return True

        if t.isalpha() and len(t) >= 6:
            vowels = set("aeiouyåäöAEIOUYÅÄÖ")
            vowel_count = sum(1 for ch in t if ch in vowels)
            if vowel_count <= 1:
                return True

        if re.search(r"[A-Za-zÅÄÖåäö][^\w\s-]+[A-Za-zÅÄÖåäö]", t):
            return True

        if re.search(r"(.)\1{3,}", t):
            return True

        return False

    def _self_overlap_or_merge_risk(self, old: str, new: str) -> bool:
        old_s = old.strip()
        new_s = new.strip()

        if not old_s or not new_s:
            return False

        if old_s.isalpha() and new_s.isalpha():
            if old_s in new_s and old_s != new_s:
                return True
            if new_s in old_s and new_s != old_s and len(new_s) <= len(old_s) - 2:
                return True

        if old_s.isalpha() and new_s.isalpha() and len(old_s) >= 6 and len(new_s) >= 4:
            if new_s.startswith(old_s[-4:]) or new_s.endswith(old_s[:4]):
                return True

        return False

    def _looks_like_bad_casing_change(self, old: str, new: str) -> bool:
        if old.lower() != new.lower():
            return False

        if old and new and old[0].islower() and new[0].isupper() and old[1:] == new[1:]:
            return False

        return old != new

    def _looks_like_spelling_degradation(self, old: str, new: str) -> bool:
        old_s = old.strip()
        new_s = new.strip()

        if len(old_s) < 5 or len(new_s) < 5:
            return False

        ratio = self._similarity_ratio(old_s, new_s)

        if ratio > 0.80 and len(new_s) < len(old_s):
            old_vowels = sum(1 for ch in old_s if ch.lower() in "aeiouyåäö")
            new_vowels = sum(1 for ch in new_s if ch.lower() in "aeiouyåäö")
            if new_vowels < old_vowels:
                return True

        return False

    # -------------------------------------------------------------------------
    # Matchning
    # -------------------------------------------------------------------------

    def _match_old_against_text(self, old: str, source_text: str) -> Literal["exact", "normalized", "rejected"]:
        if old in source_text:
            return "exact"

        normalized_old = self._normalize_text(old)
        normalized_source = self._normalize_text(source_text)

        if normalized_old and normalized_old in normalized_source:
            return "normalized"

        return "rejected"

    @staticmethod
    def _normalize_text(text: str) -> str:
        text = text.replace("\u00a0", " ")
        text = re.sub(r"\s+", " ", text)
        return text.strip()

    # -------------------------------------------------------------------------
    # Hjälpare för struktur
    # -------------------------------------------------------------------------

    def _get_docx_element_by_id(self, element_id: str) -> Optional[dict[str, Any]]:
        elements = self.json_structured_document.get("elements", [])
        for element in elements:
            if element.get("element_id") == element_id:
                return element
        return None

    def _get_element_text_for_suggestion(self, suggestion: SuggestedChange) -> str:
        element = self._get_docx_element_by_id(suggestion.element_id)
        if element is None:
            return ""
        return element.get("text", "") or ""

    # -------------------------------------------------------------------------
    # Deduplicering
    # -------------------------------------------------------------------------

    def _deduplicate_raw_suggestions(self, suggestions: list[dict[str, Any]]) -> list[dict[str, Any]]:
        seen = set()
        result = []

        for item in suggestions:
            key = self._raw_suggestion_key(item)
            if key in seen:
                continue
            seen.add(key)
            result.append(item)

        return result

    def _deduplicate_validated_suggestions(
        self, suggestions: list[SuggestedChange]
    ) -> list[SuggestedChange]:
        seen = set()
        result = []

        for item in suggestions:
            key = (
                item.element_type,
                item.element_id,
                item.footnote_id,
                item.old,
                item.new,
            )
            if key in seen:
                continue
            seen.add(key)
            result.append(item)

        return result

    def _raw_suggestion_key(self, item: dict[str, Any]):
        return (
            item.get("type"),
            item.get("element_id"),
            item.get("footnote_id"),
            item.get("old"),
            item.get("new"),
        )

    # -------------------------------------------------------------------------
    # Serialisering
    # -------------------------------------------------------------------------

    def _suggested_change_to_output_dict(self, s: SuggestedChange) -> dict[str, Any]:
        payload = {
            "type": s.element_type,
            "element_id": s.element_id,
            "old": s.old,
            "new": s.new,
        }
        if s.motivation:
            payload["motivation"] = s.motivation
        if s.footnote_id is not None:
            payload["footnote_id"] = s.footnote_id
        if not s.safe_to_apply:
            payload["safe_to_apply"] = False
            payload["safety_reason"] = s.safety_reason
        return payload

    # -------------------------------------------------------------------------
    # Typkontroller
    # -------------------------------------------------------------------------

    def _require_str(self, item: dict[str, Any], key: str) -> str:
        value = item.get(key)
        if not isinstance(value, str):
            raise ValueError(f"Missing or invalid string field: {key}")
        return value

    def _require_nonempty_str(self, item: dict[str, Any], key: str) -> str:
        value = self._require_str(item, key)
        if not value.strip():
            raise ValueError(f"Field {key} must not be empty")
        return value

    def _optional_str(self, value: Any) -> Optional[str]:
        if value is None:
            return None
        if isinstance(value, str):
            return value
        return str(value)

    def _safe_preview(self, item: dict[str, Any], max_len: int = 180) -> str:
        text = json.dumps(item, ensure_ascii=False)
        if len(text) <= max_len:
            return text
        return text[:max_len] + "..."


def main():
    num_args = len(sys.argv)
    if num_args < 6 or num_args > 7:
        print(
            f"Usage: python {os.path.basename(__file__)} "
            "<structure_file.json> <api_key> <model> <temperature> "
            "<prompt_policy_file> <custom_addition_file>"
        )
        sys.exit(1)

    try:
        filepath = sys.argv[1]
        api_key = sys.argv[2]
        model = sys.argv[3]
        temperature = float(sys.argv[4])
        policy_path = sys.argv[5]
        custom_path = sys.argv[6] if num_args > 6 else None
    except Exception as ex:
        print(f"Error on input arguments: {str(ex)}")
        sys.exit(1)

    with open(policy_path, encoding="utf-8") as f:
        base_prompt = f.read().strip()

    full_prompt = base_prompt
    if custom_path and os.path.exists(custom_path):
        with open(custom_path, encoding="utf-8") as f:
            custom = f.read().strip()
        full_prompt += "\n\n" + custom

    logger = logging.getLogger("ai-test")
    logger.setLevel(logging.INFO)
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
    logger.handlers.clear()
    logger.addHandler(handler)

    ai = JBGLangImprovSuggestorAI(
        api_key=api_key,
        model=model,
        prompt_policy=full_prompt,
        temperature=temperature,
        logger=logger,
        strict_validation=True,
        allow_normalized_matches=True,
    )
    ai.load_structure(filepath)
    ai.suggest_changes_token_aware_batching()
    suggestions_path = ai.save_as_json()
    report_path = ai.save_filter_report()

    print(f"Saved validated suggestions to: {suggestions_path}")
    print(f"Saved filter report to: {report_path}")


if __name__ == "__main__":
    main()