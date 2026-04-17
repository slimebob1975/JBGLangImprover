import json
import openai
import sys
import os
import re
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
        new = s.new or ""

        for pattern in self.TODO_PATTERNS:
            if re.search(pattern, old, flags=re.IGNORECASE):
                issues.append(SuggestionIssue(
                    severity="error" if self.reject_todo_transformations else "warning",
                    code="internal_note_rewrite",
                    message="Förslaget verkar skriva om en intern arbetsnotering eller TODO-markör."
                ))
                break

        if not old.strip() or not new.strip():
            issues.append(SuggestionIssue(
                severity="error",
                code="empty_side",
                message="Old eller new är tom efter trimning."
            ))

    def _check_length_growth(self, s: SuggestedChange, issues: list[SuggestionIssue]):
        old_len = len(s.old.strip())
        new_len = len(s.new.strip())

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
        new_sentences = self._sentence_count(s.new)

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
        new = s.new.lower()

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
        new_has_ascii_quotes = '"' in s.new or "'" in s.new

        if old_has_typographic_quotes and new_has_ascii_quotes:
            issues.append(SuggestionIssue(
                severity="warning",
                code="quote_style_shift",
                message="Förslaget byter från typografiska till raka citattecken."
            ))

    def _check_element_specific_risk(self, s: SuggestedChange, issues: list[SuggestionIssue]):
        element_type = s.element_type or ""
        old_len = len(s.old.strip())
        new_len = len(s.new.strip())

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
    """

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

            validated.append(suggested)

        return self._deduplicate_validated_suggestions(validated)

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
            key =(
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

    def _require_int(self, item: dict[str, Any], key: str) -> int:
        value = item.get(key)
        if not isinstance(value, int):
            raise ValueError(f"Missing or invalid int field: {key}")
        return value

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