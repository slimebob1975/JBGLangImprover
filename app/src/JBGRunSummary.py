"""
Samlad sammanfattning av en körning.

RunSummary är datakällan för diagnostikfilen <fil>_run_summary.json och,
i en senare fas, för avsnittet "Om klarspråkningen" i resultatdokumentet.
"""

import json
from dataclasses import dataclass, field, asdict
from datetime import datetime, timezone
from typing import Any, Optional


@dataclass
class SuggestionCounts:
    raw: int = 0                 # objekt som modellen returnerade (efter dubblettrensning)
    validated: int = 0           # godkända av schema- och matchningskontroll
    accepted: int = 0            # godkända av kvalitetsfiltret
    planned: int = 0             # ändringsplaner som kunde byggas
    applied: int = 0             # faktiskt renderade i dokumentet
    failed: int = 0              # planer som inte kunde renderas
    comments_applied: int = 0


@dataclass
class RunSummary:
    input_filename: str
    model: str
    docx_mode: str
    temperature: Optional[float] = None
    include_motivations: bool = False
    include_about_section: bool = True
    compute_readability: bool = True
    global_review: bool = False
    prompt_customized: Optional[bool] = None   # None = okänt (t.ex. CLI-körning)
    started_at: str = field(default_factory=lambda: _now())
    finished_at: Optional[str] = None
    succeeded: bool = False
    error: Optional[str] = None

    local_suggestions: SuggestionCounts = field(default_factory=SuggestionCounts)
    usage: dict[str, Any] = field(default_factory=dict)
    readability: Optional[dict[str, Any]] = None
    about_section: Optional[dict[str, Any]] = None
    # Global granskning: parts, raw, accepted, rejected, by_category, errors,
    # comments_applied. None när den globala granskningen inte kördes.
    global_findings: Optional[dict[str, Any]] = None

    def finish(self, succeeded: bool, error: Optional[str] = None) -> None:
        self.finished_at = _now()
        self.succeeded = succeeded
        self.error = error

    def to_dict(self) -> dict:
        return asdict(self)

    def save(self, path: str) -> str:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, indent=2, ensure_ascii=False)
        return path


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")
