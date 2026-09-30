"""
Samlar tokenanvändning per API-anrop och per fas (t.ex. "local", "global").

Trackern är trådsäker och fristående från OpenAI-klienten: den tar emot
responsobjektets `usage` (eller en dict med samma fält) och tål att fält saknas,
eftersom olika modeller och SDK-versioner rapporterar olika detaljnivå.
"""

import threading
from dataclasses import dataclass, field, asdict
from typing import Any, Optional


@dataclass
class UsageRecord:
    phase: str
    model: str
    prompt_tokens: int = 0
    completion_tokens: int = 0
    reasoning_tokens: int = 0
    cached_prompt_tokens: int = 0
    total_tokens: int = 0
    succeeded: bool = True


@dataclass
class UsageTotals:
    calls: int = 0
    failed_calls: int = 0
    prompt_tokens: int = 0
    completion_tokens: int = 0
    reasoning_tokens: int = 0
    cached_prompt_tokens: int = 0
    total_tokens: int = 0

    def add(self, record: UsageRecord) -> None:
        self.calls += 1
        if not record.succeeded:
            self.failed_calls += 1
        self.prompt_tokens += record.prompt_tokens
        self.completion_tokens += record.completion_tokens
        self.reasoning_tokens += record.reasoning_tokens
        self.cached_prompt_tokens += record.cached_prompt_tokens
        self.total_tokens += record.total_tokens


@dataclass
class UsageTracker:
    records: list[UsageRecord] = field(default_factory=list)

    def __post_init__(self):
        self._lock = threading.Lock()

    # ------------------------------------------------------------------
    # Registrering
    # ------------------------------------------------------------------

    def record(self, phase: str, model: str, usage: Any) -> UsageRecord:
        """Registrera ett lyckat anrop. `usage` kan vara None, en dict eller ett SDK-objekt."""
        rec = UsageRecord(phase=phase, model=model, **self._extract(usage))
        with self._lock:
            self.records.append(rec)
        return rec

    def record_failure(self, phase: str, model: str) -> UsageRecord:
        """Registrera ett anrop som misslyckades innan någon usage kunde läsas."""
        rec = UsageRecord(phase=phase, model=model, succeeded=False)
        with self._lock:
            self.records.append(rec)
        return rec

    # ------------------------------------------------------------------
    # Summering
    # ------------------------------------------------------------------

    def totals(self, phase: Optional[str] = None) -> UsageTotals:
        totals = UsageTotals()
        with self._lock:
            for rec in self.records:
                if phase is None or rec.phase == phase:
                    totals.add(rec)
        return totals

    def phases(self) -> list[str]:
        with self._lock:
            seen: list[str] = []
            for rec in self.records:
                if rec.phase not in seen:
                    seen.append(rec.phase)
            return seen

    def to_dict(self) -> dict:
        return {
            "total": asdict(self.totals()),
            "by_phase": {phase: asdict(self.totals(phase)) for phase in self.phases()},
        }

    # ------------------------------------------------------------------
    # Hjälpare
    # ------------------------------------------------------------------

    @classmethod
    def _extract(cls, usage: Any) -> dict:
        prompt = cls._int(cls._get(usage, "prompt_tokens", "input_tokens"))
        completion = cls._int(cls._get(usage, "completion_tokens", "output_tokens"))
        total = cls._int(cls._get(usage, "total_tokens")) or (prompt + completion)

        completion_details = cls._get(usage, "completion_tokens_details", "output_tokens_details")
        prompt_details = cls._get(usage, "prompt_tokens_details", "input_tokens_details")

        return {
            "prompt_tokens": prompt,
            "completion_tokens": completion,
            "total_tokens": total,
            "reasoning_tokens": cls._int(cls._get(completion_details, "reasoning_tokens")),
            "cached_prompt_tokens": cls._int(cls._get(prompt_details, "cached_tokens")),
        }

    @staticmethod
    def _get(obj: Any, *names: str) -> Any:
        if obj is None:
            return None
        for name in names:
            value = obj.get(name) if isinstance(obj, dict) else getattr(obj, name, None)
            if value is not None:
                return value
        return None

    @staticmethod
    def _int(value: Any) -> int:
        try:
            return int(value or 0)
        except (TypeError, ValueError):
            return 0
