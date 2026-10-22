"""
Datahantering och lagring.

Utgångsläge: dokument, deras diagnostikfiler och utdrag ur dokumenten i
loggarna ska inte finnas kvar längre än nödvändigt.

- keep_debug_data(): miljövariabeln JBG_KEEP_DEBUG_DATA (av som standard).
  Påslagen behålls uppladdning och JSON-filer efter körningen och loggarna
  innehåller dokumenttext, som för felsökning i en utvecklingsmiljö.
- loggable_text(): text i loggarna, eller bara dess längd.
- run_file_paths(): uppladdningen och de diagnostikfiler en körning skapar.
- delete_files(): tar bort filer och tål att de redan saknas.
"""

import os
from typing import Iterable

DEBUG_ENV_VAR = "JBG_KEEP_DEBUG_DATA"

# Diagnostikfiler som orkestreraren skriver bredvid uppladdningen.
DIAGNOSTIC_SUFFIXES = (
    "_structure.json",
    "_suggestions.json",
    "_suggestion_filter_report.json",
    "_global_findings.json",
    "_run_summary.json",
)


def keep_debug_data() -> bool:
    return os.getenv(DEBUG_ENV_VAR, "").strip().lower() in {"1", "true", "yes", "on", "ja"}


def loggable_text(text) -> str:
    """Texten i citattecken vid felsökning, annars bara längden."""
    if keep_debug_data():
        return repr(text)
    return f"<{len(text or '')} tecken>"


def run_file_paths(input_path: str) -> list[str]:
    stem = os.path.splitext(input_path)[0]
    return [input_path] + [stem + suffix for suffix in DIAGNOSTIC_SUFFIXES]


def delete_files(paths: Iterable[str], logger) -> int:
    deleted = 0
    for path in paths:
        try:
            if os.path.isfile(path):
                os.remove(path)
                deleted += 1
        except OSError as ex:
            logger.warning(f"Could not delete {os.path.basename(path)}: {ex}")
    return deleted
