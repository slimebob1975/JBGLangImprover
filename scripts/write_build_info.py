"""
Skriver build_info.json i byggsteget i GitHub Actions, innan appen packas.

Den driftsatta appen kan inte fråga git om sin version, eftersom arbetsflödet
packar den med "zip release.zip ./* -r", som utelämnar mappen .git. Därför
sparas commit, gren, körningsnummer och byggtid här och läses av appen.

Användning i arbetsflödet:  python scripts/write_build_info.py dev
"""

import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path


def build_info(environment: str, env=os.environ) -> dict:
    repository = env.get("GITHUB_REPOSITORY", "")
    server = env.get("GITHUB_SERVER_URL", "https://github.com")
    return {
        "commit": env.get("GITHUB_SHA", ""),
        "branch": env.get("GITHUB_REF_NAME", ""),
        "run_number": env.get("GITHUB_RUN_NUMBER", ""),
        "repository_url": f"{server}/{repository}" if repository else "",
        "environment": environment,
        "built_at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
    }


def main(argv=None) -> Path:
    argv = sys.argv[1:] if argv is None else argv
    environment = argv[0] if argv else os.environ.get("GITHUB_REF_NAME", "")
    path = Path(__file__).resolve().parents[1] / "build_info.json"
    path.write_text(json.dumps(build_info(environment), indent=2) + "\n", encoding="utf-8")
    print(f"Wrote {path}")
    return path


if __name__ == "__main__":
    main()
