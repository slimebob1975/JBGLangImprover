"""
Version och commit för den kod som körs, för visning under rubriken, i
körningens sammanfattning och i avsnittet "Om klarspråkningen".

- Versionsnumret läses från filen VERSION i repots rot.
- I den driftsatta appen läses commit, gren och miljö från build_info.json,
  som byggsteget skriver (scripts/write_build_info.py).
- Lokalt, utan build_info.json, frågas git om commit; finns inte git heller
  anges bara versionen och miljön "lokal".
"""

import json
import subprocess
from functools import lru_cache
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]


def _git(*args) -> str:
    try:
        result = subprocess.run(["git", *args], cwd=ROOT, capture_output=True, text=True, timeout=2)
        return result.stdout.strip() if result.returncode == 0 else ""
    except (OSError, subprocess.SubprocessError):
        return ""


def read_build_info(root: Path = ROOT) -> dict:
    version_file = root / "VERSION"
    version = version_file.read_text(encoding="utf-8").strip() if version_file.exists() else ""
    info = {"version": version or "okänd", "commit": "", "branch": "",
            "environment": "lokal", "built_at": "", "commit_url": ""}

    build_file = root / "build_info.json"
    if build_file.exists():
        try:
            data = json.loads(build_file.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            data = {}
        info.update({
            "commit": data.get("commit", ""),
            "branch": data.get("branch", ""),
            "environment": data.get("environment") or "okänd",
            "built_at": data.get("built_at", ""),
        })
        if data.get("repository_url") and info["commit"]:
            info["commit_url"] = f"{data['repository_url']}/commit/{info['commit']}"
    elif (root / ".git").exists():
        info["commit"] = _git("rev-parse", "HEAD")
        info["branch"] = _git("rev-parse", "--abbrev-ref", "HEAD")

    info["commit_short"] = info["commit"][:7]
    label = f"Version {info['version']}"
    if info["commit_short"]:
        label += f" · {info['commit_short']}"
    info["label"] = f"{label} ({info['environment']})"
    return info


@lru_cache(maxsize=1)
def get_build_info() -> dict:
    """Läses en gång per process; versionen ändras bara vid ny driftsättning."""
    return read_build_info()
