"""Frozen, deterministic revision probes; no external model or files required."""
import json
import subprocess
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "frozen_wordlevel_compare.py"
FIXTURES = ROOT / "tests" / "fixtures" / "wordlevel_frozen_v1.json"


class FrozenWordLevelComparisonTests(unittest.TestCase):
    def test_frozen_cases_and_integrity(self):
        cases = json.loads(FIXTURES.read_text(encoding="utf-8"))["cases"]
        proc = subprocess.run([sys.executable, str(SCRIPT), "--probe-only"],
                              cwd=ROOT, capture_output=True, text=True, check=True)
        results = json.loads(proc.stdout)
        self.assertEqual([c["id"] for c in cases], [r["id"] for r in results])
        for case, result in zip(cases, results):
            with self.subTest(case=case["id"]):
                self.assertTrue(result["invariants_passed"])
                if "expected_status" in case:
                    self.assertEqual(result["status"], case["expected_status"])
                if "expected_reason" in case:
                    self.assertEqual(result["fallback_reason"], case["expected_reason"])
                self.assertGreater(result["revisions"]["ins"], 0)
                self.assertGreater(result["revisions"]["del"], 0)


if __name__ == "__main__":
    unittest.main()
