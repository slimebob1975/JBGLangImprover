"""Frozen, deterministic revision probes; no external model or files required."""
import json
import os
import tempfile
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

    def test_opt_in_real_fallback_capture_and_replay(self):
        cases = json.loads(FIXTURES.read_text(encoding="utf-8"))["cases"]
        flagged = next(x for x in cases if x["id"] == "formatted_changed_boundary")
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "input.json"
            captured = Path(tmp) / "captured.json"
            source.write_text(json.dumps({"schema": 1, "cases": [flagged]}), encoding="utf-8")
            env = os.environ.copy()
            env["JBG_WORDLEVEL_CAPTURE_PATH"] = str(captured)
            proc = subprocess.run([sys.executable, str(SCRIPT), "--probe-only",
                                   "--fixture-file", str(source)], cwd=ROOT,
                                  env=env, capture_output=True, text=True, check=True)
            results = json.loads(proc.stdout)
            self.assertEqual(results[0]["fallback_reason"], "mixed_run_formatting")
            data = json.loads(captured.read_text(encoding="utf-8"))
            self.assertEqual(len(data["cases"]), 1)
            sample = data["cases"][0]
            self.assertEqual(sample["fallback_reason"], "mixed_run_formatting")
            self.assertEqual(sample["old"], flagged["old"])
            self.assertEqual(sample["new"], flagged["new"])
            self.assertIn("<w:p", sample["paragraph_xml"])
            # Exact pre-edit paragraph XML can be replayed without any AI call.
            rerun = subprocess.run([sys.executable, str(SCRIPT), "--probe-only",
                                    "--fixture-file", str(captured)], cwd=ROOT,
                                   capture_output=True, text=True, check=True)
            recaptured = json.loads(rerun.stdout)
            self.assertEqual(recaptured[0]["fallback_reason"], "mixed_run_formatting")
            self.assertTrue(recaptured[0]["invariants_passed"])


if __name__ == "__main__":
    unittest.main()
