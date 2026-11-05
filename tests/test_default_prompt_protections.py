"""Regressionstest för skyddade uppgifter i den redigerbara standardprompten."""
from pathlib import Path
import unittest

from app.main import load_prompt_parts


ROOT = Path(__file__).resolve().parents[1]


class DefaultPromptProtectionTests(unittest.TestCase):
    def test_protections_are_in_editable_default_prompt(self):
        editable, before, after = load_prompt_parts(str(ROOT / "policy" / "prompt_policy.md"))
        self.assertIn("### Skyddade namn, titlar och hänvisningar", editable)
        self.assertIn("### Mallinstruktioner och utkastmarkörer", editable)
        self.assertIn('"arbetslöshetskassor" med "a-kassor"', editable)
        self.assertIn("har företräde framför övriga", editable)
        self.assertIn("Lämna källförteckningar", editable)
        self.assertIn("Du får förbättra den omgivande löptexten", editable)
        self.assertIn("Börja skriva här", editable)
        self.assertNotIn("### Skyddade namn, titlar och hänvisningar", before + after)

    def test_locked_sections_unchanged(self):
        content = (ROOT / "policy" / "prompt_policy.md").read_text(encoding="utf-8")
        self.assertEqual(content.count("<!-- START_LOCKED -->"), 2)
        self.assertEqual(content.count("<!-- END_LOCKED -->"), 2)
        self.assertEqual(content.count("<!-- START_EDITABLE -->"), 1)
        self.assertEqual(content.count("<!-- END_EDITABLE -->"), 1)


if __name__ == "__main__":
    unittest.main()
