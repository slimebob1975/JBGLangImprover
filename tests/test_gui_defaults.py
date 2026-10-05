import inspect
import os
import unittest
from html.parser import HTMLParser
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

try:
    import fastapi  # noqa: F401
    HAS_FASTAPI = True
except ImportError:  # pragma: no cover - beror på miljön
    HAS_FASTAPI = False


class _Inputs(HTMLParser):
    def __init__(self):
        super().__init__()
        self.inputs = {}

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag == "input" and attrs.get("id"):
            self.inputs[attrs["id"]] = attrs


class GuiDefaultsTests(unittest.TestCase):
    """Standardvalen i formuläret är beslut; de ska inte ändras av misstag."""

    def setUp(self):
        parser = _Inputs()
        parser.feed((ROOT / "templates" / "index.html").read_text(encoding="utf-8"))
        self.inputs = parser.inputs

    def checked(self, element_id):
        return "checked" in self.inputs[element_id]

    def test_tracked_changes_is_the_default_result_mode(self):
        self.assertTrue(self.checked("trackedChanges"))
        self.assertFalse(self.checked("simpleMarking"))

    def test_optional_features_are_off_by_default(self):
        for element_id in ("globalReview", "computeLix", "includeAboutSection"):
            with self.subTest(element_id=element_id):
                self.assertFalse(self.checked(element_id))

    @unittest.skipUnless(HAS_FASTAPI, "fastapi saknas i miljön")
    def test_server_defaults_match_the_form(self):
        cwd = os.getcwd()
        os.chdir(ROOT)   # main.py använder relativa sökvägar från repots rot
        try:
            from app import main
        finally:
            os.chdir(cwd)
        defaults = {name: p.default.default for name, p in inspect.signature(main.upload_file).parameters.items()
                    if hasattr(p.default, "default")}
        self.assertEqual(defaults["docx_mode"], "tracked")
        self.assertEqual((defaults["global_review"], defaults["compute_lix"], defaults["include_about_section"]),
                         (False, False, False))


if __name__ == "__main__":
    unittest.main()
