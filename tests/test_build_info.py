import json
import os
import re
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from app.src.JBGAboutSectionRenderer import build_about_section_blocks
from app.src.JBGBuildInfo import read_build_info

ROOT = Path(__file__).resolve().parents[1]

try:
    from fastapi.testclient import TestClient
    HAS_FASTAPI = True
except ImportError:  # pragma: no cover - beror på miljön
    HAS_FASTAPI = False


class BuildInfoTests(unittest.TestCase):
    def test_deployed_app_reads_build_info(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "VERSION").write_text("1.4.0\n", encoding="utf-8")
            (root / "build_info.json").write_text(json.dumps({
                "commit": "a1b2c3d4e5f6", "branch": "dev", "environment": "dev",
                "repository_url": "https://github.com/iaf/JBGLangImprover",
            }), encoding="utf-8")
            info = read_build_info(root)
        self.assertEqual(info["label"], "Version 1.4.0 · a1b2c3d (dev)")
        self.assertEqual(info["commit_url"], "https://github.com/iaf/JBGLangImprover/commit/a1b2c3d4e5f6")

    def test_without_build_info_or_git_only_version_and_local(self):
        with tempfile.TemporaryDirectory() as tmp:
            (Path(tmp) / "VERSION").write_text("1.0.0", encoding="utf-8")
            info = read_build_info(Path(tmp))
        self.assertEqual(info["label"], "Version 1.0.0 (lokal)")
        self.assertEqual(info["commit_url"], "")

    def test_repository_has_a_version_file(self):
        self.assertRegex((ROOT / "VERSION").read_text(encoding="utf-8").strip(), r"^\d+\.\d+\.\d+$")


class BuildScriptAndWorkflowTests(unittest.TestCase):
    def test_build_script_records_commit_branch_and_environment(self):
        import importlib.util
        spec = importlib.util.spec_from_file_location("write_build_info", ROOT / "scripts" / "write_build_info.py")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        info = module.build_info("prod", env={
            "GITHUB_SHA": "f00dbabe1234", "GITHUB_REF_NAME": "prod", "GITHUB_RUN_NUMBER": "42",
            "GITHUB_SERVER_URL": "https://github.com", "GITHUB_REPOSITORY": "iaf/JBGLangImprover",
        })
        self.assertEqual((info["commit"], info["branch"], info["run_number"], info["environment"]),
                         ("f00dbabe1234", "prod", "42", "prod"))
        self.assertEqual(info["repository_url"], "https://github.com/iaf/JBGLangImprover")

    def test_both_workflows_write_build_info_before_packaging(self):
        for env in ("dev", "prod"):
            with self.subTest(env=env):
                text = (ROOT / ".github" / "workflows" / f"azure-deploy-langimprover-app_{env}.yml").read_text(encoding="utf-8")
                step = text.index(f"python scripts/write_build_info.py {env}")
                self.assertLess(step, text.index("zip release.zip"))


class VersionInResultsTests(unittest.TestCase):
    def test_about_section_names_the_program_version(self):
        summary = {"model": "m", "docx_mode": "tracked", "app_version": "1.4.0",
                   "app_commit": "a1b2c3d4e5f6", "app_environment": "prod",
                   "usage": {"total": {}}, "local_suggestions": {}}
        texts = [b.text for b in build_about_section_blocks(summary, tracked=True)]
        self.assertIn("Programversion: 1.4.0 (a1b2c3d, prod)", texts)

    def test_every_run_summary_records_the_version(self):
        import logging
        from docx import Document
        from app.src.JBGLanguageImprover import JBGLanguageImprover
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "r.docx"
            Document().save(path)
            logger = logging.getLogger("version-summary"); logger.propagate = False
            improver = JBGLanguageImprover(input_path=str(path), api_key="k", model="m", prompt_policy="p",
                                           temperature=1, include_motivations=False, logger=logger)
        summary = improver.run_summary
        self.assertEqual(summary.app_version, (ROOT / "VERSION").read_text(encoding="utf-8").strip())
        self.assertTrue(summary.app_environment)

    @unittest.skipUnless(HAS_FASTAPI, "fastapi saknas i miljön")
    def test_config_returns_the_version(self):
        cwd = os.getcwd()
        os.chdir(ROOT)
        try:
            from app import main
            version = TestClient(main.app).get("/config").json()["version"]
        finally:
            os.chdir(cwd)
        self.assertEqual(version["version"], (ROOT / "VERSION").read_text(encoding="utf-8").strip())
        self.assertTrue(version["label"].startswith("Version "))


if __name__ == "__main__":
    unittest.main()
