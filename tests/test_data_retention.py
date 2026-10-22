import io
import json
import logging
import os
import tempfile
import time
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

from docx import Document

from app.src import JBGLangImprovSuggestorAI as suggestor_module
from app.src.JBGDataRetention import DEBUG_ENV_VAR, loggable_text, run_file_paths
from app.src.JBGLangImprovSuggestorAI import JBGLangImprovSuggestorAI

ROOT = Path(__file__).resolve().parents[1]

try:
    from fastapi.testclient import TestClient
    HAS_FASTAPI = True
except ImportError:  # pragma: no cover - beror på miljön
    HAS_FASTAPI = False


class LoggedTextTests(unittest.TestCase):
    """Dokumenttext hamnar bara i loggarna när felsökningsdata uttryckligen ska sparas."""

    def test_only_length_by_default(self):
        with mock.patch.dict(os.environ, {DEBUG_ENV_VAR: ""}):
            self.assertEqual(loggable_text("hemlig mening"), "<13 tecken>")
        with mock.patch.dict(os.environ, {DEBUG_ENV_VAR: "1"}):
            self.assertEqual(loggable_text("hemlig mening"), "'hemlig mening'")

    def test_rejected_proposals_are_logged_without_document_text(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "s.json"
            path.write_text(json.dumps({"type": "docx", "elements": [{
                "type": "paragraph", "element_id": "paragraph_1",
                "text": "Kassan fattade ett hemligt beslut om ersättning."}]}, ensure_ascii=False), encoding="utf-8")
            logger = logging.getLogger(f"retention-{id(self)}")
            logger.propagate = False
            suggestor = JBGLangImprovSuggestorAI(api_key="k", model="m", prompt_policy="p", temperature=1, logger=logger)
            suggestor.load_structure(str(path))
            # "hemligt" -> "x": avvisas som för olikt originalet
            suggestion = {"type": "paragraph", "element_id": "paragraph_1", "old": "hemligt", "new": "x"}
            for value, should_contain in (("", False), ("1", True)):
                with self.subTest(debug=value), mock.patch.dict(os.environ, {DEBUG_ENV_VAR: value}), \
                     self.assertLogs(logger, level="WARNING") as captured:
                    suggestor._postprocess_suggestions([suggestion])
                text = "\n".join(captured.output)
                self.assertIn("reason=", text)
                self.assertEqual("hemligt" in text, should_contain)

    def test_run_file_paths(self):
        self.assertEqual(run_file_paths("uploads/abc_Rapport.docx"), [
            "uploads/abc_Rapport.docx", "uploads/abc_Rapport_structure.json", "uploads/abc_Rapport_suggestions.json",
            "uploads/abc_Rapport_suggestion_filter_report.json", "uploads/abc_Rapport_global_findings.json",
            "uploads/abc_Rapport_run_summary.json",
        ])


@unittest.skipUnless(HAS_FASTAPI, "fastapi saknas i miljön")
class AppRetentionTests(unittest.TestCase):
    """Hela flödet genom appen: uppladdning, körning, nedladdning och städning."""

    @classmethod
    def setUpClass(cls):
        cls._cwd = os.getcwd()
        os.chdir(ROOT)   # main.py använder relativa sökvägar från repots rot
        from app import main
        cls.main = main

    @classmethod
    def tearDownClass(cls):
        os.chdir(cls._cwd)

    def run_job(self, debug=""):
        main = self.main
        document = Document()
        document.add_paragraph("Inledning", style="Heading 1")
        document.add_paragraph("Myndigheten har genomfört en granskning av ärendet.")
        buffer = io.BytesIO()
        document.save(buffer)
        editable, _, _ = main.load_prompt_parts(str(ROOT / "policy" / "prompt_policy.md"))
        client = SimpleNamespace(chat=SimpleNamespace(completions=SimpleNamespace(create=lambda **k: SimpleNamespace(
            choices=[SimpleNamespace(message=SimpleNamespace(content="[]"))], usage=None))))
        with mock.patch.object(suggestor_module.openai, "OpenAI", return_value=client), \
             mock.patch.dict(os.environ, {DEBUG_ENV_VAR: debug}):
            response = TestClient(main.app).post("/upload/", data={
                "api_key": "sk-test", "model": "test-model", "editable_prompt": editable,
                "temperature": "1", "docx_mode": "simple",
            }, files={"file": ("Retention_test.docx", buffer.getvalue(),
                               "application/vnd.openxmlformats-officedocument.wordprocessingml.document")})
        self.assertEqual(response.status_code, 200)
        job_id = response.json()["job_id"]
        upload = os.path.join(main.UPLOAD_DIR, f"{job_id}_Retention_test.docx")
        self.addCleanup(lambda: [os.remove(p) for p in run_file_paths(upload) if os.path.exists(p)])
        result = os.path.join(main.RESULTS_DIR, f"{job_id}.result.docx")
        self.addCleanup(lambda: os.path.exists(result) and os.remove(result))
        return job_id, upload, result

    def test_upload_and_diagnostics_are_deleted_after_the_run(self):
        job_id, upload, result = self.run_job()
        self.assertTrue(self.main.jobs_lang[job_id]["done"])
        self.assertEqual([p for p in run_file_paths(upload) if os.path.exists(p)], [])
        self.assertTrue(os.path.exists(result))

    def test_debug_setting_keeps_upload_and_diagnostics(self):
        _, upload, _ = self.run_job(debug="1")
        self.assertTrue(os.path.exists(upload))
        self.assertTrue(os.path.exists(os.path.splitext(upload)[0] + "_structure.json"))

    def test_result_is_deleted_after_download(self):
        job_id, _, result = self.run_job()
        response = TestClient(self.main.app).get(f"/download/{job_id}")
        self.assertEqual(response.status_code, 200)
        self.assertTrue(response.content.startswith(b"PK"))
        self.assertFalse(os.path.exists(result))

    def test_cleanup_removes_old_files_and_prunes_jobs(self):
        main = self.main
        old = time.time() - 2 * 3600
        paths = [os.path.join(main.UPLOAD_DIR, "old_test_upload.docx"),
                 os.path.join(main.RESULTS_DIR, "old_test.result.docx")]
        for path in paths:
            Path(path).write_bytes(b"x")
            os.utime(path, (old, old))
        main.jobs_lang.update({
            "old-done": {"done": True, "created_at": old},
            "old-running": {"done": False, "created_at": old},
            "new-done": {"done": True, "created_at": time.time()},
        })
        self.addCleanup(lambda: [main.jobs_lang.pop(k, None) for k in ("old-done", "old-running", "new-done")])
        main.run_cleanup(logging.getLogger("cleanup-test"))
        self.assertFalse(any(os.path.exists(p) for p in paths))
        self.assertNotIn("old-done", main.jobs_lang)
        self.assertIn("old-running", main.jobs_lang)     # pågående jobb får 24 timmar
        self.assertIn("new-done", main.jobs_lang)

    def test_cleanup_runs_on_a_timer(self):
        main = self.main
        calls = []
        with mock.patch.object(main, "CLEANUP_INTERVAL_SECONDS", 0.05), \
             mock.patch.object(main, "run_cleanup", side_effect=lambda logger: calls.append(time.time())):
            with TestClient(main.app):          # startar appens lifespan
                deadline = time.time() + 2
                while len(calls) < 2 and time.time() < deadline:
                    time.sleep(0.05)
        self.assertGreaterEqual(len(calls), 2)


if __name__ == "__main__":
    unittest.main()
