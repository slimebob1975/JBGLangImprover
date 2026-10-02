import json
import logging
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock
from zipfile import ZipFile

from docx import Document
from lxml import etree

from app.src import JBGGlobalAnalyzerAI as analyzer_module
from app.src import JBGLangImprovSuggestorAI as suggestor_module
from app.src.JBGGlobalAnalyzerAI import JBGGlobalAnalyzerAI
from app.src.JBGLangImprovSuggestorAI import JBGLangImprovSuggestorAI
from app.src.JBGLanguageImprover import JBGLanguageImprover
from app.src.JBGModelClient import MODEL_CALL_MAX_RETRIES
from tests.test_about_section import build_iaf_like_document


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"


def _quiet_logger(name):
    logger = logging.getLogger(name)
    logger.handlers.clear()
    logger.addHandler(logging.NullHandler())
    logger.propagate = False
    return logger


def _client(content):
    def create(**kwargs):
        return SimpleNamespace(choices=[SimpleNamespace(message=SimpleNamespace(content=content))], usage=None)
    return SimpleNamespace(chat=SimpleNamespace(completions=SimpleNamespace(create=create)))


class RetriesInsteadOfPauseTests(unittest.TestCase):
    """Ingen fast paus mellan anropen; OpenAI-biblioteket sköter nya försök."""

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.logger = _quiet_logger(f"retries-{id(self)}")
        self.structure_path = Path(self.temp_dir.name) / "s.json"
        elements = [{"type": "paragraph", "element_id": f"paragraph_{i}",
                     "text": "En mening som upprepas för att fylla flera delar. " * 30}
                    for i in range(1, 7)]
        self.structure_path.write_text(json.dumps({"type": "docx", "elements": elements}), encoding="utf-8")

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_local_review_uses_sdk_retries_and_never_sleeps(self):
        with mock.patch.object(suggestor_module.openai, "OpenAI", return_value=_client("[]")) as factory, \
             mock.patch("time.sleep", side_effect=AssertionError("fixed pause")):
            suggestor = JBGLangImprovSuggestorAI(api_key="k", model="m", prompt_policy="p",
                                                 temperature=1, logger=self.logger)
            suggestor.load_structure(str(self.structure_path))
            suggestor.suggest_changes_token_aware_batching(max_tokens_per_call=500)
        self.assertGreater(len(suggestor._chunk_elements(suggestor.json_structured_document["elements"], 500)), 1)
        factory.assert_called_with(api_key="k", max_retries=MODEL_CALL_MAX_RETRIES)

    def test_global_review_uses_sdk_retries(self):
        with mock.patch.object(analyzer_module.openai, "OpenAI", return_value=_client('{"findings": []}')) as factory:
            JBGGlobalAnalyzerAI("k", "m", 1, self.logger, policy="p").analyze(
                {"elements": [{"type": "paragraph", "element_id": "paragraph_1", "text": "Text.", "doc_order": 1}]}
            )
        factory.assert_called_with(api_key="k", max_retries=MODEL_CALL_MAX_RETRIES)

    def test_status_message_for_one_and_several_calls(self):
        messages = []
        with mock.patch.object(suggestor_module.openai, "OpenAI", return_value=_client("[]")):
            suggestor = JBGLangImprovSuggestorAI(api_key="k", model="m", prompt_policy="p", temperature=1,
                                                 logger=self.logger, progress_callback=messages.append)
            suggestor.load_structure(str(self.structure_path))
            suggestor.suggest_changes_token_aware_batching()
            suggestor.suggest_changes_token_aware_batching(max_tokens_per_call=500)
        self.assertIn("Skickar dokumentet till språkmodellen i ett anrop.", messages)
        self.assertTrue(any(m.startswith("Skickar dokumentet till språkmodellen i ") and m.endswith(" anrop.")
                            and "ett" not in m for m in messages))
        self.assertFalse(any("Dokumentet är stort" in m for m in messages))


class RevisionIdTests(unittest.TestCase):
    """Nya spårade ändringar får id som inte krockar med ändringar i originalet."""

    def test_new_revisions_continue_after_existing_ids(self):
        with tempfile.TemporaryDirectory() as tmp:
            source, output = Path(tmp) / "src.docx", Path(tmp) / "out.docx"
            build_iaf_like_document(str(source), existing_revision_id=41)
            content = json.dumps([{
                "type": "paragraph", "element_id": "paragraph_2",
                "old": "brister i arbetet", "new": "har brister i arbetet",
                "motivation": "Tydligare.",
            }], ensure_ascii=False)
            with mock.patch.object(suggestor_module.openai, "OpenAI", return_value=_client(content)):
                improver = JBGLanguageImprover(
                    input_path=str(source), api_key="k", model="m", prompt_policy="p", temperature=1,
                    include_motivations=True, logger=_quiet_logger(f"rev-{id(self)}"), docx_mode="tracked",
                    include_about_section=False, compute_readability=False,
                )
                improver.run(output_path=str(output))
            self.assertEqual(improver.run_summary.local_suggestions.applied, 1)
            with ZipFile(output) as z:
                body = etree.fromstring(z.read("word/document.xml"))
            ids = [int(e.get(f"{W}id")) for e in body.iter(f"{W}ins", f"{W}del")
                   if e.get(f"{W}author") != "Någon"]
            self.assertTrue(ids)
            self.assertTrue(all(i > 41 for i in ids), ids)
            self.assertEqual(len(ids), len(set(ids)))


if __name__ == "__main__":
    unittest.main()
