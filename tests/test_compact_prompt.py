import json
import logging
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

from app.src import JBGLangImprovSuggestorAI as suggestor_module
from app.src.JBGLangImprovSuggestorAI import JBGLangImprovSuggestorAI


def _quiet_logger(name):
    logger = logging.getLogger(name)
    logger.handlers.clear()
    logger.addHandler(logging.NullHandler())
    logger.propagate = False
    return logger


FULL = {  # så som extraktorn sparar ett element i dag
    "type": "paragraph", "element_id": "paragraph_49",
    "text": "Granskningen visar att ett riskbaserat urval kan identifiera fler fel.",
    "empty": False, "part_name": "word/document.xml", "container_path": "/document/body/paragraph[49]",
    "footnote_id": None, "paragraph_index": 49, "contains_linebreaks": False, "contains_tabs": False,
    "may_contain_special_runs": False, "style_id": "Normal", "style_name": "Normal",
    "heading_level": None, "doc_order": 54, "generated_kind": None,
}


class CompactPromptTests(unittest.TestCase):
    """G1.6: modellen får bara det den behöver; delningen i anrop är oförändrad."""

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.logger = _quiet_logger(f"compact-{id(self)}")

    def tearDown(self):
        self.temp_dir.cleanup()

    def suggestor(self, elements):
        path = Path(self.temp_dir.name) / "s.json"
        path.write_text(json.dumps({"type": "docx", "elements": elements}, ensure_ascii=False), encoding="utf-8")
        suggestor = JBGLangImprovSuggestorAI(api_key="k", model="m", prompt_policy="p", temperature=1, logger=self.logger)
        suggestor.load_structure(str(path))
        return suggestor

    def test_only_the_fields_the_model_needs_are_sent(self):
        compact = JBGLangImprovSuggestorAI._compact_element
        self.assertEqual(compact(FULL), {"type": "paragraph", "element_id": "paragraph_49", "text": FULL["text"]})
        footnote = dict(FULL, type="footnote", element_id="footnote_3", footnote_id="4")
        self.assertEqual(compact(footnote)["footnote_id"], "4")
        heading = dict(FULL, element_id="paragraph_48", text="Resultat", heading_level=2)
        self.assertEqual(compact(heading)["heading_level"], 2)
        self.assertNotIn("heading_level", compact(FULL))

    def test_empty_elements_are_not_sent_and_an_empty_part_makes_no_call(self):
        calls = []

        def create(**kwargs):
            calls.append(kwargs["messages"][1]["content"])
            return SimpleNamespace(choices=[SimpleNamespace(message=SimpleNamespace(content="[]"))], usage=None)

        text_element = dict(FULL, text="Mening som fyller ut delen. " * 70)
        empties = [dict(FULL, element_id=f"paragraph_{i}", text="", empty=True) for i in range(100, 160)]
        suggestor = self.suggestor([text_element] + empties)
        chunks = suggestor._chunk_elements(suggestor.json_structured_document["elements"], max_tokens_per_call=600)
        self.assertGreater(len(chunks), 1)   # minst en del består bara av tomma element

        client = SimpleNamespace(chat=SimpleNamespace(completions=SimpleNamespace(create=create)))
        with mock.patch.object(suggestor_module.openai, "OpenAI", return_value=client):
            suggestor.suggest_changes_token_aware_batching(max_tokens_per_call=600)
        self.assertEqual(len(calls), 1)
        sent = json.loads(calls[0].split(": ", 1)[1])
        self.assertEqual([e["element_id"] for e in sent], ["paragraph_49"])

    def test_the_split_into_calls_is_unchanged(self):
        # Delningen mäts fortfarande på de fullständiga elementen, så varje anrop
        # innehåller samma text som före G1.6.
        elements = [dict(FULL, element_id=f"paragraph_{i}", text="Mening som fyller ut delen. " * 20)
                    for i in range(1, 13)]
        suggestor = self.suggestor(elements)
        chunks = suggestor._chunk_elements(suggestor.json_structured_document["elements"], max_tokens_per_call=1500)
        expected, current, size = [], [], len(suggestor.policy_prompt)
        for element in elements:   # samma regel som före ändringen, på hela elementet
            length = len(json.dumps(element, ensure_ascii=False))
            if current and size + length > 1500 * 4:
                expected.append(current)
                current, size = [], len(suggestor.policy_prompt)
            current.append(element["element_id"])
            size += length
        expected.append(current)
        self.assertEqual([[e["element_id"] for e in chunk] for chunk in chunks], expected)

    def test_policy_describes_the_fields(self):
        policy = Path(__file__).resolve().parents[1].joinpath("policy", "prompt_policy.md").read_text(encoding="utf-8")
        self.assertIn("footnote_id för fotnoter", policy)
        self.assertIn("heading_level för rubriker", policy)


if __name__ == "__main__":
    unittest.main()
