import json
import logging
import os
import tempfile
import threading
import time
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

from docx import Document

from app.src import JBGGlobalAnalyzerAI as analyzer_module
from app.src import JBGLangImprovSuggestorAI as suggestor_module
from app.src.JBGLangImprovSuggestorAI import JBGLangImprovSuggestorAI
from app.src.JBGLanguageImprover import JBGLanguageImprover
from app.src.JBGModelClient import DEFAULT_MAX_PARALLEL_MODEL_CALLS, max_parallel_model_calls


def _quiet_logger(name):
    logger = logging.getLogger(name)
    logger.handlers.clear()
    logger.addHandler(logging.NullHandler())
    logger.propagate = False
    return logger


class _SlowModel:
    """
    Simulerad modell: varje anrop tar en stund och föreslår en ändring i det
    första elementet i sin del. Räknar hur många anrop som pågår samtidigt.
    """

    def __init__(self, delay=0.3, fail_for=None, global_content='{"findings": []}'):
        self.delay = delay
        self.fail_for = fail_for or set()
        self.global_content = global_content
        self.lock = threading.Lock()
        self.active = self.max_active = 0
        self.local_spans = []
        self.global_spans = []

    def create(self, **kwargs):
        user = kwargs["messages"][1]["content"]
        is_global = "dokumentnivå" in kwargs["messages"][0]["content"]
        start = time.monotonic()
        with self.lock:
            self.active += 1
            self.max_active = max(self.max_active, self.active)
        try:
            time.sleep(self.delay)
            if is_global:
                content = self.global_content
            else:
                first_id = user.split('"element_id": "', 1)[1].split('"', 1)[0]
                if first_id in self.fail_for:
                    raise RuntimeError(f"simulerat fel för {first_id}")
                content = json.dumps([{"type": "paragraph", "element_id": first_id,
                                       "old": "Mening", "new": "En mening"}])
        finally:
            with self.lock:
                self.active -= 1
                (self.global_spans if is_global else self.local_spans).append((start, time.monotonic()))
        return SimpleNamespace(choices=[SimpleNamespace(message=SimpleNamespace(content=content))], usage=None)


def _structure(path, count=6):
    elements = [{"type": "paragraph", "element_id": f"paragraph_{i}",
                 "text": "Mening som fyller ut delen. " * 70} for i in range(1, count + 1)]
    path.write_text(json.dumps({"type": "docx", "elements": elements}), encoding="utf-8")


class ParallelLocalCallsTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.path = Path(self.temp_dir.name) / "s.json"
        _structure(self.path)
        self.logger = _quiet_logger(f"parallel-{id(self)}")

    def tearDown(self):
        self.temp_dir.cleanup()

    def run_suggestor(self, model, env=None, messages=None):
        client = SimpleNamespace(chat=SimpleNamespace(completions=model))
        with mock.patch.object(suggestor_module.openai, "OpenAI", return_value=client), \
             mock.patch.dict(os.environ, env or {}, clear=False):
            suggestor = JBGLangImprovSuggestorAI(api_key="k", model="m", prompt_policy="p", temperature=1,
                                                 logger=self.logger,
                                                 progress_callback=(messages.append if messages is not None else None))
            suggestor.load_structure(str(self.path))
            started = time.monotonic()
            suggestor.suggest_changes_token_aware_batching(max_tokens_per_call=600)
            return suggestor, time.monotonic() - started

    def test_calls_overlap_up_to_the_limit_and_results_keep_chunk_order(self):
        model = _SlowModel()
        suggestor, elapsed = self.run_suggestor(model)
        calls = len(model.local_spans)
        self.assertEqual(calls, 6)
        self.assertEqual(model.max_active, DEFAULT_MAX_PARALLEL_MODEL_CALLS)
        self.assertLess(elapsed, calls * model.delay * 0.75)          # klart snabbare än i följd
        self.assertEqual([s.element_id for s in suggestor.validated_suggestions],
                         [f"paragraph_{i}" for i in range(1, 7)])      # samma ordning som i följd

    def test_limit_of_one_gives_calls_in_sequence(self):
        model = _SlowModel(delay=0.05)
        suggestor, _ = self.run_suggestor(model, env={"JBG_MAX_PARALLEL_MODEL_CALLS": "1"})
        self.assertEqual(model.max_active, 1)
        self.assertEqual(len(suggestor.validated_suggestions), 6)

    def test_a_failing_call_does_not_stop_the_others(self):
        messages = []
        model = _SlowModel(delay=0.05, fail_for={"paragraph_3"})
        suggestor, _ = self.run_suggestor(model, messages=messages)
        self.assertEqual([s.element_id for s in suggestor.validated_suggestions],
                         ["paragraph_1", "paragraph_2", "paragraph_4", "paragraph_5", "paragraph_6"])
        self.assertTrue(any(m.startswith("Fel i API-anrop 3 av 6") for m in messages))

    def test_progress_counts_finished_calls(self):
        messages = []
        self.run_suggestor(_SlowModel(delay=0.05), messages=messages)
        self.assertEqual([m for m in messages if m.startswith("Klar med")],
                         [f"Klar med {k} av 6 anrop." for k in range(1, 7)])

    def test_limit_setting(self):
        for raw, expected in (("", DEFAULT_MAX_PARALLEL_MODEL_CALLS), ("2", 2), ("0", 1), ("x", DEFAULT_MAX_PARALLEL_MODEL_CALLS)):
            with self.subTest(raw=raw), mock.patch.dict(os.environ, {"JBG_MAX_PARALLEL_MODEL_CALLS": raw}):
                self.assertEqual(max_parallel_model_calls(), expected)


class GlobalAlongsideLocalTests(unittest.TestCase):
    def test_global_call_runs_while_local_calls_are_running(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "rapport.docx"
            document = Document()
            for i in range(8):
                document.add_paragraph("Mening som fyller ut delen. " * 70)
            document.save(source)

            model = _SlowModel(delay=0.3)
            client = SimpleNamespace(chat=SimpleNamespace(completions=model))
            with mock.patch.object(suggestor_module.openai, "OpenAI", return_value=client), \
                 mock.patch.object(analyzer_module.openai, "OpenAI", return_value=client), \
                 mock.patch.object(suggestor_module, "MAX_TOKEN_PER_CALL", 600):
                improver = JBGLanguageImprover(
                    input_path=str(source), api_key="k", model="m", prompt_policy="p", temperature=1,
                    include_motivations=False, logger=_quiet_logger("global-alongside"), docx_mode="simple",
                    global_review=True, include_about_section=False, compute_readability=False,
                )
                improver.run(output_path=str(Path(tmp) / "out.docx"))

            self.assertEqual(len(model.global_spans), 1)
            global_start, global_end = model.global_spans[0]
            last_local_end = max(end for _, end in model.local_spans)
            first_local_start = min(start for start, _ in model.local_spans)
            # Den globala granskningen startar innan den lokala är klar ...
            self.assertLess(global_start, last_local_end)
            self.assertLess(abs(global_start - first_local_start), model.delay)
            # ... och resultatet är på plats innan körningen fortsätter
            self.assertEqual(improver.run_summary.global_findings["accepted"], 0)
            self.assertTrue(improver.run_summary.succeeded)


if __name__ == "__main__":
    unittest.main()
