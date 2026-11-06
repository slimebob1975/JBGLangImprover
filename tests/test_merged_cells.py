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

from app.src import JBGLangImprovSuggestorAI as suggestor_module
from app.src.JBGDocumentStructureExtractor import DocumentStructureExtractor
from app.src.JBGLanguageImprover import JBGLanguageImprover
from app.src.JBGReadabilityMetrics import compute_document_readability

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"


def _quiet_logger(name):
    logger = logging.getLogger(name)
    logger.handlers.clear()
    logger.addHandler(logging.NullHandler())
    logger.propagate = False
    return logger


def merged_table_document(path):
    """
    Som tabell 10 i testrapporten (c7c86075): en rubrikrad där varje
    kolumnrubrik spänner över två kolumner, och "Program" över två rader.
    """
    document = Document()
    table = document.add_table(rows=3, cols=5)
    table.cell(0, 0).merge(table.cell(1, 0))           # "Program", två rader
    table.cell(0, 1).merge(table.cell(0, 2))           # "Kvinnor", två kolumner
    table.cell(0, 3).merge(table.cell(0, 4))           # "Män", två kolumner
    table.cell(0, 0).text = "Program"
    table.cell(0, 1).text = "Kvinnor"
    table.cell(0, 3).text = "Män och pojkar"
    for column, text in zip(range(1, 5), ("Antal", "Andel", "Antal", "Andel")):
        table.cell(1, column).text = text
    for column, text in enumerate(("Etableringsprogrammet", "370", "3,1", "170", "3,0")):
        table.cell(2, column).text = text
    document.save(path)


class MergedCellsTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.source = Path(self.temp_dir.name) / "merged.docx"
        merged_table_document(self.source)
        self.structure = DocumentStructureExtractor(str(self.source), _quiet_logger("merged")).extract()

    def tearDown(self):
        self.temp_dir.cleanup()

    def cells(self, row):
        return {e["element_id"]: e["text"] for e in self.structure["elements"]
                if e["element_id"].startswith(f"table_1_cell_{row}_")}

    def test_cells_are_numbered_like_the_xml_without_duplicates(self):
        # Rad 1 har tre verkliga celler: Program, Kvinnor, Män och pojkar
        self.assertEqual(self.cells(1), {
            "table_1_cell_1_1_p1": "Program",
            "table_1_cell_1_2_p1": "Kvinnor",
            "table_1_cell_1_3_p1": "Män och pojkar",
        })
        # Rad 2: den fortsatta cellen under "Program" är tom; övriga i ordning
        self.assertEqual(self.cells(2), {
            "table_1_cell_2_1_p1": "",
            "table_1_cell_2_2_p1": "Antal", "table_1_cell_2_3_p1": "Andel",
            "table_1_cell_2_4_p1": "Antal", "table_1_cell_2_5_p1": "Andel",
        })

    def test_merged_text_is_counted_once_in_lix(self):
        words = compute_document_readability(self.structure, []).before.words
        # Program (1), Kvinnor (1), Män och pojkar (3), fyra rubriker (4) och
        # värderaden (7; "3,1" och "3,0" är två ord var). Före rättningen räknades
        # Program, Kvinnor och Män och pojkar två gånger: 21.
        self.assertEqual(words, 16)

    def test_a_proposal_after_a_merged_cell_reaches_the_right_cell(self):
        content = json.dumps([{"type": "table_cell", "element_id": "table_1_cell_1_3_p1",
                               "old": "Män och pojkar", "new": "Pojkar och män", "motivation": "m"}],
                             ensure_ascii=False)
        client = SimpleNamespace(chat=SimpleNamespace(completions=SimpleNamespace(create=lambda **k: SimpleNamespace(
            choices=[SimpleNamespace(message=SimpleNamespace(content=content))], usage=None))))
        for mode in ("tracked", "simple"):
            with self.subTest(mode=mode):
                output = Path(self.temp_dir.name) / f"out_{mode}.docx"
                with mock.patch.object(suggestor_module.openai, "OpenAI", return_value=client):
                    improver = JBGLanguageImprover(
                        input_path=str(self.source), api_key="k", model="m", prompt_policy="p", temperature=1,
                        include_motivations=False, logger=_quiet_logger(f"merged-{mode}"), docx_mode=mode,
                        include_about_section=False, compute_readability=False)
                    improver.run(output_path=str(output))
                counts = improver.run_summary.local_suggestions
                self.assertEqual((counts.applied, counts.failed), (1, 0),
                                 [r.message for r in improver.render_results])
                row = etree.fromstring(ZipFile(output).read("word/document.xml")).find(f".//{W}tbl").find(f"{W}tr")
                third = row.findall(f"{W}tc")[2]
                if mode == "tracked":
                    accepted = "".join(third.xpath(
                        ".//w:t[not(ancestor::w:del)]/text()", namespaces={"w": W_NS}))
                    rejected = "".join(third.xpath(
                        ".//w:t[not(ancestor::w:ins)]/text() | .//w:delText/text()",
                        namespaces={"w": W_NS}))
                    self.assertEqual(accepted, "Pojkar och män")
                    self.assertEqual(rejected, "Män och pojkar")
                else:
                    combined = "".join(t.text or "" for t in third.iter(f"{W}t", f"{W}delText"))
                    self.assertIn("Män och pojkar", combined)
                    self.assertIn("Pojkar och män", combined)


if __name__ == "__main__":
    unittest.main()
