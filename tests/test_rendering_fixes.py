import json
import logging
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock
from zipfile import ZipFile

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from lxml import etree

from app.src import JBGLangImprovSuggestorAI as suggestor_module
from app.src.JBGLangImprovSuggestorAI import JBGLangImprovSuggestorAI
from app.src.JBGLanguageImprover import JBGLanguageImprover
from app.src.JBGRunEditing import append_visible_text, map_boundaries, run_visible_text

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"


def word_view(paragraph, accept: bool) -> str:
    """Styckets text efter Words 'Godkänn alla' (accept=True) eller 'Avvisa alla'."""
    out = []
    for element in paragraph.iter(f"{W}t", f"{W}delText", f"{W}br", f"{W}tab"):
        ancestors = {a.tag for a in element.iterancestors()}
        if (accept and f"{W}del" in ancestors) or (not accept and f"{W}ins" in ancestors):
            continue
        out.append({f"{W}br": "\n", f"{W}tab": "\t"}.get(element.tag, element.text or ""))
    return "".join(out)


def _quiet_logger(name):
    logger = logging.getLogger(name)
    logger.handlers.clear()
    logger.addHandler(logging.NullHandler())
    logger.propagate = False
    return logger


# ---------------- Dokument med de tre fallen ----------------

def same_run_document(path):
    """En textbit med text, radbrytning och text (vanligt i Word)."""
    document = Document()
    document.add_paragraph()._p.append(parse_xml(
        f'<w:r {nsdecls("w")}><w:t xml:space="preserve">Första raden ska stå kvar.</w:t><w:br/>'
        f'<w:t xml:space="preserve">Andra raden har ett fel.</w:t></w:r>'))
    document.save(path)


def hyperlink_document(path):
    """Testkörningen c7c86075: förslaget går in i länktexten."""
    document = Document()
    paragraph = document.add_paragraph("Våra rapporter finns publicerade på vår webbplats ")
    rid = document.part.relate_to(
        "https://www.iaf.se", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True)
    paragraph._p.append(parse_xml(
        f'<w:hyperlink {nsdecls("w", "r")} r:id="{rid}"><w:r><w:t>www.iaf.se</w:t></w:r></w:hyperlink>'))
    paragraph.add_run(".")
    document.save(path)


def line_break_document(path):
    """Testkörningen b02ce006: förslagets gamla text slutar med en radbrytning."""
    document = Document()
    paragraph = document.add_paragraph("Modeller kan vara ett verktyg för ")
    paragraph.add_run("arbetslöshetskassorna att snabbare hitta fel, som IAF gjort i granskningen.")
    paragraph.add_run().add_break()
    paragraph.add_run("Nästa rad efter brytningen.")
    document.save(path)


CASES = {
    "same_run": (same_run_document, "har ett fel", "är rättad",
                 "Första raden ska stå kvar.\nAndra raden är rättad."),
    "hyperlink": (hyperlink_document, "publicerade på vår webbplats www.iaf.se", "på vår webbplats, iaf.se",
                  "Våra rapporter finns på vår webbplats, iaf.se."),
    "line_break": (line_break_document,
                   "arbetslöshetskassorna att snabbare hitta fel, som IAF gjort i granskningen.\n",
                   "a-kassorna för att snabbare hitta fel, vilket IAF har gjort i granskningen.",
                   "Modeller kan vara ett verktyg för a-kassorna för att snabbare hitta fel, vilket IAF har "
                   "gjort i granskningen.Nästa rad efter brytningen."),
}


class RenderingFixesTests(unittest.TestCase):
    def run_case(self, name, mode):
        build, old, new, _ = CASES[name]
        tmp = Path(tempfile.mkdtemp())
        source, output = tmp / "source.docx", tmp / "output.docx"
        build(source)
        content = json.dumps([{"type": "paragraph", "element_id": "paragraph_1", "old": old, "new": new,
                               "motivation": "m"}], ensure_ascii=False)
        client = SimpleNamespace(chat=SimpleNamespace(completions=SimpleNamespace(create=lambda **k: SimpleNamespace(
            choices=[SimpleNamespace(message=SimpleNamespace(content=content))], usage=None))))
        with mock.patch.object(suggestor_module.openai, "OpenAI", return_value=client):
            improver = JBGLanguageImprover(
                input_path=str(source), api_key="k", model="m", prompt_policy="p", temperature=1,
                include_motivations=False, logger=_quiet_logger(f"{name}-{mode}"), docx_mode=mode,
                include_about_section=False, compute_readability=False)
            improver.run(output_path=str(output))
        original = etree.fromstring(ZipFile(source).read("word/document.xml")).find(f"{W}body").find(f"{W}p")
        result = etree.fromstring(ZipFile(output).read("word/document.xml")).find(f"{W}body").find(f"{W}p")
        return improver, original, result

    def test_every_case_is_applied_in_both_modes(self):
        for name in CASES:
            for mode in ("tracked", "simple"):
                with self.subTest(case=name, mode=mode):
                    improver, _, _ = self.run_case(name, mode)
                    counts = improver.run_summary.local_suggestions
                    self.assertEqual((counts.applied, counts.failed), (1, 0),
                                     [r.message for r in improver.render_results])

    def test_tracked_changes_restore_the_original_and_give_the_right_result(self):
        for name, (_, _, _, accepted) in CASES.items():
            with self.subTest(case=name):
                _, original, result = self.run_case(name, "tracked")
                self.assertEqual(word_view(result, accept=False), word_view(original, accept=True))
                self.assertEqual(word_view(result, accept=True), accepted)

    def test_text_in_the_same_run_is_never_lost(self):
        # Tidigare raderades "Första raden ska stå kvar." och radbrytningen tyst
        for mode in ("tracked", "simple"):
            with self.subTest(mode=mode):
                _, _, result = self.run_case("same_run", mode)
                text = word_view(result, accept=False)
                self.assertTrue(text.startswith("Första raden ska stå kvar.\n"), text)

    def test_the_link_is_kept_and_its_new_text_stays_linked(self):
        # Spåra ändringar: efter "Godkänn alla" är "iaf.se" fortfarande länkat
        _, _, result = self.run_case("hyperlink", "tracked")
        link = result.find(f"{W}hyperlink")
        self.assertIsNotNone(link)
        self.assertEqual(word_view(link, accept=True), "iaf.se")

        # Enkel färgmarkering: i länken står "www.iaf.se" rödstruket och "iaf.se" grönt
        _, _, result = self.run_case("hyperlink", "simple")
        link = result.find(f"{W}hyperlink")
        self.assertIsNotNone(link)
        runs = {"".join(t.text or "" for t in r.iter(f"{W}t")): r for r in link.findall(f"{W}r")}
        self.assertIsNotNone(runs["www.iaf.se"].find(f"{W}rPr/{W}strike"))
        self.assertIsNone(runs["iaf.se"].find(f"{W}rPr/{W}strike"))
        self.assertIsNotNone(runs["iaf.se"].find(f"{W}rPr/{W}color"))

    def test_line_breaks_and_tabs_are_real_elements(self):
        run = etree.Element(f"{W}r")
        append_visible_text(run, "a\tb\nc")
        self.assertEqual([child.tag.split("}")[1] for child in run], ["t", "tab", "t", "br", "t"])
        self.assertEqual(run_visible_text(run), "a\tb\nc")

    def test_new_text_follows_the_container_where_each_change_starts(self):
        old, new = "publicerade på vår webbplats www.iaf.se", "på vår webbplats, iaf.se"
        cut = map_boundaries(old, new, [len("publicerade på vår webbplats ")])[0]
        self.assertEqual((new[:cut], new[cut:]), ("på vår webbplats, ", "iaf.se"))


class LegalReferenceTests(unittest.TestCase):
    """Lagrumshänvisningar får aldrig ändras av ett språkförslag."""

    def accepted(self, text, old, new, element_type="paragraph"):
        extra = {"footnote_id": "7"} if element_type == "footnote" else {}
        path = Path(tempfile.mkdtemp()) / "s.json"
        path.write_text(json.dumps({"type": "docx", "elements": [
            {"type": element_type, "element_id": f"{element_type}_1", "text": text, **extra}]}, ensure_ascii=False),
            encoding="utf-8")
        suggestor = JBGLangImprovSuggestorAI(api_key="k", model="m", prompt_policy="p", temperature=1,
                                             logger=_quiet_logger("legal"))
        suggestor.load_structure(str(path))
        return bool(suggestor._postprocess_suggestions([
            {"type": element_type, "element_id": f"{element_type}_1", "old": old, "new": new, **extra}]))

    def test_changes_to_a_reference_are_rejected(self):
        for text, old, new, element_type in (
            ("Enligt 6 c § gäller detta.", "c §", " c ", "footnote"),     # § tas bort
            ("Enligt 6 c § gäller detta.", "6 c §", "6 b §", "paragraph"),   # paragrafen ändras
            ("Enligt 6 c § lagen gäller detta.", "c", "b", "paragraph"),     # kort "ordbyte" i hänvisningen
            ("Se 1 kap. 6 § i lagen.", "1 kap. 6 §", "6 §", "paragraph"),    # kapitlet försvinner
            ("Enligt lagen (1997:238) gäller detta.", "lagen (1997:238)", "lagen (1997:283)", "paragraph"),
        ):
            with self.subTest(old=old, new=new):
                self.assertFalse(self.accepted(text, old, new, element_type))

    def test_rewrites_that_keep_the_reference_are_accepted(self):
        self.assertTrue(self.accepted(
            "Enligt 6 c § lagen gäller detta för kassan.",
            "Enligt 6 c § lagen gäller detta för kassan.",
            "Enligt 6 c § i lagen gäller detta för a-kassan."))
        self.assertTrue(self.accepted("Enligt 6 c § gäller detta i ärendet.", "i ärendet", "för ärendet"))


if __name__ == "__main__":
    unittest.main()
