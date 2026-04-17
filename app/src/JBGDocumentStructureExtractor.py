import os
import json
import docx
import sys
from dataclasses import dataclass, asdict
from typing import Optional, Any
from lxml import etree
from zipfile import ZipFile


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": W_NS}


# ============================================================================
# Datamodeller
# ============================================================================

@dataclass
class ExtractedElement:
    type: str
    element_id: str
    text: str
    empty: bool

    part_name: str
    container_path: str

    footnote_id: Optional[str] = None
    section_index: Optional[int] = None
    header_index: Optional[int] = None
    footer_index: Optional[int] = None

    table_index: Optional[int] = None
    row_index: Optional[int] = None
    col_index: Optional[int] = None

    textbox_index: Optional[int] = None
    paragraph_index: Optional[int] = None

    contains_linebreaks: bool = False
    contains_tabs: bool = False
    may_contain_special_runs: bool = False


# ============================================================================
# Extractor
# ============================================================================

class DocumentStructureExtractor:
    def __init__(self, filepath, logger):
        self.filepath = filepath
        self.logger = logger
        self.ext = os.path.splitext(filepath)[1].lower()
        self.structure = None

    def extract(self):
        if self.ext == ".docx":
            self.structure = self._extract_docx_structure()
        else:
            raise ValueError("Unsupported file type. Use .docx only!")
        return self.structure

    # ------------------------------------------------------------------
    # DOCX
    # ------------------------------------------------------------------

    def _extract_docx_structure(self):
        doc = docx.Document(self.filepath)
        structure = {
            "type": "docx",
            "elements": []
        }

        elements: list[ExtractedElement] = []

        # 1. Paragraphs in main document
        for i, para in enumerate(doc.paragraphs, start=1):
            text = para.text or ""
            elements.append(ExtractedElement(
                type="paragraph",
                element_id=f"paragraph_{i}",
                text=text,
                empty=not bool(text.strip()),
                part_name="word/document.xml",
                container_path=f"/document/body/paragraph[{i}]",
                paragraph_index=i,
                contains_linebreaks="\n" in text,
                contains_tabs="\t" in text,
                may_contain_special_runs=self._paragraph_may_contain_special_runs(para),
            ))

        # 2. Tables in main document
        for ti, table in enumerate(doc.tables, start=1):
            for ri, row in enumerate(table.rows, start=1):
                for ci, cell in enumerate(row.cells, start=1):
                    cell_text = cell.text or ""
                    elements.append(ExtractedElement(
                        type="table_cell",
                        element_id=f"table_{ti}_cell_{ri}_{ci}",
                        text=cell_text.strip(),
                        empty=not bool(cell_text.strip()),
                        part_name="word/document.xml",
                        container_path=f"/document/body/table[{ti}]/row[{ri}]/cell[{ci}]",
                        table_index=ti,
                        row_index=ri,
                        col_index=ci,
                        contains_linebreaks="\n" in cell_text,
                        contains_tabs="\t" in cell_text,
                        may_contain_special_runs=True,  # försiktig default för tabellceller
                    ))

        # 3. Headers and footers per section
        for si, section in enumerate(doc.sections, start=1):
            header = section.header
            footer = section.footer

            if hasattr(header, "paragraphs"):
                for hi, para in enumerate(header.paragraphs, start=1):
                    text = para.text or ""
                    elements.append(ExtractedElement(
                        type="header",
                        element_id=f"header_s{si}_{hi}",
                        text=text,
                        empty=not bool(text.strip()),
                        part_name=f"word/header{si}.xml",
                        container_path=f"/header/paragraph[{hi}]",
                        section_index=si,
                        header_index=hi,
                        contains_linebreaks="\n" in text,
                        contains_tabs="\t" in text,
                        may_contain_special_runs=self._paragraph_may_contain_special_runs(para),
                    ))

            if hasattr(footer, "paragraphs"):
                for fi, para in enumerate(footer.paragraphs, start=1):
                    text = para.text or ""
                    elements.append(ExtractedElement(
                        type="footer",
                        element_id=f"footer_s{si}_{fi}",
                        text=text,
                        empty=not bool(text.strip()),
                        part_name=f"word/footer{si}.xml",
                        container_path=f"/footer/paragraph[{fi}]",
                        section_index=si,
                        footer_index=fi,
                        contains_linebreaks="\n" in text,
                        contains_tabs="\t" in text,
                        may_contain_special_runs=self._paragraph_may_contain_special_runs(para),
                    ))

        # 4. Textboxes (main document)
        textbox_counter = 1
        for pi, para in enumerate(doc.paragraphs, start=1):
            textboxes = self._extract_textboxes_from_paragraph(para)
            for tbx_local_index, box_info in enumerate(textboxes, start=1):
                textbox_text = box_info["text"]
                elements.append(ExtractedElement(
                    type="textbox",
                    element_id=f"textbox_{textbox_counter}",
                    text=textbox_text.strip(),
                    empty=not bool(textbox_text.strip()),
                    part_name="word/document.xml",
                    container_path=f"/document/body/paragraph[{pi}]/textbox[{tbx_local_index}]",
                    textbox_index=textbox_counter,
                    paragraph_index=pi,
                    contains_linebreaks="\n" in textbox_text,
                    contains_tabs="\t" in textbox_text,
                    may_contain_special_runs=True,
                ))
                textbox_counter += 1

        # 5. Footnotes
        for i, footnote_info in enumerate(self._extract_footnotes(), start=1):
            footnote_text = footnote_info["text"]
            xml_id = footnote_info["footnote_id"]

            elements.append(ExtractedElement(
                type="footnote",
                element_id=f"footnote_{i}",
                text=footnote_text.strip(),
                empty=not bool(footnote_text.strip()),
                part_name="word/footnotes.xml",
                container_path=f"/footnotes/footnote[@id='{xml_id}']",
                footnote_id=xml_id,
                contains_linebreaks="\n" in footnote_text,
                contains_tabs="\t" in footnote_text,
                may_contain_special_runs=True,
            ))

        structure["elements"] = [asdict(e) for e in elements]
        return structure

    def _paragraph_may_contain_special_runs(self, para) -> bool:
        try:
            xml = para._element
            if xml.find(".//w:footnoteReference", namespaces=NSMAP) is not None:
                return True
            if xml.find(".//w:commentReference", namespaces=NSMAP) is not None:
                return True
            if xml.find(".//w:drawing", namespaces=NSMAP) is not None:
                return True
            return False
        except Exception:
            return True

    def _extract_textboxes_from_paragraph(self, paragraph):
        """
        Returnerar textboxes i ett stycke med text + xml-referens.
        """
        textboxes = []

        drawing_elements = paragraph._element.findall(
            ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing"
        )

        for drawing in drawing_elements:
            t_elements = drawing.findall(
                ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"
            )
            full_text = "".join([t.text for t in t_elements if t.text])
            if full_text:
                textboxes.append({
                    "xml": drawing,
                    "text": full_text,
                })

        return textboxes

    def _extract_footnotes(self):
        footnotes = []

        try:
            with ZipFile(self.filepath) as docx_zip:
                if "word/footnotes.xml" not in docx_zip.namelist():
                    return footnotes

                footnotes_xml = docx_zip.read("word/footnotes.xml")
                tree = etree.fromstring(footnotes_xml)

                for footnote in tree.findall("w:footnote", NSMAP):
                    footnote_id = footnote.get(f"{{{W_NS}}}id")

                    # hoppa över separatorer
                    if footnote_id in ("-1", "0"):
                        continue

                    texts = footnote.findall(".//w:t", NSMAP)
                    full_text = "".join([t.text for t in texts if t.text])

                    if full_text.strip():
                        footnotes.append({
                            "footnote_id": footnote_id,
                            "text": full_text,
                        })

        except Exception as e:
            self.logger.warning(f"Could not extract footnotes: {e}")

        return footnotes

    # ------------------------------------------------------------------
    # Hjälpmetoder för fortsatt pipeline
    # ------------------------------------------------------------------

    @staticmethod
    def _extract_docx_elements(filepath):
        """
        Behålls tills vidare för bakåtkompatibilitet med äldre editorer.
        """
        elements = {}
        doc = docx.Document(filepath)

        for i, para in enumerate(doc.paragraphs, start=1):
            elements[f"paragraph_{i}"] = para

        for ti, table in enumerate(doc.tables, start=1):
            for ri, row in enumerate(table.rows, start=1):
                for ci, cell in enumerate(row.cells, start=1):
                    elements[f"table_{ti}_cell_{ri}_{ci}"] = cell

        for si, section in enumerate(doc.sections, start=1):
            if hasattr(section.header, "paragraphs"):
                for hi, para in enumerate(section.header.paragraphs, start=1):
                    elements[f"header_s{si}_{hi}"] = para

            if hasattr(section.footer, "paragraphs"):
                for fi, para in enumerate(section.footer.paragraphs, start=1):
                    elements[f"footer_s{si}_{fi}"] = para

        textbox_counter = 1
        extractor = DocumentStructureExtractor(filepath, logger=_NullLogger())
        for para in doc.paragraphs:
            textboxes = extractor._extract_textboxes_from_paragraph(para)
            for box in textboxes:
                elements[f"textbox_{textbox_counter}"] = box["xml"]
                textbox_counter += 1

        for i, footnote in enumerate(extractor._extract_footnotes(), start=1):
            elements[f"footnote_{i}"] = footnote

        return doc, elements

    def save_as_json(self, output_path=None):
        if not self.structure:
            self.extract()
        if not output_path:
            output_path = self.filepath + "_structure.json"
        try:
            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(self.structure, f, indent=2, ensure_ascii=False)
            return output_path
        except Exception as e:
            self.logger.error(f"Error saving JSON structure: {str(e)}")
            return None


class _NullLogger:
    def warning(self, msg): pass
    def error(self, msg): pass
    def info(self, msg): pass
    def debug(self, msg): pass


def main():
    import logging

    if len(sys.argv) != 2:
        print(f"Usage: python {os.path.basename(__file__)} <docx document>")
        sys.exit(1)

    filepath = sys.argv[1]

    logger = logging.getLogger("extractor-test")
    logger.setLevel(logging.INFO)
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger.handlers.clear()
    logger.addHandler(handler)

    extractor = DocumentStructureExtractor(filepath, logger)
    extractor.extract()
    output_json = extractor.save_as_json()
    print(f"Structure saved to: {output_json}")


if __name__ == "__main__":
    main()