import os
import json
import docx
import fitz
import sys
from lxml import etree
from zipfile import ZipFile

class DocumentStructureExtractor:
    def __init__(self, filepath, logger):
        self.filepath = filepath
        self.logger = logger
        self.ext = os.path.splitext(filepath)[1].lower()
        self.structure = None

    def extract(self):
        if self.ext == ".docx":
            self.structure = self._extract_docx_texts()
        elif self.ext == ".pdf":
            self.structure = self._extract_pdf()
        else:
            raise ValueError("Unsupported file type. Use .docx or .pdf")
        return self.structure

    def _extract_docx_texts_simple(self):
        doc = docx.Document(self.filepath)
        structure = {
            "type": "docx",
            "paragraphs": []
        }

        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            structure["paragraphs"].append({
                "paragraph": len(structure["paragraphs"]) + 1,  # visible index
                "text": text,
                "empty": not bool(text)
            })

        return structure
    
    def _extract_docx_texts(self):
        doc = docx.Document(self.filepath)
        structure = {
            "type": "docx",
            "elements": []
        }

        # Paragraphs
        for i, para in enumerate(doc.paragraphs):
            text = para.text.strip()
            structure["elements"].append({
                "type": "paragraph",
                "element_id": f"paragraph_{i+1}",
                "text": text,
                "empty": not bool(text)
            })

        # Tables
        for ti, table in enumerate(doc.tables):
            for ri, row in enumerate(table.rows):
                for ci, cell in enumerate(row.cells):
                    cell_text = cell.text.strip()
                    structure["elements"].append({
                        "type": "table_cell",
                        "element_id": f"table_{ti+1}_cell_{ri+1}_{ci+1}",
                        "text": cell_text,
                        "empty": not bool(cell_text)
                    })

        # Headers and footers
        if hasattr(doc.sections[0].header, "paragraphs"):
            for hi, para in enumerate(doc.sections[0].header.paragraphs):
                text = para.text.strip()
                structure["elements"].append({
                    "type": "header",
                    "element_id": f"header_{hi+1}",
                    "text": text,
                    "empty": not bool(text)
                })

        if hasattr(doc.sections[0].footer, "paragraphs"):
            for fi, para in enumerate(doc.sections[0].footer.paragraphs):
                text = para.text.strip()
                structure["elements"].append({
                    "type": "footer",
                    "element_id": f"footer_{fi+1}",
                    "text": text,
                    "empty": not bool(text)
                })
                
         # Textboxes
        textbox_counter = 1
        for para in doc.paragraphs:
            textbox_texts = self._extract_textbox_texts_from_paragraph(para)
            for txt in textbox_texts:
                structure["elements"].append({
                    "type": "textbox",
                    "element_id": f"textbox_{textbox_counter}",
                    "text": txt.strip(),
                    "empty": not bool(txt.strip())
                })
                textbox_counter += 1
                
        # Footnotes (if implemented via XML parsing or external lib)
        for i, (text, xml_id) in enumerate(self._extract_footnote_texts()):
             structure["elements"].append({
                 "type": "footnote",
                 "element_id": f"footnote_{i+1}",
                 "footnote_id": xml_id,
                 "text": text,
                 "empty": not bool(text)
             })

        return structure
    
    @staticmethod
    def _extract_docx_elements(filepath):
        
        elements = {}
        
        doc = docx.Document(filepath)
    
        # Paragraphs
        for i, para in enumerate(doc.paragraphs):
            elements[f"paragraph_{i+1}"] = para

        # Tables
        for ti, table in enumerate(doc.tables):
            for ri, row in enumerate(table.rows):
                for ci, cell in enumerate(row.cells):
                    elements[f"table_{ti+1}_cell_{ri+1}_{ci+1}"] = cell
        
        # Headers and footers
        if hasattr(doc.sections[0].header, "paragraphs"):
            for i, para in enumerate(doc.sections[0].header.paragraphs):
                elements[f"header_{i+1}"] = para

        if hasattr(doc.sections[0].footer, "paragraphs"):
            for fi, para in enumerate(doc.sections[0].footer.paragraphs):
                elements[f"footer_{i+1}"] = para
                
         # Textboxes
        textbox_counter = 1
        for para in doc.paragraphs:
            textboxes = DocumentStructureExtractor._extract_textboxes_from_paragraph(para)
            for box in textboxes:
                elements[f"textbox_{textbox_counter}"] = box
                textbox_counter += 1
                
        # Footnotes (if implemented via XML parsing or external lib)
        for i, footnote in enumerate(DocumentStructureExtractor._extract_footnotes(filepath)):
            elements[f"footnote_{i+1}"] = footnote 

        return doc, elements

    def _extract_textbox_texts_from_paragraph(self, paragraph):
        """
        Extract text from textboxes in the given paragraph.
        This version avoids passing 'namespaces' directly by using fully qualified names.
        """
        textbox_texts = []

        # Search for <w:drawing> elements
        drawing_elements = paragraph._element.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing")

        for drawing in drawing_elements:
            # Look for all <w:t> elements under <w:txbxContent>
            t_elements = drawing.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")
            full_text = " ".join([t.text for t in t_elements if t.text])
            if full_text:
                textbox_texts.append(full_text)

        return textbox_texts
        
    @staticmethod
    def _extract_textboxes_from_paragraph(paragraph):
        """
        Extract textboxes in the given paragraph.
        This version avoids passing 'namespaces' directly by using fully qualified names.
        """
        textboxes = []

        # Search for <w:drawing> elements
        drawing_elements = paragraph._element.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing")

        for drawing in drawing_elements:
            # Look for all <w:t> elements under <w:txbxContent>
            t_elements = drawing.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")
            full_text = " ".join([t.text for t in t_elements if t.text])
            if full_text:
                textboxes.append(drawing)

        return textboxes

    def _extract_footnote_texts(self):
        footnotes = []

        try:
            # Open the docx as a zip file
            with ZipFile(self.filepath) as docx_zip:
                if "word/footnotes.xml" not in docx_zip.namelist():
                    return footnotes  # No footnotes present

                # Read footnotes.xml
                footnotes_xml = docx_zip.read("word/footnotes.xml")
                tree = etree.fromstring(footnotes_xml)

                # Define namespaces
                namespaces = {
                    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                }

                # Find all w:footnote elements
                for footnote in tree.findall("w:footnote", namespaces):
                    footnote_id = footnote.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id")
                    
                    # Skip footnote types like separators (-1, 0)
                    if footnote_id in ("-1", "0"):
                        continue

                    # Get all text within the footnote
                    texts = footnote.findall(".//w:t", namespaces)
                    full_text = "".join([t.text for t in texts if t.text])

                    if full_text.strip():
                        footnotes.append((full_text.strip(), footnote_id))

        except Exception as e:
            self.logger.warning(f"Could not extract footnotes: {e}")

        return footnotes
        
    @staticmethod
    def _extract_footnotes(filepath):
        footnotes = []

        try:
            # Open the docx as a zip file
            with ZipFile(filepath) as docx_zip:
                if "word/footnotes.xml" not in docx_zip.namelist():
                    return footnotes  # No footnotes present

                # Read footnotes.xml
                footnotes_xml = docx_zip.read("word/footnotes.xml")
                tree = etree.fromstring(footnotes_xml)

                # Define namespaces
                namespaces = {
                    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                }

                # Find all w:footnote elements
                for footnote in tree.findall("w:footnote", namespaces):
                    footnote_id = footnote.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id")
                    
                    # Skip footnote types like separators (-1, 0)
                    if footnote_id in ("-1", "0"):
                        continue
                    else:

                        # Get all text within the footnote
                        texts = footnote.findall(".//w:t", namespaces)
                        full_text = "".join([t.text for t in texts if t.text])

                        # If the footnote has text, use it
                        if full_text.strip():
                            footnotes.append(footnote)

        except Exception as e:
            self.logger.warning(f"Could not extract footnotes: {e}")

        return footnotes


    def _extract_pdf(self):
        doc = fitz.open(self.filepath)
        structure = {"type": "pdf", "pages": []}

        for page_index, page in enumerate(doc, start=1):
            blocks = sorted(page.get_text("blocks"), key=lambda b: -b[1])  # sort top-down
            lines = [
                {"line": i + 1, "text": b[4].strip()}
                for i, b in enumerate(blocks)
                if b[4].strip()
            ]
            structure["pages"].append({
                "page": page_index,
                "lines": lines
            })

        doc.close()
        return structure

    def save_as_json(self, output_path = None):
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
    
def main():
    import logging
    
    if len(sys.argv) != 2:
        print(f"Usage: python {os.path.basename(__file__)} <docx or pdf document to generate JSON structure for>")
        sys.exit(1)
    
    filepath = sys.argv[1]
    
    # Set up logger
    logger = logging.getLogger("extratctor-test")
    logger.setLevel(logging.INFO)
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger.handlers.clear()
    logger.addHandler(handler)
    
    # Construct Extractor object
    extractor = DocumentStructureExtractor(filepath, logger)
    extractor.extract()
    output_json = extractor.save_as_json()
    print(f"Structure saved to: {output_json}")
    
if __name__ == "__main__":
    main()
