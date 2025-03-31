import os
import json
from docx import Document
import fitz  # PyMuPDF
import sys

class DocumentStructureExtractor:
    def __init__(self, filepath):
        self.filepath = filepath
        self.ext = os.path.splitext(filepath)[1].lower()
        self.structure = None

    def extract(self):
        if self.ext == ".docx":
            self.structure = self._extract_docx()
        elif self.ext == ".pdf":
            self.structure = self._extract_pdf()
        else:
            raise ValueError("Unsupported file type. Use .docx or .pdf")
        return self.structure

    def _extract_docx(self):
        doc = Document(self.filepath)
        structure = {
            "type": "docx",
            "paragraphs": []
        }
        for i, para in enumerate(doc.paragraphs, start=1):
            text = para.text.strip()
            structure["paragraphs"].append({
                "paragraph": i,
                "text": text,
                "empty": not bool(text)
            })
        return structure

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
            print(f"Error saving JSON structure: {str(e)}")
            return None
    
def main():
    
    if len(sys.argv) != 2:
        print(f"Usage: python {os.path.basename(__file__)} <docx or pdf document to generate JSON structure for>")
        sys.exit(1)
    
    filepath = sys.argv[1]
    extractor = DocumentStructureExtractor(filepath)
    extractor.extract()
    output_json = extractor.save_as_json()
    print(f"Structure saved to: {output_json}")
    
if __name__ == "__main__":
    main()
