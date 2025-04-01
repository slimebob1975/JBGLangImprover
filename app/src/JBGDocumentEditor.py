from docx import Document
from docx.shared import RGBColor
import os, sys
import json
import fitz # PyMuPDF
import re

DEBUG = False

class JBGDocumentEditor:
    
    def __init__(self, filepath, changes_path):
        """
        :param filepath: Path to the input document (.docx or .pdf)
        :param changes_path: Path to the suggested changes JSON file
        """
        self.filepath = filepath
        self.changes = self._get_changes_from_json(changes_path)  
        self.ext = os.path.splitext(filepath)[1].lower()
        self.edited_document = None

    @staticmethod
    def _get_changes_from_json(json_filepath):
        with open(json_filepath, 'r', encoding='utf-8') as f:
            return json.load(f)    
    
    def apply_changes(self):
        if self.ext == ".docx":
            self.edited_document = self._edit_docx()
        elif self.ext == ".pdf":
            self.edited_document = self._annotate_pdf()
        else:
            raise ValueError("Unsupported file type. Use .docx or .pdf")

    def save_edited_document(self, output_path=None):
            
        if self.edited_document:
            if self.ext == ".pdf":
                if output_path is None:
                    output_path = self.filepath.replace(".pdf", "_annotated.pdf")
                self.edited_document.save(output_path, garbage=4, deflate=True)
                self.edited_document.close()
            elif self.ext == ".docx":
                if output_path is None:
                    output_path = self.filepath.replace(".docx", "_edited.docx")
                self.edited_document.save(output_path)
            else:
                raise ValueError("Unsupported file type. Use.docx or.pdf")
            return output_path

    def _edit_docx(self):
        
        doc = Document(self.filepath)
        self._apply_to_empty_paragraphs(doc)
        
        applied_changes = set()

        for idx, para in enumerate(doc.paragraphs):
            para_index = idx + 1
            paragraph_text = para.text

            # Collect all changes for this paragraph
            applicable_changes = [c for c in self.changes if c.get("paragraph") == para_index]
            if not applicable_changes:
                continue

            current_text = paragraph_text
            rebuilt = [("text", current_text)]  # initial unstyled

            for change in applicable_changes:
                old, new = change["old"], change["new"]
                norm_old = self._normalize_text(old)

                new_rebuilt = []
                for part_type, part_text in rebuilt:
                    if part_type != "text":
                        new_rebuilt.append((part_type, part_text))
                        continue

                    norm_part = self._normalize_text(part_text)

                    if norm_old not in norm_part:
                        new_rebuilt.append((part_type, part_text))
                        continue

                    # Fuzzy split, but keep exact positions from original text
                    split_parts = re.split(re.escape(old), part_text)
                    for i, seg in enumerate(split_parts):
                        if seg:
                            new_rebuilt.append(("text", seg))
                        if i < len(split_parts) - 1:
                            new_rebuilt.append(("strike", old))
                            new_rebuilt.append(("insert", new))

                rebuilt = new_rebuilt
                #print(f"✅ Applied: '{old}' → '{new}' in paragraph {para_index}")

                applied_changes.add((para_index, old))

            # Clear original runs
            for _ in range(len(para.runs)):
                para._element.remove(para.runs[0]._element)

            # Add styled runs
            for kind, val in rebuilt:
                run = para.add_run(val)
                if kind == "strike":
                    run.font.strike = True
                    run.font.color.rgb = RGBColor(255, 0, 0)
                elif kind == "insert":
                    run.font.color.rgb = RGBColor(0, 128, 0)

        self._suggest_nearby_paragraphs(doc)
        
        for change in self.changes:
            key = (change.get("paragraph"), change.get("old"))
            if key not in applied_changes:
                print(f"⚠️ No match found for '{change['old']}' in paragraph {change['paragraph']}.")

        return doc

    def _suggest_nearby_paragraphs(self, doc):
        total = len(doc.paragraphs)
        for change in self.changes:
            if "paragraph" in change:
                para_num = change["paragraph"]
                if not (1 <= para_num <= total):
                    print(f"Warning: Paragraph {para_num} is out of range.")
                    continue
                expected_para = doc.paragraphs[para_num - 1].text
                if change["old"] not in expected_para:
                    for offset in [-2, -1, 1, 2]:
                        alt_idx = para_num - 1 + offset
                        if 0 <= alt_idx < total:
                            if change["old"] in doc.paragraphs[alt_idx].text:
                                print(f"Could not find '{change['old']}' in paragraph {para_num} — did you mean paragraph {alt_idx + 1}?")
                                break

    def _apply_to_empty_paragraphs(self, doc):
        for idx, para in enumerate(doc.paragraphs):
            para_index = idx + 1
            if para.text.strip() != "":
                continue

            for change in self.changes:
                if not isinstance(change, dict):
                    print(f"Skipping invalid change (not a dict): {change}")
                    continue
                elif (
                    change.get("paragraph") == para_index
                    and change.get("old", "") == ""
                    and change.get("new")
                ):
                    para.text = change["new"]
                    print(f"Filled empty paragraph {para_index} with: {change['new']}")

    @staticmethod
    def _normalize_text(text):
        # Replace all whitespace (tabs, newlines, etc.) with single spaces
        return re.sub(r'\s+', ' ', text).strip()
    
    @staticmethod
    def _clean_pdf_text(text):

        # Remove digits stuck to end of words: deltidsarbete15 → deltidsarbete
        text = re.sub(r"(?<=[a-zA-Z])\d{1,3}\b", "", text)

        # Remove digits at end of word before punctuation: deltidsarbete15. → deltidsarbete.
        text = re.sub(r"(?<=[a-zA-Z])\d{1,3}(?=[.,])", "", text)

        return text.strip()
    
    def _annotate_pdf_simple(self):
    
        doc = fitz.open(self.filepath)

        for change in self.changes:
            if "page" not in change or "old" not in change:
                continue

            page_index = change["page"] - 1
            if page_index >= len(doc):
                continue

            page = doc[page_index]
            rects = page.search_for(change["old"])
            if not rects:
                print(f"No match found for '{change['old']}' on page {change['page']}")
                continue

            for rect in rects:
                highlight = page.add_highlight_annot(rect)
                if "new" in change:
                    highlight.set_info(content=f"Suggestion: replace with '{change['new']}'")

        return doc
    
    def _annotate_pdf(self):
        doc = fitz.open(self.filepath)

        for change in self.changes:
            if "page" not in change or "old" not in change:
                continue

            page_index = change["page"] - 1
            if page_index >= len(doc):
                continue

            old = change["old"]
            new = change.get("new", "")
            target_line = change.get("line")
            page = doc[page_index]

            # Get all text blocks on page, sorted top-down
            blocks = sorted(page.get_text("blocks"), key=lambda b: -b[1])  # y1 descending
            line_texts = [(i + 1, b, b[4].strip()) for i, b in enumerate(blocks) if b[4].strip()]

            match_found = False

            # Try target line first
            if target_line:
                for line_no, block, text in line_texts:
                    if line_no == target_line and (self._normalize_text(old) in self._normalize_text(self._clean_pdf_text(text)) 
                                                   or self._normalize_text(self._clean_pdf_text(text)) in self._normalize_text(self._clean_pdf_text(old))):
                        rects = page.search_for(old, clip=fitz.Rect(block[:4]))
                        for rect in rects:
                            highlight = page.add_highlight_annot(rect)
                            if new:
                                highlight.set_info(content=f"{new}")
                        match_found = True
                        break

                # If not found, try nearby lines
                if not match_found:
                    for line_no, block, text in line_texts:
                        if abs(line_no - target_line) <= 2 and (self._normalize_text(old) in self._normalize_text(self._clean_pdf_text(text)) 
                                                  or self._normalize_text(self._clean_pdf_text(text)) in self._normalize_text(self._clean_pdf_text(old))):
                            rects = page.search_for(old, clip=fitz.Rect(block[:4]))
                            for rect in rects:
                                highlight = page.add_highlight_annot(rect)
                                if new:
                                    highlight.set_info(content=f"{new}")
                            print(f"Could not find '{old}' on page {change['page']}, line {target_line} — did you mean line {line_no}?")
                            match_found = True
                            break
            else:
                # Fallback if no line provided
                rects = page.search_for(old)
                for rect in rects:
                    highlight = page.add_highlight_annot(rect)
                    if new:
                        highlight.set_info(content=f"{new}")
                match_found = bool(rects)

            if not match_found:
                print(f"No match found for '{old}' on page {change['page']}.")

        self._deduplicate_annotations(doc)
        return doc
    
    def _deduplicate_annotations(self, doc, distance_threshold=5):
        
        num_duplicates = 0
        for page in doc:
            existing = []
            to_remove = []

            for annot in page.annots():
                if annot.type[0] != 8:  # 8 = Highlight, 1 = FreeText
                    continue

                rect = annot.rect
                content = (annot.info.get("content") or "").strip()

                is_duplicate = False
                for seen_rect, seen_text in existing:
                    if (
                        content == seen_text or
                        (abs(rect.x0 - seen_rect.x0) < distance_threshold and
                        abs(rect.y0 - seen_rect.y0) < distance_threshold and
                        abs(rect.x1 - seen_rect.x1) < distance_threshold and
                        abs(rect.y1 - seen_rect.y1) < distance_threshold)
                    ):
                        is_duplicate = True
                        break

                if is_duplicate:
                    to_remove.append(annot)
                else:
                    existing.append((rect, content))

                num_duplicates += len(to_remove)
            for annot in to_remove:
                page.delete_annot(annot)

        print(f"Removed {num_duplicates} duplicate annotations.")
        return

def main():
    if len(sys.argv) != 3:
        print(f"Usage: python {os.path.basename(__file__)} <document_path> <changes_json_path>")
        sys.exit(1)

    doc_path = sys.argv[1]
    changes_path = sys.argv[2]

    try:
        editor = JBGDocumentEditor(doc_path, changes_path)
        editor.apply_changes()
        output_path = editor.save_edited_document()
        print(f"Document processed and saved to: {output_path}")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
