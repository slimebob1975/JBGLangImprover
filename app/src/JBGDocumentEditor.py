import docx
from docx.shared import RGBColor
from docx.enum.text import WD_UNDERLINE
import datetime
import os
import sys
import uuid
import shutil
import json
import fitz # PyMuPDF
import re
import zipfile
from lxml import etree
from tempfile import mkdtemp

DEBUG = False
SUGGESTION = "Förslag"
MOTIVATION = "Motivering"
NSMAP = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

class JBGDocumentEditor:
    
    def __init__(self, filepath, changes_path, include_comments, docx_mode, logger):
        """
        :param filepath: Path to the input document (.docx or .pdf)
        :param changes_path: Path to the suggested changes JSON file
        """
        self.filepath = filepath
        self.changes = self._get_changes_from_json(changes_path)  
        self.logger = logger
        self.include_comments = include_comments
        self.docx_mode = docx_mode  # "simple" or "tracked"
        self.ext = os.path.splitext(filepath)[1].lower()
        self.edited_document = None
        self.nsmap = NSMAP

    @staticmethod
    def _get_changes_from_json(json_filepath):
        with open(json_filepath, 'r', encoding='utf-8') as f:
            return json.load(f)    
    
    import os

    def apply_changes(self):
        try:
            # Make a working copy
            original_path = self.filepath
            basename = os.path.basename(original_path)
            temp_name = f"{uuid.uuid4()}_{basename}"
            temp_dir = "uploads" if os.name == "nt" else "/tmp"
            temp_path = os.path.join(temp_dir, temp_name)

            shutil.copyfile(original_path, temp_path)
            self.filepath = temp_path  # Redirect all edits to the temp copy

            # Choose method based on file type
            if self.filepath.endswith(".docx"):
                if self.docx_mode == "tracked":
                    self.edited_document = self._edit_docx_tracked()
                else:
                    self.edited_document = self._edit_docx()
            elif self.filepath.endswith(".pdf"):
                self.edited_document = self._annotate_pdf()
            else:
                raise ValueError("Unsupported file format. Use .docx or .pdf!")

        except Exception as e:
            self.logger.error(f"❌ Failed to apply changes: {e}")
            raise

    def save_edited_document(self, output_path=None):

        if output_path is None:
            base, ext = os.path.splitext(self.filepath)
            suffix = "_edited" if ext.lower() == ".docx" else "_annotated"
            output_path = f"{base}{suffix}{ext}"

        if self.edited_document:
            try:
                self.edited_document.save(output_path)
                self.logger.info(f"💾 Saved document with visual edits or tracked changes to: {output_path}")
            except Exception as e:
                self.logger.error(f"❌ Failed to save edited document: {e}")
                raise
        else:
            try:
                if self.filepath != output_path:
                    shutil.copyfile(self.filepath, output_path)
                    self.logger.info(f"📝 Returned already-saved document with comments or annotations: {output_path}")
            except Exception as e:
                self.logger.error(f"❌ Failed to copy saved file: {e}")
                raise

        return output_path

    def _edit_docx(self):
        
        doc = docx.Document(self.filepath)
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
                self.logger.info(f"✅ Applied: '{old}' → '{new}' in paragraph {para_index}")

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
                    if self.docx_mode == "tracked":
                        run.font.color.rgb = RGBColor(255, 0, 0)
                        run.font.underline = WD_UNDERLINE.SINGLE
                    else:
                        run.font.color.rgb = RGBColor(0, 128, 0)

                    # Attach comment if available
                    if self.include_comments:
                        for change in applicable_changes:
                            if change["new"] == val and "motivation" in change:
                                run.add_comment(
                                    text=change["motivation"],
                                    author="JBG klarspråkningstjänst",
                                    initials="JBG"
                                )

        self._suggest_nearby_paragraphs(doc)
        
        for change in self.changes:
            key = (change.get("paragraph"), change.get("old"))
            if key not in applied_changes:
                self.logger.error(f"❌No match found for '{change['old']}' in paragraph {change['paragraph']}.")

        return doc
    
    def _suggest_nearby_paragraphs(self, doc):
        total = len(doc.paragraphs)
        for change in self.changes:
            if "paragraph" in change:
                para_num = change["paragraph"]
                if not (1 <= para_num <= total):
                    self.logger.warning(f"Warning: Paragraph {para_num} is out of range.")
                    continue
                expected_para = doc.paragraphs[para_num - 1].text
                if change["old"] not in expected_para:
                    for offset in [-2, -1, 1, 2]:
                        alt_idx = para_num - 1 + offset
                        if 0 <= alt_idx < total:
                            if change["old"] in doc.paragraphs[alt_idx].text:
                                self.logger.warning(f"Could not find '{change['old']}' in paragraph {para_num} — did you mean paragraph {alt_idx + 1}?")
                                break

    def _apply_to_empty_paragraphs(self, doc):
        for idx, para in enumerate(doc.paragraphs):
            para_index = idx + 1
            if para.text.strip() != "":
                continue

            for change in self.changes:
                if not isinstance(change, dict):
                    self.logger.error(f"Skipping invalid change (not a dict): {change}")
                    continue
                elif (
                    change.get("paragraph") == para_index
                    and change.get("old", "") == ""
                    and change.get("new")
                ):
                    para.text = change["new"]
                    self.logger.info(f"Filled empty paragraph {para_index} with: {change['new']}")

    
    def _edit_docx_tracked(self):
        """
        Applies simple markup first, saves to a temporary file,
        then converts that markup to native tracked changes.
        """
        # First, apply standard markup
        self.edited_document = self._edit_docx()

        # Save the interim version to disk
        temp_base, _ = os.path.splitext(self.filepath)
        intermediate_path = f"{temp_base}_intermediate.docx"
        self.edited_document.save(intermediate_path)
        self.logger.info(f"📄 Saved intermediate docx with markup to: {intermediate_path}")

        # Convert to tracked changes based on visual styles
        tracked_doc_path = self._convert_markup_to_tracked(intermediate_path)

        # Re-open and return final document for saving
        final_doc = docx.Document(tracked_doc_path)
        return final_doc

    
    def _convert_markup_to_tracked(self, input_docx_path):
        """
        Post-processes a .docx file to wrap colored/struck/underlined text into native Word
        tracked-change XML tags: <w:del> and <w:ins>.
        Returns the filepath of the modified document.
        """
        try:
            temp_dir = mkdtemp()
            tracked_docx_path = os.path.join(temp_dir, os.path.basename(input_docx_path).replace("_intermediate", "_tracked"))

            with zipfile.ZipFile(input_docx_path, 'r') as zin:
                zin.extractall(temp_dir)

            document_xml_path = os.path.join(temp_dir, "word", "document.xml")
            parser = etree.XMLParser(remove_blank_text=True)
            with open(document_xml_path, "rb") as f:
                tree = etree.parse(f, parser)

            for run in tree.xpath("//w:r", namespaces=self.nsmap):
                rpr = run.find("w:rPr", namespaces=self.nsmap)
                if rpr is None:
                    continue

                color = rpr.find("w:color", namespaces=self.nsmap)
                strike = rpr.find("w:strike", namespaces=self.nsmap)
                underline = rpr.find("w:u", namespaces=self.nsmap)

                color_val = color.get(f"{{{self.nsmap['w']}}}val") if color is not None else None
                is_strike = strike is not None
                is_underline = underline is not None

                parent = run.getparent()
                if color_val == "FF0000" and is_strike:
                    wrapper = etree.Element(f"{{{self.nsmap['w']}}}del")
                    wrapper.set(f"{{{self.nsmap['w']}}}author", "JBG")
                    wrapper.set(f"{{{self.nsmap['w']}}}date", datetime.utcnow().isoformat())
                    parent.replace(run, wrapper)
                    wrapper.append(run)
                elif color_val == "FF0000" and is_underline:
                    wrapper = etree.Element(f"{{{self.nsmap['w']}}}ins")
                    wrapper.set(f"{{{self.nsmap['w']}}}author", "JBG")
                    wrapper.set(f"{{{self.nsmap['w']}}}date", datetime.utcnow().isoformat())
                    parent.replace(run, wrapper)
                    wrapper.append(run)

            with open(document_xml_path, "wb") as f:
                tree.write(f, pretty_print=True, xml_declaration=True, encoding="UTF-8")

            with zipfile.ZipFile(tracked_docx_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for root, _, files in os.walk(temp_dir):
                    for filename in files:
                        filepath = os.path.join(root, filename)
                        archive_name = os.path.relpath(filepath, temp_dir)
                        zout.write(filepath, archive_name)

            self.logger.info(f"🔁 Converted markup to tracked changes in: {tracked_docx_path}")
            return tracked_docx_path
        except Exception as ex:
            self.logger.warning(f"🚧 Tracked changes generation not fully functioning yet: {ex}.")
            return input_docx_path
    
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
            motivation = change.get("motivation", "")
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
                                if self.include_comments:
                                    highlight.set_info(content=f"{SUGGESTION}: {new} \n\n{MOTIVATION}: {motivation}")
                                else:
                                    highlight.set_info(content=f"{SUGGESTION}: {new}")
                        match_found = True
                        self.logger.info(f"✅ Applied: '{old}' → '{new}' on page {page_index + 1}, line {target_line + 1}")
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
                                    if self.include_comments:
                                        highlight.set_info(content=f"{SUGGESTION}: {new} \n\n{MOTIVATION}: {motivation}")
                                    else:
                                        highlight.set_info(content=f"{SUGGESTION}: {new}")
                            self.logger.warning(f"Could not find '{old}' on page {change['page']}, line {target_line} — did you mean line {line_no}?")
                            match_found = True
                            break
            else:
                # Fallback if no line provided
                rects = page.search_for(old)
                for rect in rects:
                    highlight = page.add_highlight_annot(rect)
                    if new:
                        if self.include_comments:
                            highlight.set_info(content=f"{SUGGESTION}: {new} \n\n{MOTIVATION}: {motivation}")
                        else:
                            highlight.set_info(content=f"{SUGGESTION}: {new}")
                match_found = bool(rects)

            if not match_found:
                self.logger.error(f"❌No match found for '{old}' on page {change['page']}.")

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

        self.logger.info(f"Removed {num_duplicates} duplicate annotations.")
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
