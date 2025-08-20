import docx
from docx.shared import RGBColor
from docx.table import _Cell
from datetime import datetime, timezone
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
import tempfile
from zipfile import ZipFile

from copy import deepcopy
from difflib import ndiff
from thefuzz import fuzz
try:
    from app.src.JBGDocumentStructureExtractor import DocumentStructureExtractor
except ModuleNotFoundError:
    from JBGDocumentStructureExtractor import DocumentStructureExtractor
    
# Settings
DEBUG = True
SUGGESTION = "Förslag"
MOTIVATION = "Motivering"
NSMAP = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
TEXT_SIM_SCORE_THRESHOLD = 90.0

class JBGDocumentEditor:
    
    def __init__(self, filepath, changes_path, include_motivations, logger):
        """
        :param filepath: Path to the input document (.docx or .pdf)
        :param changes_path: Path to the suggested changes JSON file
        """
        self.filepath = filepath
        self.changes = self._get_changes_from_json(changes_path)  
        self.logger = logger
        self.include_motivations = include_motivations
        self.footnote_changes = []
        self.ext = os.path.splitext(filepath)[1].lower()
        self.edited_document = None
        self.nsmap = NSMAP

    @staticmethod
    def _get_changes_from_json(json_filepath):
        with open(json_filepath, 'r', encoding='utf-8') as f:
            return json.load(f)    

    def apply_changes(self):
        try:
            # Choose method based on file type
            if self.filepath.endswith(".docx"):
                self._apply_changes_docx()
            elif self.filepath.endswith(".pdf"):
                self._apply_changes_pdf()
            else:
                raise ValueError("Unsupported file format. Use .docx or .pdf!")

        except Exception as e:
            self.logger.error(f"❌ Failed to apply changes: {e}")
            raise
        
    def _apply_changes_pdf(self):
        self.edited_document = self._annotate_pdf()
        
    def _apply_changes_docx(self):
        
        # Step 1: Apply visual markup
        markup_doc = self._edit_docx()

        # Step 2: Save intermediate version (used in both modes)
        intermediate_path = self._save_edited_document(doc=markup_doc, suffix="_intermediate")
        self.filepath = intermediate_path  # point everything to this file

        # Step 3: Patch footnotes.xml inside the ZIP
        if self.footnote_changes:
            self._edit_footnote_texts()
        
        self.edited_document = docx.Document(intermediate_path)
 
    def _save_edited_document(self, doc=None, suffix="_edited"):
        base, ext = os.path.splitext(self.filepath)
        output_path = f"{base}{suffix}{ext}"
        doc = doc or self.edited_document
        doc.save(output_path)
        self.logger.info(f"💾 Saved: {output_path}")
        return output_path

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
        doc, elements = DocumentStructureExtractor._extract_docx_elements(self.filepath)
        applied_changes = set()

        for change in self.changes:
            element_id = change.get("element_id")
            old = change.get("old")
            new = change.get("new")
            motivation = change.get("motivation")

            if not element_id or element_id not in elements:
                self.logger.warning(f"⚠️ Skipping change — unknown or missing element_id: {element_id}")
                continue

            element = elements[element_id]
            if DEBUG: self.logger.debug(f" -- Considering element: {str(element)} with id {element_id}")

            # Standard Word elements with .text and .runs
            if hasattr(element, "text") and hasattr(element, "runs"):
                if DEBUG: self.logger.debug(f" -- Element has text and runs attribute")
                original_text = element.text  or ""
                if not original_text and isinstance(element, etree._Element):
                    original_text = self._get_joined_text_from_xml_element(element)
                if DEBUG: self.logger.debug(f" -- Original text was: {original_text}")

                normalized_old = self._normalize_text(old)
                normalized_origin = self._normalize_text(original_text)
                if normalized_old not in normalized_origin:
                    self.logger.debug(f"⚠️ Not identical match for '{old}' and '{original_text}' in {element_id}")
                    sim_score = fuzz.ratio(normalized_old, normalized_origin)
                    self.logger.debug(f"⚠️ Comparison similarity score: {sim_score} %")
                    if float(sim_score) < TEXT_SIM_SCORE_THRESHOLD:
                        self.logger.error(f"❌ No enough match for '{old}' in {element_id} (fuzzy match similarity score: {sim_score}")
                        continue
                
                # Build diff
                diffed = self._diff_words(old, new)
                
                # Capture footnoteReferences before clearing runs
                footnote_refs = []
                for run in element.runs:
                    xml_run = run._element
                    if xml_run.find("w:footnoteReference", namespaces=self.nsmap) is not None:
                        footnote_refs.append(deepcopy(xml_run))

                # Clear runs
                for _ in range(len(element.runs)):
                    element._element.remove(element.runs[0]._element)

                # Add formatted runs
                for kind, val in diffed:
                    run = element.add_run(val)
                    if kind == "strike":
                        run.font.strike = True
                        run.font.color.rgb = RGBColor(255, 0, 0)
                    elif kind == "insert":
                        run.font.color.rgb = RGBColor(0, 128, 0)

                # Rebuild footnote references
                for ref in footnote_refs:
                    element._element.append(ref)
                
                # Add motivation comment
                if self.include_motivations and motivation:
                    if "footer" in element_id or "header" in element_id:
                        self.logger.warning(f"⚠️ Skipping comment for {element_id} (unsupported in header/footer)")
                    else:
                        try:
                            element.add_comment(
                                text=motivation,
                                author="JBG klarspråkningstjänst",
                                initials="JBG"
                            )
                        except Exception as e:
                            self.logger.warning(f"⚠️ Could not add comment to {element_id}: {e}")

                applied_changes.add(element_id)
                self.logger.info(f"✅ Applied: '{old}' → '{new}' in {element_id}")
                continue
            
            # Handle table cells
            elif isinstance(element, _Cell):
                if DEBUG:
                    self.logger.debug(f" -- Element is _Cell")
                cell_handled = False

                normalized_old = self._normalize_text(old)
                normalized_cell = self._normalize_text(
                    "\n".join(para.text.strip() for para in element.paragraphs)
                )
                sim_score = fuzz.ratio(normalized_old, normalized_cell)
                self.logger.debug(f"⚠️ Comparison similarity score: {sim_score} %")

                if float(sim_score) < TEXT_SIM_SCORE_THRESHOLD:
                    self.logger.error(
                        f"❌ Not enough match for '{old}' in {element_id} (fuzzy score: {sim_score})"
                    )
                    continue

                old_lines = old.strip().splitlines()
                new_lines = new.strip().splitlines()
                paras = element.paragraphs

                if len(old_lines) != len(paras):
                    self.logger.warning(
                        f"⚠️ Line count mismatch in {element_id}: {len(old_lines)} lines vs {len(paras)} paragraphs"
                    )
                else:
                    for i, (para, old_line, new_line) in enumerate(zip(paras, old_lines, new_lines)):
                        norm_old_line = self._normalize_text(old_line)
                        norm_para_line = self._normalize_text(para.text)
                        if norm_old_line in norm_para_line:
                            self.logger.info(
                                f"✅ Match: '{norm_old_line}' == '{norm_para_line}' in row {i} of Cell {element_id}"
                            )
                            diffed = self._diff_words(old_line, new_line)
                            for _ in range(len(para.runs)):
                                para._element.remove(para.runs[0]._element)
                            for kind, val in diffed:
                                run = para.add_run(val)
                                if kind == "strike":
                                    run.font.strike = True
                                    run.font.color.rgb = RGBColor(255, 0, 0)
                                elif kind == "insert":
                                    run.font.color.rgb = RGBColor(0, 128, 0)
                        else:
                            self.logger.warning(
                                f"⚠️ Could not match old line to paragraph {i+1} in {element_id}"
                            )

                    if self.include_motivations and motivation:
                        try:
                            paras[0].add_comment(
                                text=motivation,
                                author="JBG klarspråkningstjänst",
                                initials="JBG",
                            )
                        except Exception as e:
                            self.logger.warning(
                                f"⚠️ Could not add comment to first paragraph in table cell {element_id}: {e}"
                            )

                    applied_changes.add(element_id)
                    self.logger.info(f"✅ Applied in table cell {element_id}")
                    cell_handled = True

                if not cell_handled:
                    self.logger.error(f"❌ No match found in table cell {element_id}")

                continue

            # Handle raw XML elements that are footnotes
            elif isinstance(element, etree._Element) and element_id.startswith("footnote"):
                if DEBUG: self.logger.debug("-- Element is raw XML and footnote")

                # Collect all footnote changes for later
                self.footnote_changes.append({
                    "element_id": element_id,
                    "footnote_id": change.get("footnote_id"),
                    "old": old,
                    "new": new,
                    "motivation": motivation
                })
                self.logger.info(f"Stored footnote '{element_id}' with new text '{new}' for later insertions")
            
            # Handle raw XML elements that are not footnotes (not supported)
            else:
                self.logger.warning(f"⚠️ Unsupported element type for {element_id}: {type(element)}")

        # Final reporting (mninus footnotes)
        for change in self.changes:
            element_id = change.get("element_id")
            if element_id not in applied_changes and not element_id.startswith("footnote"):
                self.logger.error(f"❌ Change not applied: {change}")

        return doc

    
    def _diff_words(self, old, new):
        """
        Returns a list of tuples: (type, text)
        type ∈ {"text", "insert", "strike"}
        """
        diff = list(ndiff(old.split(), new.split()))
        result = []

        for d in diff:
            tag, word = d[0], d[2:]
            if tag == " ":
                result.append(("text", word + " "))
            elif tag == "-":
                result.append(("strike", word + " "))
            elif tag == "+":
                result.append(("insert", word + " "))

        return result
    
    def _get_joined_text_from_xml_element(self, element):
        texts = element.findall(".//w:t", namespaces=self.nsmap)
        return "".join(t.text for t in texts if t.text)

    def _edit_footnote_texts(self):
        """
        Reopens the saved .docx as a ZIP archive and directly patches footnotes.xml.
        """
        if not self.footnote_changes:
            self.logger.info("ℹ️ No footnote changes to apply.")
            return
        else:
            self.logger.info(f"Considering {len(self.footnote_changes)} footnote changes...")

        with ZipFile(self.filepath, 'r') as zin:
            with tempfile.TemporaryDirectory() as temp_dir:
                zin.extractall(temp_dir)
                footnote_path = os.path.join(temp_dir, "word", "footnotes.xml")

                if not os.path.exists(footnote_path):
                    self.logger.warning("⚠️ footnotes.xml not found in document.")
                    return

                parser = etree.XMLParser(ns_clean=True, recover=True)
                tree = etree.parse(footnote_path, parser)
                ns = self.nsmap

                for change in self.footnote_changes:
                    element_id = change["element_id"]
                    footnote_id = str(change["footnote_id"])
                    old = change["old"]
                    new = change["new"]
                    self.logger.debug(f"Footnote {footnote_id} should have '{old}' replaced with '{new}'")

                    xpath_expr = f".//w:footnote[@w:id='{footnote_id}']"
                    matches = tree.xpath(xpath_expr, namespaces=ns)

                    if not matches:
                        self.logger.warning(f"⚠️ Could not locate footnote ID {footnote_id} in XML.")
                        continue

                    footnote_node = matches[0]
                    text_nodes = footnote_node.xpath(".//w:t", namespaces=ns)
                    full_text = ''.join(t.text or '' for t in text_nodes)

                    
                    normalized_old = self._normalize_text(old)
                    normalized_origin = self._normalize_text(full_text)
                    if normalized_old not in normalized_origin:
                        self.logger.debug(f"⚠️ Could not find exact match for old text in footnote {element_id}:'{old}' and '{full_text}'")
                        sim_score = fuzz.ratio(normalized_old, normalized_origin)
                        self.logger.debug(f"⚠️ Comparison similarity score: {sim_score} %")
                        if float(sim_score) < TEXT_SIM_SCORE_THRESHOLD:
                            self.logger.error(f"❌ No enough match for '{old}' in {element_id} (fuzzy match similarity score: {sim_score}")
                            continue

                    # Apply naive replacement
                    try:
                        updated_text = full_text.replace(old, new)

                        # Redistribute updated text into original <w:t> nodes
                        remaining = updated_text
                        for i, t in enumerate(text_nodes):
                            original_len = len(t.text or "")
                            if i < len(text_nodes) - 1:
                                t.text = remaining[:original_len]
                                remaining = remaining[original_len:]
                            else:
                                t.text = remaining

                        self.logger.info(f"✅ Patched footnote {footnote_id}: '{old}' → '{new}'")

                    except Exception as e:
                        self.logger.error(f"❌ Failed to update footnote {footnote_id}: {e}")

                # Write modified footnotes.xml back
                with open(footnote_path, 'wb') as fout:
                    tree.write(fout, pretty_print=True, xml_declaration=True, encoding='utf-8')

                # Repackage the .docx
                with ZipFile(self.filepath, 'w') as zout:
                    for root, _, files in os.walk(temp_dir):
                        for filename in files:
                            file_path = os.path.join(root, filename)
                            arc_path = os.path.relpath(file_path, temp_dir)
                            zout.write(file_path, arc_path)

                self.logger.info("📦 Updated .docx file with modified footnotes.")
    
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

            # Get all text blocks on page, sorted top→down.
            # PyMuPDF coordinates increase downward, so smaller y means closer to the top.
            blocks = sorted(page.get_text("blocks"), key=lambda b: b[1])  # y1 ascending (top→down)
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
                                if self.include_motivations:
                                    highlight.set_info(content=f"{SUGGESTION}: {new} \n\n{MOTIVATION}: {motivation}")
                                else:
                                    highlight.set_info(content=f"{SUGGESTION}: {new}")
                        match_found = True
                        self.logger.info(f"✅ Applied: '{old}' → '{new}' on page {page_index + 1}, line {target_line}")
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
                                    if self.include_motivations:
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
                        if self.include_motivations:
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

            annots = page.annots()
            if not annots:
                continue
            for annot in annots:
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
