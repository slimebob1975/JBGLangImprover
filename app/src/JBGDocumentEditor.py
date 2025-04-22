import docx
from docx.shared import RGBColor
from docx.enum.text import WD_UNDERLINE
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
from copy import deepcopy

DEBUG = True
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
        
        # Check if the produced document is valid
        self._validate_docx_integrity(os.path.dirname(tracked_doc_path))

        # Re-open and return final document for saving
        final_doc = docx.Document(tracked_doc_path)
        return final_doc
    
    def _enable_tracked_changes_settings(self, settings_xml_path):
       
        import xml.etree.ElementTree as ET
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        ET.register_namespace("w", ns["w"])

        tree = ET.parse(settings_xml_path)
        root = tree.getroot()

        # Only add if not already present
        if root.find("w:trackRevisions", namespaces=ns) is None:
            track = ET.Element(f"{{{ns['w']}}}trackRevisions")
            root.insert(0, track)
            tree.write(settings_xml_path, xml_declaration=True, encoding="UTF-8")
            self.logger.info("📌 Enabled <w:trackRevisions/> in settings.xml")

    
    def _convert_markup_to_tracked(self, input_docx_path):
        """
        Converts simple visual markups in a .docx file (like red strike/underline)
        into native Word tracked changes using <w:del> and <w:ins>.
        Returns the filepath of the modified document.
        """
        try:
            temp_dir = mkdtemp()
            tracked_docx_path = os.path.join(
                temp_dir, os.path.basename(input_docx_path).replace("_intermediate", "_tracked")
            )

            with zipfile.ZipFile(input_docx_path, 'r') as zin:
                zin.extractall(temp_dir)
                
            # Enable tracked changes in settings.xml
            settings_path = os.path.join(temp_dir, "word", "settings.xml")
            if os.path.exists(settings_path):
                self._enable_tracked_changes_settings(settings_path)

            # Start working with the document!
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

                if color_val == "FF0000":
                    if is_underline:
                        result = self._handle_insertions_and_comments(run)
                    elif is_strike:
                        result = self._handle_deletions(run)
                    if result == -1:
                        continue
                    
            # Write back modified XML
            with open(document_xml_path, "wb") as f:
                tree.write(f, pretty_print=True, xml_declaration=True, encoding="UTF-8")

            # 🧼 Optional: clean invalid tracked change leftovers
            document_xml_path = os.path.join(temp_dir, "word", "document.xml")
            self._sanitize_document_xml(document_xml_path)
            self._cleanup_docx_metadata(temp_dir)
            self._ensure_minimal_comments_xml(temp_dir)
            self._rebuild_document_rels_if_invalid(temp_dir)
            self._validate_docx_integrity(temp_dir)
            self._ensure_minimal_comments_structure(temp_dir)
            
            # Zip back into a new docx
            with zipfile.ZipFile(tracked_docx_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for root, _, files in os.walk(temp_dir):
                    for filename in files:
                        filepath = os.path.join(root, filename)
                        archive_name = os.path.relpath(filepath, temp_dir)
                        zout.write(filepath, archive_name)

            self.logger.info(f"🔁 Converted markup to tracked changes in: {tracked_docx_path}")
            return tracked_docx_path

        except Exception as ex:
            self.logger.warning(f"🚧 Tracked changes generation not fully functioning yet: {ex}")
            return input_docx_path
    
    def _handle_insertions_and_comments(self, run):
        parent = run.getparent()
        insertion_index = parent.index(run)

        # Clone and clean run
        run_copy = deepcopy(run)
        comment_ref = run_copy.find("w:commentReference", namespaces=self.nsmap)
        if comment_ref is not None:
            run_copy.remove(comment_ref)

        rpr_copy = run_copy.find("w:rPr", namespaces=self.nsmap)
        if rpr_copy is not None:
            rpr_copy.clear()
        else:
            etree.SubElement(run_copy, f"{{{self.nsmap['w']}}}rPr")

        # Wrap in a tracked-change element
        wrapper = etree.Element(f"{{{self.nsmap['w']}}}ins")
        wrapper.set(f"{{{self.nsmap['w']}}}author", "JBG Klarspråkningstjänst")
        wrapper.set(f"{{{self.nsmap['w']}}}date", datetime.now(timezone.utc).isoformat())

        # Always wrap <w:r> inside <w:del> or <w:ins>
        wrapper.append(run_copy)

        # Replace run
        parent.remove(run)
        parent.insert(insertion_index, wrapper)

        # Handle comment (only for insertions)
        if comment_ref is not None:
            comment_ref_run = etree.Element(f"{{{self.nsmap['w']}}}r")
            comment_ref_run.append(deepcopy(comment_ref))
            parent.insert(insertion_index + 1, comment_ref_run)
            
        return 0
            
    def _handle_deletions(self, run):

        parent = run.getparent()  # Expected to be <w:p>
        insertion_index = parent.index(run)

        # Clone the run
        run_copy = deepcopy(run)

        # Remove commentReference if any
        comment_ref = run_copy.find("w:commentReference", namespaces=self.nsmap)
        if comment_ref is not None:
            run_copy.remove(comment_ref)

        # Clean formatting
        rpr = run_copy.find("w:rPr", namespaces=self.nsmap)
        if rpr is not None:
            rpr.clear()
        else:
            etree.SubElement(run_copy, f"{{{self.nsmap['w']}}}rPr")

        # Extract the <w:t>
        text_elem = run_copy.find("w:t", namespaces=self.nsmap)
        if text_elem is None:
            return -1

        text_value = text_elem.text or ""

        # Create <w:delText> to replace <w:t>
        del_text_elem = etree.Element(f"{{{self.nsmap['w']}}}delText")
        del_text_elem.text = text_value
        run_copy.remove(text_elem)
        run_copy.append(del_text_elem)

        # Wrap in <w:del>
        del_tag = etree.Element(f"{{{self.nsmap['w']}}}del")
        del_tag.set(f"{{{self.nsmap['w']}}}author", "JBG Klarspråkningstjänst")
        del_tag.set(f"{{{self.nsmap['w']}}}date", datetime.now(timezone.utc).isoformat())
        del_tag.set(f"{{{self.nsmap['w']}}}rsidDel", "00000000")
        del_tag.append(run_copy)

        # Ensure parent <w:p> has rsidDel too
        if parent.tag.endswith("p") and f"{{{self.nsmap['w']}}}rsidDel" not in parent.attrib:
            parent.set(f"{{{self.nsmap['w']}}}rsidDel", "00000000")

        # Replace <w:r> with <w:del>
        parent.remove(run)
        parent.insert(insertion_index, del_tag)

        return 0
    
    def _validate_docx_integrity(self, docx_unzipped_dir):
        import xml.etree.ElementTree as ET
        import os

        ns = {
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
            "ct": "http://schemas.openxmlformats.org/package/2006/content-types"
        }

        # 0. Check that deletions are proper (they cause most problems)
        document_path = os.path.join(docx_unzipped_dir, "word", "document.xml")
        self._validate_deletion_structure(document_path)
        
        # 1. Extract all comment IDs from document.xml
        doc_tree = ET.parse(document_path)
        comment_refs = doc_tree.findall(".//w:commentReference", namespaces=ns)
        comment_ids_in_doc = {int(c.attrib[f"{{{ns['w']}}}id"]) for c in comment_refs}

        # Also check commentRangeStart and commentRangeEnd consistency
        range_starts = doc_tree.findall(".//w:commentRangeStart", namespaces=ns)
        range_ends = doc_tree.findall(".//w:commentRangeEnd", namespaces=ns)
        range_start_ids = {int(r.attrib[f"{{{ns['w']}}}id"]) for r in range_starts}
        range_end_ids = {int(r.attrib[f"{{{ns['w']}}}id"]) for r in range_ends}
        if range_start_ids != range_end_ids:
            raise Exception("❌ Comment range start/end IDs do not match.")

        # 2. Load comments.xml and check IDs
        comments_path = os.path.join(docx_unzipped_dir, "word", "comments.xml")
        if not os.path.exists(comments_path):
            raise Exception("❌ comments.xml is missing.")

        comments_tree = ET.parse(comments_path)
        comment_elems = comments_tree.findall(".//w:comment", namespaces=ns)
        comment_ids_defined = {int(c.attrib[f"{{{ns['w']}}}id"]) for c in comment_elems}

        # 3. Validate that every used ID is defined
        undefined_ids = comment_ids_in_doc - comment_ids_defined
        if undefined_ids:
            raise Exception(f"❌ Undefined comment IDs found in document.xml: {undefined_ids}")

        # 4. Check document.xml.rels for comment relationship
        rels_path = os.path.join(docx_unzipped_dir, "word", "_rels", "document.xml.rels")
        rels_tree = ET.parse(rels_path)
        rels = rels_tree.findall(".//rel:Relationship", namespaces=ns)
        comment_rels = [
            r for r in rels
            if r.attrib.get("Type") == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
        ]
        if not comment_rels:
            raise Exception("❌ Missing relationship to comments.xml in document.xml.rels")

        # 5. Check [Content_Types].xml
        content_types_path = os.path.join(docx_unzipped_dir, "[Content_Types].xml")
        ct_tree = ET.parse(content_types_path)
        overrides = ct_tree.findall(".//ct:Override", namespaces=ns)
        if not any(o.attrib.get("PartName") == "/word/comments.xml" for o in overrides):
            raise Exception("❌ Missing override for comments.xml in [Content_Types].xml")

        if DEBUG:
            unknown_tags = []
            for elem in doc_tree.iter():
                if not elem.tag.startswith(f"{{{ns['w']}}}"):
                    unknown_tags.append(elem.tag)
            if unknown_tags:
                self.logger.warning(f"🔍 Found {len(unknown_tags)} unknown tags in document.xml.")
                for tag in set(unknown_tags):
                    self.logger.warning(f"⚠️ Unknown element tag: {tag}")

        self.logger.info("✅ DOCX tracked-change integrity check passed.")

    def _validate_deletion_structure(self, document_xml_path):

        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        parser = etree.XMLParser(remove_blank_text=True)
        tree = etree.parse(document_xml_path, parser)
        root = tree.getroot()

        invalid_del_blocks = []

        for del_elem in root.xpath(".//w:del", namespaces=ns):
            parent = del_elem.getparent()
            if parent.tag != f"{{{ns['w']}}}p":
                invalid_del_blocks.append("Invalid placement (not inside <w:p>)")
                continue

            runs = del_elem.findall("w:r", namespaces=ns)
            if not runs:
                invalid_del_blocks.append("Missing <w:r> inside <w:del>")
                continue

            for run in runs:
                if run.find("w:delText", namespaces=ns) is None:
                    invalid_del_blocks.append("Missing <w:delText> inside <w:r> in <w:del>")

        if invalid_del_blocks:
            self.logger.warning(f"⚠️ Found {len(invalid_del_blocks)} invalid <w:del> structures.")
            raise Exception("❌ Invalid <w:del> structures detected in document.xml.")

        self.logger.info("✅ All <w:del> structures are valid (location + delText).")

    def _sanitize_document_xml(self, document_xml_path):

        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

        parser = etree.XMLParser(remove_blank_text=True)
        tree = etree.parse(document_xml_path, parser)
        root = tree.getroot()

        modified = False

        for del_elem in root.xpath(".//w:del", namespaces=ns):
            for run in del_elem.xpath(".//w:r", namespaces=ns):
                t = run.find("w:t", namespaces=ns)
                if t is not None:
                    # Convert to <w:delText>
                    text_val = t.text or ""
                    run.remove(t)
                    del_text = etree.Element(f"{{{ns['w']}}}delText")
                    del_text.text = text_val
                    run.append(del_text)
                    modified = True

        # Remove empty <w:r> elements
        for run in root.xpath("//w:r", namespaces=ns):
            if len(run) == 0:
                parent = run.getparent()
                parent.remove(run)
                modified = True
                
        # Log unknown attributes in <w:del>, <w:ins>
        for elem in root.xpath(".//w:del | .//w:ins", namespaces=ns):
            for attr in elem.attrib:
                if not attr.startswith(f"{{{ns['w']}}}"):
                    self.logger.warning(f"⚠️ Unexpected attribute in {elem.tag}: {attr}")

        if modified:
            with open(document_xml_path, "wb") as f:
                tree.write(f, pretty_print=True, xml_declaration=True, encoding="UTF-8")
            self.logger.info("🧼 Sanitized document.xml for invalid <w:t> in <w:del> or empty runs.")

    def _cleanup_docx_metadata(self, docx_unzipped_dir):
        import xml.etree.ElementTree as ET

        ns_ct = {"ct": "http://schemas.openxmlformats.org/package/2006/content-types"}
        ns_rel = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}

        # 1. Clean [Content_Types].xml
        ct_path = os.path.join(docx_unzipped_dir, "[Content_Types].xml")
        if os.path.exists(ct_path):
            tree = ET.parse(ct_path)
            root = tree.getroot()
            removed = 0
            for part in root.findall("ct:Override", namespaces=ns_ct):
                name = part.attrib.get("PartName", "")
                if "commentsExtended.xml" in name or "commentsIds.xml" in name:
                    root.remove(part)
                    removed += 1
            if removed:
                tree.write(ct_path, xml_declaration=True, encoding="utf-8")
                self.logger.info(f"🧹 Removed {removed} ghost overrides from [Content_Types].xml")

        # 2. Clean word/_rels/document.xml.rels
        rels_path = os.path.join(docx_unzipped_dir, "word", "_rels", "document.xml.rels")
        if os.path.exists(rels_path):
            tree = ET.parse(rels_path)
            root = tree.getroot()
            removed = 0
            for rel in root.findall("rel:Relationship", namespaces=ns_rel):
                target = rel.attrib.get("Target", "")
                if "commentsExtended.xml" in target or "commentsIds.xml" in target:
                    root.remove(rel)
                    removed += 1
            if removed:
                tree.write(rels_path, xml_declaration=True, encoding="utf-8")
                self.logger.info(f"🧹 Removed {removed} ghost relationships from document.xml.rels")
    
    def _rebuild_document_rels_if_invalid(self, docx_unzipped_dir):
        import os
        from lxml import etree

        rels_path = os.path.join(docx_unzipped_dir, "word", "_rels", "document.xml.rels")
        os.makedirs(os.path.dirname(rels_path), exist_ok=True)

        try:
            parser = etree.XMLParser(remove_blank_text=True)
            etree.parse(rels_path, parser)
            self.logger.info("✅ document.xml.rels is valid.")
        except Exception as ex:
            self.logger.warning(f"⚠️ document.xml.rels missing or invalid, regenerating: {ex}")

            nsmap = {None: "http://schemas.openxmlformats.org/package/2006/relationships"}
            root = etree.Element("Relationships", nsmap=nsmap)

            part_counter = 1

            def add_rel(target, rtype):
                nonlocal part_counter
                if os.path.exists(os.path.join(docx_unzipped_dir, "word", target)):
                    etree.SubElement(root, "Relationship", {
                        "Id": f"rId{part_counter}",
                        "Type": f"http://schemas.openxmlformats.org/officeDocument/2006/relationships/{rtype}",
                        "Target": target
                    })
                    self.logger.info(f"🔗 Added relationship for: {target}")
                    part_counter += 1

            # Common Word document parts
            if os.path.exists(os.path.join(docx_unzipped_dir, "word", "comments.xml")):
                add_rel("comments.xml", "comments")            
            add_rel("styles.xml", "styles")
            add_rel("settings.xml", "settings")
            add_rel("fontTable.xml", "fontTable")
            add_rel("webSettings.xml", "webSettings")
            add_rel("theme/theme1.xml", "theme")

            tree = etree.ElementTree(root)
            tree.write(rels_path, pretty_print=True, xml_declaration=True, encoding="UTF-8")

            self.logger.info("🛠️ Rebuilt document.xml.rels with detected relationships.")

    def _ensure_minimal_comments_xml(self, docx_unzipped_dir):
        from lxml import etree
        import os

        comments_path = os.path.join(docx_unzipped_dir, "word", "comments.xml")
        if not os.path.exists(comments_path):
            nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            root = etree.Element(f"{{{nsmap['w']}}}comments", nsmap=nsmap)
            tree = etree.ElementTree(root)

            os.makedirs(os.path.dirname(comments_path), exist_ok=True)
            tree.write(comments_path, pretty_print=True, xml_declaration=True, encoding="UTF-8")
            self.logger.info("🧾 Created minimal empty comments.xml as placeholder.")

    def _ensure_minimal_comments_structure(self, docx_unzipped_dir):
        """
        Ensures that the comments.xml exists and is minimal (even if include_comments=False),
        and purges ghost references to removed or unused comment parts.
        """
        import os
        from lxml import etree

        ns_w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        ns_ct = "http://schemas.openxmlformats.org/package/2006/content-types"
        ns_rel = "http://schemas.openxmlformats.org/package/2006/relationships"

        # Ensure comments.xml exists (even if empty)
        comments_path = os.path.join(docx_unzipped_dir, "word", "comments.xml")
        if not os.path.exists(comments_path):
            root = etree.Element(f"{{{ns_w}}}comments", nsmap={"w": ns_w})
            tree = etree.ElementTree(root)
            os.makedirs(os.path.dirname(comments_path), exist_ok=True)
            tree.write(comments_path, pretty_print=True, xml_declaration=True, encoding="UTF-8")
            self.logger.info("🧾 Created fallback comments.xml")

        # Remove unused references from [Content_Types].xml
        ct_path = os.path.join(docx_unzipped_dir, "[Content_Types].xml")
        if os.path.exists(ct_path):
            tree = etree.parse(ct_path)
            root = tree.getroot()
            for override in root.findall(f".//{{{ns_ct}}}Override"):
                name = override.get("PartName", "")
                if "commentsExtended.xml" in name or "commentsIds.xml" in name:
                    root.remove(override)
                    self.logger.info(f"🧹 Removed ghost content type: {name}")
            tree.write(ct_path, pretty_print=True, xml_declaration=True, encoding="utf-8")

        # Remove ghost relationships in document.xml.rels
        rels_path = os.path.join(docx_unzipped_dir, "word", "_rels", "document.xml.rels")
        if os.path.exists(rels_path):
            tree = etree.parse(rels_path)
            root = tree.getroot()
            for rel in root.findall(f".//{{{ns_rel}}}Relationship"):
                target = rel.get("Target", "")
                if "commentsExtended.xml" in target or "commentsIds.xml" in target:
                    root.remove(rel)
                    self.logger.info(f"🧹 Removed ghost relationship: {target}")
            tree.write(rels_path, pretty_print=True, xml_declaration=True, encoding="utf-8")

    
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
