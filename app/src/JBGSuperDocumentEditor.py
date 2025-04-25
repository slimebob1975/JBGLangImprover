import os
import shutil
import uuid
import docx
from app.src.JBGDocumentEditor import JBGDocumentEditor
from app.src.JBGDocxRepairer import AutoDocxRepairer
from app.src.JBGDocxInternalValidator import DocxInternalValidator
from lxml import etree
from tempfile import mkdtemp
import zipfile
from datetime import datetime, timezone
from uuid import uuid4

DEBUG = True
REPAIR_ON = True
REQUIRED_STYLES = {"Normal", "DefaultParagraphFont", "TableNormal", "CommentText", "InsertedText", "DeletedText"}

class JBGSuperDocumentEditor:
    def __init__(self, filepath, changes_path, include_motivations, docx_mode, logger):
        self.filepath = filepath
        self.changes_path = changes_path
        self.include_motivations = include_motivations
        self.docx_mode = docx_mode
        self.logger = logger
        if DEBUG:
            self.logger.info(f" The filepath in __init__ for JBGSuperDocumentEditor is: {self.filepath}")

        self.ext = os.path.splitext(filepath)[1].lower()
        if self.ext == ".pdf":
            self.editor = PDFDocumentEditor(filepath, changes_path, include_motivations, logger)
        elif self.ext == ".docx":
            if docx_mode == "tracked":
                self.editor = DocxTrackedChangesEditor(filepath, changes_path, include_motivations, logger)
            else:
                self.editor = DocxSimpleMarkupEditor(filepath, changes_path, include_motivations, logger)
        else:
            raise ValueError("Unsupported file format")
        
        self.nsmap = {
            'w': "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            'w14': "http://schemas.microsoft.com/office/word/2010/wordml"
        }

    def apply_changes(self):
        return self.editor.apply_changes()

    def save_edited_document(self, output_path=None):
        return self.editor.save(output_path)


class PDFDocumentEditor:
    def __init__(self, filepath, changes_path, include_motivations, logger):
        from app.src.JBGDocumentEditor import JBGDocumentEditor  # reuse PDF logic from legacy class
        self.editor = JBGDocumentEditor(filepath, changes_path, include_motivations, "simple", logger)

    def apply_changes(self):
        self.editor.apply_changes()  # this populates edited_document
        return self.editor.edited_document

    def save(self, output_path=None):
        return self.editor.save_edited_document(output_path)


class DocxDocumentEditor:
    def __init__(self, filepath, changes_path, include_motivations, logger):
        self.filepath = filepath
        self.changes_path = changes_path
        self.include_motivations = include_motivations
        self.logger = logger
        self.temp_path = self._copy_to_temp()

    def _copy_to_temp(self):
        basename = os.path.basename(self.filepath)
        temp_name = f"{uuid.uuid4()}_{basename}"
        temp_dir = "uploads" if os.name == "nt" else "/tmp"
        temp_path = os.path.join(temp_dir, temp_name)
        shutil.copyfile(self.filepath, temp_path)
        return temp_path

    def _get_changes(self):
        import json
        with open(self.changes_path, encoding="utf-8") as f:
            return json.load(f)


class DocxSimpleMarkupEditor(DocxDocumentEditor):
    def apply_changes(self):
        legacy = JBGDocumentEditor(self.temp_path, self.changes_path, self.include_motivations, "simple", self.logger)
        self.doc = legacy._edit_docx()
        return self.doc

    def save(self, output_path=None):
        if not output_path:
            output_path = self.temp_path.replace(".docx", "_edited.docx")
        self.doc.save(output_path)
        return output_path


class DocxTrackedChangesEditor(DocxSimpleMarkupEditor):
    def apply_changes(self):
        
        # Initially, we use simple markup from legacy editor implementation
        doc = super().apply_changes()  # visual markup
        interim_path = self.temp_path.replace(".docx", "_interim.docx")
        doc.save(interim_path)

        #legacy = JBGDocumentEditor(interim_path, self.changes_path, self.include_motivations, "tracked", self.logger)
        #final_path = legacy._convert_markup_to_tracked(interim_path)
        try:
            final_path = self._convert_markup_to_tracked(interim_path)
        except Exception as ex:
            self.logger.warning(f"🧨 SuperEditor XML manipulation failed: {ex}")
            raise EditorProcessingException("Tracked changes manipulation failed.")
        
        # Validate before repair
        validator = DocxInternalValidator(final_path)
        errors = validator.validate()
        if errors:
            self.logger.warning("🧪 Pre-repair validation issues detected:")
            for e in errors:
                self.logger.warning(f"  {e}")
        
        # Try to repair
        if REPAIR_ON:
            if DEBUG:
                with zipfile.ZipFile(final_path, 'r') as z:
                    for f in z.namelist():
                        if f.startswith('customXml'):
                            print("✅ Found in tracked docx:", f)
            repairer = AutoDocxRepairer(logger=self.logger)
            repaired_path = repairer.repair(final_path)
            if DEBUG:
                with zipfile.ZipFile(repaired_path, 'r') as z:
                    for f in z.namelist():
                        if f.startswith('customXml'):
                            print("✅ Still present after repair:", f)
        else:
            repaired_path = final_path
        self.doc = docx.Document(repaired_path)
        self.final_path = repaired_path
        return self.doc
    
    def _convert_markup_to_tracked(self, input_docx_path: str) -> str:
        """
        Full rewrite of tracked conversion:
        - Convert red+strike → <w:del><w:r><w:delText>...</w:delText></w:r></w:del>
        - Convert green → <w:ins><w:r><w:t>...</w:t></w:r></w:ins>
        """
        self.logger.info(f"_convert_markup_to_tracked has input file {input_docx_path}")
        def random_hex_string(length=8):
            return os.urandom(length).hex()
        
        nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        tracked_id = 1
        temp_dir = mkdtemp()
        output_path = os.path.join(temp_dir, os.path.basename(input_docx_path).replace("_interim", "_tracked"))
        
        # Unpack
        with zipfile.ZipFile(input_docx_path, "r") as zin:
            zin.extractall(temp_dir)
            
        # Repair the interim package by merging missing files, if needed
        self._merge_missing_parts(self.filepath, temp_dir)
        
        # Ensure there relationships
        self._ensure_theme_relationship(temp_dir)

        # Start manipulating the XML structure
        doc_path = os.path.join(temp_dir, "word", "document.xml")
        parser = etree.XMLParser(remove_blank_text=True)
        tree = etree.parse(doc_path, parser)
        root = tree.getroot()

        for run in root.xpath("//w:r", namespaces=nsmap):
            rpr = run.find("w:rPr", namespaces=nsmap)
            if rpr is None:
                continue

            # Detect red color (FF0000) and strike or underline
            color = rpr.find("w:color", namespaces=nsmap)
            strike = rpr.find("w:strike", namespaces=nsmap)

            color_val = color.get(f"{{{nsmap['w']}}}val") if color is not None else None

            is_red_strike = color_val == "FF0000" and strike is not None
            is_green_insert = color_val == "008000"

            if not is_red_strike and not is_green_insert:
                continue

            parent = run.getparent()
            index = parent.index(run)

            # Clean and copy run
            run_copy = etree.Element(f"{{{nsmap['w']}}}r")
            text_elem = run.find("w:t", namespaces=nsmap)
            if text_elem is None:
                continue
            
            # Check if this run includes a commentReference
            comment_ref = run.find("w:commentReference", namespaces=nsmap)

            text = text_elem.text or ""
            if is_red_strike:
                deltext = etree.Element(f"{{{nsmap['w']}}}delText")
                deltext.text = text
                run_copy.append(deltext)
                wrapper = etree.Element(f"{{{nsmap['w']}}}del")
            elif is_green_insert:
                ins_text = etree.Element(f"{{{nsmap['w']}}}t")
                ins_text.text = text
                run_copy.append(ins_text)
                wrapper = etree.Element(f"{{{nsmap['w']}}}ins")

            wrapper.set(f"{{{nsmap['w']}}}id", str(tracked_id))
            wrapper.set(f"{{{nsmap['w']}}}author", "JBG Klarspråkningstjänst")
            wrapper.set(f"{{{nsmap['w']}}}date", datetime.now(timezone.utc).isoformat())
            rsid = random_hex_string() # Like "00A10B2F"
            wrapper.set(f"{{{nsmap['w']}}}rsidR", rsid)
            wrapper.set(f"{{{nsmap['w']}}}{'rsidDel' if is_red_strike else 'rsidIns'}", rsid)
            
            tracked_id += 1

            wrapper.append(run_copy)
            parent.remove(run)
            parent.insert(index, wrapper)
            
            # Re-insert comment reference if it was originally attached
            if comment_ref is not None and self.include_motivations:
                comment_ref_run = etree.Element(f"{{{nsmap['w']}}}r")
                comment_ref_run.append(etree.fromstring(etree.tostring(comment_ref)))
                parent.insert(index + 1, comment_ref_run)

        for para in root.xpath("//w:p", namespaces=nsmap):
            para.set("{http://schemas.microsoft.com/office/word/2010/wordml}paraId", uuid4().hex[:8])
            para.set("{http://schemas.microsoft.com/office/word/2010/wordml}textId", uuid4().hex[:8])
            para.set(f"{{{nsmap['w']}}}rsidR", "00A10B2F")
            para.set(f"{{{nsmap['w']}}}rsidP", "00A10B2F")
        
        # Write back
        tree.write(doc_path, pretty_print=True, encoding="UTF-8", xml_declaration=True)
        self.logger.info("🔁 Replaced visual markups with tracked changes XML.")
        
        # Check no overlapping tags
        if root.xpath("//w:ins//w:del", namespaces=nsmap):
            raise EditorProcessingException("Invalid nesting: <w:del> inside <w:ins>")
        if root.xpath("//w:del//w:ins", namespaces=nsmap):
            raise EditorProcessingException("Invalid nesting: <w:ins> inside <w:del>")
        
        # Ensure all required files exist 
        required_paths = [
            os.path.join(temp_dir, "word", "document.xml"),
            os.path.join(temp_dir, "word", "comments.xml"),
            os.path.join(temp_dir, "word", "settings.xml"),
            os.path.join(temp_dir, "word", "_rels", "document.xml.rels")
        ]
        for path in required_paths:
            if not os.path.exists(path):
                self.logger.warning(f"⚠️ Required file missing: {path}")
                
        # Add final patches to critical XML structures
        doc_path = os.path.join(temp_dir, "word", "document.xml")
        settings_path = os.path.join(temp_dir, "word", "settings.xml")

        self._final_patch_document_xml(doc_path)
        self._final_patch_settings_xml(settings_path)

        # Repack
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for root_dir, _, files in os.walk(temp_dir):
                for file in files:
                    full_path = os.path.join(root_dir, file)
                    archive_name = os.path.relpath(full_path, temp_dir)
                    zout.write(full_path, archive_name)

        # After zipfile.ZipFile(...) repack
        if DEBUG:
            with zipfile.ZipFile(output_path, 'r') as zcheck:
                for f in zcheck.namelist():
                    print("✅ Included in final ZIP:", f)
        
        self.logger.info(f"✅ Tracked DOCX saved: {output_path}")
        
        return output_path
    
    # Ensure original file structure is preserved by copying missing original files
    # on demand
    
    def _ensure_theme_relationship(self, tmp_dir):
        rels_path = os.path.join(tmp_dir, "word", "_rels", "document.xml.rels")
        if not os.path.exists(rels_path):
            return

        nsmap = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
        tree = etree.parse(rels_path)
        root = tree.getroot()

        existing = [r.get("Type") for r in root.findall("rel:Relationship", namespaces=nsmap)]
        theme_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"

        if theme_type not in existing:
            etree.SubElement(root, "Relationship", {
                "Id": "rIdTheme",
                "Type": theme_type,
                "Target": "theme/theme1.xml"
            })
            tree.write(rels_path, pretty_print=True, encoding="UTF-8", xml_declaration=True)
            self.logger.info("🔗 Added missing theme relationship to document.xml.rels")
    
    # Helper to detect incomplete styles.xml
    def _is_incomplete_styles_file(self, styles_path):
        if not os.path.exists(styles_path) or os.stat(styles_path).st_size == 0:
            return True
        try:
            tree = etree.parse(styles_path)
            root = tree.getroot()
            found_styles = {s.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId")
                            for s in root.xpath("//w:style", namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})}
            missing = REQUIRED_STYLES - found_styles
            return bool(missing)
        except Exception:
            return True

    # Helper to patch incomplete styles.xml
    def _patch_or_inject_styles(self, dst_path):
        ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        nsmap = {"w": ns}
        tracked_styles = {
            "InsertedText": "character",
            "DeletedText": "character",
            "CommentText": "character"
        }

        if not os.path.exists(dst_path):
            # No styles at all? Then inject full minimal
            self._inject_minimal_styles_xml(dst_path)
            return

        try:
            parser = etree.XMLParser(remove_blank_text=True)
            tree = etree.parse(dst_path, parser)
            root = tree.getroot()
        except Exception as e:
            self.logger.warning(f"⚠️ Failed to parse styles.xml: {e}")
            self._inject_minimal_styles_xml(dst_path)
            return

        existing_ids = {s.get(f"{{{ns}}}styleId") for s in root.findall(".//w:style", namespaces=nsmap)}
        added = 0

        for style_id, style_type in tracked_styles.items():
            if style_id not in existing_ids:
                style = etree.SubElement(root, f"{{{ns}}}style", {
                    f"{{{ns}}}type": style_type,
                    f"{{{ns}}}styleId": style_id
                })
                etree.SubElement(style, f"{{{ns}}}name", {f"{{{ns}}}val": style_id})
                added += 1

        if added > 0:
            tree.write(dst_path, pretty_print=True, encoding="UTF-8", xml_declaration=True)
            self.logger.info(f"➕ Patched {added} tracked-change styles into styles.xml")
        else:
            self.logger.info("✅ styles.xml already contains all tracked-change styles")

    
    # Helper to inject minimal styles.xml
    def _inject_minimal_styles_xml(self, dst_path):
        ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        root = etree.Element(f"{{{ns}}}styles", nsmap={"w": ns})

        style_defs = [
            ("Normal", "paragraph"),
            ("DefaultParagraphFont", "character"),
            ("TableNormal", "table"),
            ("CommentText", "character"),
            ("InsertedText", "character"),
            ("DeletedText", "character")
        ]

        for style_id, style_type in style_defs:
            style = etree.SubElement(root, f"{{{ns}}}style", {
                f"{{{ns}}}type": style_type,
                f"{{{ns}}}styleId": style_id
            })
            etree.SubElement(style, f"{{{ns}}}name", {f"{{{ns}}}val": style_id})

        os.makedirs(os.path.dirname(dst_path), exist_ok=True)
        etree.ElementTree(root).write(dst_path, pretty_print=True, encoding="UTF-8", xml_declaration=True)

    # Helper to inject trackRevisions and rsids into settings.xml
    def _fix_or_inject_settings_xml(self, dst_path):
        ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        if not os.path.exists(dst_path):
            root = etree.Element(f"{{{ns}}}settings", nsmap={"w": ns})
        else:
            tree = etree.parse(dst_path)
            root = tree.getroot()

        if root.find(f".//{{{ns}}}trackRevisions") is None:
            track = etree.Element(f"{{{ns}}}trackRevisions")
            root.insert(0, track)

        if root.find(f".//{{{ns}}}rsids") is None:
            rsids = etree.Element(f"{{{ns}}}rsids")
            etree.SubElement(rsids, f"{{{ns}}}rsidRoot", {f"{{{ns}}}val": "00A10B2F"})
            etree.SubElement(rsids, f"{{{ns}}}rsid", {f"{{{ns}}}val": "00A10B2F"})
            root.insert(1, rsids)

        if root.find(f".//{{{ns}}}compat") is None:
            compat = etree.Element(f"{{{ns}}}compat")
            etree.SubElement(compat, f"{{{ns}}}compatSetting", {
                f"{{{ns}}}name": "compatibilityMode",
                f"{{{ns}}}val": "15"
            })
            root.append(compat)

        os.makedirs(os.path.dirname(dst_path), exist_ok=True)
        etree.ElementTree(root).write(dst_path, pretty_print=True, encoding="UTF-8", xml_declaration=True)

    def _merge_missing_parts(self, source_docx_path, temp_dir):
        backup_dir = mkdtemp()
        with zipfile.ZipFile(source_docx_path, 'r') as zin:
            zin.extractall(backup_dir)

        required_files = [
            "word/styles.xml",
            "word/settings.xml",
            "word/fontTable.xml",
            "word/webSettings.xml",
            "word/_rels/document.xml.rels",
            "word/theme/theme1.xml"
        ]

        for rel_path in required_files:
            src = os.path.join(backup_dir, rel_path)
            dst = os.path.join(temp_dir, rel_path)

            if "styles.xml" in rel_path:
                if (not os.path.exists(dst)) or (os.path.exists(src) and self._is_incomplete_styles_file(dst)):
                    self._patch_or_inject_styles(dst)
                    self.logger.info(f"🧬 Injected minimal styles.xml with required tracked-change styles")
                elif not os.path.exists(dst) and os.path.exists(src):
                    os.makedirs(os.path.dirname(dst), exist_ok=True)
                    shutil.copyfile(src, dst)
                    self.logger.info(f"📥 Recovered missing file: {rel_path}")

            elif "settings.xml" in rel_path:
                if (not os.path.exists(dst)) or (os.path.exists(src)):
                    if os.path.exists(src) and not os.path.exists(dst):
                        os.makedirs(os.path.dirname(dst), exist_ok=True)
                        shutil.copyfile(src, dst)
                        self.logger.info(f"📥 Recovered missing file: {rel_path}")
                    self._fix_or_inject_settings_xml(dst)
                    self.logger.info(f"🛠 Ensured <w:trackRevisions> and <w:rsids> in settings.xml")

            else:
                if not os.path.exists(dst) and os.path.exists(src):
                    os.makedirs(os.path.dirname(dst), exist_ok=True)
                    shutil.copyfile(src, dst)
                    self.logger.info(f"📥 Recovered missing file: {rel_path}")

        # SPECIAL CASE: Copy whole customXml/ if it exists
        src_customxml = os.path.join(backup_dir, "customXml")
        dst_customxml = os.path.join(temp_dir, "customXml")
        if os.path.exists(src_customxml):
            shutil.copytree(src_customxml, dst_customxml, dirs_exist_ok=True)
            self.logger.info("📂 Copied customXml/ folder from original.")

        shutil.rmtree(backup_dir)

    def _final_patch_document_xml(self, doc_path):
        parser = etree.XMLParser(remove_blank_text=True)
        tree = etree.parse(doc_path, parser)
        root = tree.getroot()

        for para in root.xpath('//w:body/w:p', namespaces=self.nsmap):
            para.set('{http://schemas.microsoft.com/office/word/2010/wordml}paraId', uuid4().hex[:8])
            para.set('{http://schemas.microsoft.com/office/word/2010/wordml}textId', uuid4().hex[:8])
            para.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidR', uuid4().hex[:8])
            para.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidP', uuid4().hex[:8])
            para.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidRPr', uuid4().hex[:8])

        tree.write(doc_path, xml_declaration=True, encoding='utf-8', pretty_print=True)

    def _final_patch_settings_xml(self, settings_path):
        parser = etree.XMLParser(remove_blank_text=True)
        tree = etree.parse(settings_path, parser)
        root = tree.getroot()
        ns = self.nsmap['w']

        # Add <w:compat> if missing
        if root.find(f"w:compat", namespaces=self.nsmap) is None:
            compat = etree.SubElement(root, f"{{{ns}}}compat")
            cs = etree.SubElement(compat, f"{{{ns}}}compatSetting")
            cs.set(f"{{{ns}}}name", "compatibilityMode")
            cs.set(f"{{{ns}}}val", "15")

        # Add <w:rsids> if missing
        if root.find(f"w:rsids", namespaces=self.nsmap) is None:
            rsids = etree.SubElement(root, f"{{{ns}}}rsids")
            rsid_val = uuid4().hex[:8]
            etree.SubElement(rsids, f"{{{ns}}}rsidRoot", val=rsid_val)
            etree.SubElement(rsids, f"{{{ns}}}rsid", val=rsid_val)

        tree.write(settings_path, xml_declaration=True, encoding='utf-8', pretty_print=True)

    def save(self, output_path=None):
        if not output_path:
            output_path = self.final_path
        else:
            self.doc.save(output_path)
        return output_path

class EditorProcessingException(Exception):
    """Raised when the Super Editor fails and should defer to the legacy editor."""
    pass


