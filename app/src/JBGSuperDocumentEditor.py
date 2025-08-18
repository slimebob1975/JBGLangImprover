import os
import shutil
import uuid
import docx
from app.src.JBGDocumentEditor import JBGDocumentEditor
from app.src.JBGDocxRepairer import AutoDocxRepairer
from app.src.JBGDocxInternalValidator import DocxInternalValidator
from app.src.JBGDocxQualityChecker import JBGDocxQualityChecker
from lxml import etree
from tempfile import mkdtemp
import zipfile
from datetime import datetime, timezone
from uuid import uuid4
import random

DEBUG = True
INTERNAL_REPAIR_ON = True
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
        
        self.nsmap = {
            'w': "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            'w14': "http://schemas.microsoft.com/office/word/2010/wordml"
        }

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
        validator_pre = DocxInternalValidator(final_path)
        errors_pre = validator_pre.validate()

        if errors_pre:
            self.logger.warning("⚠️ Pre-repair validation issues detected:")
            for e in errors_pre:
                self.logger.warning(f"  {e}")
        else:
            self.logger.info("✅ No pre-repair validation issues detected.")

        # Attempt repair if needed
        if errors_pre and INTERNAL_REPAIR_ON:
            repairer = AutoDocxRepairer(logger=self.logger)
            repaired_path = repairer.repair(final_path)

            # Validate AFTER repair
            validator_post = DocxInternalValidator(repaired_path)
            errors_post = validator_post.validate()

            if errors_post:
                self.logger.warning("⚠️ Post-repair validation issues detected:")
                for e in errors_post:
                    self.logger.warning(f"  {e}")
            else:
                self.logger.info("✅ No post-repair validation issues detected.")
        else:
            repaired_path = final_path

        self.doc = docx.Document(repaired_path)
        self.final_path = repaired_path
        
        # Run QC Check after everything
        self.logger.info("🔍 Running final DOCX Quality Control check...")
        checker = JBGDocxQualityChecker(self.final_path, self.logger)
        checker.quality_control_docx()
        
        return self.doc
    
    def _convert_markup_to_tracked(self, input_docx_path: str) -> str:
        """
        Rewritten: 
        - Full tracked change conversion.
        - Guaranteed styles patch AFTER all XML work.
        """
        self.logger.info(f"_convert_markup_to_tracked has input file {input_docx_path}")

        nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        tracked_id = 1
        temp_dir = mkdtemp()
        output_path = os.path.join(temp_dir, os.path.basename(input_docx_path).replace("_interim", "_tracked"))

        with zipfile.ZipFile(input_docx_path, "r") as zin:
            zin.extractall(temp_dir)

        # Merge missing parts and ensure theme relationsships
        self._merge_missing_parts(self.filepath, temp_dir)
        self._ensure_theme_relationship(temp_dir)
        
        # ✅ Validate/patch styles.xml
        styles_path = os.path.join(temp_dir, "word", "styles.xml")
        settings_path = os.path.join(temp_dir, "word", "settings.xml")

        # Patch styles
        self._patch_or_inject_styles(styles_path)
        self._ensure_required_styles_in_document_xml(temp_dir)

        # Fix settings
        self._fix_or_inject_settings_xml(settings_path)

        # Inject realistic RSIDs, pass current temp output
        available_rsids = self._inject_multiple_rsid_entries(settings_path, docx_path=temp_dir)

        doc_path = os.path.join(temp_dir, "word", "document.xml")
        parser = etree.XMLParser(remove_blank_text=True)
        tree = etree.parse(doc_path, parser)
        root = tree.getroot()

        for run in root.xpath("//w:r", namespaces=nsmap):
            rpr = run.find("w:rPr", namespaces=nsmap)
            if rpr is None:
                continue

            color = rpr.find("w:color", namespaces=nsmap)
            strike = rpr.find("w:strike", namespaces=nsmap)

            color_val = color.get(f"{{{nsmap['w']}}}val") if color is not None else None

            is_red_strike = color_val == "FF0000" and strike is not None
            is_green_insert = color_val == "008000"

            if not is_red_strike and not is_green_insert:
                continue

            parent = run.getparent()
            index = parent.index(run)

            run_copy = etree.Element(f"{{{nsmap['w']}}}r")
            text_elem = run.find("w:t", namespaces=nsmap)
            if text_elem is None:
                continue

            comment_ref = run.find("w:commentReference", namespaces=nsmap)

            text = text_elem.text or ""
            if text and not text.endswith((" ", ".", ",", ";", ":", "!", "?", "”", "’")):
                text += " "
            if is_red_strike:
                deltext = etree.Element(f"{{{nsmap['w']}}}delText", nsmap={"xml": "http://www.w3.org/XML/1998/namespace"})
                deltext.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                deltext.text = text
                run_copy.append(deltext)
                wrapper = etree.Element(f"{{{nsmap['w']}}}del")
            elif is_green_insert:
                ins_text = etree.Element(f"{{{nsmap['w']}}}t", nsmap={"xml": "http://www.w3.org/XML/1998/namespace"})
                ins_text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                ins_text.text = text
                run_copy.append(ins_text)
                wrapper = etree.Element(f"{{{nsmap['w']}}}ins")

            wrapper.set(f"{{{nsmap['w']}}}id", str(tracked_id))
            wrapper.set(f"{{{nsmap['w']}}}author", "JBG Klarspråkningstjänst")
            wrapper.set(f"{{{nsmap['w']}}}date", datetime.now(timezone.utc).isoformat())
            wrapper.set(f"{{{nsmap['w']}}}rsidR", random.choice(available_rsids))
            wrapper.set(f"{{{nsmap['w']}}}{'rsidDel' if is_red_strike else 'rsidIns'}", random.choice(available_rsids))

            tracked_id += 1

            wrapper.append(run_copy)
            parent.remove(run)
            parent.insert(index, wrapper)

            if comment_ref is not None and self.include_motivations:
                comment_ref_run = etree.Element(f"{{{nsmap['w']}}}r")
                comment_ref_run.append(etree.fromstring(etree.tostring(comment_ref)))
                parent.insert(index + 1, comment_ref_run)

        # Try to insert rsid's in a way that mimic human behavior
        current_rsid = random.choice(available_rsids)
        paragraph_counter = 0
        for para in root.xpath("//w:p", namespaces=nsmap):
            para.set("{http://schemas.microsoft.com/office/word/2010/wordml}paraId", uuid4().hex[:8])
            para.set("{http://schemas.microsoft.com/office/word/2010/wordml}textId", uuid4().hex[:8])

            para.set(f"{{{nsmap['w']}}}rsidR", current_rsid)
            para.set(f"{{{nsmap['w']}}}rsidP", current_rsid)
            para.set(f"{{{nsmap['w']}}}rsidRPr", current_rsid)

            paragraph_counter += 1
            if paragraph_counter >= random.randint(2, 4):
                current_rsid = random.choice(available_rsids)
                paragraph_counter = 0

        # Save modified document.xml
        tree.write(doc_path, pretty_print=True, encoding="UTF-8", xml_declaration=True)
        self.logger.info("🔁 Replaced visual markups with tracked changes XML.")
        
        # Clean up the xml tree
        self._clean_xml_tree(doc_path)

        # Repackage everything
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for root_dir, _, files in os.walk(temp_dir):
                for file in files:
                    full_path = os.path.join(root_dir, file)
                    archive_name = os.path.relpath(full_path, temp_dir)
                    zout.write(full_path, archive_name)

        self.logger.info(f"✅ Tracked DOCX saved: {output_path}")
        return output_path
    
    def _clean_xml_tree(self, doc_path):
        """
        Clean up empty or invalid XML elements to ensure Word compatibility.
        """
        parser = etree.XMLParser(remove_blank_text=True)
        tree = etree.parse(doc_path, parser)
        root = tree.getroot()
        nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

        # Remove empty proofErr, noProof, etc
        for tag in ["proofErr", "noProof"]:
            for elem in root.xpath(f"//w:{tag}", namespaces=nsmap):
                parent = elem.getparent()
                if parent is not None:
                    parent.remove(elem)

        # Remove empty run properties
        for rpr in root.xpath("//w:rPr", namespaces=nsmap):
            if len(rpr) == 0:
                parent = rpr.getparent()
                if parent is not None:
                    parent.remove(rpr)

        tree.write(doc_path, pretty_print=True, encoding="UTF-8", xml_declaration=True)
        self.logger.info("🧹 Cleaned up empty XML tags in document.xml.")

    
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
    def _patch_or_inject_styles(self, styles_path):
        """
        Ensure critical styles exist inside styles.xml.
        If missing, inject Normal, DefaultParagraphFont, TableNormal, CommentText, InsertedText, DeletedText.
        """
        ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        nsmap = {"w": ns}

        if not os.path.exists(styles_path):
            self.logger.warning(f"⚠️ styles.xml missing — creating minimal styles.xml")
            root = etree.Element(f"{{{ns}}}styles", nsmap=nsmap)
            tree = etree.ElementTree(root)
        else:
            parser = etree.XMLParser(remove_blank_text=True)
            try:
                tree = etree.parse(styles_path, parser)
                root = tree.getroot()
            except Exception as ex:
                self.logger.warning(f"⚠️ styles.xml corrupted ({ex}) — recreating minimal styles.xml")
                root = etree.Element(f"{{{ns}}}styles", nsmap=nsmap)
                tree = etree.ElementTree(root)

        existing_styles = {s.get(f"{{{ns}}}styleId") for s in root.findall(".//w:style", namespaces=nsmap)}

        # Required base styles
        required_styles = [
            ("Normal", "paragraph", "Normal", True),
            ("DefaultParagraphFont", "character", "Default Paragraph Font", False),
            ("TableNormal", "table", "Table Normal", False),
            ("CommentText", "character", "Comment Text", False),
            ("InsertedText", "character", "Inserted Text", False),
            ("DeletedText", "character", "Deleted Text", False),
        ]

        added = 0
        for style_id, style_type, style_name, is_default in required_styles:
            if style_id not in existing_styles:
                style = etree.Element(f"{{{ns}}}style", {
                    f"{{{ns}}}type": style_type,
                    f"{{{ns}}}styleId": style_id
                })
                if is_default:
                    style.set(f"{{{ns}}}default", "1")

                etree.SubElement(style, f"{{{ns}}}name", {f"{{{ns}}}val": style_name})

                if style_id == "Normal":
                    etree.SubElement(style, f"{{{ns}}}qFormat")

                if style_id == "TableNormal":
                    tblpr = etree.SubElement(style, f"{{{ns}}}tblPr")
                    # Insert minimal table properties expected by Word
                    etree.SubElement(tblpr, f"{{{ns}}}tblStyleRowBandSize", {f"{{{ns}}}val": "1"})
                    etree.SubElement(tblpr, f"{{{ns}}}tblStyleColBandSize", {f"{{{ns}}}val": "1"})

                root.append(style)
                added += 1

        if added > 0 or not os.path.exists(styles_path):
            tree.write(styles_path, pretty_print=True, encoding="UTF-8", xml_declaration=True)
            self.logger.info(f"📝 Added {added} missing styles into styles.xml.")
        else:
            self.logger.info(f"✅ All critical styles already present in styles.xml.")

    
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
        
        self.logger.info(f"🧬 Injected minimal styles.xml with required tracked-change styles")

    def _validate_or_patch_styles(self, styles_path):
        """
        Ensure styles.xml includes all critical Word styles (Normal, TableNormal, etc.).
        Inject defaults if missing.
        """

        ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        nsmap = {"w": ns}

        required_styles = [
            ("Normal", "paragraph", True),
            ("DefaultParagraphFont", "character", False),
            ("TableNormal", "table", False),
            ("CommentText", "character", False),
            ("InsertedText", "character", False),
            ("DeletedText", "character", False),
        ]

        if not os.path.exists(styles_path):
            self.logger.warning(f"⚠️ styles.xml missing completely. Injecting minimal version.")
            root = etree.Element(f"{{{ns}}}styles", nsmap={"w": ns})
        else:
            try:
                parser = etree.XMLParser(remove_blank_text=True)
                tree = etree.parse(styles_path, parser)
                root = tree.getroot()
            except Exception as e:
                self.logger.warning(f"⚠️ Failed to parse styles.xml: {e}. Rebuilding minimal version.")
                root = etree.Element(f"{{{ns}}}styles", nsmap={"w": ns})

        existing_ids = {s.get(f"{{{ns}}}styleId") for s in root.findall(".//w:style", namespaces=nsmap)}

        added = 0
        for style_id, style_type, is_default in required_styles:
            if style_id not in existing_ids:
                style = etree.Element(f"{{{ns}}}style", {
                    f"{{{ns}}}type": style_type,
                    f"{{{ns}}}styleId": style_id
                })
                if is_default:
                    style.set(f"{{{ns}}}default", "1")

                etree.SubElement(style, f"{{{ns}}}name", {f"{{{ns}}}val": style_id})
                if style_id == "Normal":
                    etree.SubElement(style, f"{{{ns}}}qFormat")
                root.append(style)
                self.logger.info(f"➕ Injected missing style: {style_id}")
                added += 1

        # Save back if any changes
        if added > 0 or not os.path.exists(styles_path):
            tree = etree.ElementTree(root)
            tree.write(styles_path, pretty_print=True, encoding="UTF-8", xml_declaration=True)
            self.logger.info(f"✅ styles.xml updated with {added} missing styles.")
        else:
            self.logger.info("✅ All critical styles already present in styles.xml.")

    # Helper to inject trackRevisions and rsids into settings.xml
    def _fix_or_inject_settings_xml(self, settings_path):
        """
        Make sure <w:settings.xml> contains <w:trackRevisions> and a <w:rsids> block.
        BUT don't generate RSIDs here! _inject_multiple_rsid_entries() will handle RSIDs separately.
        """
        import os
        from lxml import etree

        ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        if not os.path.exists(settings_path):
            self.logger.warning(f"⚠️ settings.xml missing, creating minimal.")
            root = etree.Element(f"{{{ns}}}settings", nsmap={"w": ns})
        else:
            parser = etree.XMLParser(remove_blank_text=True)
            tree = etree.parse(settings_path, parser)
            root = tree.getroot()

        # Ensure <trackRevisions> exists
        if root.find(f"./{{{ns}}}trackRevisions") is None:
            etree.SubElement(root, f"{{{ns}}}trackRevisions")
            self.logger.info("📝 Added <w:trackRevisions> to settings.xml.")

        # Ensure <rsids> exists (empty for now)
        if root.find(f"./{{{ns}}}rsids") is None:
            etree.SubElement(root, f"{{{ns}}}rsids")
            self.logger.info("📝 Added empty <w:rsids> block to settings.xml.")

        # Ensure <compat> exists
        if root.find(f"./{{{ns}}}compat") is None:
            etree.SubElement(root, f"{{{ns}}}compat")
            self.logger.info("📝 Added <w:compat> to settings.xml.")
        
        # Ensure <updateFields> exists
        if root.find(f"./{{{ns}}}updateFields") is None:
            etree.SubElement(root, f"{{{ns}}}updateFields", {f"{{{ns}}}val": "true"})
            self.logger.info("📝 Added <w:updateFields w:val='true'/> to settings.xml.")

        tree = etree.ElementTree(root)
        tree.write(settings_path, pretty_print=True, encoding="UTF-8", xml_declaration=True)

        self.logger.info("🛠 Ensured <w:trackRevisions> and empty <w:rsids> exist in settings.xml.")


    def _ensure_required_styles_in_document_xml(self, docx_tmp_dir):
        """
        Make sure document.xml paragraphs and runs reference canonical styles like Normal and DefaultParagraphFont.
        """
        document_path = os.path.join(docx_tmp_dir, "word", "document.xml")
        
        if not os.path.exists(document_path):
            self.logger.warning(f"⚠️ document.xml missing — cannot fix styles.")
            return

        nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

        parser = etree.XMLParser(remove_blank_text=True)
        tree = etree.parse(document_path, parser)
        root = tree.getroot()

        fixed_paragraphs = 0

        for para in root.xpath("//w:body/w:p", namespaces=nsmap):
            ppr = para.find("w:pPr", namespaces=nsmap)
            if ppr is None:
                ppr = etree.Element("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr")
                para.insert(0, ppr)

            # Check if paragraph already has a pStyle
            pstyle = ppr.find("w:pStyle", namespaces=nsmap)
            if pstyle is None:
                pstyle = etree.Element("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle")
                pstyle.attrib["{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"] = "Normal"
                ppr.insert(0, pstyle)
                fixed_paragraphs += 1

        if fixed_paragraphs > 0:
            tree.write(document_path, pretty_print=True, xml_declaration=True, encoding="UTF-8")
            self.logger.info(f"✅ Injected Normal style into {fixed_paragraphs} paragraphs in document.xml")
        else:
            self.logger.info("✅ All paragraphs already have paragraph styles defined.")
    
    def _inject_fresh_styles_xml(self, styles_path):
        """
        Completely rebuilds a minimal but valid styles.xml required by Word.
        """
        ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        nsmap = {"w": ns}
        
        root = etree.Element(f"{{{ns}}}styles", nsmap={"w": ns})

        style_defs = [
            ("Normal", "paragraph", "Normal", True),
            ("DefaultParagraphFont", "character", "Default Paragraph Font", False),
            ("TableNormal", "table", "Table Normal", False),
            ("CommentText", "character", "Comment Text", False),
            ("InsertedText", "character", "Inserted Text", False),
            ("DeletedText", "character", "Deleted Text", False)
        ]

        for style_id, style_type, style_name, is_default in style_defs:
            style = etree.SubElement(root, f"{{{ns}}}style", {
                f"{{{ns}}}type": style_type,
                f"{{{ns}}}styleId": style_id
            })
            if is_default:
                style.set(f"{{{ns}}}default", "1")
            etree.SubElement(style, f"{{{ns}}}name", {f"{{{ns}}}val": style_name})
            if style_id == "Normal":
                etree.SubElement(style, f"{{{ns}}}qFormat")

        # Save it
        os.makedirs(os.path.dirname(styles_path), exist_ok=True)
        tree = etree.ElementTree(root)
        tree.write(styles_path, pretty_print=True, encoding="UTF-8", xml_declaration=True)

        self.logger.info(f"🧬 Injected a fresh, valid styles.xml at {styles_path}")

    def _inject_multiple_rsid_entries(self, settings_path, count=None, docx_path=None):
        """
        Creates multiple RSIDs randomly and injects into settings.xml properly.
        Returns the list of generated RSID values.
        """
        def random_hex_string(length=8):
            return os.urandom(length).hex()

        ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        if not os.path.exists(settings_path):
            self.logger.warning(f"⚠️ settings.xml missing, creating minimal settings.xml first.")
            root = etree.Element(f"{{{ns}}}settings", nsmap={"w": ns})
        else:
            parser = etree.XMLParser(remove_blank_text=True)
            tree = etree.parse(settings_path, parser)
            root = tree.getroot()

        # Remove old <w:rsids> if exists
        for elem in root.findall(f"{{{ns}}}rsids"):
            root.remove(elem)

        if count is None:
            count = self._estimate_rsid_count_from_document(docx_path=docx_path)

        rsids = etree.Element(f"{{{ns}}}rsids")
        rsid_list = [random_hex_string()[:8] for _ in range(count)]
        etree.SubElement(rsids, f"{{{ns}}}rsidRoot", {f"{{{ns}}}val": rsid_list[0]})
        for rsid_val in rsid_list:
            etree.SubElement(rsids, f"{{{ns}}}rsid", {f"{{{ns}}}val": rsid_val})

        root.insert(0, rsids)

        tree = etree.ElementTree(root)
        tree.write(settings_path, pretty_print=True, encoding="UTF-8", xml_declaration=True)

        self.logger.info(f"🔗 Injected {len(rsid_list)} RSID entries into {settings_path}")
        
        return rsid_list

    def _estimate_rsid_count_from_document(self, docx_path=None):
        """
        Analyze document.xml to estimate RSID count more realistically.
        """
        estimated_count = 50
        docx_path = docx_path or self.filepath

        try:
            if os.path.isdir(docx_path):
                doc_xml_path = os.path.join(docx_path, "word", "document.xml")
                if not os.path.exists(doc_xml_path):
                    self.logger.warning(f"⚠️ document.xml missing inside extracted folder {docx_path}")
                    return estimated_count

                parser = etree.XMLParser(remove_blank_text=True)
                tree = etree.parse(doc_xml_path, parser)
                paragraphs = tree.xpath('//w:p', namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
                para_count = len(paragraphs)
            else:
                with zipfile.ZipFile(docx_path, 'r') as zin:
                    if 'word/document.xml' not in zin.namelist():
                        self.logger.warning(f"⚠️ document.xml missing inside {docx_path}")
                        return estimated_count
                    with zin.open('word/document.xml') as doc_file:
                        parser = etree.XMLParser(remove_blank_text=True)
                        tree = etree.parse(doc_file, parser)
                        paragraphs = tree.xpath('//w:p', namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
                        para_count = len(paragraphs)

            if para_count < 20:
                estimated_count = 50
            elif para_count < 100:
                estimated_count = 100
            else:
                estimated_count = 200

            self.logger.info(f"📄 Detected {para_count} paragraphs → Suggest {estimated_count} RSIDs.")
        except Exception as ex:
            self.logger.warning(f"⚠️ Failed to estimate paragraph count: {ex}")

        return estimated_count

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
                if not os.path.exists(dst) and os.path.exists(src):
                    shutil.copyfile(src, dst)
                    self.logger.info(f"📥 Recovered missing file: {rel_path}")
                elif os.path.exists(dst):
                    self._patch_or_inject_styles(dst)
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


