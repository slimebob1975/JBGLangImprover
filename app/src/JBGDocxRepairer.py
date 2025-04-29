import os
import sys
import platform
import argparse
import zipfile
from tempfile import mkdtemp
from lxml import etree
import shutil
from app.src.JBGDocxInternalValidator import DocxInternalValidator

NSMAP = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
REQUIRED_STYLES = [
            ("Normal", "paragraph", "Normal", True),
            ("DefaultParagraphFont", "character", "Default Paragraph Font", False),
            ("TableNormal", "table", "Table Normal", False),
            ("CommentText", "character", "Comment Text", False),
            ("InsertedText", "character", "Inserted Text", False),
            ("DeletedText", "character", "Deleted Text", False)
            ]
    
class AutoDocxRepairer:
    def __init__(self, logger=None):
        self.logger = logger or print
        self.repairer = self._choose_repairer()

    def _log(self, msg):
        if callable(self.logger):
            self.logger(msg)
        else:
            print(msg)

    def _choose_repairer(self):
        if platform.system() == "Windows":
            try:
                import win32com.client
                # Test if Word COM is available
                win32com.client.Dispatch("Word.Application")
                self._log("🪟 Using WordRepairer (Windows COM automation).")
                return WordRepairer(logger=self.logger, enabled=True)
            except Exception as e:
                self._log(f"⚠️ Failed to initialize Word COM: {e} — using XML fallback.")
        else:
            self._log("🐧 Non-Windows platform — using XML fallback.")
        
        return DocxXmlRepairer(logger=self.logger)

    def repair(self, input_path, output_path=None):
        
        output_path = output_path or input_path

        repaired_file = self.repairer.repair(input_path, output_path)
        
        validator = DocxInternalValidator(repaired_file)
        errors = validator.validate()
        if errors:
            self._log("⚠️ Validator flagged issues post-repair:")
            for e in errors:
                self._log(f"  {e}")
        
        return repaired_file 

class WordRepairer:
    def __init__(self, logger=None, enabled=True):
        self.logger = logger or print
        self.enabled = enabled and platform.system() == "Windows"
        self.word = None

    def _log(self, msg):
        if callable(self.logger):
            self.logger(msg)
        else:
            print(msg)

    def repair(self, input_path, output_path=None):
        if not self.enabled:
            self._log("🔒 WordRepairer skipped (not enabled or not on Windows).")
            return None
        
        try:
            # Import where needed
            import win32com.client
        except ImportError:
            self._log("❌ win32com.client not available — fallback required.")
            return DocxXmlRepairer(logger=self.logger).repair(input_path, output_path)
        else:
            self._log("✅ Win32com.client available. Proceeding...")
        
        # Check file exists
        if not os.path.exists(input_path):
            self._log(f"❌ File not found: {input_path}")
            return None
        
        # Normalize paths
        input_path = os.path.abspath(input_path)
        if output_path:
            output_path = os.path.abspath(output_path)

        if not os.path.isfile(input_path) or not input_path.lower().endswith(".docx"):
            self._log(f"❌ Invalid file: {input_path}")
            return None

        try:
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = False
            self.word.DisplayAlerts = False

            self._log(f"🔧 Attempting to repair: {input_path}")
            
            doc = self.word.Documents.Open(FileName=input_path, OpenAndRepair=True)

            if not output_path:
                output_path = input_path

            doc.SaveAs(output_path)
            doc.Close()
            self._log(f"✅ Repaired file saved to: {output_path}")
            return output_path

        except Exception as ex:
            self._log(f"❌ Windows COM automation failed to repair file {input_path}: {ex}. Using XML fallback")
            try:
                xml_repairer = DocxXmlRepairer(logger=self.logger)
                return xml_repairer.repair(input_path, output_path)
            except Exception as ex:
                raise Exception(f"❌ XML fallback failed to repair file {input_path}: {ex}.")

        finally:
            if self.word:
                self.word.Quit()
                self.word = None

    def repair_batch(self, input_dir, output_dir=None):
        if not self.enabled:
            self._log("🔒 WordRepairer batch skipped (not on Windows).")
            return []

        repaired_files = []
        for filename in os.listdir(input_dir):
            if filename.lower().endswith(".docx"):
                input_path = os.path.join(input_dir, filename)
                output_path = os.path.join(output_dir, filename) if output_dir else None
                result = self.repair(input_path, output_path)
                if result:
                    repaired_files.append(result)

        return repaired_files
    
class DocxXmlRepairer:
    
    def __init__(self, logger=None):
        self.logger = logger or print
        self.required_styles = REQUIRED_STYLES

    def repair(self, docx_path, output_path=None):
        
        try:
            tmp_dir = mkdtemp()
            with zipfile.ZipFile(docx_path, "r") as zin:
                zin.extractall(tmp_dir)
            zin.close()

            comments_path = os.path.join(tmp_dir, "word", "comments.xml")
            if not os.path.exists(comments_path):
                self.logger.warning("⚠️ comments.xml missing — creating minimal version.")
                self._add_minimal_comments_to_dir(tmp_dir)

            else:
                try:
                    etree.parse(comments_path)
                    self.logger.info("✅ comments.xml is valid.")
                except etree.XMLSyntaxError:
                    self.logger.warning("⚠️ comments.xml is malformed — repairing.")
                    self._repair_comments_in_dir(tmp_dir)

            # Make all the hard work...
            self._add_comment_infra_files(tmp_dir)
            self._ensure_extended_comment_relationships(tmp_dir)
            self._ensure_content_type_overrides(tmp_dir)
            self._remove_orphan_comment_refs(tmp_dir)
            self._ensure_all_comment_ids_exist(tmp_dir)
            self._ensure_comment_relationship(tmp_dir)
            self._enable_track_revisions(tmp_dir)
            self._ensure_rsid_definitions(tmp_dir)
            self._ensure_styles_definitions(tmp_dir) 
            
            # Ensure all required files exist 
            required_paths = [
                os.path.join(tmp_dir, "word", "document.xml"),
                os.path.join(tmp_dir, "word", "comments.xml"),
                os.path.join(tmp_dir, "word", "settings.xml"),
                os.path.join(tmp_dir, "word", "_rels", "document.xml.rels")
            ]
            for path in required_paths:
                if not os.path.exists(path):
                    self.logger.warning(f"⚠️ Required file missing: {path}")

            # Final output path
            if not output_path or os.path.abspath(output_path) == os.path.abspath(docx_path):
                output_path = docx_path.replace(".docx", "_repaired.docx")

            try:
               with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
                    for root_dir, _, files in os.walk(tmp_dir):
                        for file in files:
                            full_path = os.path.join(root_dir, file)
                            archive_name = os.path.relpath(full_path, tmp_dir)
                            zout.write(full_path, archive_name)
            except PermissionError as pe:
                self.logger.error(f"❌ Permission denied during zipping: {pe}")
                raise
            else:
                try:
                    shutil.rmtree(tmp_dir)
                    self.logger.info(f"🧹 Cleaned up temporary directory: {tmp_dir}")
                except Exception as cleanup_err:
                    self.logger.warning(f"⚠️ Failed to clean temp dir {tmp_dir}: {cleanup_err}")

            self.logger.info(f"🔧 Repaired and saved as: {output_path}")
            
            return output_path

        except zipfile.BadZipFile:
            raise Exception("❌ .docx file is not a valid ZIP archive — unrecoverable.")
    
    # Empty placeholder for now -- need to add functionality here
    def _repair_comments(self, docx_path, output_path=None):
        if not output_path:
            return docx_path
        else:
            return output_path
        
    def _remove_orphan_comment_refs(self, tmp_dir):
        
        doc_path = os.path.join(tmp_dir, "word", "document.xml")
        comments_path = os.path.join(tmp_dir, "word", "comments.xml")
        parser = etree.XMLParser(remove_blank_text=True)

        try:
            doc_tree = etree.parse(doc_path, parser)
            comment_elems = etree.parse(comments_path).xpath("//w:comment", namespaces=NSMAP)
            valid_ids = {c.get(f"{{{NSMAP['w']}}}id") for c in comment_elems}

            for ref in doc_tree.xpath("//w:commentReference", namespaces=NSMAP):
                if ref.get(f"{{{NSMAP['w']}}}id") not in valid_ids:
                    parent = ref.getparent()
                    parent.remove(ref)
                    self.logger.info(f"🧽 Removed orphan commentReference: ID={ref.get('w:id')}")

            doc_tree.write(doc_path, pretty_print=True, encoding="UTF-8", xml_declaration=True)

        except Exception as e:
            self.logger.warning(f"⚠️ Failed to sanitize comment references: {e}")
        else:
            self.logger.info(f" ✅ No orphan commentReference:s to remove.")

    def _add_minimal_comments_to_dir(self, tmp_dir):
        """
        Creates a minimal comments.xml in the given already-extracted tmp_dir.
        This version avoids re-extracting the docx file and re-zipping here.
        """

        comments_path = os.path.join(tmp_dir, "word", "comments.xml")
        os.makedirs(os.path.dirname(comments_path), exist_ok=True)

        nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        root = etree.Element(f"{{{nsmap['w']}}}comments", nsmap=nsmap)
        etree.ElementTree(root).write(comments_path, pretty_print=True, encoding="UTF-8", xml_declaration=True)

        self.logger.info(f"🧾 Created minimal comments.xml at: {comments_path}")
    
    def _ensure_comment_relationship(self, tmp_dir):
     
        rels_path = os.path.join(tmp_dir, "word", "_rels", "document.xml.rels")
        if not os.path.exists(rels_path):
            return  # Let Word fix this later

        tree = etree.parse(rels_path)
        root = tree.getroot()
        nsmap = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
        
        # Ensure only one relationship per type
        existing = [r.get("Target") for r in root.findall("rel:Relationship", namespaces=nsmap)]
        if "comments.xml" not in existing:
            etree.SubElement(root, "Relationship", {
                "Id": "rIdComments",
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                "Target": "comments.xml"
            })
            tree.write(rels_path, pretty_print=True, encoding="UTF-8", xml_declaration=True)
            self.logger.info("🔗 Added relationship to comments.xml")


        if not any("comments.xml" in r.get("Target", "") for r in root.findall("rel:Relationship", namespaces=nsmap)):
            etree.SubElement(root, "Relationship", {
                "Id": "rIdComments",
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
                "Target": "comments.xml"
            })
            tree.write(rels_path, pretty_print=True, encoding="UTF-8", xml_declaration=True)
            self.logger.info("🔗 Added missing relationship to comments.xml")
        else:
            self.logger.info(f" ✅ Seems like all comments has their proper relationsships.")        
            
    def _add_comment_infra_files(self, tmp_dir):

        word_dir = os.path.join(tmp_dir, "word")
        os.makedirs(word_dir, exist_ok=True)

        def safe_write(filename, content_func):
            path = os.path.join(word_dir, filename)
            if os.path.exists(path):
                self.logger.info(f"📄 {filename} already exists — skipped.")
                return
            tree = content_func()
            tree.write(path, xml_declaration=True, encoding="UTF-8", pretty_print=True)
            self.logger.info(f"🧾 Created {filename}")

        # 1. commentsExtended.xml
        def comments_ex_tree():
            root = etree.Element("{http://schemas.microsoft.com/office/word/2010/wordml}commentsEx")
            return etree.ElementTree(root)

        # 2. commentsIds.xml
        def comments_ids_tree():
            ns = "http://schemas.microsoft.com/office/word/2016/wordml"
            root = etree.Element(f"{{{ns}}}commentsIds")
            etree.SubElement(root, f"{{{ns}}}commentId", {
                f"{{{ns}}}paraId": "00000000",
                f"{{{ns}}}durableId": "{00000000-0000-0000-0000-000000000000}"
            })
            return etree.ElementTree(root)

        # 3. people.xml
        def people_tree():
            root = etree.Element("{http://schemas.microsoft.com/office/2006/metadata/customXml}personList")
            return etree.ElementTree(root)

        # Apply all safely
        safe_write("commentsExtended.xml", comments_ex_tree)
        safe_write("commentsIds.xml", comments_ids_tree)
        safe_write("people.xml", people_tree)
        
    def _ensure_extended_comment_relationships(self, tmp_dir):

        rels_path = os.path.join(tmp_dir, "word", "_rels", "document.xml.rels")
        if not os.path.exists(rels_path):
            self.logger.warning("⚠️ document.xml.rels missing — skipping relationship injection.")
            return

        tree = etree.parse(rels_path)
        root = tree.getroot()
        nsmap = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}

        existing_targets = {
            rel.attrib.get("Target")
            for rel in root.findall("rel:Relationship", namespaces=nsmap)
        }

        required = [
            ("commentsExtended.xml", "http://schemas.microsoft.com/office/2011/relationships/commentsExtended", "rIdExt"),
            ("commentsIds.xml", "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds", "rIdIds"),
            ("people.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/person", "rIdPeople"),
        ]

        for target, rtype, rid in required:
            if target in existing_targets:
                self.logger.info(f"🔗 Relationship for {target} already exists — skipped.")
            else:
                etree.SubElement(root, "Relationship", {
                    "Id": rid,
                    "Type": rtype,
                    "Target": target
                })
                self.logger.info(f"🔗 Added relationship for: {target}")

        tree.write(rels_path, pretty_print=True, encoding="UTF-8", xml_declaration=True)

    def _ensure_content_type_overrides(self, tmp_dir):

        ct_path = os.path.join(tmp_dir, "[Content_Types].xml")
        if not os.path.exists(ct_path):
            self.logger.warning("⚠️ [Content_Types].xml missing — skipping override injection.")
            return

        tree = etree.parse(ct_path)
        root = tree.getroot()
        ns_ct = "http://schemas.openxmlformats.org/package/2006/content-types"

        existing_parts = {
            override.attrib.get("PartName")
            for override in root.findall(f".//{{{ns_ct}}}Override")
        }

        overrides = [
            ("/word/commentsExtended.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"),
            ("/word/commentsIds.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"),
            ("/word/people.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml"),
        ]

        for part_name, content_type in overrides:
            if part_name in existing_parts:
                self.logger.info(f"🧾 Content override for {part_name} already exists — skipped.")
            else:
                etree.SubElement(root, f"{{{ns_ct}}}Override", {
                    "PartName": part_name,
                    "ContentType": content_type
                })
                self.logger.info(f"🧾 Added content type override: {part_name}")

        tree.write(ct_path, pretty_print=True, encoding="UTF-8", xml_declaration=True)
        
    def _enable_track_revisions(self, tmp_dir):
        settings_path = os.path.join(tmp_dir, "word", "settings.xml")
        if not os.path.exists(settings_path):
            self.logger.warning("⚠️ settings.xml not found — skipping trackRevisions.")
            return

        tree = etree.parse(settings_path)
        root = tree.getroot()
        nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

        if root.find("w:trackRevisions", namespaces=nsmap) is None:
            track = etree.Element("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}trackRevisions")
            root.insert(0, track)
            tree.write(settings_path, pretty_print=True, encoding="UTF-8", xml_declaration=True)
            self.logger.info("📝 Added <w:trackRevisions> to settings.xml.")

    def _ensure_all_comment_ids_exist(self, tmp_dir):
        doc_path = os.path.join(tmp_dir, "word", "document.xml")
        comments_path = os.path.join(tmp_dir, "word", "comments.xml")
        if not os.path.exists(comments_path):
            return

        parser = etree.XMLParser(remove_blank_text=True)
        doc_tree = etree.parse(doc_path, parser)
        comments_tree = etree.parse(comments_path, parser)
        comments_root = comments_tree.getroot()

        existing_ids = {
            c.get(f"{{{NSMAP['w']}}}id") for c in comments_root.findall("w:comment", namespaces=NSMAP)
        }

        for ref in doc_tree.xpath("//w:commentReference", namespaces=NSMAP):
            cid = ref.get(f"{{{NSMAP['w']}}}id")
            if cid not in existing_ids:
                dummy = etree.Element(f"{{{NSMAP['w']}}}comment", nsmap=NSMAP)
                dummy.set(f"{{{NSMAP['w']}}}id", cid)
                dummy.set(f"{{{NSMAP['w']}}}author", "JBG")
                dummy.set(f"{{{NSMAP['w']}}}date", datetime.now(timezone.utc).isoformat())
                dummy.set(f"{{{NSMAP['w']}}}initials", "JBG")
                p = etree.SubElement(dummy, f"{{{NSMAP['w']}}}p")
                r = etree.SubElement(p, f"{{{NSMAP['w']}}}r")
                t = etree.SubElement(r, f"{{{NSMAP['w']}}}t")
                t.text = "Kommentar saknades, skapad automatiskt."
                comments_root.append(dummy)
                self.logger.info(f"🧩 Added dummy comment for id={cid}")

        comments_tree.write(comments_path, pretty_print=True, encoding="UTF-8", xml_declaration=True)

    def _ensure_rsid_definitions(self, tmp_dir):
        settings_path = os.path.join(tmp_dir, "word", "settings.xml")
        if not os.path.exists(settings_path):
            self.logger.warning("⚠️ settings.xml not found — skipping rsid definitions.")
            return

        parser = etree.XMLParser(remove_blank_text=True)
        tree = etree.parse(settings_path, parser)
        root = tree.getroot()
        ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        existing_rsids = root.find(f"{{{ns}}}rsids")
        
        if existing_rsids is not None:
            self.logger.info("✅ RSID definitions already exist — no need to overwrite.")
            return  # Do not touch if already there!

        # If no rsids exist, create a realistic one
        rsids = etree.Element(f"{{{ns}}}rsids")
        rsid_list = [os.urandom(4).hex() for _ in range(random.randint(8, 15))]

        etree.SubElement(rsids, f"{{{ns}}}rsidRoot", {f"{{{ns}}}val": rsid_list[0]})
        for rsid in rsid_list:
            etree.SubElement(rsids, f"{{{ns}}}rsid", {f"{{{ns}}}val": rsid})

        # Insert it early inside settings.xml
        root.insert(0, rsids)

        tree = etree.ElementTree(root)
        tree.write(settings_path, pretty_print=True, encoding="UTF-8", xml_declaration=True)

        self.logger.info(f"📝 Injected {len(rsid_list)} RSID entries into settings.xml.")
        
    def _ensure_styles_definitions(self, tmp_dir):
        """
        Ensure that styles.xml exists and contains mandatory base styles (Normal, TableNormal, etc.).
        """

        ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        nsmap = {"w": ns}
        styles_path = os.path.join(tmp_dir, "word", "styles.xml")

        # Check if styles.xml exists
        if not os.path.exists(styles_path):
            self.logger.warning(f"⚠️ styles.xml missing — creating minimal version.")

            # Create minimal root
            root = etree.Element(f"{{{ns}}}styles", nsmap=nsmap)
            tree = etree.ElementTree(root)
        else:
            parser = etree.XMLParser(remove_blank_text=True)
            try:
                tree = etree.parse(styles_path, parser)
                root = tree.getroot()
            except Exception as ex:
                self.logger.warning(f"⚠️ styles.xml corrupted ({ex}) — recreating.")
                root = etree.Element(f"{{{ns}}}styles", nsmap=nsmap)
                tree = etree.ElementTree(root)

        existing_styles = {s.get(f"{{{ns}}}styleId") for s in root.findall(".//w:style", namespaces=nsmap)}

        # Define required base styles
        
        added = 0
        for style_id, style_type, style_name, is_default in self.required_styles:
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
                root.append(style)
                added += 1

        # Save only if we added anything
        if added > 0 or not os.path.exists(styles_path):
            tree.write(styles_path, pretty_print=True, encoding="UTF-8", xml_declaration=True)
            self.logger.info(f"📝 Added {added} missing styles into styles.xml during repair.")
        else:
            self.logger.info(f"✅ All critical styles already present in styles.xml during repair.")


def main():
    parser = argparse.ArgumentParser(description="Repair .docx files using Word automation (Windows only).")
    parser.add_argument("input", help="Path to .docx file or folder containing .docx files")
    parser.add_argument("-o", "--output", help="Output file or folder (optional)")
    args = parser.parse_args()

    if platform.system() != "Windows":
        print("❌ This script only works on Windows (requires Microsoft Word).")
        sys.exit(1)

    repairer = WordRepairer()

    if os.path.isdir(args.input):
        output_dir = args.output if args.output and os.path.isdir(args.output) else None
        repaired = repairer.repair_batch(args.input, output_dir)
        print(f"✅ Repaired {len(repaired)} files.")
    elif os.path.isfile(args.input):
        repairer.repair(args.input, args.output)
    else:
        print("❌ Invalid input path provided.")


if __name__ == "__main__":
    main()
