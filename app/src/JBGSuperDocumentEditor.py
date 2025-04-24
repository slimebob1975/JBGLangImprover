import os
import shutil
import uuid
import docx
from copy import deepcopy
from app.src.JBGDocumentEditor import JBGDocumentEditor
from app.src.JBGDocxRepairer import AutoDocxRepairer
from lxml import etree
from tempfile import mkdtemp
import zipfile
from datetime import datetime, timezone

class JBGSuperDocumentEditor:
    def __init__(self, filepath, changes_path, include_motivations, docx_mode, logger):
        self.filepath = filepath
        self.changes_path = changes_path
        self.include_motivations = include_motivations
        self.docx_mode = docx_mode
        self.logger = logger

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

        repairer = AutoDocxRepairer(logger=self.logger)
        repaired_path = repairer.repair(final_path)

        self.doc = docx.Document(repaired_path)
        self.final_path = repaired_path
        return self.doc
    
    def _convert_markup_to_tracked(self, input_docx_path: str) -> str:
        """
        Full rewrite of tracked conversion:
        - Convert red+strike → <w:del><w:r><w:delText>...</w:delText></w:r></w:del>
        - Convert green → <w:ins><w:r><w:t>...</w:t></w:r></w:ins>
        """
        nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        tracked_id = 1
        temp_dir = mkdtemp()
        output_path = os.path.join(temp_dir, os.path.basename(input_docx_path).replace("_interim", "_tracked"))

        with zipfile.ZipFile(input_docx_path, "r") as zin:
            zin.extractall(temp_dir)

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
            tracked_id += 1

            wrapper.append(run_copy)
            parent.remove(run)
            parent.insert(index, wrapper)
            
            # Re-insert comment reference if it was originally attached
            if comment_ref is not None and self.include_motivations:
                comment_ref_run = etree.Element(f"{{{nsmap['w']}}}r")
                comment_ref_run.append(etree.fromstring(etree.tostring(comment_ref)))
                parent.insert(index + 1, comment_ref_run)

        tree.write(doc_path, pretty_print=True, encoding="UTF-8", xml_declaration=True)
        self.logger.info("🔁 Replaced visual markups with tracked changes XML.")

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for root_dir, _, files in os.walk(temp_dir):
                for file in files:
                    full_path = os.path.join(root_dir, file)
                    archive_name = os.path.relpath(full_path, temp_dir)
                    zout.write(full_path, archive_name)

        self.logger.info(f"✅ Tracked DOCX saved: {output_path}")
        return output_path
    
    def save(self, output_path=None):
        if not output_path:
            output_path = self.final_path
        else:
            self.doc.save(output_path)
        return output_path

class EditorProcessingException(Exception):
    """Raised when the Super Editor fails and should defer to the legacy editor."""
    pass


