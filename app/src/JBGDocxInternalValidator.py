# EXTENDED DocxInternalValidator TO SUPPORT XML WELL-FORMEDNESS CHECKING

import zipfile
import os
from lxml import etree
import shutil

class DocxInternalValidator:
    def __init__(self, docx_path):
        self.docx_path = docx_path
        self.errors = []
        self.temp_dir = None

    def _extract_docx(self):
        from tempfile import mkdtemp
        self.temp_dir = mkdtemp()
        with zipfile.ZipFile(self.docx_path, 'r') as zin:
            zin.extractall(self.temp_dir)

    def _validate_relationships(self):
        rels_path = os.path.join(self.temp_dir, "word", "_rels", "document.xml.rels")
        if not os.path.exists(rels_path):
            self.errors.append("❌ Missing document.xml.rels")
            return

        tree = etree.parse(rels_path)
        root = tree.getroot()
        rel_targets = {r.get("Target") for r in root.findall(".//", namespaces={'': "http://schemas.openxmlformats.org/package/2006/relationships"})}

        required_targets = [
            "styles.xml", "settings.xml", "comments.xml", "fontTable.xml", "webSettings.xml"
        ]
        for target in required_targets:
            if not any(target in t for t in rel_targets):
                self.errors.append(f"⚠️ Missing relationship to {target}")

    def _validate_styles(self):
        styles_path = os.path.join(self.temp_dir, "word", "styles.xml")
        if not os.path.exists(styles_path):
            self.errors.append("❌ Missing styles.xml")
            return

        tree = etree.parse(styles_path)
        root = tree.getroot()

        style_ids = {s.get("w:styleId") for s in root.findall(".//w:style", namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})}

        required_styles = ["Normal", "DefaultParagraphFont", "TableNormal", "CommentText", "InsertedText", "DeletedText"]

        for style in required_styles:
            if style not in style_ids:
                self.errors.append(f"⚠️ Missing style: {style}")

    def _validate_settings(self):
        settings_path = os.path.join(self.temp_dir, "word", "settings.xml")
        if not os.path.exists(settings_path):
            self.errors.append("❌ Missing settings.xml")
            return

        tree = etree.parse(settings_path)
        root = tree.getroot()
        nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

        if root.find("w:trackRevisions", namespaces=nsmap) is None:
            self.errors.append("⚠️ Missing <w:trackRevisions> in settings.xml")

    def _validate_comments(self):
        comments_path = os.path.join(self.temp_dir, "word", "comments.xml")
        if not os.path.exists(comments_path):
            return

        tree = etree.parse(comments_path)
        root = tree.getroot()
        comment_ids = {c.get("w:id") for c in root.findall(".//w:comment", namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})}

        document_path = os.path.join(self.temp_dir, "word", "document.xml")
        if not os.path.exists(document_path):
            self.errors.append("❌ Missing document.xml for comment validation")
            return

        doc_tree = etree.parse(document_path)
        refs = doc_tree.findall(".//w:commentReference", namespaces={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
        for ref in refs:
            ref_id = ref.get("w:id")
            if ref_id not in comment_ids:
                self.errors.append(f"⚠️ commentReference id={ref_id} has no corresponding comment.")

    def _check_wellformed_parts(self):
        with zipfile.ZipFile(self.docx_path, 'r') as z:
            for item in z.namelist():
                if item.endswith('.xml'):
                    try:
                        data = z.read(item)
                        etree.fromstring(data)
                    except etree.XMLSyntaxError as e:
                        self.errors.append(f"❌ {item} is broken: {e}")

    def validate(self):
        self._extract_docx()
        self._validate_relationships()
        self._validate_styles()
        self._validate_settings()
        self._validate_comments()
        self._check_wellformed_parts()
        shutil.rmtree(self.temp_dir)
        return self.errors

if __name__ == "__main__":
    import sys
    validator = DocxInternalValidator(sys.argv[1])
    problems = validator.validate()
    if not problems:
        print("✅ No structural problems detected!")
    else:
        print("\n".join(problems))
