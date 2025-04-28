import zipfile
import os
import shutil
import xml.etree.ElementTree as ET

class JBGDocxQualityChecker:
    
    def __init__(self, doxc_path, logger=None):
        self.docx_path = doxc_path
        self.logger = logger

    def _unzip_docx(self, extract_to):
        if os.path.exists(extract_to):
            shutil.rmtree(extract_to)
        os.makedirs(extract_to, exist_ok=True)
        with zipfile.ZipFile(self.docx_path, 'r') as zip_ref:
            zip_ref.extractall(extract_to)

    def _check_critical_styles(self, styles_path):
        critical = {'Normal', 'DefaultParagraphFont', 'TableNormal', 'CommentText', 'InsertedText', 'DeletedText'}
        present = set()
        try:
            tree = ET.parse(styles_path)
            root = tree.getroot()
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            for style in root.findall('.//w:style', ns):
                style_id = style.attrib.get(f'{{{ns["w"]}}}styleId')
                if style_id:
                    present.add(style_id)
        except Exception as e:
            self.logger.error(f"Error parsing styles.xml: {e}")
        missing = critical - present
        return missing

    def _check_settings(self, settings_path):
        track_revisions = False
        rsids = set()
        try:
            tree = ET.parse(settings_path)
            root = tree.getroot()
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            if root.find('w:trackRevisions', ns) is not None:
                track_revisions = True
            for rsid in root.findall('w:rsids/w:rsid', ns):
                val = rsid.attrib.get(f'{{{ns["w"]}}}val')
                if val:
                    rsids.add(val)
        except Exception as e:
            self.logger.error(f"Error parsing settings.xml: {e}")
        return track_revisions, rsids

    def quality_control_docx(self):
        extract_dir = self.docx_path.replace('.docx', '_extract')
        self._unzip_docx(extract_dir)
        
        styles_path = os.path.join(extract_dir, 'word/styles.xml')
        settings_path = os.path.join(extract_dir, 'word/settings.xml')
        
        missing_styles = self._check_critical_styles(styles_path) if os.path.exists(styles_path) else set()
        track_revisions, rsids = self._check_settings(settings_path) if os.path.exists(settings_path) else (False, set())
        
        self.logger.info("=== DOCX Quality Control Report ===")
        self.logger.info(f"Document: {os.path.basename(self.docx_path)}")
        self.logger.info("------------------------------------")
        self.logger.info(f"Critical Styles Missing: {missing_styles}")
        self.logger.info(f"Track Changes Enabled: {track_revisions}")
        self.logger.info(f"Number of RSIDs: {len(rsids)}")
        if len(rsids) < 10:
            self.logger.warning("⚠️ Warning: RSIDs look sparse — consider enriching the editing history.")
        else:
            self.logger.info("✅ RSIDs look realistic.")
        self.logger.info("------------------------------------")

        shutil.rmtree(extract_dir)

if __name__ == "__main__":
    import sys
    if len(sys.argv) != 2:
        print("Usage: python JBGDocxQualityChecker.py your_document.docx")
    else:
        checker = JBGDocxQualityChecker(sys.argv[1])
        checker.quality_control_docx()
