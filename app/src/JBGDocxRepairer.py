import os
import sys
import platform
import argparse
import zipfile
from tempfile import mkdtemp
from lxml import etree
    
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
        return self.repairer.repair(input_path, output_path)


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

    def repair(self, docx_path, output_path=None):
        try:
            with zipfile.ZipFile(docx_path, "r") as zin:
                zin.testzip()  # Check for basic corruption
                if any("word/comments.xml" in f for f in zin.namelist()):
                    self.logger.info("✅ comments.xml exists.")
                    return docx_path  # Nothing to fix
                else:
                    self.logger.warning("⚠️ comments.xml missing — creating minimal version.")
                    return self._add_minimal_comments(docx_path, output_path)

        except zipfile.BadZipFile:
            raise Exception("❌ .docx file is not a valid ZIP archive — unrecoverable.")

    def _add_minimal_comments(self, path, output_path=None):
       
        tmp_dir = mkdtemp()
        with zipfile.ZipFile(path, "r") as zin:
            zin.extractall(tmp_dir)

        comments_path = os.path.join(tmp_dir, "word", "comments.xml")
        os.makedirs(os.path.dirname(comments_path), exist_ok=True)

        nsmap = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        root = etree.Element(f"{{{nsmap['w']}}}comments", nsmap=nsmap)
        etree.ElementTree(root).write(comments_path, pretty_print=True, encoding="UTF-8", xml_declaration=True)

        if not output_path:
            output_path = path.replace(".docx", "_repaired.docx")

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for root_dir, _, files in os.walk(tmp_dir):
                for file in files:
                    full_path = os.path.join(root_dir, file)
                    archive_name = os.path.relpath(full_path, tmp_dir)
                    zout.write(full_path, archive_name)

        self.logger(f"🔧 Repaired and saved as: {output_path}")
        return output_path
    
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
