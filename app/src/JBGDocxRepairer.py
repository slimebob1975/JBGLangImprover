import os
import sys
import platform
import argparse

if platform.system() == "Windows":
    import win32com.client

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
            raise Exception(f"❌ Failed to repair file {input_path}: {ex}")

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
