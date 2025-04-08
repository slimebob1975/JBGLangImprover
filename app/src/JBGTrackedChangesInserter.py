import zipfile
import pytz
from lxml import etree
from datetime import datetime
import re

DEFAULT_NAMESPACE = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

class JBGTrackedChangesInserter:
    def __init__(self, docx_path, logger, namespace = DEFAULT_NAMESPACE, \
        author="JBG klarspråkningstjänst", timezone="Europe/Stockholm"):
        self.docx_path = docx_path
        self.logger = logger
        self.namespace = namespace
        self.author = author
        self.tz = pytz.timezone(timezone)
        self.timestamp = datetime.now(self.tz).isoformat()
        self._load_document_xml()

    def _load_document_xml(self):
        with zipfile.ZipFile(self.docx_path, 'r') as zin:
            self.document_xml = zin.read('word/document.xml')
            self.other_files = {name: zin.read(name) for name in zin.namelist() if name != 'word/document.xml'}

        self.doc_tree = etree.fromstring(self.document_xml)

    def _make_run(self, text):
        r = etree.Element('{%s}r' % self.namespace['w'])
        t = etree.Element('{%s}t' % self.namespace['w'], attrib={"{http://www.w3.org/XML/1998/namespace}space": "preserve"})
        t.text = text
        r.append(t)
        return r

    def _make_deleted_run(self, text):
        r = etree.Element('{%s}r' % self.namespace['w'])
        del_text = etree.Element('{%s}delText' % self.namespace['w'], attrib={"{http://www.w3.org/XML/1998/namespace}space": "preserve"})
        del_text.text = text
        r.append(del_text)
        return r

    def enable_track_changes(self):
        with zipfile.ZipFile(self.docx_path, 'r') as zin:
            settings_xml = zin.read('word/settings.xml')
            settings_other = {name: zin.read(name) for name in zin.namelist() if name != 'word/settings.xml'}

        tree = etree.fromstring(settings_xml)
        if tree.find('w:trackRevisions', self.namespace) is None:
            track = etree.Element('{%s}trackRevisions' % self.namespace['w'])
            tree.insert(0, track)

        updated_settings = etree.tostring(tree, xml_declaration=True, encoding='UTF-8')

        with zipfile.ZipFile(self.docx_path, 'w') as zout:
            for name, content in settings_other.items():
                zout.writestr(name, content)
            zout.writestr('word/settings.xml', updated_settings)

    def apply_to_paragraph(self, para, old_text, new_text):
        try:
            # Recombine paragraph text from current runs
            current_runs = para.xpath('.//w:r', namespaces=self.namespace)
            full_text = "".join(
                t.text for r in current_runs for t in r.xpath('.//w:t', namespaces=self.namespace) if t.text
            )

            norm_old = self._normalize_text(old_text)
            norm_full = self._normalize_text(full_text)

            if norm_old not in norm_full:
                if self.logger:
                    self.logger.warning(f"❌ Normalized match not found: '{old_text}' not in paragraph")
                    self.logger.debug(f"🧾 Full paragraph: {full_text}")
                return False

            # Split original paragraph into 3 logical parts
            before, _, after = full_text.partition(old_text)

            # Clear only current runs (safely)
            runs_to_remove = para.xpath('./w:r', namespaces=self.namespace)
            for run in runs_to_remove:
                if any(child.tag.endswith("delText") or child.tag.endswith("ins") for child in run):
                    continue  # skip already modified content
                try:
                    para.remove(run)
                except Exception as e:
                    if self.logger:
                        self.logger.error(f"❌ Failed to remove original run: {etree.tostring(run)}")
                    raise e


            # Rebuild paragraph
            if before.strip():
                para.append(self._make_run(before))

            del_el = etree.Element('{%s}del' % self.namespace['w'], attrib={
                '{%s}author' % self.namespace['w']: self.author,
                '{%s}date' % self.namespace['w']: self.timestamp
            })
            del_el.append(self._make_deleted_run(old_text))
            para.append(del_el)

            ins_el = etree.Element('{%s}ins' % self.namespace['w'], attrib={
                '{%s}author' % self.namespace['w']: self.author,
                '{%s}date' % self.namespace['w']: self.timestamp
            })
            ins_el.append(self._make_run(new_text))
            para.append(ins_el)

            if after.strip():
                para.append(self._make_run(after))

            return True

        except Exception as e:
            if self.logger:
                self.logger.exception(f"🔥 Exception during tracked change: '{old_text}' → '{new_text}'")
            raise

    def apply_tracked_replacement(self, old_text, new_text):
        change_count = 0
        for para in self.doc_tree.xpath('//w:p', namespaces=self.namespace):
            if self.apply_to_paragraph(para, old_text, new_text):
                change_count += 1

        return change_count

    def save(self, output_path=None):
        new_doc_xml = etree.tostring(self.doc_tree, xml_declaration=True, encoding='UTF-8')
        if output_path is None:
            output_path = self.docx_path

        with zipfile.ZipFile(output_path, 'w') as zout:
            for name, content in self.other_files.items():
                zout.writestr(name, content)
            zout.writestr('word/document.xml', new_doc_xml)
            
    @staticmethod
    def _normalize_text(text):
        return re.sub(r'\s+', ' ', text).strip()
            
def main():
    import sys
    import os
    import shutil

    if len(sys.argv) < 4:
        print(f"Usage: python {os.path.basename(__file__)} <input.docx> <old_text> <new_text> [<paragraph_number>]")
        sys.exit(1)

    input_path = sys.argv[1]
    old_text = sys.argv[2]
    new_text = sys.argv[3]
    target_para = int(sys.argv[4]) if len(sys.argv) == 5 else None

    if not input_path.endswith(".docx") or not os.path.exists(input_path):
        print("❌ Invalid file path or not a .docx file.")
        sys.exit(1)

    # Create modified copy
    output_path = input_path.replace(".docx", "_modified.docx")
    shutil.copyfile(input_path, output_path)

    # Run track change insertion
    inserter = JBGTrackedChangesInserter(output_path)
    inserter.enable_track_changes()

    if target_para:
        paras = inserter.doc_tree.xpath('//w:p', namespaces=DEFAULT_NAMESPACE)
        if 1 <= target_para <= len(paras):
            para = paras[target_para - 1]
            success = inserter.apply_to_paragraph(para, old_text, new_text)
            print(f"✅ Change {'applied' if success else 'not found'} in paragraph {target_para}")
        else:
            print(f"⚠️ Invalid paragraph number. Document has {len(paras)} paragraphs.")
    else:
        count = inserter.apply_tracked_replacement(old_text, new_text)
        print(f"✅ Total tracked changes applied: {count}")

    inserter.save()
    print(f"📄 Saved as: {output_path}")

if __name__ == "__main__":
    main()

