import zipfile
import pytz
from lxml import etree
from datetime import datetime
from copy import deepcopy

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

class JBGTrackChangesActivatorInserter:
    def __init__(self, docx_path, author="JBG klarspråkningstjänst", timezone="Europe/Stockholm"):
        self.docx_path = docx_path
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
        r = etree.Element('{%s}r' % NS['w'])
        t = etree.Element('{%s}t' % NS['w'], attrib={"{http://www.w3.org/XML/1998/namespace}space": "preserve"})
        t.text = text
        r.append(t)
        return r

    def _make_deleted_run(self, text):
        r = etree.Element('{%s}r' % NS['w'])
        del_text = etree.Element('{%s}delText' % NS['w'], attrib={"{http://www.w3.org/XML/1998/namespace}space": "preserve"})
        del_text.text = text
        r.append(del_text)
        return r

    def enable_track_changes(self):
        with zipfile.ZipFile(self.docx_path, 'r') as zin:
            settings_xml = zin.read('word/settings.xml')
            settings_other = {name: zin.read(name) for name in zin.namelist() if name != 'word/settings.xml'}

        tree = etree.fromstring(settings_xml)
        if tree.find('w:trackRevisions', NS) is None:
            track = etree.Element('{%s}trackRevisions' % NS['w'])
            tree.insert(0, track)

        updated_settings = etree.tostring(tree, xml_declaration=True, encoding='UTF-8')

        with zipfile.ZipFile(self.docx_path, 'w') as zout:
            for name, content in settings_other.items():
                zout.writestr(name, content)
            zout.writestr('word/settings.xml', updated_settings)

    def apply_to_paragraph(self, para, old_text, new_text):
        runs = para.xpath('.//w:r', namespaces=NS)
        for run in runs:
            text_el = run.find('w:t', NS)
            if text_el is not None and old_text in text_el.text:
                full_text = text_el.text
                before, _, after = full_text.partition(old_text)

                para.remove(run)

                if before.strip():
                    para.append(self._make_run(before))

                del_el = etree.Element('{%s}del' % NS['w'], attrib={
                    '{%s}author' % NS['w']: self.author,
                    '{%s}date' % NS['w']: self.timestamp
                })
                del_el.append(self._make_deleted_run(old_text))
                para.append(del_el)

                ins_el = etree.Element('{%s}ins' % NS['w'], attrib={
                    '{%s}author' % NS['w']: self.author,
                    '{%s}date' % NS['w']: self.timestamp
                })
                ins_el.append(self._make_run(new_text))
                para.append(ins_el)

                if after.strip():
                    para.append(self._make_run(after))
                return True  # One match per paragraph
        return False

    def apply_tracked_replacement(self, old_text, new_text):
        change_count = 0
        for para in self.doc_tree.xpath('//w:p', namespaces=NS):
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
    inserter = JBGTrackChangesActivatorInserter(output_path)
    inserter.enable_track_changes()

    if target_para:
        paras = inserter.doc_tree.xpath('//w:p', namespaces=NS)
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

