import json
import os
import sys
from JBGDocumentStructureExtractor import DocumentStructureExtractor
from JBGLangImprovSuggestorAI import JBGLangImprovSuggestorAI
from JBGDocumentEditor import JBGDocumentEditor

class JBGLanguageImprover:
        
    def __init__(self, input_path, key_file, policy_file):
        self.input_path = input_path
        self.key_file = key_file
        self.policy_file = policy_file
        self.structure_json = input_path.replace(os.path.splitext(input_path)[1], "_structure.json")
        self.suggestions_json = input_path.replace(os.path.splitext(input_path)[1], "_suggestions.json")

    def run(self):
        
        print("🔍 Extracting structure...")
        extractor = DocumentStructureExtractor(self.input_path)
        extractor.extract()
        self.structure_json = extractor.save_as_json()

        print("🧠 Generating suggestions with AI...")
        ai = JBGLangImprovSuggestorAI(self.key_file, self.policy_file)
        ai.load_structure(self.structure_json)
        ai.suggest_changes_token_aware_batching()
        self.suggestions_json = ai.save_as_json()

        print("✏️ Applying suggestions to document...")
        editor = JBGDocumentEditor(self.input_path, self.suggestions_json)
        editor.apply_changes()
        output_path = editor.save_edited_document()

        print(f"✅ Final improved document saved to: {output_path}")
        return output_path

def main():
    
    if len(sys.argv) != 4:
        print(f"Usage: python {os.path.basename(__file__)} <document_path> <azure_key_file> <policy_file>")
        sys.exit(1)

    improver = JBGLanguageImprover(
        input_path=sys.argv[1],
        key_file=sys.argv[2],
        policy_file=sys.argv[3]
    )
    improver.run()
    
if __name__ == "__main__":
    main()