import os
import sys
from app.src.JBGDocumentStructureExtractor import DocumentStructureExtractor
from app.src.JBGLangImprovSuggestorAI import JBGLangImprovSuggestorAI
from app.src.JBGDocumentEditor import JBGDocumentEditor
import logging

class JBGLanguageImprover:
        
    def __init__(self, input_path, key_file, policy_file, logger):
        self.input_path = input_path
        self.key_file = key_file
        self.policy_file = policy_file
        self.logger = logger
        self.structure_json = input_path.replace(os.path.splitext(input_path)[1], "_structure.json")
        self.suggestions_json = input_path.replace(os.path.splitext(input_path)[1], "_suggestions.json")

    def run(self):
        
        self.logger.info("🔍 Extracting structure...")
        extractor = DocumentStructureExtractor(self.input_path, self.logger)
        extractor.extract()
        self.structure_json = extractor.save_as_json()

        self.logger.info("🧠 Generating suggestions with AI...")
        ai = JBGLangImprovSuggestorAI(self.key_file, self.policy_file, self.logger)
        ai.load_structure(self.structure_json)
        ai.suggest_changes_token_aware_batching()
        self.suggestions_json = ai.save_as_json()

        self.logger.info("✏️ Applying suggestions to document...")
        editor = JBGDocumentEditor(self.input_path, self.suggestions_json, self.logger)
        editor.apply_changes()
        output_path = editor.save_edited_document()

        # End logging and close logging file
        self.logger.info(f"✅ Final improved document saved to: {output_path}")
            
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