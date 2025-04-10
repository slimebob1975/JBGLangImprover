import os
import sys
from app.src.JBGDocumentStructureExtractor import DocumentStructureExtractor
from app.src.JBGLangImprovSuggestorAI import JBGLangImprovSuggestorAI
from app.src.JBGDocumentEditor import JBGDocumentEditor

import logging

class JBGLanguageImprover:
        
    def __init__(self, input_path, api_key, model, prompt_policy, temperature, include_comments, docx_mode, logger):
        self.input_path = input_path
        self.api_key = api_key
        self.model = model
        self.prompt_policy = prompt_policy
        self.temperature = temperature
        self.include_comments = include_comments
        self.docx_mode = docx_mode
        self.logger = logger
        self.structure_json = input_path.replace(os.path.splitext(input_path)[1], "_structure.json")
        self.suggestions_json = input_path.replace(os.path.splitext(input_path)[1], "_suggestions.json")

    def run(self):
        
        self.logger.info("🔍 Extracting structure...")
        extractor = DocumentStructureExtractor(self.input_path, self.logger)
        extractor.extract()
        self.structure_json = extractor.save_as_json()

        self.logger.info("🧠 Generating suggestions with AI...")
        ai = JBGLangImprovSuggestorAI(self.api_key, self.model, self.prompt_policy, self.temperature, self.logger)
        ai.load_structure(self.structure_json)
        ai.suggest_changes_token_aware_batching()
        self.suggestions_json = ai.save_as_json()

        self.logger.info("✏️ Applying suggestions to document...")
        editor = JBGDocumentEditor(self.input_path, self.suggestions_json, self.include_comments, self.docx_mode, self.logger)
        editor.apply_changes()
        output_path = editor.save_edited_document()

        # End logging and close logging file
        self.logger.info(f"✅ Final improved document saved to: {output_path}")
            
        return output_path

def main():
    
    if len(sys.argv) != 6:
        print(f"Usage: python {os.path.basename(__file__)} <document_path> <api_key> <model> <prompt_policy_file> <custom_addition_file>")
        sys.exit(1)

    input_path = sys.argv[1]
    api_key = sys.argv[2]
    model = sys.argv[3]
    policy_path = sys.argv[4]
    custom_path = sys.argv[5]

    # Load and merge prompts
    with open(policy_path, encoding="utf-8") as f:
        base_prompt = f.read().strip()
    full_prompt = base_prompt
    if os.path.exists(custom_path):
        with open(custom_path, encoding="utf-8") as f:
            custom = f.read().strip()
        full_prompt += "\n\n" + custom

    # Set up logger
    logger = logging.getLogger("test-run")
    logger.setLevel(logging.DEBUG)
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger.handlers.clear()
    logger.addHandler(handler)

    improver = JBGLanguageImprover(
        input_path=input_path,
        api_key=api_key,
        model=model,
        prompt_policy=full_prompt,
        logger=logger
    )
    improver.run()

if __name__ == "__main__":
    main()
