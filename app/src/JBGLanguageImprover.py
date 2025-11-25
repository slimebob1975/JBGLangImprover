import os
import sys
try:
    from app.src.JBGDocumentStructureExtractor import DocumentStructureExtractor
    from app.src.JBGLangImprovSuggestorAI import JBGLangImprovSuggestorAI
    from app.src.JBGSuperDocumentEditor import JBGSuperDocumentEditor
    from app.src.JBGDocumentEditor import JBGDocumentEditor
    from app.src.JBGDocxRepairer import AutoDocxRepairer
except ModuleNotFoundError as ex:
    print(f"Some modules could not be imported: {str(ex)}")
    from JBGDocumentStructureExtractor import DocumentStructureExtractor
    from JBGLangImprovSuggestorAI import JBGLangImprovSuggestorAI
    from JBGSuperDocumentEditor import JBGSuperDocumentEditor
    from JBGDocumentEditor import JBGDocumentEditor
    from JBGDocxRepairer import AutoDocxRepairer

import logging

class JBGLanguageImprover:
        
    def __init__(self, input_path, api_key, model, prompt_policy, temperature, include_motivations, docx_mode, logger, progress_callback=None):
        self.input_path = input_path
        self.api_key = api_key
        self.model = model
        self.prompt_policy = prompt_policy
        self.temperature = temperature
        self.include_motivations = include_motivations
        self.docx_mode = docx_mode
        self.logger = logger
        self.progress_callback = progress_callback
        self.structure_json = input_path.replace(os.path.splitext(input_path)[1], "_structure.json")
        self.suggestions_json = input_path.replace(os.path.splitext(input_path)[1], "_suggestions.json")

    def run(self):

        self._report("🔍 Analyserar dokumentets struktur...")
        extractor = DocumentStructureExtractor(self.input_path, self.logger)
        extractor.extract()
        self.structure_json = extractor.save_as_json()

        self._report("🧠 Skickar dokumentet till språkmodellen för förslag...")
        ai = JBGLangImprovSuggestorAI(
            self.api_key,
            self.model,
            self.prompt_policy,
            self.temperature,
            self.logger,
            progress_callback=self.progress_callback,
        )
        ai.load_structure(self.structure_json)
        ai.suggest_changes_token_aware_batching()
        self.suggestions_json = ai.save_as_json()

        self._report("✏️ Förslagen appliceras på dokumentet...")
        try:
            editor = JBGSuperDocumentEditor(
                self.input_path,
                self.suggestions_json,
                self.include_motivations,
                self.docx_mode,
                self.logger,
            )
            editor.apply_changes()
            output_path = editor.save_edited_document()
        except Exception as ex:
            self.logger.warning(
                f"⚠️ JBGSuperDocumentEditor misslyckades: {str(ex)}. Använder fallback-implementation."
            )
            editor = JBGDocumentEditor(self.input_path, self.suggestions_json, self.include_motivations, self.logger)
            editor.apply_changes()
            output_path = editor.save_edited_document()

            # Attempt repairs (included in the SuperEditor above)
            if editor.ext == ".docx":
                try:
                    repair_path = output_path
                    self._report("🔧 Försöker reparera Word-dokumentet om det är delvis korrupt...")
                    repair_path = AutoDocxRepairer(logger=self.logger).repair(repair_path)
                except Exception as ex:
                    self.logger.info(f"❌ Misslyckades med att reparera Word-dokumentet. Orsak: {str(ex)}")
                else:
                    output_path = repair_path

        self._report(f"✅ Färdigt. Det förbättrade dokumentet sparades.")

        return output_path
    
    def _report(self, message: str):
        self.logger.info(message)
        if self.progress_callback is not None:
            try:
                self.progress_callback(message)
            except Exception as ex:
                self.logger.warning(f"Progress callback failed: {ex}")

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
