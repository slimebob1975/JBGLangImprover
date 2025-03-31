import json
import openai
import sys
import os
import re

class JBGLangImprovSuggestorAI:
    def __init__(self, key_file, policy_file):
        
        with open(key_file, 'r') as f:
            self.keys = json.load(f)
        
        with open(policy_file, 'r', encoding='utf-8') as f:
            self.policy_prompt = f.read()

        self.file_path = None
        self.json_structured_document = None
        self.json_suggestions = None
        
    def load_structure(self, filepath):
        self.file_path = filepath
        self.json_structured_document = json.load(open(filepath, "r", encoding="utf-8"))
        if not self.json_structured_document:
            print(f"Error: Could not load JSON document from {filepath}")
            sys.exit(1)
            
    def save_as_json(self, output_path = None):
        if not self.json_suggestions:
            self.suggest_changes()
        if not output_path:
            output_path = self.file_path + "_suggestions.json"
        try:
            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(self.json_suggestions, f, indent=2, ensure_ascii=False)
            return output_path
        except Exception as e:
            print(f"Error saving JSON structure: {str(e)}")
            return None
    
    def suggest_changes(self):

        client = openai.OpenAI(api_key=self.keys["api_key"])
        
        messages = [
            {"role": "system", "content": self.policy_prompt},
            {"role": "user", "content": f"Här är det dokument som ska granskas: {self.json_structured_document}."}
        ]

        try:
            response = client.chat.completions.create(
                model=self.keys["model"],
                messages=messages,
                temperature=0.7
            )
            suggestions = response.choices[0].message.content
            cleaned_suggestions = self._clean_json_response(suggestions)
            self.json_suggestions = json.loads(cleaned_suggestions)
        except Exception as e:
            print(f"Error during OpenAI suggestion generation: {e}")
            self.json_suggestions = None
    import re

    def _clean_json_response(self, raw_text):
        
        # Remove code fences if present
        if raw_text.strip().startswith("```"):
            cleaned = re.sub(r"^```(?:json)?", "", raw_text.strip())
            cleaned = re.sub(r"```$", "", cleaned.strip())
            return cleaned.strip()
        return raw_text
        
def main():
    
    if len(sys.argv) != 3:
        print(f"Usage: python {os.path.basename(__file__)} <document JSON structure file> <policy file>")
        sys.exit(1)
    
    structure_path = sys.argv[1]
    policy_file = sys.argv[2]
    key_file = "./openai/azure_keys.json"

    suggestor_ai = JBGLangImprovSuggestorAI(key_file, policy_file)
    suggestor_ai.load_structure(structure_path)
    suggestor_ai.suggest_changes()
    output_path = structure_path.replace(".json", "_suggestions.json")
    suggestor_ai.save_as_json(output_path)

    print(f"Suggestions saved to: {output_path}")

if __name__ == "__main__":
    main()
