import json
import openai
import sys
import os
import re
import time
import logging

MAX_TOKEN_PER_CALL = 3000

class JBGLangImprovSuggestorAI:
    def __init__(self, key_file, policy_file, logger):
        
        with open(key_file, 'r') as f:
            self.keys = json.load(f)
        
        with open(policy_file, 'r', encoding='utf-8') as f:
            self.policy_prompt = f.read()

        self.logger = logger
        self.file_path = None
        self.json_structured_document = None
        self.json_suggestions = None
        
    def load_structure(self, filepath):
        self.file_path = filepath
        self.json_structured_document = json.load(open(filepath, "r", encoding="utf-8"))
        if not self.json_structured_document:
            self.logger.error(f"Error: Could not load JSON document from {filepath}")
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
            self.logger.error(f"Error saving JSON structure: {str(e)}")
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
            self.logger.error(f"Error during OpenAI suggestion generation: {e}")
            self.json_suggestions = None

    def suggest_changes_token_aware_batching(self, max_tokens_per_call=MAX_TOKEN_PER_CALL):
        client = openai.OpenAI(api_key=self.keys["api_key"])
        model_name = self.keys["model"]

        system_msg = {"role": "system", "content": self.policy_prompt}
        structure = self.json_structured_document

        if structure["type"] == "docx":
            elements = structure["paragraphs"]
        elif structure["type"] == "pdf":
            elements = [
                {"page": page["page"], "line": line["line"], "text": line["text"]}
                for page in structure["pages"] for line in page["lines"]
            ]
        else:
            raise ValueError("Unsupported document type.")

        chunks = []
        current_chunk = []
        current_char_count = len(self.policy_prompt)

        for elem in elements:
            elem_text = json.dumps(elem, ensure_ascii=False)
            if current_char_count + len(elem_text) > max_tokens_per_call * 4:
                chunks.append(current_chunk)
                current_chunk = [elem]
                current_char_count = len(self.policy_prompt) + len(elem_text)
            else:
                current_chunk.append(elem)
                current_char_count += len(elem_text)

        if current_chunk:
            chunks.append(current_chunk)

        self.logger.info(f"🔹 Sending {len(chunks)} separate API requests due to size.")

        all_suggestions = []
        first = True
        for i, chunk in enumerate(chunks):
            
            # Avoid congestion
            if not first:
                time.sleep(5)
            else:
                first = False
                
            user_prompt = f"Här är en del av dokumentet som ska granskas: {json.dumps(chunk, ensure_ascii=False)}"
            messages = [system_msg, {"role": "user", "content": user_prompt}]

            try:
                response = client.chat.completions.create(
                    model=model_name,
                    messages=messages,
                    temperature=0.7
                )
                suggestions = response.choices[0].message.content
                cleaned = self._clean_json_response(suggestions)
                parsed = json.loads(cleaned)
                all_suggestions.extend(parsed)
            except Exception as e:
                self.logger.error(f"Error in chunk {i+1}: {e}")

        self.json_suggestions = all_suggestions

    
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
