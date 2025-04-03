import json
import openai
import sys
import os
import re
import time
import logging

MAX_TOKEN_PER_CALL = 3000

class JBGLangImprovSuggestorAI:
    
    def __init__(self, api_key, model, prompt_policy, logger):
        
        self.api_key = api_key
        self.model = model
        self.policy_prompt = prompt_policy
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

        self.logger.info(f"The used prompt policy:\n{str(self.policy_prompt)}\n\n")

        client = openai.OpenAI(api_key=self.api_key)
        
        messages = [
            {"role": "system", "content": self.policy_prompt},
            {"role": "user", "content": f"Här är det dokument som ska granskas: {self.json_structured_document}."}
        ]

        try:
            response = client.chat.completions.create(
                model=self.model,
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
        
        self.logger.info(f"The used prompt policy:\n{str(self.policy_prompt)}\n\n")
        
        client = openai.OpenAI(api_key=self.api_key)
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
                    model=self.model,
                    messages=messages,
                    temperature=0.7
                )
                suggestions = response.choices[0].message.content
                cleaned = self._clean_json_response(suggestions)
                parsed = json.loads(cleaned)
                all_suggestions.extend(parsed)
                self.logger.info(f"✅ Executed API requests #{i+1}.")
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
    
    if len(sys.argv) != 6:
        print(f"Usage: python {os.path.basename(__file__)} <structure_file.json> <api_key> <model> <prompt_policy_file> <custom_addition_file>")
        sys.exit(1)

    filepath = sys.argv[1]
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
    logger = logging.getLogger("ai-test")
    logger.setLevel(logging.INFO)
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger.handlers.clear()
    logger.addHandler(handler)

    ai = JBGLangImprovSuggestorAI(api_key, model, full_prompt, logger)
    ai.load_structure(filepath)
    ai.suggest_changes_token_aware_batching()
    ai.save_as_json()

if __name__ == "__main__":
    main()

