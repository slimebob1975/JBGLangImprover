# 📘 JBG Language Improvement System

This project provides a fully functional web-based system for **language improvement of `.docx` and `.pdf` documents** using OpenAI models (like `gpt-4o`, `gpt-4`, `gpt-3.5-turbo`, etc.). It is designed to help authors, editors, and public agencies improve clarity and readability in line with plain language principles like *klarspråk*.

---

## ✨ What the System Does

1. **Extracts document structure**: Paragraphs for `.docx`, and page-line text layout for `.pdf`.
2. **Sends structured text to OpenAI models** for suggestions, including tracked changes, clarifications, and better wording — based on a policy-driven prompt.
3. **Annotates PDF files** with comments and highlights suggested changes.
4. **Edits DOCX files** by marking old text with strikethrough + red color and inserting improved text in green.
5. **Logs each session** to a timestamped log file for debugging and traceability.

---

## ⚙️ Features

- 🔐 User-provided API key (never stored server-side)
- 🔄 Model selection for cost/performance control
- 📝 Customizable prompt additions per session
- 💾 Downloadable output with improved language
- 🧠 Session logs for advanced usage
- 🧽 Auto-cleaning of old uploads and logs

---

## 📁 Folder Structure Overview

```
language-improver/
├── app/
│   ├── main.py                     # FastAPI app
│   └── src/
│       ├── JBGLanguageImprover.py
│       ├── JBGLangImprovSuggestorAI.py
│       ├── JBGDocumentEditor.py
│       └── JBGDocumentStructureExtractor.py
├── policy/
│   ├── prompt_policy.md            # ✅ Base prompt file (required)
│
├── templates/
│   └── index.html                  # HTML frontend
├── static/
│   ├── styles/
│   │   └── styles.css
│   ├── javascript/
│   │   └── script.js
├── uploads/                        # Temp files, deleted after 24h
├── logs/                           # Session logs
```

---

## 🧾 Creating Your Prompt Policy

The core of this tool is its **prompt policy**, defined in `policy/prompt_policy.md`.

To create your own:
1. Create a file in the `policy/` folder.
2. Write instructions describing what kind of improvements you'd like (e.g., "Use clear language. Avoid abbreviations. Rewrite long sentences.")
3. You can also extend the base prompt with session-specific notes via the GUI.

Example contents of `prompt_policy.md` in Swedish:

```markdown
Klarspråk är språk som är vårdat, enkelt och begripligt. Undvik förkortningar och svåra uttryck. Förklara tekniska termer om möjligt. Korta ner långa meningar.
```

---

## 🚀 Running Locally with Uvicorn

To run the app locally for testing:

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Start the server:
   ```bash
   uvicorn app.main:app --reload
   ```

3. Open your browser and visit:
   ```
   http://127.0.0.1:8000
   ```

From here, you can upload a document, enter your OpenAI API key, choose a model, and optionally refine the prompt before improving your language!

---

## 👥 Contributing

This system is extensible — feel free to contribute new prompt strategies, support for more file types, or enhancements to model interaction logic.

---
