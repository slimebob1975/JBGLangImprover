
# JBG Language Improvement System

A web-based system for **klarspråksgranskning av Word-dokument (`.docx`)** using OpenAI models. The system extracts structured text from a document, asks a language model for policy-driven improvement suggestions, filters and validates the suggestions, and produces an edited Word document with visible proposed changes.

The current setup is optimized for Swedish plain-language review in formal public-sector texts, especially reports and similar documents.

---

## What the System Does

1. **Extracts Word document structure**The app extracts text from Word elements such as paragraphs, headings, table cells, headers, footers and footnotes.
2. **Sends structured text to OpenAI models**Each text element is sent to an OpenAI model together with a prompt policy. The model returns JSON suggestions containing:

   - `old`
   - `new`
   - `motivation`
   - element metadata such as `type`, `element_id` and, where relevant, `footnote_id`
3. **Uses a policy-driven prompt**The prompt policy is stored in `policy/prompt_policy.md`. It is split into locked and editable sections so the stable JSON/output requirements can be separated from the klarspråk rules that may be adjusted over time.
4. **Filters and validates suggestions**Suggested changes are checked before they are applied. The system can flag issues such as:

   - weak locality
   - multi-sentence rewrites
   - long replacement spans
   - spelling or formatting risks
   - changes that are difficult to anchor safely in the document
5. **Applies accepted suggestions to DOCX files**The editor marks removed text and inserted text in the Word document so that the user can review proposed changes.
6. **Writes logs and reports**Each run produces diagnostic output such as:

   - the raw suggestions JSON
   - the suggestion filter report
   - session logs
   - output document files

---

## Current Prompt Strategy

The prompt is designed for Swedish **klarspråk**: language that is clear, simple, correct and appropriate for the intended audience.

The current prompt gives priority to **klarspråksnytta** over strict edit locality. This means the model may suggest larger rewrites when they clearly improve comprehension, but it should avoid unnecessary rewrites when the original text is already clear and correct.

### Key prompt principles

The prompt tells the model to:

- adapt the text to the intended reader
- structure information logically
- simplify sentence structure and word choice
- follow Swedish writing rules and public-sector writing conventions
- use active voice when appropriate
- prefer verbs over nominalizations
- avoid unnecessary commas by making sentences clearer
- explain specialist terms when useful
- write out abbreviations when needed
- use clear and meaningful headings
- keep the tone professional, clear and accessible

### Change threshold

The model should only suggest changes when they provide a clear improvement in:

- comprehensibility
- clarity
- correctness

Purely stylistic variation without a clear benefit should be avoided.

### Locality and larger rewrites

The prompt normally asks for changes on word, phrase or sentence level. However, larger rewrites are allowed when:

- the text is difficult to understand
- the sentence structure is unnecessarily complex
- the structure makes the text harder to read

When making a larger rewrite, the model should:

- preserve the same information
- avoid adding new information or interpretations
- avoid introducing more paragraph breaks than necessary

### Terminology rules

The prompt currently includes specific terminology guidance:

- Prefer `a-kassor` and `a-kassorna` when the text uses `arbetslöshetskassor` or `arbetslöshetkassorna`.
- Do **not** rewrite `arbetslöshetsförsäkringen`; keep the word intact regardless of form or context.

---

## Prompt Policy File

The prompt policy lives here:

```text
policy/prompt_policy.md
```

The file uses comment markers to separate locked and editable parts:

```markdown
<!-- START_LOCKED -->
Stable role, input/output and JSON-format requirements.
<!-- END_LOCKED -->

<!-- START_EDITABLE -->
Klarspråk rules, terminology rules and tuning instructions.
<!-- END_EDITABLE -->
```

The locked sections should normally contain stable system behavior such as:

- the assistant role
- input format
- required output format
- JSON requirements
- field requirements

The editable section should contain policy decisions that may be tuned between runs, such as:

- klarspråk principles
- terminology preferences
- how aggressive rewrites should be
- when to prefer local edits
- when larger rewrites are justified

---

## Output JSON Format

The model should return a JSON array. Each object describes one suggested change.

Example:

```json
[
  {
    "type": "paragraph",
    "element_id": "paragraph_8",
    "old": "gammal text",
    "new": "ny text",
    "motivation": "Motivering till förändringen."
  },
  {
    "type": "footnote",
    "element_id": "footnote_3",
    "footnote_id": "4",
    "old": "Gammal text i fotnoten.",
    "new": "Ny text i fotnoten.",
    "motivation": "Anledning till ändrad text."
  }
]
```

The `old` value must match text in the relevant document element. The `new` value is the proposed replacement text.

---

## Suggestion Filtering and Diagnostics

The suggestion filter report is used to inspect which suggestions were accepted and which issues were detected.

Typical warning codes include:

| Code                       | Meaning                                                       |
| -------------------------- | ------------------------------------------------------------- |
| `weak_locality`          | The `old` span is long, which may indicate a broad rewrite. |
| `multi_sentence_rewrite` | The suggestion rewrites several sentences.                    |
| `spelling_degradation`   | The suggestion may introduce a spelling problem.              |
| `anchor_risk`            | The change may be difficult to locate safely in the document. |

Warnings do not necessarily mean that a suggestion is wrong. They are useful for tuning the prompt, filters and rendering pipeline.

---

## Features

- User-provided OpenAI API key
- Model selection for cost and quality control
- Policy-driven prompt file
- Optional session-specific prompt additions
- Structured JSON suggestion output
- Suggestion validation and filter reports
- DOCX editing with visible proposed changes
- Session logs for debugging and traceability
- Temporary upload and log handling

---

## Folder Structure Overview

```text
language-improver/
├── app/
│   ├── main.py
│   └── src/
│       ├── JBGLanguageImprover.py
│       ├── JBGLangImprovSuggestorAI.py
│       ├── JBGDocumentEditor.py
│       └── JBGDocumentStructureExtractor.py
├── policy/
│   └── prompt_policy.md
├── templates/
│   └── index.html
├── static/
│   ├── styles/
│   │   └── styles.css
│   └── javascript/
│       └── script.js
├── uploads/
├── logs/
└── README.md
```

---

## Running Locally with Uvicorn

1. Install dependencies:

```bash
pip install -r requirements.txt
```

2. Start the server:

```bash
uvicorn app.main:app --reload
```

3. Open the app:

```text
http://127.0.0.1:8000
```

From there you can upload a Word document, enter your OpenAI API key, choose a model and run the language improvement workflow.

---

## Recommended Prompt Tuning Workflow

When changing the prompt, change one thing at a time and compare:

1. the raw suggestions JSON
2. the suggestion filter report
3. the session log
4. the edited Word document

Useful questions when evaluating a run:

- Did the number of suggestions change?
- Did the model suggest more local or broader rewrites?
- Did important klarspråk improvements disappear?
- Did the model introduce unnecessary stylistic changes?
- Did the filter report flag more `weak_locality` or `multi_sentence_rewrite` warnings?
- Did any accepted suggestions fail during rendering?

This makes it easier to separate prompt effects from filtering and rendering issues.

---

## Contributing

The system is designed to be extensible. Useful areas for improvement include:

- stronger suggestion validation
- better Swedish spelling and morphology checks
- safer anchoring of long replacements
- better handling of tables and footnotes
- prompt variants for different document types
- evaluation scripts for comparing suggestion files between runs
