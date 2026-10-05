# JBG Language Improvement System

A web service for **klarspråksgranskning av Word-dokument (`.docx`)** with OpenAI models. It reads the document's text and structure, asks a language model for policy-driven improvements, checks every proposal against the text, and returns a new Word file with the proposals as tracked changes (or simple colour markup). Optionally, it also reviews the document as a whole, computes readability (LIX) and adds a section describing the review.

The service is built for Swedish plain-language review of formal public-sector texts, especially reports. The original file is never changed; the user always gets a new file back.

---

## How a run works

1. **Extraction.** The document is read into a structure of elements: body paragraphs, headings (with level), table cells, text boxes, headers, footers and footnotes, each with a stable `element_id` and its place in the reading order. The text is read the way Word shows it with markup: other authors' pending insertions are included, pending deletions are not.
   Not reviewed: paragraphs whose text is entirely hidden (e.g. template notes), and automatically generated content (tables of contents, lists of figures, indexes, bibliographies), which is replaced by a placeholder. A previously added "Om klarspråkningen" section is also skipped.
2. **Local review.** The text is sent to the model in parts (at most about 8,000 tokens per call), up to four calls at a time. The model returns proposals (`old` → `new` with a motivation) for single elements.
3. **Global review** *(optional)*. At the same time, the model reads the whole document once and reports document-level findings, such as unnecessary repetitions or conclusions that need support. These become Word comments; the text is never changed by them.
4. **Checks.** Every proposal and finding is checked against the actual text before it is used (see [Checks](#checks)).
5. **Readability** *(optional)*. LIX before and after is computed without any model call.
6. **Rendering.** Accepted proposals are written into the actual parts of the Word file (body, tables, text boxes, headers, footers, footnotes) as tracked changes or simple markup. Comments, the "Om klarspråkningen" section and the LIX comment are added last.

A failure in an optional step (global review, LIX, the section) never stops the run; the locally improved document is still delivered.

---

## Using the service

The form has five panels:

| Panel | Options | Default |
|---|---|---|
| **Ladda upp din text** | the `.docx` file | – |
| **Inställningar för språkmodellen** | API key, model, the editable part of the local prompt ("Anpassa den lokala promptinstruktionen") | – |
| **Dokumentinställningar** | "Granska dokumentets struktur och upplägg" (global review), "Beräkna LIX-värde före och efter språkgranskning" | both off |
| **Hur ska resultatet se ut?** | "Enkel färgmarkering" (deleted text struck through in red, new text in green, no comments) or "Spåra ändringar och infogade kommentarer" (tracked changes with a comment motivating each proposal) | tracked changes |
| **Övrigt** | "Lägg till ”Om klarspråkningen” sist" | off |

The editable prompt only affects the local review; the global review has its own fixed policy. When the service runs behind Azure App Service authentication, the page shows "Inloggad som: …" (otherwise "okänd användare"); the name is for display only and is never used for authorization.

The form stays locked after a run; reload the page (Ctrl-Shift-R) to start a new one.

---

## Local review

### Prompt strategy

The prompt is designed for Swedish **klarspråk**: language that is clear, simple, correct and appropriate for the intended audience. It gives priority to **klarspråksnytta** over strict edit locality: larger rewrites are allowed when they clearly improve comprehension, but unnecessary rewrites of clear and correct text should be avoided.

The prompt tells the model to:

- adapt the text to the intended reader and structure information logically
- simplify sentence structure and word choice, prefer active voice and verbs over nominalizations
- follow Swedish writing rules and public-sector conventions
- explain specialist terms and write out abbreviations when useful
- use clear, meaningful headings without a final full stop
- keep the tone professional, clear and accessible

Changes should only be suggested when they clearly improve comprehensibility, clarity or correctness. Larger rewrites must preserve the information, add no new information or interpretations, and not add unnecessary paragraph breaks.

Terminology rules currently include: prefer `a-kassor`/`a-kassorna` over `arbetslöshetskassor`/`arbetslöshetskassorna`, and never rewrite `arbetslöshetsförsäkringen`.

### The prompt policy file

`policy/prompt_policy.md` is split by comment markers into locked and editable parts:

```markdown
<!-- START_LOCKED -->
Stable role, input/output and JSON-format requirements.
<!-- END_LOCKED -->

<!-- START_EDITABLE -->
Klarspråk rules, terminology rules and tuning instructions.
<!-- END_EDITABLE -->
```

The locked parts hold the stable contract (role, input format, output format); the editable part holds the klarspråk rules and is shown in the form, where the user can adjust it for a single run. The run summary records whether the prompt was customized.

### Input and output

For each element the model receives only `type`, `element_id` and `text`, plus `footnote_id` for footnotes and `heading_level` for headings. Empty elements are not sent. The model returns a JSON array:

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

`old` must match text in the element; `new` is the proposed replacement.

---

## Global review

Enabled by "Granska dokumentets struktur och upplägg". The model reads a compact outline of the whole document (headings with level, element ids and text, in reading order) and returns findings in six categories. Each finding becomes one Word comment, by the separate author "JBG Klarspråkningstjänst (global granskning)", so the comments can be shown or hidden on their own in Word.

| Category | Comment starts with | Places | At most |
|---|---|---|---|
| `repetition` | Onödig upprepning. | 2 or more | 10 |
| `inconsistency` | Inkonsekvent påstående. | 2 or more, both quoted | 10 |
| `error` | Troligt fel. | 1 or more | 10 |
| `disposition` | Förslag om disposition. | 1 or more, on a section heading | 5 |
| `heading` | Förslag om rubrik. | 1 or more, on a heading | 5 |
| `conclusion` | Slutsats som behöver stöd. | 1 or more | 5 |

The policy is in `policy/global_prompt_policy.md` (locked output format, editable rules per category). Only the policy shapes what the model looks for; the rules below are enforced in code, whatever the model returns:

- Every element id must exist, and every quote must be found in the text (quotes with "…" are matched part by part). For inconsistencies, the conflicting statement must also be quoted and found.
- Comments on headings must sit on a real heading; disposition findings must sit on a section heading in the body text.
- **Template conventions:** nothing in the cover area (everything before the first section heading) gets a comment, and disposition is only suggested within a chapter, never as a reordering of top-level chapters.
- Placeholders for generated content (e.g. a table of contents) can never be part of a finding.

Rejected findings are saved with a reason in `<file>_global_findings.json`. A finding is anchored as a comment around the whole paragraph; findings in text boxes are anchored at the paragraph holding the text box, and footnotes are never used as anchors. Very long documents (over about 60,000 tokens) are reviewed in parts split at top-level headings.

---

## Checks

Before a local proposal is used, the service:

- trims it to the smallest changed span, without ever cutting inside a word (dots in abbreviations such as "bl.a." count as part of the word)
- rejects spans that start or end inside a word in the actual text, new text glued onto a neighbouring word, and obviously garbled text (valid forms such as "IAF:s", "t.ex." and web addresses are recognized)
- rejects replacements too different from the original (`low_similarity`), with an exception for pure comma and semicolon edits
- ignores changes that only add whitespace

Rejected proposals are logged with their reason. Accepted proposals then pass a quality filter that mainly adds warnings to `<file>_suggestion_filter_report.json`:

| Code | Meaning |
|---|---|
| `weak_locality` | The `old` span is long, which suggests a broad rewrite. |
| `multi_sentence_rewrite` | The proposal rewrites several sentences. |
| `oversized_rewrite` | The new text is much longer than the original. |
| `stylistic_rewrite` | A stylistic word swap rather than a correction. |
| `functional_shift` | The proposal seems to change what the text says or does, not just its language. |
| `internal_note_rewrite` | The proposal rewrites an internal note or TODO marker. |
| `quote_style_shift` | Typographic quotation marks replaced by straight ones. |
| `sensitive_element_growth` | The text grows noticeably in a sensitive element (header, footer, footnote, table cell, text box). |

Warnings do not mean a proposal is wrong; they help when tuning the prompt and the checks.

---

## Readability (LIX)

LIX before and after is computed deterministically from the structure JSON and the suggestions JSON:

```bash
python -m app.src.JBGReadabilityMetrics <file>_structure.json [<file>_suggestions.json]
```

LIX = words/sentences + 100 × long words (more than six letters)/words. Body paragraphs, table cells and text boxes are counted; headings, headers, footers, footnotes, hidden text and generated content are not. The value is reported for the whole document and per top-level chapter. The "after" value assumes that every accepted proposal is applied. LIX is an indicator, not a quality score: a clearer rewrite can raise LIX when short words disappear.

---

## "Om klarspråkningen"

When enabled, a section is added last in the document, on a new page:

- run facts: date, model, result mode, whether the prompt was customized
- tables: model calls and tokens per phase (including tokens re-used from cache), local proposals, global findings per category, and LIX for the whole document and each chapter

The section follows the same rules in every document: it is placed after all existing content, has an unnumbered heading in a heading style taken from the document (numbering is switched off if the style is numbered; bold text if the document has no heading styles), and in tracked mode it is one tracked insertion, so rejecting it restores the original document exactly. A hidden bookmark marks it, so a new run skips and replaces it instead of adding a second one.

When the section is off but LIX is on, LIX is added as a comment on the document's title instead (the first paragraph in the Title style, otherwise the first body paragraph with text).

---

## Output files

Each run writes, next to the uploaded file (cleaned up automatically):

| File | Content |
|---|---|
| `<file>_structure.json` | the extracted elements, plus `excluded` (hidden and generated paragraphs) |
| `<file>_suggestions.json` | the accepted local proposals |
| `<file>_suggestion_filter_report.json` | quality warnings per proposal |
| `<file>_global_findings.json` | global findings, accepted and rejected (with reasons) |
| `<file>_run_summary.json` | settings, token usage per phase, proposal counts, LIX, section and comment results |
| `logs/<job_id>.log` | the session log, including every rejected proposal and its reason |

---

## Configuration

Settings in `.env` (see `.env_template`):

| Variable | Purpose |
|---|---|
| `FRAME_ANCESTORS` | Allowed parent pages when the service is embedded in an iframe |
| `APP_TITLE` | Title shown in the page |
| `JBG_MAX_PARALLEL_MODEL_CALLS` | Simultaneous local model calls (default 4; 1 = in sequence) |

Model calls are retried automatically on rate limits, server errors and timeouts (4 retries, handled by the OpenAI library). The logged-in user name comes from the `X-MS-CLIENT-PRINCIPAL-NAME` header set by Azure App Service authentication.

Endpoints: `/` (the page), `/config`, `/me`, `/get_editable_prompt/`, `/upload/`, `/status/{job_id}`, `/download/{job_id}`, `/healthz`.

---

## Folder structure

```text
JBGLangImprover/
├── app/
│   ├── main.py                            web app and endpoints
│   └── src/
│       ├── JBGLanguageImprover.py         orchestrates a run
│       ├── JBGDocumentStructureExtractor.py
│       ├── JBGContentClassifier.py        hidden text and generated content
│       ├── JBGLangImprovSuggestorAI.py    local review and checks
│       ├── JBGGlobalAnalyzerAI.py         global review and its validation
│       ├── JBGModelClient.py              OpenAI client, retries, parallel limit
│       ├── JBGUsageTracker.py             token counting
│       ├── JBGChangePlanner.py, JBGTokenDiffEngine.py
│       ├── JBGTrackedChangesRenderer.py, JBGSimpleMarkupRenderer.py
│       ├── JBGCommentsRenderer.py, JBGGlobalFindingsRenderer.py
│       ├── JBGAboutSectionRenderer.py     "Om klarspråkningen"
│       ├── JBGReadabilityMetrics.py, JBGReadabilityCommentRenderer.py
│       ├── JBGRunSummary.py, JBGRevisionIds.py
│       └── JBGDocxPackage.py, JBGDocumentPartAdapter.py,
│           JBGHeaderFooterPartAdapter.py, JBGFootnotesPartAdapter.py
├── policy/
│   ├── prompt_policy.md                   local review
│   └── global_prompt_policy.md            global review
├── templates/index.html
├── static/ (styles, javascript)
├── tests/
├── BACKLOG.md
└── README.md
```

---

## Running locally

```bash
pip install -r requirements.txt
uvicorn app.main:app --reload
```

Then open `http://127.0.0.1:8000`, upload a Word document, enter an OpenAI API key and choose a model.

## Tests

The suite (170 tests) builds temporary Word documents and uses a simulated model, so no API key is needed. It covers extraction (including hidden text, generated content and other authors' tracked changes), the checks on local proposals, the global review and its validation rules, rendering in both modes (body, tables, text boxes, headers, footers, footnotes, comments), the "Om klarspråkningen" section across many document shapes (including a simulated "reject all"), LIX, parallel calls, the form's defaults and the `/me` endpoint. Run it from the repository root:

```bash
python -m unittest discover -s tests -v
```

Tests that need FastAPI are skipped automatically where it is not installed.

---

## Tuning the prompts

Change one thing at a time and compare the output files of two runs on the same document: the suggestions JSON, the filter report, the global findings (including the rejected ones and their reasons), the log, and the Word document itself. Useful questions:

- Did the number or kind of proposals change? Did important improvements disappear?
- Did the model suggest broader rewrites or more stylistic changes?
- Did more proposals get rejected, and for which reasons?
- For the global review: are the findings useful, respectfully worded and in the right category?

Model output varies between runs of the same document, so judge a change over several runs rather than one.

---

## Known limitations and plans

Open issues and planned improvements are tracked in `BACKLOG.md`, for example text inside content controls (not yet read), proposals rejected as too different from the original (dates, synonyms), and a possible simplification of the result options.
