# Backlog – JBGLangImprover

## Known issues

- [ ] The comment module does add comments to footnotes changes, but they do not render properly in Word
- [ ] `CommentsRenderer` can only anchor comments to tracked revisions, not to a paragraph or a span without a change

## Planned improvments

### EPIC: Global (document-level) language improvement layer

All current proposals are local (one element, one `old` -> `new`). Add an optional
second layer that looks at the document as a whole. The layer is off by default,
runs after the local layer, and must never block or degrade the local result: if
the global step fails, the run still delivers the locally improved document.

#### Phase 0 - Foundations

- [x] G0.1 Capture token usage per API call (`prompt_tokens`, `completion_tokens`, reasoning tokens when present) and aggregate per phase (local/global) in a `UsageTracker`
- [x] G0.2 Extend `DocumentStructureExtractor` with `style_id`, `style_name`, `heading_level` (outline level) and `doc_order` (true body order incl. tables) without changing existing `element_id`s
- [x] G0.3 Add a `RunSummary` data object that collects model, usage, counts (raw/validated/accepted/applied/failed suggestions) and metrics, and save it as `<job>_run_summary.json`
- [x] G0.4 Regression tests: existing IDs unchanged, headings detected for both English (`Heading1`) and Swedish (`Rubrik1`) style IDs and for custom styles with `w:outlineLvl`

#### Phase 1 - LIX before/after (deterministic, no LLM)

- [x] G1.1 New module `JBGReadabilityMetrics.py`: LIX = words/sentences + 100 * long words (> 6 letters)/words, plus word, sentence and long-word counts
- [x] G1.2 Compute LIX "before" from the structure JSON only, and LIX "after" from structure JSON + accepted suggestions JSON by applying the anchored `old` -> `new` spans in memory (reuse `ChangePlanner` anchors, skip overlapping conflicts)
- [x] G1.3 Decide and document the text scope: body paragraphs, table cells and textboxes by default; headings, headers/footers and footnotes reported separately or excluded
  - Decision: headings, headers/footers and footnotes are excluded; "after" assumes all accepted suggestions are applied
- [ ] G1.6 Reduce prompt tokens: send only the fields the model needs (`element_id`, `type`, `text`, `heading_level`, `footnote_id`) instead of the full element dict
- [x] G1.4 Report LIX for the whole document and per top-level section (depends on G0.2), with the usual interpretation bands (very easy ... very difficult)
- [x] G1.5 Unit tests with hand-calculated Swedish reference texts

#### Phase 2 - GUI option and plumbing

- [ ] G2.1 Add an optional checkbox "Granska även dokumentet som helhet (global granskning)" with an information label and tooltip explaining what it does and that it costs more time and tokens
- [ ] G2.2 Add a checkbox "Lägg till avsnittet Om klarspråkningen" (default on) with tooltip
- [ ] G2.3 Send `global_review` and `include_about_section` via `FormData`; add them as `Form(False)`/`Form(True)` in `main.py`, pass them to `JBGLanguageImprover`, log them, and include both ids in `lockUI()`
- [ ] G2.4 Show progress messages for the global phase in the status polling

#### Phase 3 - Global analyzer (LLM)

- [ ] G3.1 New policy file `policy/global_prompt_policy.md` with locked/editable sections and a strict JSON output schema for findings: `category`, `severity`, `element_ids`, `quote`, `description`, `proposal`, optional `old`/`new`, optional `proposed_order`
- [ ] G3.2 New module `JBGGlobalAnalyzerAI.py` that sends a compact outline (headings, element ids, text) to the model; for large documents use map-reduce (per-section digests first, then a whole-document pass) within a token budget
- [ ] G3.3 Validate findings: all `element_ids` exist, `quote` matches the source text, categories are known, deduplicate against local suggestions
- [ ] G3.4 Findings that contain a safe single-element `old`/`new` are converted to `SuggestedChange` (tagged `origin=global`) and go through the existing filter/planner/renderer
- [ ] G3.5 Save `<job>_global_findings.json` next to the other diagnostic files

Finding categories, to be implemented in this order:

- [ ] G3.a Unnecessary repetitions across different parts of the document
- [ ] G3.b Inconsistencies (terminology, numbers, dates, names, abbreviations) and factual/logical errors inside the document
- [ ] G3.c Section and heading order: logical flow from a whole-document perspective, including heading wording and level
- [ ] G3.d Erroneous, irrelevant or unsupported (baseless) conclusions, with a reference to what is missing in the text
- [ ] G3.e Further checks (e.g. missing summary, undefined abbreviations at first use, promised content that never appears)

#### Phase 4 - Rendering of global findings

- [ ] G4.1 Extend `CommentsRenderer` with paragraph-level anchoring (`commentRangeStart/End` around a whole paragraph) so findings without a text change can be shown as Word comments
- [ ] G4.2 Never anchor global comments inside footnotes (see known issue); anchor to the referencing body paragraph instead
- [ ] G4.3 Simple markup mode: list global findings in the "Om klarspråkningen" section with references to headings
- [ ] G4.4 Section reordering is proposed only (comment + proposed order), never applied automatically

#### Phase 5 - "Om klarspråkningen" section

- [ ] G5.1 Append a new section with heading "Om klarspråkningen" at the end of the main body (before the final `w:sectPr`), preceded by a page break, using the document's own heading style (resolve via `styles.xml`, fallback to bold)
- [ ] G5.2 Content: date, markup mode, GPT model, number of API calls, tokens sent and received (local/global/total), LIX before and after (with the note that "after" assumes that all proposals are accepted), number of local proposals (validated/accepted/applied), number of global findings per category (only when global review was used), and whether the prompt was customized
- [ ] G5.3 Exclude the section from extraction, LIX and review if the output document is processed again (marker via bookmark or custom style)
- [ ] G5.4 Tests: section is added once, the document round-trips, tracked mode stays valid in Word

#### Phase 6 - Documentation and evaluation

- [ ] G6.1 Update README (options, new JSON files, output format, known limitations)
- [ ] G6.2 Evaluation set of 3-5 real reports: compare findings, false positives, cost and run time with and without the global layer

## Solved

- [x] Capture token usage per API call (`UsageTracker`) and report it in `<job>_run_summary.json`
- [x] Add heading/style metadata and true document order (`doc_order`) to the structure JSON without changing `element_id`s
- [x] Resolve and render headers/footers through their real OOXML relationship targets
- [x] Add regression coverage for shared section stories, variants, fields, markup and tracked changes
