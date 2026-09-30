# Backlog – JBGLangImprover

## Known issues

- [ ] The comment module does add comments to footnotes changes, but they do not render properly in Word
- [ ] Safety checks reject some legitimate rewrites, e.g. `2024-09-24` -> `den 24 september 2024` and `rättslig grund` -> `lagstöd` (`low_similarity`), and many multi-sentence rewrites (`old_looks_truncated`, `replacement_causes_obvious_duplication`); 43 of 134 raw suggestions were rejected in test run 93dbc428
- [ ] Span minimization can cut a proposal into word fragments that the safety checks then reject, so good corrections are lost. Test run ba144f3d: `bl.a.` -> `bland annat` arrived as `'.a.'` -> `'and annat'` (`old_looks_truncated`); the misspelling `diagranm` and `STs` -> `ST:s` were also rejected. Confirm with `<job>_suggestion_filter_report.json`
- [ ] `TrackedChangesRenderer` starts revision ids at 1 regardless of existing revisions in the document, so ids can collide with tracked changes already in the source
- [ ] Simple markup fails on some long rewrites in paragraphs with manual line breaks ("Invalid last_local_end" in `JBGSimpleMarkupRenderer`; test run b02ce006, paragraph_53 and paragraph_115). Tracked changes already handles anchors at line breaks; port the same fix and add a regression test
- [ ] Global findings vary considerably between runs of the same document (e.g. a within-section repetition found in two runs and missing in the third); consider a lower temperature for the global call where the model allows it
- [ ] Check whether template notes such as "Ta ej bort denna avsnittsbrytning!!" (test run b391e3ac, paragraph_127, flagged as "Troligt fel") are hidden or white text in the template. If so, the extractor should skip hidden text (`w:vanish`) so it is neither reviewed nor counted in LIX
- [ ] Fixed 5 s pause between API calls (about 50 s of a 5.5 min run with 11 calls); replace with retry/backoff on rate-limit errors

## Planned improvments

### GUI

- [ ] Roll back the unlocking of the form after a run: keep all controls locked when a run finishes or fails, so a new run requires reloading the page (consistent with the footer hint about Ctrl-Shift-R). Remove the `unlockUI()` calls in `script.js` (submit error, status error, download `finally`); `lockUI()` can keep its `lockedByUI` bookkeeping or return to the original version

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

- [x] G2.1 Add an optional checkbox "Granska även dokumentet som helhet (global granskning)" with an information label and tooltip explaining what it does and that it costs more time and tokens
  - Implemented as "Granska dokumentets struktur och upplägg" under the heading "Globala inställningar", unchecked by default; enabled in Phase 3 (first increment: repetitions)
- [x] G2.2 Add a checkbox "Lägg till avsnittet Om klarspråkningen" (default on) with tooltip
- [x] G2.3 Send `global_review` and `include_about_section` via `FormData`; add them as `Form(False)`/`Form(True)` in `main.py`, pass them to `JBGLanguageImprover`, log them, and include both ids in `lockUI()`
  - `global_review` runs the global analyzer (Phase 3)
- [x] G2.4 Show progress messages for the global phase in the status polling
- [x] G2.5 Make LIX optional: checkbox "Beräkna LIX-värde före och efter granskning" under "Övrigt" (default on), sent as `compute_lix`; when on, LIX is saved in the run summary even without the "Om klarspråkningen" section
- [x] G2.6 GUI layout: compact vertical spacing, left-aligned radio buttons and checkboxes, section headings "Globala inställningar" and "Övrigt", prompt label "Anpassa den lokala promptinstruktionen"

#### Phase 3 - Global analyzer (LLM)

- [x] G3.1 New policy file `policy/global_prompt_policy.md` with locked/editable sections and a strict JSON output schema for findings: `category`, `severity`, `element_ids`, `quote`, `description`, `proposal`, optional `old`/`new`, optional `proposed_order`
- [x] G3.2 New module `JBGGlobalAnalyzerAI.py` that sends a compact outline (headings, element ids, text) to the model; for large documents use map-reduce (per-section digests first, then a whole-document pass) within a token budget
  - First version: one call for documents up to about 60 000 tokens; longer documents are split at top-level headings and analyzed part by part
- [ ] G3.2b Map-reduce for very long documents, so repetitions across parts are found (per-part digests, then a whole-document pass)
- [x] G3.3 Validate findings: all `element_ids` exist, `quote` matches the source text, categories are known, deduplicate against local suggestions
  - Done except deduplication against local suggestions (only relevant once findings carry `old`/`new`, see G3.4)
- [ ] G3.4 Findings that contain a safe single-element `old`/`new` are converted to `SuggestedChange` (tagged `origin=global`) and go through the existing filter/planner/renderer
- [x] G3.5 Save `<job>_global_findings.json` next to the other diagnostic files

Finding categories, to be implemented in this order:

- [x] G3.a Unnecessary repetitions across different parts of the document
  - Tuning 1 (test run 6ea91865): all 6 accepted findings were repetitions between levels of detail (summary, conclusions, chapter overview, detailed results), which is intentional in layered reports. The policy now flags only repetitions at the same level, in both directions. Quotes with "…" are matched fragment by fragment (3 genuine findings had been rejected)
  - Comments start with the category label as a sentence ("Onödig upprepning. …"; next category e.g. "Inkonsekvent påstående. …") and use their own author, "JBG Klarspråkningstjänst (global granskning)", so they can be shown or hidden separately in Word
  - Evaluation (test run b02ce006, with tuning 1 active): 9 findings, 7 of them between different chapters (e.g. summary vs conclusions), i.e. the model does not follow the "same level only" rule in the policy. The findings were judged useful as they are, so no filter is applied in code. If the noise becomes a problem, the deterministic option is to accept a repetition only when all its places are in the same top-level chapter (rejected findings would still be saved with a reason)
- [x] G3.b Inconsistencies (terminology, numbers, dates, names, abbreviations) and factual/logical errors inside the document
  - Category `inconsistency`, comment label "Inkonsekvent påstående. …", same global call as repetitions. The conflicting statement must be quoted in `related_quote` and is verified against the text; the comment shows it ("Jämför med: avsnittet ”X” (”…”)")
  - The policy excludes rounding of the same value, values for different periods/groups/methods, compatible intervals, stylistic variation, and chapter-number references (automatic numbering is not in the text)
  - Factual errors that need knowledge outside the document are not covered
  - Category `error`, comment label "Troligt fel. …": obvious errors in a single place (leftover draft text, placeholders, broken sentences, calculation errors within a paragraph). One place is enough (`MIN_LOCATIONS`); spelling, grammar and wording stay with the local review. Added after test run 7c5dfd4d, where leftover text in a text box ("ha diagranm%") was filed as an inconsistency
- [x] G3.c Section and heading order: logical flow from a whole-document perspective, including heading wording and level
  - Two categories, both comment-only and anchored on a heading (checked in code, `anchor_not_a_heading`): `disposition` ("Förslag om disposition. …": order, content under the wrong chapter, heading level vs content; optional `proposed_order`, verified against the document's real headings) and `heading` ("Förslag om rubrik. …": does the heading describe the section below it; optional example wording in the proposal)
  - Respectful tone and protected standard sections (förord, sammanfattning, inledning, källor, bilagor) are set in the policy; at most 5 findings per category, enforced in code (`MAX_PER_CATEGORY`)
  - A broken transition before a heading is only used as support for a disposition finding, never as a finding of its own
  - Disposition findings must sit on a body section heading (`anchor_not_a_section_heading`), and `proposed_order` may only name such headings. Test run b391e3ac proposed moving the template fact box "IAF:s tillsyn" (a table before the Förord); headings in fact boxes, tables and text boxes are not sections. Heading suggestions may still sit on fact-box headings
  - Whether a claim in a message heading is supported by the text belongs to G3.d
- [ ] G3.d Erroneous, irrelevant or unsupported (baseless) conclusions, with a reference to what is missing in the text
- [ ] G3.e Further checks (e.g. missing summary, undefined abbreviations at first use, promised content that never appears)

#### Phase 4 - Rendering of global findings

- [x] G4.1 Extend `CommentsRenderer` with paragraph-level anchoring (`commentRangeStart/End` around a whole paragraph) so findings without a text change can be shown as Word comments
- [x] G4.2 Never anchor global comments inside footnotes (see known issue); anchor to the referencing body paragraph instead
  - Findings are anchored at body paragraphs and table cells; text boxes are anchored at their host paragraph; footnotes, headers and footers are not part of the global review
- [ ] G4.3 Simple markup mode: list global findings in the "Om klarspråkningen" section with references to headings
- [ ] G4.4 Section reordering is proposed only (comment + proposed order), never applied automatically

#### Phase 5 - "Om klarspråkningen" section

- [x] G5.1 Append a new section with heading "Om klarspråkningen" at the end of the main body (before the final `w:sectPr`), preceded by a page break, using the document's own heading style (resolve via `styles.xml`, fallback to bold)
- [x] G5.2 Content: date, markup mode, GPT model, number of API calls, tokens sent and received (local/global/total), LIX before and after (with the note that "after" assumes that all proposals are accepted), number of local proposals (validated/accepted/applied), number of global findings per category (only when global review was used), and whether the prompt was customized
- [x] G5.3 Exclude the section from extraction, LIX and review if the output document is processed again (marker via bookmark or custom style)
- [x] G5.4 Tests: section is added once, the document round-trips, tracked mode stays valid in Word
  - Decisions: placed last (after the back cover), inserted as one tracked insertion in tracked mode, unnumbered heading style resolved from the document (Källor/Bilaga heading first), hidden bookmark `_JBG_OmKlarsprakningen`, replaced on a new run
  - The same three rules apply to every document, regardless of template: (1) last in the body, on a new page; (2) one tracked insertion in tracked mode; (3) unnumbered heading, enforced with a direct `numId 0` override even when the style or `numbering.xml` numbers it
  - Regression matrix: IAF-like, plain, numbered heading via style, numbered heading via `numbering.xml`, English references, ends with table, ends with content control, ends with section break, empty document, last paragraph without formatting
  - The global findings line is added in Phase 3
- [ ] G5.5 Verify in Word with a real IAF report: heading style, page break after the back cover, and header/footer/background of the new page
- [ ] G5.6 Known limitation: the section inherits the page layout of the document's last Word section (columns, orientation, headers/footers, background). If this is a problem in practice, insert a real section break with a clean single-column layout instead of a page break

#### Phase 6 - Documentation and evaluation

- [ ] G6.1 Update README (options, new JSON files, output format, known limitations)
- [ ] G6.2 Evaluation set of 3-5 real reports: compare findings, false positives, cost and run time with and without the global layer

## Solved

- [x] The local review ran twice when the model returned zero suggestions (`save_as_json` re-ran the whole review), doubling time and token cost
- [x] GUI: the form is unlocked when a run finishes or fails, so a new run does not require reloading the page (controls that were already disabled stay disabled) — to be reverted, see Planned improvments / GUI
- [x] `CommentsRenderer` can anchor a comment to a whole paragraph (`add_paragraph_comment`), not only to tracked revisions (G4.1)
- [x] "Om klarspråkningen": cached prompt tokens are shown ("varav N återanvända (lägre kostnad)") so the cost is not overestimated
- [x] LIX: manual line breaks inside a sentence are no longer counted as sentence boundaries
- [x] LIX: sections start only at body headings; headings in tables/textboxes, captions and empty chapters no longer create sections
- [x] LIX: `lix_delta` is computed from the rounded values so it matches the displayed before/after values
- [x] Capture token usage per API call (`UsageTracker`) and report it in `<job>_run_summary.json`
- [x] Add heading/style metadata and true document order (`doc_order`) to the structure JSON without changing `element_id`s
- [x] Resolve and render headers/footers through their real OOXML relationship targets
- [x] Add regression coverage for shared section stories, variants, fields, markup and tracked changes
