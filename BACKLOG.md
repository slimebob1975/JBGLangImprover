# Backlog – JBGLangImprover

## Known issues

- [ ] The comment module does add comments to footnotes changes, but they do not render properly in Word
- [ ] A proposal that drops a symbol from a legal reference can pass the character-similarity check when the rest looks alike, e.g. `c §` -> ` c ` (found while testing; pre-existing, not caused by the equivalence exceptions). Consider rejecting proposals that remove `§` or change digits in references
- [ ] Simple markup fails on some long rewrites in paragraphs with manual line breaks ("Invalid last_local_end" in `JBGSimpleMarkupRenderer`; test run b02ce006, paragraph_53 and paragraph_115). Tracked changes already handles anchors at line breaks; port the same fix and add a regression test
- [ ] Cleanup: `upload_file_old` in `app/main.py` is no longer routed and can be removed
- [ ] Global findings vary considerably between runs of the same document (e.g. a within-section repetition found in two runs and missing in the third); consider a lower temperature for the global call where the model allows it
- [ ] The extractor does not read text inside body-level content controls (`w:sdt`), because it only reads paragraphs and tables directly in the body, so such text is neither reviewed locally nor globally. Generated content in content controls (tables of contents) is already handled as a placeholder (G3.f). Before fixing: check whether the templates put real text (cover fields, standard text) in content controls. The fix needs new element ids (e.g. `sdt_2_p1`) so existing ids do not shift, and support in both renderers and the comment anchoring. Together with this: recognize Word's own template markers, generic and reliable when present — locked content controls (`w:lock`), document protection with editable ranges (`w:permStart`/`w:permEnd`, where everything outside is standard content) and placeholder text (`w:showingPlcHdr`) — and keep such content out of the global (and possibly local) review

## Planned improvments

### GUI

- [x] Tracked changes ("Spåra ändringar och infogade kommentarer") is the default result mode in the form and in `/upload/`; the form defaults are pinned by `tests/test_gui_defaults.py`
- [ ] In the long run, possibly make "Spåra ändringar och infogade kommentarer" the only result mode (it is the default since this patch). The GUI could then be simplified from the two radio buttons to two checkboxes: "Spåra ändringar" and "Infogade kommentarer" (the latter maps to the existing `include_motivations` option, today hidden and always on). Before removing simple markup: check whether anyone depends on it, and decide what the comments checkbox means without tracked changes (simple markup has no comments today)
- [x] Show "Inloggad som: …" at the top of the page, like the sister service: `/me` returns the name from Azure App Service authentication (header `X-MS-CLIENT-PRINCIPAL-NAME`, URL-decoded, max 200 characters) and `script.js` shows it under the subtitle. Without a login (e.g. locally) the line reads "Inloggad som: okänd användare", like the sister service; if the call fails, the line stays hidden. Display only: the name is never used for authorization

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
- [x] G1.6 Reduce prompt tokens: the local review sends only `type`, `element_id` and `text`, plus `footnote_id` for footnotes and `heading_level` for headings (both described in the policy's locked input section); empty elements are not sent, and a part with only empty elements makes no call. The split into calls is unchanged (still measured on the full element data), so each call contains the same text as before. Measured on test run de5618be: document payload 176,603 -> 45,799 characters (-74%), whole local prompt -60%, estimated about 62,000 -> 25,000 prompt tokens per run. Fewer, larger calls could be tried later as a separate experiment
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
  - Moved to "Globala inställningar" under the structure checkbox, unchecked by default. "Lägg till ”Om klarspråkningen” sist" is also unchecked by default (GUI and `main.py` defaults)
  - Without the section, LIX is added as a comment on the document title: the first paragraph in the Title style, otherwise the first body paragraph with text (fact boxes and text boxes are skipped)
- [x] G2.6 GUI layout: compact vertical spacing, left-aligned radio buttons and checkboxes, section headings "Globala inställningar" and "Övrigt", prompt label "Anpassa den lokala promptinstruktionen"
  - The form is split into five grey panels (`fieldset` + `legend`, class `settings-panel`, same look as the sister service): "Ladda upp din text", "Inställningar för språkmodellen", "Dokumentinställningar" (renamed from "Globala inställningar"), "Hur ska resultatet se ut?" and "Övrigt". The form's own white card is removed
  - Tooltips reviewed against the service: file (new file, original unchanged), model (used for both reviews), local prompt (includes text boxes, does not affect the document review), structure review (no "kommer senare", text is never changed), LIX ("after" assumes all proposals accepted), simple markup (struck through in red, no comments), tracked changes (with motivating comments), "Om klarspråkningen" (global findings, LIX only if computed)
  - The title, subtitle and "Inloggad som" line are aligned with the left edge of the panels (same width and centering as the form) at any screen width, with the same vertical spacing between them as the sister service (about 18, 16 and 14 px)
  - LIX checkbox label changed to "Beräkna LIX-värde före och efter språkgranskning" for clarity

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
  - Template conventions, without relying on template files, numbering or style names (decided after test runs 2991c745 and cc61bb1e): (A) the cover area — everything before the first section heading in reading order (title, cover boxes, colophon) — never gets a global comment (`cover_material`), and cover elements are dropped from a finding's related places; (B) disposition is only suggested within a chapter: findings anchored on a top-level heading are rejected (`top_level_order`), and `proposed_order` is kept only if all its headings are subsections in the anchor's chapter. Both are in the policy too. Boilerplate in the body without Word markers (e.g. "Vi är IAF") cannot be recognized from a single document
  - Verified in Word (test run de5618be): a heading with two comments (a heading suggestion plus the paragraph below it with a repetition and a conclusion comment) is still easy to read, so no limit per paragraph is needed
- [x] G3.d Erroneous, irrelevant or unsupported (baseless) conclusions, with a reference to what is missing in the text
  - Category `conclusion`, comment label "Slutsats som behöver stöd. …": conclusions that claim more than the results show (sample to all, cause from co-variation, "visar" vs "tyder på"), conclusions without any result behind them, and conclusions that do not answer the stated purpose. A conclusion that contradicts a result stays an inconsistency
  - Comment on the conclusion itself (paragraph, fact box or message heading; one place is enough). The supporting result may be cited in `related_quote`, verified like other quotes ("Jämför med: …"); an unverifiable quote is dropped, not the finding
  - Only support within the document is judged; conclusions backed by a cited source, explicit assessments with stated reasoning, and recommendations are excluded. Respectful wording, at most 5 findings (`MAX_PER_CATEGORY`)
- [x] G3.f Do not comment on automatically generated content: tables of contents, lists of figures or tables, indexes and bibliographies are recognized in three forms (a content control of the gallery type "Table of Contents"/"Bibliographies", a `TOC`/`INDEX`/`BIBLIOGRAPHY` field spanning several paragraphs, or paragraphs in Word's `toc 1-9`/`table of figures`/`index` styles). Each block becomes one placeholder element of type `generated` (e.g. "[Innehållsförteckning, skapas automatiskt]") in its place in the reading order. The global review sees it, so the heading above is not seen as empty, but it can never be part of a finding (`generated_content`); the local review never sees it and LIX does not count it. The policy says not to comment on generated content or on headings that only introduce it. Does not depend on reading the text inside content controls
- [ ] G3.g Irrelevant text: a new global category `irrelevant` for text under a section or heading that is out of context, i.e. does not fit the section it is in and does not serve the document's purpose (e.g. a digression, a leftover passage from another report, details unrelated to the section's topic). Reported as "Irrelevanta textavsnitt" in the "Om klarspråkningen" table; the comment label could be "Irrelevant textavsnitt. …" (or a softer wording such as "Textavsnitt som kanske inte hör hit. …" — to decide). Design points:
  - Boundaries to existing categories: `heading` covers a heading that misdescribes its section as a whole; `disposition` covers content that belongs elsewhere in the document. `irrelevant` is for a passage that does not fit its section and has no better place in the document
  - Anchored on the passage itself (body paragraph, table cell or text box); one place is enough, and the section heading may be given as a related place
  - Same rules as the other categories: respectful wording, quote verified against the text, no comments on the cover area or on generated content, at most 5 findings (`MAX_PER_CATEGORY`)
  - Risk of false positives is high (background, definitions and examples can look off-topic), so the policy should only flag clear cases and the category should be evaluated on several reports before relying on it
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
- [x] G5.7 Numbers in tables: "Anrop och tokens" (one column per phase plus total), "Förslag", "Iakttagelser från den globala granskningen" (per category) and "Läsbarhet (LIX)" (whole document and each chapter: before, after, change). Run facts stay as short lines. Tables are tracked row by row, so rejecting the section still restores the original; the extractor skips the section's tables as well as its paragraphs on a rerun
  - Verified in Word (test run de5618be): rejecting the tracked section removes its tables along with the text
- [ ] G5.5 Verify in Word with a real IAF report: heading style, page break after the back cover, and header/footer/background of the new page
- [ ] G5.6 Known limitation: the section inherits the page layout of the document's last Word section (columns, orientation, headers/footers, background). If this is a problem in practice, insert a real section break with a clean single-column layout instead of a page break

#### Phase 6 - Documentation and evaluation

- [x] G6.1 Update README (options, new JSON files, output format, known limitations): rewritten to describe the service as it is now — the five form panels and defaults, the local and global review, the checks with the actual warning codes (the old table listed an `anchor_risk` code that does not exist), LIX, "Om klarspråkningen", output files, configuration, endpoints and tests
- [ ] G6.2 Evaluation set of 3-5 real reports: compare findings, false positives, cost and run time with and without the global layer

## Solved

- [x] Legitimate rewrites rejected as `low_similarity` (test runs 93dbc428, 2991c745). The similarity rules compare characters, not meaning, so two narrow exceptions now apply to both of them (general and footnote/text box): (1) the same date in another form, e.g. `2024-09-24` -> `den 24 september 2024`, with nothing else changed; (2) short word swaps of 1–3 letter-only words with a reasonable length ratio, e.g. `rättslig grund` -> `lagstöd`, `skedde` -> `gjordes`, unless the swap adds or removes a negation (inte, ej, aldrig, ingen …). Re-validating all 30 `low_similarity` rejections from 12 test runs: 3 good proposals recovered (including a comma removal from the previous fix); the 8 still rejected are garbled minus signs, word deletions and colon edits
- [x] The garbled-text check rejected correct Swedish (test run 2991c745): new texts containing "IAF:s" were rejected as `corrupted_new_text` because of the letter–punctuation–letter pattern. Valid forms are now set aside before the check: genitive after an abbreviation (IAF:s, ST:s), dotted abbreviations followed by a space or punctuation (t.ex., bl.a., fr.o.m.), and web and e-mail addresses. The vowel guess (e.g. "skrift" flagged) is removed. Garbled text such as "kontrollernaVisar", "ut.Nästa" or "re,sultat" is still rejected
- [x] Removing a comma (`,` -> nothing) or changing it to a semicolon was rejected as `low_similarity`; pure punctuation edits on commas and semicolons now pass the similarity check. Other one-character replacements (e.g. a minus sign replaced by a space) are still rejected
- [x] Good local proposals were lost by span trimming and the truncation check (test runs ba144f3d and 2cfb505b). Trimming now never cuts inside a word on either the `old` or the `new` side (dots inside abbreviations count as part of the word), and a fragment is detected by where the span lies in the element text instead of by guessing from word length, vowels or punctuation. Now accepted: `bl.a.` -> `bland annat`, `t.ex.` -> `till exempel`, `ha diagranm%` -> `har diagram`, `IAF` -> `Inspektionen för arbetslöshetsförsäkringen (IAF)`, and comma insertions after short words such as `dock` -> `dock,` (all previously rejected). Proposals that only add whitespace are ignored. Real fragments (e.g. `.a.` or `ex.` inside an abbreviation, new text glued onto a neighbouring word) are still rejected
- [x] Model calls run in parallel: the local review sends at most 4 calls at a time (`JBG_MAX_PARALLEL_MODEL_CALLS`, default 4, 1 = in sequence), and the global review starts right after extraction and runs alongside them. Only the calls run in threads; answers are parsed and validated in chunk order, so proposals come out in the same order as before. Progress shows finished calls ("Klar med 3 av 7 anrop."). A failing call does not stop the others. In test run 2cfb505b the 7 local calls took 4 min 31 s in sequence
- [x] Local proposals were lost in paragraphs with other authors' pending tracked insertions ("Anchor text mismatch"): body and table-cell paragraphs now use the same text model as the renderers (text in `w:ins` included, `w:del` excluded). Existing comments and tracked changes from other authors are kept in both modes
- [x] Fully hidden paragraphs (e.g. the template note "Ta ej bort denna avsnittsbrytning!!") are left out of the review and LIX; hidden is resolved as in Word (direct formatting, character style, paragraph style, document defaults). Ids of other paragraphs are unchanged; partly hidden paragraphs are kept as they are. Applies to body paragraphs, table cells and text boxes. Excluded ids are listed in the structure JSON under `excluded`
- [x] The local review ran twice when the model returned zero suggestions (`save_as_json` re-ran the whole review), doubling time and token cost
- [x] GUI: the form is unlocked when a run finishes or fails (later reverted, see below)
- [x] GUI: the form stays locked after a run finishes or fails; a new run requires reloading the page (consistent with the footer hint about Ctrl-Shift-R)
- [x] `TrackedChangesRenderer` numbers new revisions from the highest revision id already in the document (shared helper `JBGRevisionIds.max_revision_id`, also used by the "Om klarspråkningen" section), so they cannot collide with tracked changes in the original
- [x] No fixed 5 s pause between API calls: the OpenAI library already retries on rate limits (429), server errors, timeouts and dropped connections, with growing waits and respect for Retry-After. The number of retries is set in one place (`JBGModelClient.MODEL_CALL_MAX_RETRIES = 4`). Saves about 5 s per extra call
- [x] `CommentsRenderer` can anchor a comment to a whole paragraph (`add_paragraph_comment`), not only to tracked revisions (G4.1)
- [x] "Om klarspråkningen": cached prompt tokens are shown ("varav N återanvända (lägre kostnad)") so the cost is not overestimated
- [x] LIX: manual line breaks inside a sentence are no longer counted as sentence boundaries
- [x] LIX: sections start only at body headings; headings in tables/textboxes, captions and empty chapters no longer create sections
- [x] LIX: `lix_delta` is computed from the rounded values so it matches the displayed before/after values
- [x] Capture token usage per API call (`UsageTracker`) and report it in `<job>_run_summary.json`
- [x] Add heading/style metadata and true document order (`doc_order`) to the structure JSON without changing `element_id`s
- [x] Resolve and render headers/footers through their real OOXML relationship targets
- [x] Add regression coverage for shared section stories, variants, fields, markup and tracked changes
