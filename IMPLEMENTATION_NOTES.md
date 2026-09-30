# Header/footer, table and safety-filter implementation notes

This change completes the first modernization step agreed for JBGLangImprover.

## What changed

- Header and footer parts are resolved through `word/_rels/document.xml.rels`.
- Each physical story part is extracted once, even when inherited/shared by several sections.
- Default, first-page and even-page references are tracked independently.
- Every extracted story paragraph receives its actual `part_name` and an exact `container_path`.
- `ChangeTarget` carries the container path into rendering.
- A dedicated `HeaderFooterPartAdapter` locates the target paragraph in any safe Word XML part and verifies its OOXML root type.
- Simple markup and tracked changes now render header/footer plans.
- Motivation comments can anchor to tracked changes in header/footer parts.
- Nested textbox paragraphs no longer leak their runs into an enclosing paragraph model.
- Each paragraph in a table cell is extracted as its own element, using the ID
  `table_T_cell_R_C_pP`, so a suggestion cannot span OOXML paragraph boundaries.
- Renderers locate the exact table-cell paragraph while retaining compatibility
  with the legacy `table_T_cell_R_C` ID for the first paragraph.
- `spelling_degradation` is limited to probable loss of one or two internal
  characters in a single word. Phrase rewrites, terminology changes and suffix
  changes are no longer rejected by this heuristic.
- Textboxes are indexed from their containing `w:drawing` elements rather than
  from every `w:txbxContent`, because Word may keep duplicate fallback copies.
- Each textbox paragraph uses `textbox_N_pP`; saved legacy `textbox_N` and
  `table_T_cell_R_C` suggestions resolve to the first paragraph.
- Multi-run tracked changes calculate offsets across all visible run children,
  including tabs and manual line breaks.
- Long-text similarity disables `SequenceMatcher`'s character-inappropriate
  `autojunk` heuristic.

## Regression coverage

Run from the repository root:

```bash
python -m unittest discover -s tests -v
```

The tests create temporary Word files and cover:

- a relationship target moved to `word/storyParts/customHeader.xml`
- inherited/shared stories across two sections
- default, first-page and even-page headers
- a paragraph inside a header table
- preservation of a PAGE field
- simple markup, tracked changes and comments in headers/footers
- reopening generated packages with `python-docx`
- extraction and rendering of the second paragraph in a table cell
- legacy table-cell ID compatibility
- accepted real-world plain-language rewrites and rejected typo-like deletions
- textbox drawing order, multiple textbox paragraphs and legacy IDs
- tracked changes ending at a manual line break
- long-text similarity without false `low_similarity` rejection

## Real-document verification

The supplied Testrapport fixture and its 77 saved suggestions were replayed
through extraction, planning and tracked-change rendering. All 77 plans were
applied, and the resulting package reopened successfully with `python-docx`.

## Remaining manual verification

The generated packages pass the automated structural and round-trip checks. Before production deployment, open representative output documents in the supported Microsoft Word versions and verify visual layout, tracked-change review and comment presentation. The test suite does not automate Microsoft Word itself.
