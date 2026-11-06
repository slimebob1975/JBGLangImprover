from dataclasses import dataclass
import json
import os
from pathlib import Path
from difflib import SequenceMatcher
from datetime import datetime, timezone
from typing import Optional

from lxml import etree

try:
    from app.src.JBGDocxPackage import W_NS, DocxPackage
    from app.src.JBGDataRetention import loggable_text
    from app.src.JBGRunEditing import SpanSegment, append_visible_text, plan_span_edit, replace_runs
    from app.src.JBGRevisionIds import max_revision_id
    from app.src.JBGTokenDiffEngine import TokenDiffEngine
except ModuleNotFoundError:
    from JBGDocxPackage import W_NS, DocxPackage
    from JBGDataRetention import loggable_text
    from JBGRunEditing import SpanSegment, append_visible_text, plan_span_edit, replace_runs
    from JBGRevisionIds import max_revision_id
    from JBGTokenDiffEngine import TokenDiffEngine

try:
    from app.src.JBGDocumentPartAdapter import DocumentPartAdapter, ParagraphNode
except ModuleNotFoundError:
    from JBGDocumentPartAdapter import DocumentPartAdapter, ParagraphNode

try:
    from app.src.JBGHeaderFooterPartAdapter import HeaderFooterPartAdapter
except ModuleNotFoundError:
    from JBGHeaderFooterPartAdapter import HeaderFooterPartAdapter

try:
    from app.src.JBGFootnotesPartAdapter import FootnotesPartAdapter, FootnoteNode
except ModuleNotFoundError:
    from JBGFootnotesPartAdapter import FootnotesPartAdapter, FootnoteNode

try:
    from app.src.JBGChangePlanner import ChangePlan, DiffSegment
except ModuleNotFoundError:
    from JBGChangePlanner import ChangePlan, DiffSegment


XML_NS = "http://www.w3.org/XML/1998/namespace"


@dataclass
class RenderResult:
    plan: ChangePlan
    applied: bool
    message: str
    anchor_part_name: Optional[str] = None
    anchor_kind: Optional[str] = None   # "ins" / "del"
    anchor_revision_id: Optional[str] = None
    anchor_end_kind: Optional[str] = None
    anchor_end_revision_id: Optional[str] = None
    word_level_status: Optional[str] = None
    word_level_fallback_reason: Optional[str] = None


class TrackedChangesRenderer:
    DOCUMENT_ELEMENT_TYPES = {"paragraph", "table_cell", "textbox"}
    FOOTNOTE_ELEMENT_TYPES = {"footnote"}
    STORY_ELEMENT_TYPES = {"header", "footer"}
    SUPPORTED_ELEMENT_TYPES = (
        DOCUMENT_ELEMENT_TYPES | FOOTNOTE_ELEMENT_TYPES | STORY_ELEMENT_TYPES
    )

    def __init__(self, package: DocxPackage, logger, author: str = "JBG Klarspråkningstjänst", word_level_changes: bool = True):
        self.package = package
        self.logger = logger
        self.author = author
        self.word_level_changes = bool(word_level_changes)
        self.document_adapter = DocumentPartAdapter(package, logger)
        self.story_adapter = HeaderFooterPartAdapter(package, logger)

        self.footnotes_adapter = None
        if self.package.part_exists("word/footnotes.xml"):
            self.footnotes_adapter = FootnotesPartAdapter(package, logger)

        # Fortsätt från det högsta id som redan finns, så att nya ändringar inte
        # krockar med spårade ändringar i originalet (kommentarer förankras via id).
        self.change_id_counter = max_revision_id(package) + 1
        self._ensure_track_revisions_enabled()

    # ------------------------------------------------------------------
    # Publikt API
    # ------------------------------------------------------------------

    def apply_plan(self, plan: ChangePlan) -> RenderResult:
        try:
            if plan.target.element_type not in self.SUPPORTED_ELEMENT_TYPES:
                return RenderResult(
                    plan=plan,
                    applied=False,
                    message=(
                        "TrackedChangesRenderer v6 supports only "
                        f"{sorted(self.SUPPORTED_ELEMENT_TYPES)}, got {plan.target.element_type}"
                    ),
                )

            if plan.target.element_type in self.DOCUMENT_ELEMENT_TYPES:
                if plan.target.part_name != "word/document.xml":
                    return RenderResult(
                        plan=plan,
                        applied=False,
                        message=(
                            "TrackedChangesRenderer v6 expected word/document.xml for "
                            f"{plan.target.element_type}, got {plan.target.part_name}"
                        ),
                    )

                anchor_info = self._apply_document_plan(plan)
                self.package.write_document_tree(self.document_adapter.tree)

                return RenderResult(
                    plan=plan,
                    applied=True,
                    message="Applied tracked changes",
                    anchor_part_name="word/document.xml",
                    anchor_kind=anchor_info["anchor_kind"],
                    anchor_revision_id=anchor_info["anchor_revision_id"],
                    word_level_status=anchor_info.get("word_level_status"),
                    word_level_fallback_reason=anchor_info.get("word_level_fallback_reason"),
                    anchor_end_kind=anchor_info.get("anchor_end_kind"),
                    anchor_end_revision_id=anchor_info.get("anchor_end_revision_id"),
                )

            if plan.target.element_type in self.STORY_ELEMENT_TYPES:
                anchor_info = self._apply_story_plan(plan)
                self.story_adapter.write_tree(plan.target.part_name)
                return RenderResult(
                    plan=plan,
                    applied=True,
                    message="Applied tracked changes",
                    anchor_part_name=plan.target.part_name,
                    anchor_kind=anchor_info["anchor_kind"],
                    anchor_revision_id=anchor_info["anchor_revision_id"],
                    word_level_status=anchor_info.get("word_level_status"),
                    word_level_fallback_reason=anchor_info.get("word_level_fallback_reason"),
                    anchor_end_kind=anchor_info.get("anchor_end_kind"),
                    anchor_end_revision_id=anchor_info.get("anchor_end_revision_id"),
                )

            if plan.target.element_type in self.FOOTNOTE_ELEMENT_TYPES:
                if plan.target.part_name != "word/footnotes.xml":
                    return RenderResult(
                        plan=plan,
                        applied=False,
                        message=(
                            "TrackedChangesRenderer v6 expected word/footnotes.xml for "
                            f"{plan.target.element_type}, got {plan.target.part_name}"
                        ),
                    )

                if self.footnotes_adapter is None:
                    return RenderResult(
                        plan=plan,
                        applied=False,
                        message="Document has no footnotes.xml part",
                    )

                anchor_info = self._apply_footnote_plan(plan)
                self.package.write_footnotes_tree(self.footnotes_adapter.tree)

                return RenderResult(
                    plan=plan,
                    applied=True,
                    message="Applied tracked changes",
                    anchor_part_name="word/footnotes.xml",
                    anchor_kind=anchor_info["anchor_kind"],
                    anchor_revision_id=anchor_info["anchor_revision_id"],
                    word_level_status=anchor_info.get("word_level_status"),
                    word_level_fallback_reason=anchor_info.get("word_level_fallback_reason"),
                    anchor_end_kind=anchor_info.get("anchor_end_kind"),
                    anchor_end_revision_id=anchor_info.get("anchor_end_revision_id"),
                )

            return RenderResult(plan=plan, applied=False, message="Unsupported element type")

        except Exception as ex:
            return RenderResult(plan=plan, applied=False, message=str(ex))

    def apply_plans(self, plans: list[ChangePlan]) -> list[RenderResult]:
        results = []
        self._fallback_fixtures = []

        grouped: dict[tuple, list[ChangePlan]] = {}

        for plan in plans:
            if plan.target.element_type not in self.SUPPORTED_ELEMENT_TYPES:
                results.append(RenderResult(
                    plan=plan,
                    applied=False,
                    message=(
                        "TrackedChangesRenderer v6 supports only "
                        f"{sorted(self.SUPPORTED_ELEMENT_TYPES)}, got {plan.target.element_type}"
                    ),
                ))
                continue

            key = (
                plan.target.part_name,
                plan.target.element_type,
                plan.target.element_id,
                plan.target.footnote_id,
            )
            grouped.setdefault(key, []).append(plan)

        for _, group in grouped.items():
            valid_group = []
            for plan in group:
                if "overlapping_change_conflict" in plan.notes:
                    results.append(RenderResult(
                        plan=plan,
                        applied=False,
                        message="Skipped due to overlapping_change_conflict",
                    ))
                else:
                    valid_group.append(plan)

            if not valid_group:
                continue

            valid_group.sort(key=lambda p: p.anchor.start, reverse=True)

            for plan in valid_group:
                # Opt-in: capture the exact pre-edit OOXML only for genuine
                # formatting/container fallbacks, never for ordinary runs.
                snapshot = self._fallback_fixture_snapshot(plan)
                result = self.apply_plan(plan)
                results.append(result)
                if (snapshot is not None and result.applied and
                        result.word_level_fallback_reason in
                        {"mixed_run_formatting", "multiple_edit_segments"}):
                    snapshot["fallback_reason"] = result.word_level_fallback_reason
                    self._fallback_fixtures.append(snapshot)

                if result.applied:
                    if plan.target.part_name == "word/document.xml":
                        self.document_adapter.refresh()
                    elif plan.target.part_name == "word/footnotes.xml" and self.footnotes_adapter is not None:
                        self.footnotes_adapter.refresh()
                    elif plan.target.element_type in self.STORY_ELEMENT_TYPES:
                        self.story_adapter.refresh(plan.target.part_name)

        self._write_fallback_fixtures()
        return results

    def _fallback_fixture_snapshot(self, plan: ChangePlan):
        """Return a lossless paragraph snapshot, without changing the renderer.

        The snapshot includes source text and raw XML. It may contain private
        document content, so it is exported only with explicit opt-in.
        """
        if (not self.word_level_changes or
                not os.environ.get("JBG_WORDLEVEL_CAPTURE_PATH") or
                plan.target.element_type != "paragraph" or
                plan.target.part_name != "word/document.xml"):
            return None
        try:
            model = self.document_adapter.get_paragraph_model_for_plan(plan)
            paragraph = model.paragraph_element
            # Do not attempt to replay paragraphs that already contain edits.
            # They depend on the order of other proposals.
            if any(el.tag in {f"{{{W_NS}}}ins", f"{{{W_NS}}}del"}
                   for el in paragraph.iter()):
                return None
            old = model.visible_text
            # Reject formats that the standalone replay model cannot represent
            # exactly. Footnote references are retained as immutable anchors.
            if any(el.tag in {f"{{{W_NS}}}tab", f"{{{W_NS}}}br", f"{{{W_NS}}}drawing"}
                   for el in paragraph.iter()):
                return None
            if (plan.anchor.start < 0 or plan.anchor.end > len(old) or
                    old[plan.anchor.start:plan.anchor.end] != plan.old_text):
                return None
            if old.count(plan.old_text) != 1:
                return None
            return {
                "id": str(plan.target.element_id),
                "source_text": old,
                "old": plan.old_text,
                "new": plan.new_text,
                "anchor_start": plan.anchor.start,
                "anchor_end": plan.anchor.end,
                "paragraph_xml": etree.tostring(paragraph, encoding="unicode"),
            }
        except Exception as exc:
            self.logger.debug("Word-level fixture snapshot skipped: %s", exc)
            return None

    def _write_fallback_fixtures(self):
        path = os.environ.get("JBG_WORDLEVEL_CAPTURE_PATH")
        if not path or not self._fallback_fixtures:
            return
        try:
            target = Path(path)
            # Do not silently overwrite existing captures from earlier runs.
            if target.exists():
                self.logger.warning("Word-level capture file exists; not overwriting: %s", target)
                return
            target.parent.mkdir(parents=True, exist_ok=True)
            payload = {
                "schema": 1,
                "description": "Sensitive pre-render OOXML snapshots of actual fallback proposals; keep local.",
                "cases": self._fallback_fixtures,
            }
            with target.open("x", encoding="utf-8") as stream:
                json.dump(payload, stream, ensure_ascii=False, indent=2)
                stream.write("\n")
        except OSError as exc:
            self.logger.warning("Could not export word-level diagnostics: %s", exc)

    # ------------------------------------------------------------------
    # Hög-nivå per elementtyp
    # ------------------------------------------------------------------

    def _apply_document_plan(self, plan: ChangePlan) -> dict:
        if plan.target.element_type == "table_cell":
            return self._apply_table_cell_plan(plan)

        located = self.document_adapter.locate_plan_nodes(plan)
        model = located["paragraph_model"]
        paragraph = model.paragraph_element
        overlapping_nodes = located["overlapping_nodes"]

        return self._apply_tracked_change_to_paragraph_element(
            paragraph=paragraph,
            visible_text=model.visible_text,
            overlapping_nodes=overlapping_nodes,
            anchor_start=plan.anchor.start,
            anchor_end=plan.anchor.end,
            old_text=plan.old_text,
            new_text=plan.new_text,
            diff_segments=plan.diff_segments,
        )

    def _apply_story_plan(self, plan: ChangePlan) -> dict:
        located = self.story_adapter.locate_plan_nodes(plan)
        model = located["paragraph_model"]
        return self._apply_tracked_change_to_paragraph_element(
            paragraph=model.paragraph_element,
            visible_text=model.visible_text,
            overlapping_nodes=located["overlapping_nodes"],
            anchor_start=plan.anchor.start,
            anchor_end=plan.anchor.end,
            old_text=plan.old_text,
            new_text=plan.new_text,
            diff_segments=plan.diff_segments,
        )

    def _apply_table_cell_plan(self, plan: ChangePlan) -> dict:
        try:
            located = self.document_adapter.locate_plan_nodes(plan)
            model = located["paragraph_model"]
            paragraph = model.paragraph_element
            overlapping_nodes = located["overlapping_nodes"]

            return self._apply_tracked_change_to_paragraph_element(
                paragraph=paragraph,
                visible_text=model.visible_text,
                overlapping_nodes=overlapping_nodes,
                anchor_start=plan.anchor.start,
                anchor_end=plan.anchor.end,
                old_text=plan.old_text,
                new_text=plan.new_text,
                diff_segments=plan.diff_segments,
            )
        except Exception:
            pass

        cell = self._find_table_cell_element(plan.target.element_id)
        if cell is None:
            raise ValueError(f"Could not resolve table cell for {plan.target.element_id}")

        paragraphs = cell.findall(f"./{{{W_NS}}}p")
        if not paragraphs:
            raise ValueError(f"No paragraphs found in table cell {plan.target.element_id}")

        for paragraph in paragraphs:
            paragraph_model = self._build_paragraph_model_from_element(paragraph)
            anchor = self._find_anchor_in_text(plan.old_text, paragraph_model["visible_text"])
            if anchor is None:
                continue

            overlapping_nodes = self._get_overlapping_nodes_for_generic_paragraph(
                nodes=paragraph_model["nodes"],
                anchor_start=anchor["start"],
                anchor_end=anchor["end"],
            )
            if not overlapping_nodes:
                continue

            return self._apply_tracked_change_to_paragraph_element(
                paragraph=paragraph,
                visible_text=paragraph_model["visible_text"],
                overlapping_nodes=overlapping_nodes,
                anchor_start=anchor["start"],
                anchor_end=anchor["end"],
                old_text=plan.old_text,
                new_text=plan.new_text,
                diff_segments=plan.diff_segments,
            )

        raise ValueError(f"No matching paragraph found in table cell {plan.target.element_id}")

    def _apply_footnote_plan(self, plan: ChangePlan) -> dict:
        if self.footnotes_adapter is None:
            raise ValueError("Document has no footnotes.xml part")

        model = self.footnotes_adapter.get_footnote_model_for_plan(plan)
        paragraph_models = self._build_footnote_paragraph_models(model)

        best = None
        for pm in paragraph_models:
            anchor = self._find_anchor_in_text(plan.old_text, pm["editable_text"])
            if anchor is None:
                continue
            best = (pm, anchor)
            break

        if best is None:
            raise ValueError(
                f"Could not locate footnote text in editable footnote paragraphs for footnote_id={plan.target.footnote_id}"
            )

        pm, anchor = best

        overlapping_nodes = self._get_overlapping_nodes_for_generic_paragraph(
            nodes=pm["editable_nodes"],
            anchor_start=anchor["start"],
            anchor_end=anchor["end"],
        )

        if not overlapping_nodes:
            raise ValueError("No overlapping footnote text nodes found for editable footnote paragraph")

        return self._apply_tracked_change_to_paragraph_element(
            paragraph=pm["paragraph"],
            visible_text=pm["editable_text"],
            overlapping_nodes=overlapping_nodes,
            anchor_start=anchor["start"],
            anchor_end=anchor["end"],
            old_text=plan.old_text,
            new_text=plan.new_text,
            diff_segments=plan.diff_segments,
        )

    # ------------------------------------------------------------------
    # Gemensam låg-nivå
    # ------------------------------------------------------------------

    def _apply_tracked_change_to_paragraph_element(
        self,
        paragraph: etree._Element,
        visible_text: str,
        overlapping_nodes: list,
        anchor_start: int,
        anchor_end: int,
        old_text: str,
        new_text: str,
        diff_segments=None,
    ) -> dict:
        matched = visible_text[anchor_start:anchor_end]
        if not self._texts_match_leniently(matched, old_text):
            raise ValueError(
                f"Anchor text mismatch. Expected old_text={loggable_text(old_text)}, found={loggable_text(matched)}"
            )

        first_node = overlapping_nodes[0]
        last_node = overlapping_nodes[-1]

        if first_node.run_element is None or last_node.run_element is None:
            raise ValueError("Anchor overlaps node(s) without run_element")

        return self._rewrite_span(
            paragraph=paragraph,
            first_node=first_node,
            last_node=last_node,
            anchor_start=anchor_start,
            anchor_end=anchor_end,
            old_text=old_text,
            new_text=new_text,
            diff_segments=diff_segments,
        )

    # ------------------------------------------------------------------
    # Footnote helpers
    # ------------------------------------------------------------------

    def _build_footnote_paragraph_models(self, footnote_model) -> list[dict]:
        models = []

        paragraphs = footnote_model.paragraphs
        for paragraph in paragraphs:
            paragraph_nodes = [n for n in footnote_model.nodes if n.paragraph_element is paragraph]

            full_text_parts = []
            editable_text_parts = []
            editable_nodes = []

            seen_editable_text = False
            editable_cursor = 0

            for node in paragraph_nodes:
                if node.kind in {"text", "tab", "linebreak"}:
                    full_text_parts.append(node.text)

                if node.kind == "footnote_ref" and not seen_editable_text:
                    continue

                if node.kind == "text":
                    seen_editable_text = True
                    editable_nodes.append(self._clone_node_with_new_span(node, editable_cursor, editable_cursor + len(node.text)))
                    editable_text_parts.append(node.text)
                    editable_cursor += len(node.text)
                elif node.kind in {"tab", "linebreak"}:
                    if seen_editable_text:
                        editable_nodes.append(self._clone_node_with_new_span(node, editable_cursor, editable_cursor + len(node.text)))
                        editable_text_parts.append(node.text)
                        editable_cursor += len(node.text)

            models.append({
                "paragraph": paragraph,
                "full_text": "".join(full_text_parts),
                "editable_text": "".join(editable_text_parts),
                "editable_nodes": editable_nodes,
            })

        return models

    def _clone_node_with_new_span(self, node, start: int, end: int):
        node_type = node.__class__
        return node_type(
            kind=node.kind,
            text=node.text,
            element=node.element,
            run_element=node.run_element,
            paragraph_element=node.paragraph_element,
            start=start,
            end=end,
            meta=dict(node.meta),
        )

    # ------------------------------------------------------------------
    # Generic helpers
    # ------------------------------------------------------------------

    def _build_paragraph_model_from_element(self, paragraph: etree._Element) -> dict:
        nodes = []
        visible_parts = []
        cursor = 0

        for run in paragraph.findall(f"./{{{W_NS}}}r"):
            for child in run:
                if child.tag == f"{{{W_NS}}}t":
                    text = child.text or ""
                    node = ParagraphNode(
                        kind="text",
                        text=text,
                        element=child,
                        run_element=run,
                        start=cursor,
                        end=cursor + len(text),
                    )
                    nodes.append(node)
                    visible_parts.append(text)
                    cursor += len(text)

                elif child.tag == f"{{{W_NS}}}tab":
                    node = ParagraphNode(
                        kind="tab",
                        text="\t",
                        element=child,
                        run_element=run,
                        start=cursor,
                        end=cursor + 1,
                    )
                    nodes.append(node)
                    visible_parts.append("\t")
                    cursor += 1

                elif child.tag in {f"{{{W_NS}}}br", f"{{{W_NS}}}cr"}:
                    node = ParagraphNode(
                        kind="linebreak",
                        text="\n",
                        element=child,
                        run_element=run,
                        start=cursor,
                        end=cursor + 1,
                    )
                    nodes.append(node)
                    visible_parts.append("\n")
                    cursor += 1

        return {
            "paragraph": paragraph,
            "visible_text": "".join(visible_parts),
            "nodes": nodes,
        }

    def _get_overlapping_nodes_for_generic_paragraph(self, nodes: list, anchor_start: int, anchor_end: int) -> list:
        overlapping_nodes = []
        for node in nodes:
            if node.kind not in {"text", "tab", "linebreak"}:
                continue
            if max(node.start, anchor_start) < min(node.end, anchor_end):
                overlapping_nodes.append(node)
        return overlapping_nodes

    def _find_anchor_in_text(self, old_text: str, source_text: str) -> Optional[dict]:
        idx = source_text.find(old_text)
        if idx != -1:
            return {"start": idx, "end": idx + len(old_text)}
        return None

    def _texts_match_leniently(self, found: str, expected: str) -> bool:
        if found == expected:
            return True
        if found.strip() == expected.strip():
            return True
        if expected.startswith(found) and len(expected) - len(found) <= 2:
            return True
        return False

    def _find_table_cell_element(self, element_id: Optional[str]) -> Optional[etree._Element]:
        if not element_id:
            return None

        parts = element_id.split("_")
        if len(parts) not in {5, 6} or parts[0] != "table" or parts[2] != "cell":
            return None

        try:
            table_index = int(parts[1])
            row_index = int(parts[3])
            col_index = int(parts[4])
        except Exception:
            return None

        root = self.document_adapter.root
        tables = root.xpath("//w:body/w:tbl", namespaces={"w": W_NS})
        if table_index < 1 or table_index > len(tables):
            return None

        table = tables[table_index - 1]
        rows = table.findall(f"{{{W_NS}}}tr")
        if row_index < 1 or row_index > len(rows):
            return None

        row = rows[row_index - 1]
        cells = row.findall(f"{{{W_NS}}}tc")
        if col_index < 1 or col_index > len(cells):
            return None

        return cells[col_index - 1]

    # ------------------------------------------------------------------
    # Single-run case
    # ------------------------------------------------------------------

    def _rewrite_span(
        self,
        paragraph: etree._Element,
        first_node,
        last_node,
        anchor_start: int,
        anchor_end: int,
        old_text: str,
        new_text: str,
        diff_segments=None,
    ) -> dict:
        """
        Ändrar styckets text [anchor_start, anchor_end), oavsett om ändringen
        ligger i en textbit, flera textbitar eller går in i en hyperlänk
        (se JBGRunEditing). Text som inte ändras i berörda textbitar behålls,
        med radbrytningar och tabbar.
        """
        plan = plan_span_edit(paragraph, first_node, last_node, anchor_start, anchor_end, new_text)
        if not self._texts_match_leniently(plan.actual_old, old_text):
            raise ValueError(
                f"Span text mismatch. Expected {loggable_text(old_text)}, got {loggable_text(plan.actual_old)}"
            )

        # A span replacement recreates runs from visible text. It cannot
        # recreate a footnote reference, a field, a drawing or other opaque
        # OOXML child. Reject the WHOLE proposal before editing any segment,
        # rather than silently dropping a reference from the document. This
        # protection applies to both word-level and legacy span mode.
        unsafe = self._unsafe_span_run_elements(plan.segments)
        if unsafe:
            # A standalone footnote-reference run can act as an immutable
            # boundary only if an unchanged, unique 16-character context
            # straddles its position in BOTH original and proposed text.
            # Otherwise continue to reject the entire proposal (v5 policy).
            if unsafe == ("footnoteReference",) and len(plan.segments) == 1:
                split = self._split_at_stable_footnote_reference(plan.segments[0])
                if split is not None:
                    plan.segments = split
                    unsafe = self._unsafe_span_run_elements(plan.segments)
            if unsafe:
                raise ValueError("unsafe_span_run_content:" + ",".join(unsafe))

        # The original span renderer remains the only fallback. Try a second,
        # strictly lossless diff only when the planner's normalized diff cannot
        # reproduce the input byte-for-byte (notably NBSP and Unicode spaces).
        fallback_reason = None
        if self.word_level_changes:
            if len(plan.segments) != 1:
                # A stable footnote split can still contain independent,
                # lossless token-level changes on either side. Validate EVERY
                # piece before modifying XML; never rewrite through the
                # reference or through nested containers such as hyperlinks.
                prepared = self._prepare_immutable_boundary_word_edits(paragraph, plan.segments)
                if prepared is None:
                    prepared = self._prepare_bookmark_trailing_space_word_edits(paragraph, plan.segments)
                if prepared is not None:
                    first_result = last_result = None
                    for piece, exact in prepared:
                        item = self._rewrite_segmented_run(piece, exact)
                        if first_result is None:
                            first_result = item
                        last_result = item
                    first_result["anchor_end_kind"] = last_result["anchor_end_kind"]
                    first_result["anchor_end_revision_id"] = last_result["anchor_end_revision_id"]
                    first_result["word_level_status"] = "word_level"
                    return first_result
                fallback_reason = "multiple_edit_segments"
            else:
                fallback_reason = self._word_diff_safety_reason(old_text, new_text, diff_segments)
                if fallback_reason == "diff_text_mismatch":
                    exact_segments = TokenDiffEngine().build_diff(old_text, new_text, lossless=True)
                    exact_reason = self._word_diff_safety_reason(old_text, new_text, exact_segments)
                    if exact_reason is None:
                        diff_segments = exact_segments
                        fallback_reason = None
                    else:
                        fallback_reason = f"exact_{exact_reason}"
                if fallback_reason == "diff_no_common_text":
                    # A single-token replacement without any equal tokens is
                    # already represented by the legacy span renderer as one
                    # deletion and one insertion.  Reclassifying it as word
                    # level would inflate the proposal metric without making
                    # Word's revision any more precise.  Keep the established
                    # renderer and distinguish this visually equivalent case
                    # from genuinely unresolved no-common-token proposals.
                    segment = plan.segments[0]
                    if self._no_common_equivalent_span(paragraph, segment, old_text, new_text, diff_segments):
                        fallback_reason = "diff_no_common_text_equivalent_span"
                if fallback_reason is None:
                    segment = plan.segments[0]
                    fallback_reason = self._word_run_safety_reason(paragraph, segment, old_text, new_text)
                    if fallback_reason == "mixed_run_formatting":
                        # Only cross a formatting boundary when an unchanged,
                        # unique text window pins that boundary in both views.
                        # Each revised run then retains its own run properties.
                        pieces = self._stable_formatted_run_pieces(segment)
                        if pieces is None:
                            pieces = self._isolated_run_word_pieces(segment)
                        if pieces is not None:
                            prepared = []
                            for piece in pieces:
                                if piece.old == piece.new:
                                    continue
                                exact = TokenDiffEngine().build_diff(piece.old, piece.new, lossless=True)
                                if (not exact or any(d.kind not in {"equal", "insert", "delete"}
                                                     for d in exact)
                                        or "".join(d.text for d in exact if d.kind != "insert") != piece.old
                                        or "".join(d.text for d in exact if d.kind != "delete") != piece.new):
                                    prepared = None
                                    break
                                prepared.append((piece, exact))
                            if prepared:
                                first_result = last_result = None
                                for piece, exact in prepared:
                                    item = self._rewrite_segmented_run(piece, exact)
                                    if first_result is None:
                                        first_result = item
                                    last_result = item
                                first_result["anchor_end_kind"] = last_result["anchor_end_kind"]
                                first_result["anchor_end_revision_id"] = last_result["anchor_end_revision_id"]
                                first_result["word_level_status"] = "word_level"
                                return first_result
                    if fallback_reason is None:
                        result = self._rewrite_segmented_run(segment, diff_segments)
                        result["word_level_status"] = "word_level"
                        return result

        anchor = None
        last_anchor = None
        for segment in plan.segments:
            if segment.old == segment.new:
                # An unchanged side of a protected reference stays byte-for-
                # byte untouched, including all formatting and bookmark data.
                continue
            new_elements = []
            inserted_wrapper = deleted_wrapper = None
            if segment.before:
                new_elements.append(self._clone_run_with_text(segment.first_run, segment.before))
            if segment.old:
                deleted_wrapper = self._make_deleted_wrapper(segment.old, source_run=segment.first_run)
                new_elements.append(deleted_wrapper)
            if segment.new:
                inserted_wrapper = self._make_inserted_wrapper(segment.new, source_run=segment.first_run)
                new_elements.append(inserted_wrapper)
            if segment.after:
                new_elements.append(self._clone_run_with_text(segment.last_run, segment.after))
            replace_runs(segment, new_elements)
            if inserted_wrapper is not None or deleted_wrapper is not None:
                current_anchor = self._select_comment_anchor(inserted_wrapper, deleted_wrapper)
                if anchor is None:
                    anchor = current_anchor
                last_anchor = current_anchor

        if anchor is None:
            raise ValueError("No revision produced by span replacement")
        anchor_element, anchor_kind = anchor
        result = {
            "anchor_kind": anchor_kind,
            "anchor_revision_id": self._get_revision_id(anchor_element),
        }
        # For a protected boundary, comment anchors must cover both sides
        # if both contain edits. The preserved footnote remains between them.
        if len(plan.segments) > 1 and last_anchor is not None:
            last_element, last_kind = last_anchor
            result["anchor_end_kind"] = last_kind
            result["anchor_end_revision_id"] = self._get_revision_id(last_element)
        if self.word_level_changes:
            result["word_level_status"] = "span_fallback"
            result["word_level_fallback_reason"] = fallback_reason
        return result

    @classmethod
    def _prepare_immutable_boundary_word_edits(cls, paragraph, segments):
        """Preflight a proven standalone-footnote split, without XML writes.

        The existing protected-reference splitter is the only source of such
        segments. Each changed side must consist of one direct plain run,
        be exactly lossless and have its own unchanged token anchor. A side
        without changes is kept verbatim. We intentionally do not support
        hyperlinks, multiple runs or formatting changes here.
        """
        if len(segments) != 2 or any(s.container is not paragraph for s in segments):
            return None
        left, right = segments
        if (len(left.runs) != 1 or len(right.runs) != 1
                or left.last_run.getnext() is None
                or left.last_run.getnext().tag != f"{{{W_NS}}}r"
                or left.last_run.getnext().find(f"{{{W_NS}}}footnoteReference") is None
                or left.last_run.getnext().getnext() is not right.first_run):
            return None
        prepared = []
        for segment in segments:
            if segment.old == segment.new:
                continue
            exact = TokenDiffEngine().build_diff(segment.old, segment.new, lossless=True)
            if cls._word_diff_safety_reason(segment.old, segment.new, exact) is not None:
                return None
            if cls._word_run_safety_reason(paragraph, segment, segment.old, segment.new) is not None:
                return None
            prepared.append((segment, exact))
        # A comment needs at least one anchor, and both sides have already
        # been checked before the first run is replaced.
        return prepared or None

    @classmethod
    def _prepare_bookmark_trailing_space_word_edits(cls, paragraph, segments):
        """Preserve a heading bookmark when only its outside trailing space is removed.

        The source span must end in a whitespace-only run immediately after
        one bookmarkEnd; all other changed text must be in contiguous plain
        direct runs immediately before that bookmarkEnd. Never move or
        reconstruct the bookmark: replace only the text runs. No generic
        splitting of changed bookmark boundaries is attempted here.
        """
        if len(segments) != 2:
            return None
        main, suffix = segments
        if (main.container is not paragraph or suffix.container is not paragraph
                or not main.runs or len(suffix.runs) != 1
                or suffix.before or suffix.after or not suffix.old
                or not suffix.old.isspace() or suffix.new != ""):
            return None
        boundary = main.last_run.getnext()
        if (boundary is None or boundary.tag != f"{{{W_NS}}}bookmarkEnd"
                or boundary.getnext() is not suffix.first_run):
            return None
        bookmark_id = boundary.get(f"{{{W_NS}}}id")
        starts = paragraph.xpath("./w:bookmarkStart[@w:id=$id]",
                                 namespaces={"w": W_NS}, id=bookmark_id)
        ends = paragraph.xpath("./w:bookmarkEnd[@w:id=$id]",
                               namespaces={"w": W_NS}, id=bookmark_id)
        if not bookmark_id or len(starts) != 1 or ends != [boundary]:
            return None
        children = list(paragraph)
        if not (children.index(starts[0]) < children.index(main.first_run)
                and children.index(main.last_run) < children.index(boundary)):
            return None
        # Do not allow intervening runs/markers or heterogeneous run properties.
        for first, second in zip(main.runs, main.runs[1:]):
            if first.getnext() is not second:
                return None
        for piece in (main, suffix):
            if cls._word_run_safety_reason(paragraph, piece, piece.old, piece.new) is not None:
                return None
        main_diff = TokenDiffEngine().build_diff(main.old, main.new, lossless=True)
        if cls._word_diff_safety_reason(main.old, main.new, main_diff) is not None:
            return None
        suffix_diff = TokenDiffEngine().build_diff(suffix.old, suffix.new, lossless=True)
        if ("".join(d.text for d in suffix_diff if d.kind != "insert") != suffix.old
                or any(d.kind != "delete" for d in suffix_diff)):
            return None
        return [(main, main_diff), (suffix, suffix_diff)]

    @staticmethod
    def _split_at_stable_footnote_reference(segment):
        """Plan two text edits around one untouched standalone footnote run.

        Require a unique 16-character equal context crossing the exact
        reference boundary. This deliberately refuses cases where AI rewrote
        the phrase next to the reference or where placement is ambiguous.
        No XML is changed until the caller has completed all safety checks.
        """
        reference_tag = f"{{{W_NS}}}footnoteReference"
        allowed_text = {f"{{{W_NS}}}rPr", f"{{{W_NS}}}t"}
        run_list = segment.runs
        if len(run_list) < 3 or segment.container.tag != f"{{{W_NS}}}p":
            return None
        reference_positions = []
        for index, run in enumerate(run_list):
            children = [child for child in run if child.tag != f"{{{W_NS}}}rPr"]
            if len(children) == 1 and children[0].tag == reference_tag:
                if not children[0].get(f"{{{W_NS}}}id"):
                    return None
                reference_positions.append(index)
            elif any(child.tag not in allowed_text for child in run):
                return None
        if len(reference_positions) != 1:
            return None
        ref_pos = reference_positions[0]
        if ref_pos == 0 or ref_pos == len(run_list) - 1:
            return None
        # Do not rewrite through hyperlinks or interleaved comments.
        if any(run.getparent() is not segment.container for run in run_list):
            return None
        if any(run_list[k].getnext() is not run_list[k + 1]
               for k in range(len(run_list) - 1)):
            return None
        def text_of(run):
            return "".join(child.text or "" for child in run if child.tag == f"{{{W_NS}}}t")
        # segment.old excludes unchanged text before/after the edit. The
        # reference offset above is measured from the *whole first run*, so
        # subtract the untouched prefix before mapping it into segment.old.
        # Otherwise a reference after "metoder" can drift into "Sammantaget"
        # when the first run contains text before the edit anchor.
        complete_text = "".join(text_of(run) for run in run_list)
        if complete_text != segment.before + segment.old + segment.after:
            return None
        boundary = sum(len(text_of(run)) for run in run_list[:ref_pos]) - len(segment.before)
        old, new = segment.old, segment.new
        if boundary < 0 or boundary > len(old):
            return None
        if boundary < 8 or len(old) - boundary < 8:
            return None
        context = old[boundary - 8:boundary + 8]
        if old.count(context) != 1 or new.count(context) != 1:
            return None
        cut = None
        for kind, a, b, c, d in SequenceMatcher(None, old, new, autojunk=False).get_opcodes():
            if kind == "equal" and a <= boundary - 8 and b >= boundary + 8:
                cut = c + (boundary - a)
                break
        if cut is None or new[cut - 8:cut + 8] != context:
            return None
        left_runs, right_runs = run_list[:ref_pos], run_list[ref_pos + 1:]
        left = SpanSegment(container=segment.container, runs=left_runs,
                           old=old[:boundary], new=new[:cut],
                           before=segment.before, after="",
                           first_run=left_runs[0], last_run=left_runs[-1])
        right = SpanSegment(container=segment.container, runs=right_runs,
                            old=old[boundary:], new=new[cut:],
                            before="", after=segment.after,
                            first_run=right_runs[0], last_run=right_runs[-1])
        return [left, right]

    @staticmethod
    def _isolated_run_word_pieces(segment):
        """Conservatively edit individual runs without moving formatting boundaries.

        Every lossless token edit must fit within exactly one original text run.
        Edits that cross runs, ambiguous insertions at a run boundary, nested
        containers, or non-text run children remain in span fallback. Untouched
        runs are not rebuilt. This is particularly useful just before a footnote
        whose reference must remain entirely outside the edited text.
        """
        runs = segment.runs
        W = f"{{{W_NS}}}"
        if len(runs) < 2 or segment.container.tag != W + "p":
            return None
        texts = []
        for run in runs:
            if (run.tag != W + "r" or run.getparent() is not segment.container
                    or any(ch.tag not in {W + "rPr", W + "t"} for ch in run)):
                return None
            value = "".join(ch.text or "" for ch in run if ch.tag == W + "t")
            if not value:
                return None
            texts.append(value)
        if any(a.getnext() is not b for a, b in zip(runs, runs[1:])):
            return None
        if "".join(texts) != segment.before + segment.old + segment.after:
            return None
        exact = TokenDiffEngine().build_diff(segment.old, segment.new, lossless=True)
        if (not exact or
                "".join(d.text for d in exact if d.kind != "insert") != segment.old or
                "".join(d.text for d in exact if d.kind != "delete") != segment.new):
            return None
        # Edits are grouped until the next equal token. This keeps a delete
        # and its replacement insertion on the SAME side of a run boundary.
        edits = []
        old_cursor = 0
        group = []
        group_start = 0
        for diff in exact + [DiffSegment(kind="equal", text="")]:
            if diff.kind == "equal":
                if group:
                    removed = "".join(d.text for d in group if d.kind == "delete")
                    added = "".join(d.text for d in group if d.kind == "insert")
                    edits.append((group_start, group_start + len(removed), added))
                    group = []
                old_cursor += len(diff.text)
            else:
                if not group:
                    group_start = old_cursor
                group.append(diff)
                if diff.kind == "delete":
                    old_cursor += len(diff.text)
        if not edits:
            return None
        limits = []
        cursor = -len(segment.before)
        for value in texts:
            limits.append((cursor, cursor + len(value)))
            cursor += len(value)
        assigned = {}
        for start, end, replacement in edits:
            candidates = [i for i, (a, b) in enumerate(limits)
                          if (a <= start < end <= b if end > start
                              else a < start < b)]
            if len(candidates) != 1:
                return None
            index = candidates[0]
            offset = start - limits[index][0]
            assigned.setdefault(index, []).append((offset, end - start, replacement))
        # Adjacent edited text runs need an independently fixed formatting
        # boundary; leave that case to the existing stable-boundary planner.
        # A trailing whitespace-only run may be deleted next to an edited
        # text run, because no lexical text crosses the run boundary.
        changed_indices = sorted(assigned)
        if any(b == a + 1 and texts[a].strip() and texts[b].strip()
               for a, b in zip(changed_indices, changed_indices[1:])):
            return None
        pieces = []
        for index, edits_in_run in sorted(assigned.items()):
            source = texts[index]
            updated = source
            for offset, count, inserted in reversed(edits_in_run):
                updated = updated[:offset] + inserted + updated[offset + count:]
            if updated == source:
                continue
            run = runs[index]
            pieces.append(SpanSegment(container=segment.container, runs=[run],
                                      old=source, new=updated, before="", after="",
                                      first_run=run, last_run=run))
        if not pieces:
            return None
        # Reconstruct the complete replacement before the first XML write.
        combined = list(texts)
        for piece in pieces:
            combined[runs.index(piece.first_run)] = piece.new
        if "".join(combined) != segment.before + segment.new + segment.after:
            return None
        return pieces

    @staticmethod
    def _stable_formatted_run_pieces(segment):
        """Conservatively split one edit over plain differently formatted runs.

        Boundaries must be pinned by unique equal text on *both* sides.
        Allow partial outer runs only after exact prefix/suffix verification.
        Refuse markup or ambiguous/revised boundaries.
        This is deliberately narrower than general multi-run editing.
        """
        runs = segment.runs
        if (len(runs) < 2 or segment.container.tag != f"{{{W_NS}}}p"):
            return None
        allowed = {f"{{{W_NS}}}rPr", f"{{{W_NS}}}t"}
        if any(run.tag != f"{{{W_NS}}}r" or run.getparent() is not segment.container
               or any(child.tag not in allowed for child in run)
               or not any(child.tag == f"{{{W_NS}}}t" for child in run)
               for run in runs):
            return None
        if any(runs[i].getnext() is not runs[i + 1] for i in range(len(runs) - 1)):
            return None
        texts = ["".join(child.text or "" for child in run
                         if child.tag == f"{{{W_NS}}}t") for run in runs]
        old, new = segment.old, segment.new
        # The edit can start/end inside its outer runs. Check the *whole*
        # source before translating formatting boundaries into edit offsets.
        # Otherwise untouched prefixes drift across revision boundaries.
        if not all(texts) or "".join(texts) != segment.before + old + segment.after:
            return None
        # A boundary can only be reused when its surrounding text is fixed.
        # Try four characters per side first; three is the strict minimum.
        matcher = SequenceMatcher(None, old, new, autojunk=False)
        opcodes = matcher.get_opcodes()
        old_cuts = []
        pos = 0
        for value in texts[:-1]:
            pos += len(value)
            cut = pos - len(segment.before)
            # Every source run must own actual changed-span text. A run
            # consisting only of untouched prefix/suffix is not split here.
            if not 0 < cut < len(old):
                return None
            old_cuts.append(cut)
        new_cuts = []
        for pos in old_cuts:
            # Prefer the original 4+4 proof. A *unique*, unchanged 3+3
            # window also pins the same exact run boundary when a nearby
            # edit leaves fewer than four characters on either side. Never
            # infer where revised text belongs: the entire window must be
            # inside one SequenceMatcher equal opcode in both texts.
            mapped = None
            for radius in (4, 3):
                if pos < radius or len(old) - pos < radius:
                    continue
                context = old[pos - radius:pos + radius]
                if old.count(context) != 1 or new.count(context) != 1:
                    continue
                for kind, i1, i2, j1, j2 in opcodes:
                    if kind != "equal" or i1 > pos - radius or i2 < pos + radius:
                        continue
                    candidate = j1 + pos - i1
                    if new[candidate - radius:candidate + radius] == context:
                        mapped = candidate
                        break
                if mapped is not None:
                    break
            if mapped is None:
                return None
            new_cuts.append(mapped)
        if any(left >= right for left, right in zip(new_cuts, new_cuts[1:])):
            return None
        pieces = []
        boundaries_old = [0] + old_cuts + [len(old)]
        boundaries_new = [0] + new_cuts + [len(new)]
        for i, run in enumerate(runs):
            pieces.append(SpanSegment(container=segment.container, runs=[run],
                                      old=old[boundaries_old[i]:boundaries_old[i + 1]],
                                      new=new[boundaries_new[i]:boundaries_new[i + 1]],
                                      before=segment.before if i == 0 else "",
                                      after=segment.after if i == len(runs) - 1 else "",
                                      first_run=run, last_run=run))
        return pieces

    @staticmethod
    def _unsafe_span_run_elements(segments):
        """Name non-text run children the span writer would silently discard.

        Plain tabs and line breaks are reproducible; a break with attributes
        (for example a page break) is not. Run properties are cloned.
        Never include document text or field contents in the diagnostic.
        """
        allowed = {f"{{{W_NS}}}rPr", f"{{{W_NS}}}t", f"{{{W_NS}}}tab",
                   f"{{{W_NS}}}br", f"{{{W_NS}}}cr"}
        unsafe = set()
        for segment in segments:
            for run in segment.runs:
                for child in run:
                    if child.tag not in allowed or (
                        child.tag in {f"{{{W_NS}}}tab", f"{{{W_NS}}}br", f"{{{W_NS}}}cr"}
                        and (child.attrib or len(child))
                    ):
                        unsafe.add(etree.QName(child).localname)
        return tuple(sorted(unsafe))

    @staticmethod
    def _word_run_safety_reason(paragraph, segment, old_text, new_text):
        """Return a reason instead of rewriting OOXML we cannot preserve.

        No hyperlinks, fields, drawings, bookmarks, comments, tab/linebreak
        elements or mixed formatting. Those use the established span renderer.
        """
        first, last = segment.first_run, segment.last_run
        if (first is None or last is None or first.getparent() is not paragraph
                or last.getparent() is not paragraph):
            return "non_direct_runs"
        children = list(paragraph)
        try:
            start, end = children.index(first), children.index(last)
        except ValueError:
            return "run_not_in_paragraph"
        if end < start:
            return "run_order"
        runs = children[start:end + 1]
        if any(run.tag != f"{{{W_NS}}}r" for run in runs):
            return "interleaved_markup"
        # Classify disallowed run children without changing eligibility. All
        # categories below still use the existing span fallback. This makes it
        # possible to prioritize future changes using real document evidence.
        allowed = {f"{{{W_NS}}}rPr", f"{{{W_NS}}}t"}
        run_children = [ch for run in runs for ch in run if ch.tag not in allowed]
        if run_children:
            tags = {etree.QName(ch).localname for ch in run_children}
            if tags & {"tab", "br", "cr"}:
                return "complex_run_tabs_or_breaks"
            if tags & {"fldChar", "instrText", "drawing", "pict", "object",
                        "footnoteReference", "endnoteReference", "sym"}:
                return "complex_run_fields_or_objects"
            return "complex_run_other_elements"
        if any(not any(ch.tag == f"{{{W_NS}}}t" for ch in run) for run in runs):
            return "complex_run_without_text"
        properties = [etree.tostring(run.find(f"{{{W_NS}}}rPr"))
                      if run.find(f"{{{W_NS}}}rPr") is not None else None
                      for run in runs]
        if any(value != properties[0] for value in properties[1:]):
            return "mixed_run_formatting"
        source = "".join("".join(ch.text or "" for ch in run
                                 if ch.tag == f"{{{W_NS}}}t") for run in runs)
        if (segment.old != old_text or segment.new != new_text or
                source != segment.before + old_text + segment.after):
            return "non_exact_run_text"
        return None

    @staticmethod
    def _no_common_equivalent_span(paragraph, segment, old_text, new_text, segments):
        """Identify safe replacements whose existing span revision is optimal.

        This is *diagnostics only*: never use this predicate to bypass OOXML
        safety or silently count a visually unchanged revision as improvement.
        """
        import re
        if not segments or [item.kind for item in segments if item.text] != ["delete", "insert"]:
            return False
        # Only one lexical word on each side; longer rewrites are not claimed
        # to have optimal granularity merely because their tokens differ.
        word = r"[^\W_]+(?:-[^\W_]+)*"
        if not re.fullmatch(word, old_text, re.UNICODE) or not re.fullmatch(word, new_text, re.UNICODE):
            return False
        return TrackedChangesRenderer._word_run_safety_reason(
            paragraph, segment, old_text, new_text) is None

    @staticmethod
    def _word_diff_safety_reason(old_text, new_text, segments):
        """A stable, specific reason for rejecting a token-level edit."""
        if old_text == new_text:
            return "diff_unchanged"
        if not segments:
            return "diff_empty"
        if any(seg.kind not in {"equal", "insert", "delete"} for seg in segments):
            return "diff_unsupported_operation"
        if ("".join(seg.text for seg in segments if seg.kind != "insert") != old_text
                or "".join(seg.text for seg in segments if seg.kind != "delete") != new_text):
            return "diff_text_mismatch"
        if not any(seg.kind != "equal" and seg.text for seg in segments):
            return "diff_no_changes"
        if not any(seg.kind == "equal" and seg.text for seg in segments):
            return "diff_no_common_text"
        return None

    @staticmethod
    def _safe_word_segments(old_text, new_text, segments):
        # Backwards-compatible eligibility API used by earlier test cases.
        return TrackedChangesRenderer._word_diff_safety_reason(old_text, new_text, segments) is None

    def _rewrite_segmented_run(self, segment, diff_segments):
        source_run = segment.first_run
        new_elements = []
        if segment.before:
            new_elements.append(self._clone_run_with_text(source_run, segment.before))
        revisions = []
        for diff in diff_segments:
            if not diff.text:
                continue
            if diff.kind == "equal":
                new_elements.append(self._clone_run_with_text(source_run, diff.text))
            elif diff.kind == "delete":
                wrapper = self._make_deleted_wrapper(diff.text, source_run=source_run)
                new_elements.append(wrapper)
                revisions.append((wrapper, "del"))
            else:
                wrapper = self._make_inserted_wrapper(diff.text, source_run=source_run)
                new_elements.append(wrapper)
                revisions.append((wrapper, "ins"))
        if segment.after:
            new_elements.append(self._clone_run_with_text(source_run, segment.after))
        # Build all XML nodes before altering the original paragraph.
        replace_runs(segment, new_elements)
        first, first_kind = revisions[0]
        last, last_kind = revisions[-1]
        return {
            "anchor_kind": first_kind,
            "anchor_revision_id": self._get_revision_id(first),
            "anchor_end_kind": last_kind,
            "anchor_end_revision_id": self._get_revision_id(last),
        }

    def _select_comment_anchor(
        self,
        inserted_wrapper: Optional[etree._Element],
        deleted_wrapper: Optional[etree._Element],
    ) -> tuple[Optional[etree._Element], Optional[str]]:
        if inserted_wrapper is not None:
            return inserted_wrapper, "ins"
        if deleted_wrapper is not None:
            return deleted_wrapper, "del"
        return None, None

    def _get_revision_id(self, element: Optional[etree._Element]) -> Optional[str]:
        if element is None:
            return None
        return element.get(f"{{{W_NS}}}id")

    def _ensure_track_revisions_enabled(self) -> None:
        tree = self.package.read_settings_tree(create_if_missing=True)
        root = tree.getroot()

        track = root.find(f"./{{{W_NS}}}trackRevisions")
        if track is None:
            root.insert(0, etree.Element(f"{{{W_NS}}}trackRevisions"))
            self.package.write_settings_tree(tree)
            self.logger.info("Enabled w:trackRevisions in settings.xml")

    def _next_change_id(self) -> str:
        current = str(self.change_id_counter)
        self.change_id_counter += 1
        return current

    def _current_word_timestamp(self) -> str:
        return datetime.now(timezone.utc).replace(microsecond=0).isoformat()

    def _clone_run_with_text(self, source_run: etree._Element, text: str) -> etree._Element:
        new_run = etree.Element(f"{{{W_NS}}}r")

        rpr = source_run.find(f"./{{{W_NS}}}rPr")
        if rpr is not None:
            new_run.append(etree.fromstring(etree.tostring(rpr)))

        append_visible_text(new_run, text)
        return new_run

    def _make_deleted_wrapper(self, text: str, source_run: Optional[etree._Element] = None) -> etree._Element:
        wrapper = etree.Element(f"{{{W_NS}}}del")
        wrapper.set(f"{{{W_NS}}}id", self._next_change_id())
        wrapper.set(f"{{{W_NS}}}author", self.author)
        wrapper.set(f"{{{W_NS}}}date", self._current_word_timestamp())

        run = etree.SubElement(wrapper, f"{{{W_NS}}}r")

        if source_run is not None:
            rpr = source_run.find(f"./{{{W_NS}}}rPr")
            if rpr is not None:
                run.append(etree.fromstring(etree.tostring(rpr)))

        append_visible_text(run, text, text_tag="delText")
        return wrapper

    def _make_inserted_wrapper(self, text: str, source_run: Optional[etree._Element] = None) -> etree._Element:
        wrapper = etree.Element(f"{{{W_NS}}}ins")
        wrapper.set(f"{{{W_NS}}}id", self._next_change_id())
        wrapper.set(f"{{{W_NS}}}author", self.author)
        wrapper.set(f"{{{W_NS}}}date", self._current_word_timestamp())

        run = etree.SubElement(wrapper, f"{{{W_NS}}}r")

        if source_run is not None:
            rpr = source_run.find(f"./{{{W_NS}}}rPr")
            if rpr is not None:
                run.append(etree.fromstring(etree.tostring(rpr)))

        append_visible_text(run, text)
        return wrapper
