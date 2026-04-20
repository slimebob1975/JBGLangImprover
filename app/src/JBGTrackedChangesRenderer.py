from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Optional

from lxml import etree

try:
    from app.src.JBGDocxPackage import W_NS, DocxPackage
except ModuleNotFoundError:
    from JBGDocxPackage import W_NS, DocxPackage

try:
    from app.src.JBGDocumentPartAdapter import DocumentPartAdapter, ParagraphNode
except ModuleNotFoundError:
    from JBGDocumentPartAdapter import DocumentPartAdapter, ParagraphNode

try:
    from app.src.JBGFootnotesPartAdapter import FootnotesPartAdapter, FootnoteNode
except ModuleNotFoundError:
    from JBGFootnotesPartAdapter import FootnotesPartAdapter, FootnoteNode

try:
    from app.src.JBGChangePlanner import ChangePlan
except ModuleNotFoundError:
    from JBGChangePlanner import ChangePlan


XML_NS = "http://www.w3.org/XML/1998/namespace"


@dataclass
class RenderResult:
    plan: ChangePlan
    applied: bool
    message: str
    anchor_part_name: Optional[str] = None
    anchor_kind: Optional[str] = None   # "ins" / "del"
    anchor_revision_id: Optional[str] = None


class TrackedChangesRenderer:
    DOCUMENT_ELEMENT_TYPES = {"paragraph", "table_cell", "textbox"}
    FOOTNOTE_ELEMENT_TYPES = {"footnote"}
    SUPPORTED_ELEMENT_TYPES = DOCUMENT_ELEMENT_TYPES | FOOTNOTE_ELEMENT_TYPES

    def __init__(self, package: DocxPackage, logger, author: str = "JBG Klarspråkningstjänst"):
        self.package = package
        self.logger = logger
        self.author = author
        self.document_adapter = DocumentPartAdapter(package, logger)

        self.footnotes_adapter = None
        if self.package.part_exists("word/footnotes.xml"):
            self.footnotes_adapter = FootnotesPartAdapter(package, logger)

        self.change_id_counter = 1
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
                )

            return RenderResult(plan=plan, applied=False, message="Unsupported element type")

        except Exception as ex:
            return RenderResult(plan=plan, applied=False, message=str(ex))

    def apply_plans(self, plans: list[ChangePlan]) -> list[RenderResult]:
        results = []

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
                result = self.apply_plan(plan)
                results.append(result)

                if result.applied:
                    if plan.target.part_name == "word/document.xml":
                        self.document_adapter.refresh()
                    elif plan.target.part_name == "word/footnotes.xml" and self.footnotes_adapter is not None:
                        self.footnotes_adapter.refresh()

        return results

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
    ) -> dict:
        matched = visible_text[anchor_start:anchor_end]
        if not self._texts_match_leniently(matched, old_text):
            raise ValueError(
                f"Anchor text mismatch. Expected old_text={old_text!r}, found={matched!r}"
            )

        first_node = overlapping_nodes[0]
        last_node = overlapping_nodes[-1]

        if first_node.run_element is None or last_node.run_element is None:
            raise ValueError("Anchor overlaps node(s) without run_element")

        first_run = first_node.run_element
        last_run = last_node.run_element

        if first_run is last_run:
            return self._rewrite_single_run_case(
                paragraph=paragraph,
                run=first_run,
                node=first_node,
                anchor_start=anchor_start,
                anchor_end=anchor_end,
                old_text=old_text,
                new_text=new_text,
            )

        return self._rewrite_multi_run_case(
            paragraph=paragraph,
            overlapping_nodes=overlapping_nodes,
            anchor_start=anchor_start,
            anchor_end=anchor_end,
            old_text=old_text,
            new_text=new_text,
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
        if len(parts) != 5 or parts[0] != "table" or parts[2] != "cell":
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

    def _rewrite_single_run_case(
        self,
        paragraph: etree._Element,
        run: etree._Element,
        node,
        anchor_start: int,
        anchor_end: int,
        old_text: str,
        new_text: str,
    ) -> dict:
        if node.kind != "text":
            raise ValueError("Single-run case currently supports only text nodes")

        node_text = node.text
        local_start = anchor_start - node.start
        local_end = anchor_end - node.start

        if local_start < 0 or local_end > len(node_text):
            raise ValueError("Invalid local offsets in single-run case")

        before_text = node_text[:local_start]
        middle_text = node_text[local_start:local_end]
        after_text = node_text[local_end:]

        if not self._texts_match_leniently(middle_text, old_text):
            raise ValueError(
                f"Single-run split mismatch. Expected {old_text!r}, got {middle_text!r}"
            )

        insert_index = paragraph.index(run)
        paragraph.remove(run)

        new_elements = []
        inserted_wrapper = None
        deleted_wrapper = None

        if before_text:
            new_elements.append(self._clone_run_with_text(run, before_text))

        if old_text:
            deleted_wrapper = self._make_deleted_wrapper(old_text, source_run=run)
            new_elements.append(deleted_wrapper)

        if new_text:
            inserted_wrapper = self._make_inserted_wrapper(new_text, source_run=run)
            new_elements.append(inserted_wrapper)

        if after_text:
            new_elements.append(self._clone_run_with_text(run, after_text))

        for offset, elem in enumerate(new_elements):
            paragraph.insert(insert_index + offset, elem)

        anchor_element, anchor_kind = self._select_comment_anchor(inserted_wrapper, deleted_wrapper)
        return {
            "anchor_kind": anchor_kind,
            "anchor_revision_id": self._get_revision_id(anchor_element),
        }

    # ------------------------------------------------------------------
    # Multi-run case
    # ------------------------------------------------------------------

    def _rewrite_multi_run_case(
        self,
        paragraph: etree._Element,
        overlapping_nodes: list,
        anchor_start: int,
        anchor_end: int,
        old_text: str,
        new_text: str,
    ) -> dict:
        first_text_node = self._find_nearest_text_node_forward(overlapping_nodes, 0)
        last_text_node = self._find_nearest_text_node_backward(overlapping_nodes, len(overlapping_nodes) - 1)

        if first_text_node is None or last_text_node is None:
            raise ValueError("Multi-run case could not find text boundary nodes")

        first_node = first_text_node
        last_node = last_text_node

        if first_node.run_element is None or last_node.run_element is None:
            raise ValueError("Resolved text boundary node(s) without run_element")

        first_run = first_node.run_element
        last_run = last_node.run_element

        first_local_start = anchor_start - first_node.start
        last_local_end = anchor_end - last_node.start

        if first_local_start < 0 or first_local_start > len(first_node.text):
            raise ValueError("Invalid first_local_start")
        if last_local_end < 0 or last_local_end > len(last_node.text):
            raise ValueError("Invalid last_local_end")

        first_before = first_node.text[:first_local_start]
        last_after = last_node.text[last_local_end:]

        changed_runs = self._get_runs_between(paragraph, first_run, last_run)
        if not changed_runs:
            raise ValueError("Could not resolve changed run span")

        actual_old = self._reconstruct_old_text_across_runs(
            changed_runs=changed_runs,
            first_boundary_node=first_node,
            last_boundary_node=last_node,
            first_local_start=first_local_start,
            last_local_end=last_local_end,
        )

        if not self._texts_match_leniently(actual_old, old_text):
            run_dump = [self._extract_visible_text_from_run(r) for r in changed_runs]
            raise ValueError(
                f"Multi-run split mismatch. Expected {old_text!r}, got {actual_old!r}. Runs={run_dump!r}"
            )

        insert_index = paragraph.index(first_run)

        for run in changed_runs:
            paragraph.remove(run)

        new_elements = []
        inserted_wrapper = None
        deleted_wrapper = None

        if first_before:
            new_elements.append(self._clone_run_with_text(first_run, first_before))

        if old_text:
            deleted_wrapper = self._make_deleted_wrapper(old_text, source_run=first_run)
            new_elements.append(deleted_wrapper)

        if new_text:
            inserted_wrapper = self._make_inserted_wrapper(new_text, source_run=first_run)
            new_elements.append(inserted_wrapper)

        if last_after:
            new_elements.append(self._clone_run_with_text(last_run, last_after))

        for offset, elem in enumerate(new_elements):
            paragraph.insert(insert_index + offset, elem)

        anchor_element, anchor_kind = self._select_comment_anchor(inserted_wrapper, deleted_wrapper)
        return {
            "anchor_kind": anchor_kind,
            "anchor_revision_id": self._get_revision_id(anchor_element),
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

    def _find_nearest_text_node_forward(self, nodes: list, start_index: int):
        for i in range(start_index, len(nodes)):
            if nodes[i].kind == "text":
                return nodes[i]
        return None

    def _find_nearest_text_node_backward(self, nodes: list, start_index: int):
        for i in range(start_index, -1, -1):
            if nodes[i].kind == "text":
                return nodes[i]
        return None

    def _reconstruct_old_text_across_runs(
        self,
        changed_runs: list[etree._Element],
        first_boundary_node,
        last_boundary_node,
        first_local_start: int,
        last_local_end: int,
    ) -> str:
        parts = []

        for run in changed_runs:
            run_text = self._extract_visible_text_from_run(run)

            if run is first_boundary_node.run_element and run is last_boundary_node.run_element:
                parts.append(run_text[first_local_start:last_local_end])
            elif run is first_boundary_node.run_element:
                parts.append(run_text[first_local_start:])
            elif run is last_boundary_node.run_element:
                parts.append(run_text[:last_local_end])
            else:
                parts.append(run_text)

        return "".join(parts)

    def _extract_visible_text_from_run(self, run: etree._Element) -> str:
        parts = []
        for child in run:
            if child.tag == f"{{{W_NS}}}t":
                parts.append(child.text or "")
            elif child.tag == f"{{{W_NS}}}tab":
                parts.append("\t")
            elif child.tag in {f"{{{W_NS}}}br", f"{{{W_NS}}}cr"}:
                parts.append("\n")
        return "".join(parts)

    def _get_runs_between(self, paragraph: etree._Element, first_run: etree._Element, last_run: etree._Element) -> list[etree._Element]:
        runs = paragraph.findall(f"./{{{W_NS}}}r")
        if first_run not in runs or last_run not in runs:
            return []

        start_idx = runs.index(first_run)
        end_idx = runs.index(last_run)

        if start_idx > end_idx:
            return []

        return runs[start_idx:end_idx + 1]

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

        t = etree.SubElement(new_run, f"{{{W_NS}}}t")
        t.set(f"{{{XML_NS}}}space", "preserve")
        t.text = text
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

        del_text = etree.SubElement(run, f"{{{W_NS}}}delText")
        del_text.set(f"{{{XML_NS}}}space", "preserve")
        del_text.text = text
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

        t = etree.SubElement(run, f"{{{W_NS}}}t")
        t.set(f"{{{XML_NS}}}space", "preserve")
        t.text = text
        return wrapper