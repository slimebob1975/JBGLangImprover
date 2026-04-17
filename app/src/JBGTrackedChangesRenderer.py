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
    from app.src.JBGChangePlanner import ChangePlan
except ModuleNotFoundError:
    from JBGChangePlanner import ChangePlan


XML_NS = "http://www.w3.org/XML/1998/namespace"


@dataclass
class RenderResult:
    plan: ChangePlan
    applied: bool
    message: str


class TrackedChangesRenderer:
    """
    Tracked changes renderer v2.2

    Stöd:
    - paragraph i word/document.xml
    - table_cell i word/document.xml
    - run-splitting på textnoder
    - riktiga <w:del> och <w:ins>
    - trackRevisions i settings.xml

    Inte stöd ännu:
    - footnotes
    - textbox
    - comments
    - headers/footers
    """

    SUPPORTED_ELEMENT_TYPES = {"paragraph", "table_cell"}

    def __init__(self, package: DocxPackage, logger, author: str = "JBG Klarspråkningstjänst"):
        self.package = package
        self.logger = logger
        self.author = author
        self.document_adapter = DocumentPartAdapter(package, logger)
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
                        "TrackedChangesRenderer v2.2 supports only "
                        f"{sorted(self.SUPPORTED_ELEMENT_TYPES)}, got {plan.target.element_type}"
                    ),
                )

            if plan.target.part_name != "word/document.xml":
                return RenderResult(
                    plan=plan,
                    applied=False,
                    message=f"TrackedChangesRenderer v2.2 supports only word/document.xml, got {plan.target.part_name}",
                )

            self._apply_document_plan(plan)
            self.package.write_document_tree(self.document_adapter.tree)

            return RenderResult(plan=plan, applied=True, message="Applied tracked changes")

        except Exception as ex:
            return RenderResult(plan=plan, applied=False, message=str(ex))

    def apply_plans(self, plans: list[ChangePlan]) -> list[RenderResult]:
        results = []

        grouped: dict[str, list[ChangePlan]] = {}

        for plan in plans:
            if plan.target.element_type not in self.SUPPORTED_ELEMENT_TYPES:
                results.append(RenderResult(
                    plan=plan,
                    applied=False,
                    message=(
                        "TrackedChangesRenderer v2.2 supports only "
                        f"{sorted(self.SUPPORTED_ELEMENT_TYPES)}, got {plan.target.element_type}"
                    ),
                ))
                continue

            key = plan.target.element_id
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
                    self.document_adapter.refresh()

        return results

    # ------------------------------------------------------------------
    # Core rendering
    # ------------------------------------------------------------------

    def _apply_document_plan(self, plan: ChangePlan) -> None:
        located = self.document_adapter.locate_plan_nodes(plan)
        model = located["paragraph_model"]
        paragraph = model.paragraph_element
        overlapping_nodes: list[ParagraphNode] = located["overlapping_nodes"]

        matched = model.visible_text[plan.anchor.start:plan.anchor.end]
        if matched != plan.old_text and matched.strip() != plan.old_text.strip():
            raise ValueError(
                f"Anchor text mismatch. Expected old_text={plan.old_text!r}, found={matched!r}"
            )

        first_node = overlapping_nodes[0]
        last_node = overlapping_nodes[-1]

        if first_node.run_element is None or last_node.run_element is None:
            raise ValueError("Anchor overlaps node(s) without run_element")

        first_run = first_node.run_element
        last_run = last_node.run_element

        if first_run is last_run:
            self._rewrite_single_run_case(
                paragraph=paragraph,
                run=first_run,
                node=first_node,
                anchor_start=plan.anchor.start,
                anchor_end=plan.anchor.end,
                old_text=plan.old_text,
                new_text=plan.new_text,
            )
            return

        self._rewrite_multi_run_case(
            paragraph=paragraph,
            overlapping_nodes=overlapping_nodes,
            anchor_start=plan.anchor.start,
            anchor_end=plan.anchor.end,
            old_text=plan.old_text,
            new_text=plan.new_text,
        )

    # ------------------------------------------------------------------
    # Single-run case
    # ------------------------------------------------------------------

    def _rewrite_single_run_case(
        self,
        paragraph: etree._Element,
        run: etree._Element,
        node: ParagraphNode,
        anchor_start: int,
        anchor_end: int,
        old_text: str,
        new_text: str,
    ) -> None:
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

        if middle_text != old_text and middle_text.strip() != old_text.strip():
            raise ValueError(
                f"Single-run split mismatch. Expected {old_text!r}, got {middle_text!r}"
            )

        insert_index = paragraph.index(run)
        paragraph.remove(run)

        new_elements = []

        if before_text:
            new_elements.append(self._clone_run_with_text(run, before_text))

        if old_text:
            new_elements.append(self._make_deleted_wrapper(old_text, source_run=run))

        if new_text:
            new_elements.append(self._make_inserted_wrapper(new_text, source_run=run))

        if after_text:
            new_elements.append(self._clone_run_with_text(run, after_text))

        for offset, elem in enumerate(new_elements):
            paragraph.insert(insert_index + offset, elem)

    # ------------------------------------------------------------------
    # Multi-run case
    # ------------------------------------------------------------------

    def _rewrite_multi_run_case(
        self,
        paragraph: etree._Element,
        overlapping_nodes: list[ParagraphNode],
        anchor_start: int,
        anchor_end: int,
        old_text: str,
        new_text: str,
    ) -> None:
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

        if actual_old != old_text and actual_old.strip() != old_text.strip():
            if old_text.startswith(actual_old) and len(old_text) - len(actual_old) <= 2:
                actual_old = old_text
            else:
                run_dump = [self._extract_visible_text_from_run(r) for r in changed_runs]
                raise ValueError(
                    f"Multi-run split mismatch. Expected {old_text!r}, got {actual_old!r}. Runs={run_dump!r}"
                )

        insert_index = paragraph.index(first_run)

        for run in changed_runs:
            paragraph.remove(run)

        new_elements = []

        if first_before:
            new_elements.append(self._clone_run_with_text(first_run, first_before))

        if old_text:
            new_elements.append(self._make_deleted_wrapper(old_text, source_run=first_run))

        if new_text:
            new_elements.append(self._make_inserted_wrapper(new_text, source_run=first_run))

        if last_after:
            new_elements.append(self._clone_run_with_text(last_run, last_after))

        for offset, elem in enumerate(new_elements):
            paragraph.insert(insert_index + offset, elem)

    # ------------------------------------------------------------------
    # Boundary helpers
    # ------------------------------------------------------------------

    def _find_nearest_text_node_forward(
        self,
        nodes: list[ParagraphNode],
        start_index: int,
    ) -> Optional[ParagraphNode]:
        for i in range(start_index, len(nodes)):
            if nodes[i].kind == "text":
                return nodes[i]
        return None

    def _find_nearest_text_node_backward(
        self,
        nodes: list[ParagraphNode],
        start_index: int,
    ) -> Optional[ParagraphNode]:
        for i in range(start_index, -1, -1):
            if nodes[i].kind == "text":
                return nodes[i]
        return None

    # ------------------------------------------------------------------
    # Reconstruction helpers
    # ------------------------------------------------------------------

    def _reconstruct_old_text_across_runs(
        self,
        changed_runs: list[etree._Element],
        first_boundary_node: ParagraphNode,
        last_boundary_node: ParagraphNode,
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

    def _get_runs_between(
        self,
        paragraph: etree._Element,
        first_run: etree._Element,
        last_run: etree._Element,
    ) -> list[etree._Element]:
        runs = paragraph.findall(f"./{{{W_NS}}}r")
        if first_run not in runs or last_run not in runs:
            return []

        start_idx = runs.index(first_run)
        end_idx = runs.index(last_run)

        if start_idx > end_idx:
            return []

        return runs[start_idx:end_idx + 1]

    # ------------------------------------------------------------------
    # Settings / document infrastructure
    # ------------------------------------------------------------------

    def _ensure_track_revisions_enabled(self) -> None:
        tree = self.package.read_settings_tree(create_if_missing=True)
        root = tree.getroot()

        track = root.find(f"./{{{W_NS}}}trackRevisions")
        if track is None:
            root.insert(0, etree.Element(f"{{{W_NS}}}trackRevisions"))
            self.package.write_settings_tree(tree)
            self.logger.info("Enabled w:trackRevisions in settings.xml")

    # ------------------------------------------------------------------
    # Element factories
    # ------------------------------------------------------------------

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