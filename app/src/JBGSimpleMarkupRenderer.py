from dataclasses import dataclass
from typing import Optional

from lxml import etree

try:
    from app.src.JBGDocxPackage import W_NS, DocxPackage
    from app.src.JBGDataRetention import loggable_text
    from app.src.JBGRunEditing import append_visible_text, plan_span_edit, replace_runs
except ModuleNotFoundError:
    from JBGDocxPackage import W_NS, DocxPackage
    from JBGDataRetention import loggable_text
    from JBGRunEditing import append_visible_text, plan_span_edit, replace_runs

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
    anchor_kind: Optional[str] = None
    anchor_revision_id: Optional[str] = None


class SimpleMarkupRenderer:
    """
    Enkel markup-renderer med visuell markering:

    - gammal text: röd + genomstruken
    - ny text: blå

    Stöd:
    - paragraph i word/document.xml
    - table_cell i word/document.xml
    - textbox i word/document.xml
    - footnote i word/footnotes.xml
    - header/footer i den verkliga OOXML-delen
    """

    DOCUMENT_ELEMENT_TYPES = {"paragraph", "table_cell", "textbox"}
    FOOTNOTE_ELEMENT_TYPES = {"footnote"}
    STORY_ELEMENT_TYPES = {"header", "footer"}
    SUPPORTED_ELEMENT_TYPES = (
        DOCUMENT_ELEMENT_TYPES | FOOTNOTE_ELEMENT_TYPES | STORY_ELEMENT_TYPES
    )

    # Klassiska granskningsfärger i Word-liknande stil
    DELETE_COLOR = "FF0000"  # röd
    INSERT_COLOR = "0000FF"  # blå

    def __init__(self, package: DocxPackage, logger):
        self.package = package
        self.logger = logger
        self.document_adapter = DocumentPartAdapter(package, logger)
        self.story_adapter = HeaderFooterPartAdapter(package, logger)

        self.footnotes_adapter = None
        if self.package.part_exists("word/footnotes.xml"):
            self.footnotes_adapter = FootnotesPartAdapter(package, logger)

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
                        "SimpleMarkupRenderer supports only "
                        f"{sorted(self.SUPPORTED_ELEMENT_TYPES)}, got {plan.target.element_type}"
                    ),
                )

            if plan.target.element_type in self.DOCUMENT_ELEMENT_TYPES:
                if plan.target.part_name != "word/document.xml":
                    return RenderResult(
                        plan=plan,
                        applied=False,
                        message=(
                            "SimpleMarkupRenderer expected word/document.xml for "
                            f"{plan.target.element_type}, got {plan.target.part_name}"
                        ),
                    )

                self._apply_document_plan(plan)
                self.package.write_document_tree(self.document_adapter.tree)
                return RenderResult(plan=plan, applied=True, message="Applied simple markup")

            if plan.target.element_type in self.STORY_ELEMENT_TYPES:
                self._apply_story_plan(plan)
                self.story_adapter.write_tree(plan.target.part_name)
                return RenderResult(plan=plan, applied=True, message="Applied simple markup")

            if plan.target.element_type in self.FOOTNOTE_ELEMENT_TYPES:
                if plan.target.part_name != "word/footnotes.xml":
                    return RenderResult(
                        plan=plan,
                        applied=False,
                        message=(
                            "SimpleMarkupRenderer expected word/footnotes.xml for "
                            f"{plan.target.element_type}, got {plan.target.part_name}"
                        ),
                    )

                if self.footnotes_adapter is None:
                    return RenderResult(
                        plan=plan,
                        applied=False,
                        message="Document has no footnotes.xml part",
                    )

                self._apply_footnote_plan(plan)
                self.package.write_footnotes_tree(self.footnotes_adapter.tree)
                return RenderResult(plan=plan, applied=True, message="Applied simple markup")

            return RenderResult(plan=plan, applied=False, message="Unsupported element type")

        except Exception as ex:
            return RenderResult(plan=plan, applied=False, message=str(ex))

    def apply_plans(self, plans: list[ChangePlan]) -> list[RenderResult]:
        results: list[RenderResult] = []
        grouped: dict[tuple, list[ChangePlan]] = {}

        for plan in plans:
            if plan.target.element_type not in self.SUPPORTED_ELEMENT_TYPES:
                results.append(RenderResult(
                    plan=plan,
                    applied=False,
                    message=(
                        "SimpleMarkupRenderer supports only "
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
                    elif plan.target.element_type in self.STORY_ELEMENT_TYPES:
                        self.story_adapter.refresh(plan.target.part_name)

        return results

    # ------------------------------------------------------------------
    # Hög-nivå per elementtyp
    # ------------------------------------------------------------------

    def _apply_document_plan(self, plan: ChangePlan) -> None:
        if plan.target.element_type == "table_cell":
            self._apply_table_cell_plan(plan)
            return

        located = self.document_adapter.locate_plan_nodes(plan)
        model = located["paragraph_model"]
        paragraph = model.paragraph_element
        overlapping_nodes = located["overlapping_nodes"]

        self._apply_simple_change_to_paragraph_element(
            paragraph=paragraph,
            visible_text=model.visible_text,
            overlapping_nodes=overlapping_nodes,
            anchor_start=plan.anchor.start,
            anchor_end=plan.anchor.end,
            old_text=plan.old_text,
            new_text=plan.new_text,
        )

    def _apply_story_plan(self, plan: ChangePlan) -> None:
        located = self.story_adapter.locate_plan_nodes(plan)
        model = located["paragraph_model"]
        self._apply_simple_change_to_paragraph_element(
            paragraph=model.paragraph_element,
            visible_text=model.visible_text,
            overlapping_nodes=located["overlapping_nodes"],
            anchor_start=plan.anchor.start,
            anchor_end=plan.anchor.end,
            old_text=plan.old_text,
            new_text=plan.new_text,
        )

    def _apply_table_cell_plan(self, plan: ChangePlan) -> None:
        try:
            located = self.document_adapter.locate_plan_nodes(plan)
            model = located["paragraph_model"]
            paragraph = model.paragraph_element
            overlapping_nodes = located["overlapping_nodes"]

            self._apply_simple_change_to_paragraph_element(
                paragraph=paragraph,
                visible_text=model.visible_text,
                overlapping_nodes=overlapping_nodes,
                anchor_start=plan.anchor.start,
                anchor_end=plan.anchor.end,
                old_text=plan.old_text,
                new_text=plan.new_text,
            )
            return
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

            self._apply_simple_change_to_paragraph_element(
                paragraph=paragraph,
                visible_text=paragraph_model["visible_text"],
                overlapping_nodes=overlapping_nodes,
                anchor_start=anchor["start"],
                anchor_end=anchor["end"],
                old_text=plan.old_text,
                new_text=plan.new_text,
            )
            return

        raise ValueError(f"No matching paragraph found in table cell {plan.target.element_id}")

    def _apply_footnote_plan(self, plan: ChangePlan) -> None:
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

        self._apply_simple_change_to_paragraph_element(
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

    def _apply_simple_change_to_paragraph_element(
        self,
        paragraph: etree._Element,
        visible_text: str,
        overlapping_nodes: list,
        anchor_start: int,
        anchor_end: int,
        old_text: str,
        new_text: str,
    ) -> None:
        matched = visible_text[anchor_start:anchor_end]
        if not self._texts_match_leniently(matched, old_text):
            raise ValueError(
                f"Anchor text mismatch. Expected old_text={loggable_text(old_text)}, found={loggable_text(matched)}"
            )

        first_node = overlapping_nodes[0]
        last_node = overlapping_nodes[-1]

        if first_node.run_element is None or last_node.run_element is None:
            raise ValueError("Anchor overlaps node(s) without run_element")

        self._rewrite_span(
            paragraph=paragraph,
            first_node=first_node,
            last_node=last_node,
            anchor_start=anchor_start,
            anchor_end=anchor_end,
            old_text=old_text,
            new_text=new_text,
        )

    # ------------------------------------------------------------------
    # Boundary guard
    # ------------------------------------------------------------------

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
    ) -> None:
        """
        Markerar ändringen av styckets text [anchor_start, anchor_end):
        borttagen text rödstruken, ny text grön. Fungerar för en eller flera
        textbitar och för ändringar som går in i en hyperlänk (se
        JBGRunEditing). Text som inte ändras behålls, med radbrytningar och tabbar.
        """
        plan = plan_span_edit(paragraph, first_node, last_node, anchor_start, anchor_end, new_text)
        if not self._texts_match_leniently(plan.actual_old, old_text):
            raise ValueError(
                f"Span text mismatch. Expected {loggable_text(old_text)}, got {loggable_text(plan.actual_old)}"
            )

        for segment in plan.segments:
            new_elements = []
            source = segment.first_run
            if segment.before:
                new_elements.append(self._clone_run_with_text(source, segment.before))
            if segment.old:
                new_elements.append(self._clone_run_with_deleted_markup(source, segment.old))
            if segment.old and segment.new:
                separator = self._make_separator_text(segment.old, segment.new)
                if separator:
                    new_elements.append(self._clone_run_with_neutral_separator(source, separator))
            if segment.new:
                new_elements.append(self._clone_run_with_inserted_markup(source, segment.new))
            if segment.new and segment.after:
                separator = self._make_separator_text(segment.new, segment.after)
                if separator:
                    new_elements.append(self._clone_run_with_neutral_separator(segment.last_run, separator))
            if segment.after:
                new_elements.append(self._clone_run_with_text(segment.last_run, segment.after))
            replace_runs(segment, new_elements)

    # ------------------------------------------------------------------
    # Footnote helpers
    # ------------------------------------------------------------------

    def _build_footnote_paragraph_models(self, footnote_model) -> list[dict]:
        models = []

        paragraphs = footnote_model.paragraphs
        for paragraph in paragraphs:
            paragraph_nodes = [n for n in footnote_model.nodes if n.paragraph_element is paragraph]

            editable_text_parts = []
            editable_nodes = []
            seen_editable_text = False
            editable_cursor = 0

            for node in paragraph_nodes:
                if node.kind == "footnote_ref" and not seen_editable_text:
                    continue

                if node.kind == "text":
                    seen_editable_text = True
                    editable_nodes.append(self._clone_node_with_new_span(
                        node,
                        editable_cursor,
                        editable_cursor + len(node.text),
                    ))
                    editable_text_parts.append(node.text)
                    editable_cursor += len(node.text)

                elif node.kind in {"tab", "linebreak"}:
                    if seen_editable_text:
                        editable_nodes.append(self._clone_node_with_new_span(
                            node,
                            editable_cursor,
                            editable_cursor + len(node.text),
                        ))
                        editable_text_parts.append(node.text)
                        editable_cursor += len(node.text)

            models.append({
                "paragraph": paragraph,
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

    def _get_overlapping_nodes_for_generic_paragraph(
        self,
        nodes: list,
        anchor_start: int,
        anchor_end: int,
    ) -> list:
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
    # Run factories
    # ------------------------------------------------------------------

    def _clone_run_with_text(self, source_run: etree._Element, text: str) -> etree._Element:
        new_run = etree.Element(f"{{{W_NS}}}r")

        rpr = source_run.find(f"./{{{W_NS}}}rPr")
        if rpr is not None:
            new_run.append(etree.fromstring(etree.tostring(rpr)))

        append_visible_text(new_run, text)
        return new_run

    def _clone_run_with_deleted_markup(self, source_run: etree._Element, text: str) -> etree._Element:
        new_run = etree.Element(f"{{{W_NS}}}r")

        rpr = self._copy_or_create_rpr(source_run)

        strike = etree.SubElement(rpr, f"{{{W_NS}}}strike")
        strike.set(f"{{{W_NS}}}val", "true")

        color = etree.SubElement(rpr, f"{{{W_NS}}}color")
        color.set(f"{{{W_NS}}}val", self.DELETE_COLOR)

        new_run.append(rpr)

        append_visible_text(new_run, text)
        return new_run

    def _clone_run_with_inserted_markup(self, source_run: etree._Element, text: str) -> etree._Element:
        new_run = etree.Element(f"{{{W_NS}}}r")

        rpr = self._copy_or_create_rpr(source_run)

        color = etree.SubElement(rpr, f"{{{W_NS}}}color")
        color.set(f"{{{W_NS}}}val", self.INSERT_COLOR)

        new_run.append(rpr)

        append_visible_text(new_run, text)
        return new_run

    def _copy_or_create_rpr(self, source_run: etree._Element) -> etree._Element:
        rpr = source_run.find(f"./{{{W_NS}}}rPr")
        if rpr is not None:
            return etree.fromstring(etree.tostring(rpr))
        return etree.Element(f"{{{W_NS}}}rPr")
    

# ------------------------------------------------------------------
# Hjälprutiner
# ------------------------------------------------------------------

    def _is_word_char(self, ch: str) -> bool:
        return ch.isalnum() or ch == "_"

    def _needs_space_between(self, left: str, right: str) -> bool:
        if not left or not right:
            return False

        left_char = left[-1]
        right_char = right[0]

        if left_char.isspace() or right_char.isspace():
            return False

        if self._is_word_char(left_char) and self._is_word_char(right_char):
            return True

        if left_char in ".:;!?" and right_char.isalpha():
            return True

        return False


    def _make_separator_text(self, left: str, right: str) -> str:
        return " " if self._needs_space_between(left, right) else ""


    def _clone_run_with_neutral_separator(self, source_run: etree._Element, text: str) -> etree._Element:
        """
        Skapar en neutral separator-run utan markupfärg/strikethrough.
        Behåller grundformattering men utan granskningsmarkering.
        """
        new_run = etree.Element(f"{{{W_NS}}}r")

        rpr = source_run.find(f"./{{{W_NS}}}rPr")
        if rpr is not None:
            clean_rpr = etree.fromstring(etree.tostring(rpr))

            # ta bort ev. color/strike/highlight om sådant redan finns
            for child in list(clean_rpr):
                if child.tag in {
                    f"{{{W_NS}}}color",
                    f"{{{W_NS}}}strike",
                    f"{{{W_NS}}}dstrike",
                    f"{{{W_NS}}}highlight",
                    f"{{{W_NS}}}u",
                }:
                    clean_rpr.remove(child)

            new_run.append(clean_rpr)

        append_visible_text(new_run, text)
        return new_run
