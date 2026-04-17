from dataclasses import dataclass, field
from typing import Optional, Literal

from lxml import etree

try:
    from app.src.JBGDocxPackage import W_NS, NSMAP, DocxPackage
except ModuleNotFoundError:
    from JBGDocxPackage import W_NS, NSMAP, DocxPackage

try:
    from app.src.JBGChangePlanner import ChangePlan
except ModuleNotFoundError:
    from JBGChangePlanner import ChangePlan


# ============================================================================
# Datamodeller för document.xml
# ============================================================================

@dataclass
class ParagraphNode:
    """
    En nod i ett stycke, antingen textbärande eller specialnod.
    """
    kind: Literal[
        "text",
        "tab",
        "linebreak",
        "footnote_reference",
        "comment_reference",
        "drawing",
        "field_char",
        "instr_text",
        "other"
    ]
    text: str
    element: etree._Element
    run_element: Optional[etree._Element]
    start: int
    end: int
    meta: dict = field(default_factory=dict)


@dataclass
class ParagraphModel:
    """
    Modell av ett stycke i document.xml, med renderad synlig text
    och en sekvens av text/specialnoder.
    """
    paragraph_element: etree._Element
    nodes: list[ParagraphNode]
    visible_text: str
    container_path: Optional[str] = None


# ============================================================================
# Adapter
# ============================================================================

class DocumentPartAdapter:
    """
    Adapter för word/document.xml.

    Första versionens ansvar:
    - läsa document.xml via DocxPackage
    - hitta paragraph/table_cell/textbox i huvuddel
    - bygga ParagraphModel
    - lokalisera planens textankare mot paragraph-noder
    - exponera säkra hjälpmetoder för senare renderers

    Den här versionen modifierar inte XML ännu.
    """

    def __init__(self, package: DocxPackage, logger):
        self.package = package
        self.logger = logger
        self.tree = self.package.read_document_tree()
        self.root = self.tree.getroot()

    # ------------------------------------------------------------------
    # Publika metoder
    # ------------------------------------------------------------------

    def refresh(self) -> None:
        self.tree = self.package.read_document_tree()
        self.root = self.tree.getroot()

    def get_paragraph_model_for_plan(self, plan: ChangePlan) -> ParagraphModel:
        """
        Hitta rätt paragraph-liknande container för en plan och bygg en modell.
        Stöd i första versionen:
        - paragraph
        - header/footer hanteras senare i egen adapter
        - table_cell
        - textbox
        """
        target = plan.target

        if target.element_type == "paragraph":
            paragraph = self._find_main_document_paragraph(plan)
            return self._build_paragraph_model(paragraph, target_path=f"element:{target.element_id}")

        if target.element_type == "table_cell":
            paragraph = self._find_table_cell_paragraph(plan)
            return self._build_paragraph_model(paragraph, target_path=f"element:{target.element_id}")

        if target.element_type == "textbox":
            paragraph = self._find_textbox_paragraph(plan)
            return self._build_paragraph_model(paragraph, target_path=f"element:{target.element_id}")

        raise ValueError(
            f"DocumentPartAdapter does not support element_type={target.element_type}"
        )

    def locate_plan_nodes(self, plan: ChangePlan) -> dict:
        """
        Returnerar nodintervall för planens ankare inom paragraph-modellen.
        Används senare av renderers för kirurgisk XML-redigering.
        """
        model = self.get_paragraph_model_for_plan(plan)
        anchor_start = plan.anchor.start
        anchor_end = plan.anchor.end

        overlapping_nodes = []
        for node in model.nodes:
            if node.kind not in {"text", "tab", "linebreak"}:
                continue
            if max(node.start, anchor_start) < min(node.end, anchor_end):
                overlapping_nodes.append(node)

        if not overlapping_nodes:
            raise ValueError(
                f"No overlapping text nodes found for anchor in {plan.target.element_id}"
            )

        return {
            "paragraph_model": model,
            "anchor_start": anchor_start,
            "anchor_end": anchor_end,
            "overlapping_nodes": overlapping_nodes,
            "first_node": overlapping_nodes[0],
            "last_node": overlapping_nodes[-1],
        }

    def debug_plan_location(self, plan: ChangePlan) -> dict:
        """
        Hjälpmetod för test/debug.
        """
        located = self.locate_plan_nodes(plan)
        model = located["paragraph_model"]

        return {
            "element_id": plan.target.element_id,
            "element_type": plan.target.element_type,
            "visible_text": model.visible_text,
            "anchor": {
                "start": plan.anchor.start,
                "end": plan.anchor.end,
                "matched_text": plan.anchor.matched_text,
            },
            "overlapping_nodes": [
                {
                    "kind": n.kind,
                    "text": n.text,
                    "start": n.start,
                    "end": n.end,
                }
                for n in located["overlapping_nodes"]
            ],
        }

    # ------------------------------------------------------------------
    # Hitta målcontainer
    # ------------------------------------------------------------------

    def _find_main_document_paragraph(self, plan: ChangePlan) -> etree._Element:
        """
        Förväntar sig element_id enligt paragraph_N.
        """
        element_id = plan.target.element_id or ""
        try:
            index = int(element_id.split("_")[1])
        except Exception as ex:
            raise ValueError(f"Invalid paragraph element_id: {element_id}") from ex

        paragraphs = self.root.xpath("//w:body/w:p", namespaces=NSMAP)
        if index < 1 or index > len(paragraphs):
            raise ValueError(f"Paragraph index out of range: {index}")

        return paragraphs[index - 1]

    def _find_table_cell_paragraph(self, plan: ChangePlan) -> etree._Element:
        """
        Första versionen använder första stycket i tabellcellen.
        element_id-format: table_T_cell_R_C
        """
        element_id = plan.target.element_id or ""
        parts = element_id.split("_")
        if len(parts) != 5 or parts[0] != "table" or parts[2] != "cell":
            raise ValueError(f"Invalid table_cell element_id: {element_id}")

        try:
            table_index = int(parts[1])
            row_index = int(parts[3])
            col_index = int(parts[4])
        except Exception as ex:
            raise ValueError(f"Invalid table indices in {element_id}") from ex

        tables = self.root.xpath("//w:body/w:tbl", namespaces=NSMAP)
        if table_index < 1 or table_index > len(tables):
            raise ValueError(f"Table index out of range: {table_index}")

        table = tables[table_index - 1]
        rows = table.findall(f"{{{W_NS}}}tr")
        if row_index < 1 or row_index > len(rows):
            raise ValueError(f"Row index out of range: {row_index}")

        row = rows[row_index - 1]
        cells = row.findall(f"{{{W_NS}}}tc")
        if col_index < 1 or col_index > len(cells):
            raise ValueError(f"Column index out of range: {col_index}")

        cell = cells[col_index - 1]
        paragraphs = cell.findall(f".//{{{W_NS}}}p")
        if not paragraphs:
            raise ValueError(f"No paragraph found in table cell {element_id}")

        return paragraphs[0]

    def _find_textbox_paragraph(self, plan: ChangePlan) -> etree._Element:
        """
        Första versionen använder första paragraph i textboxens txbxContent.
        element_id-format: textbox_N
        """
        element_id = plan.target.element_id or ""
        try:
            textbox_index = int(element_id.split("_")[1])
        except Exception as ex:
            raise ValueError(f"Invalid textbox element_id: {element_id}") from ex

        txbx_contents = self.root.xpath("//w:txbxContent", namespaces=NSMAP)
        if textbox_index < 1 or textbox_index > len(txbx_contents):
            raise ValueError(f"Textbox index out of range: {textbox_index}")

        txbx = txbx_contents[textbox_index - 1]
        paragraphs = txbx.findall(f".//{{{W_NS}}}p")
        if not paragraphs:
            raise ValueError(f"No paragraph found in textbox {element_id}")

        return paragraphs[0]

    # ------------------------------------------------------------------
    # Bygg paragraph model
    # ------------------------------------------------------------------

    def _build_paragraph_model(self, paragraph_element: etree._Element, target_path: Optional[str] = None) -> ParagraphModel:
        nodes: list[ParagraphNode] = []
        visible_parts: list[str] = []
        cursor = 0

        for run in paragraph_element.findall(f".//{{{W_NS}}}r"):
            run_nodes, cursor = self._extract_run_nodes(run, cursor)
            nodes.extend(run_nodes)
            for n in run_nodes:
                if n.kind in {"text", "tab", "linebreak"}:
                    visible_parts.append(n.text)

        visible_text = "".join(visible_parts)

        return ParagraphModel(
            paragraph_element=paragraph_element,
            nodes=nodes,
            visible_text=visible_text,
            container_path=target_path,
        )

    def _extract_run_nodes(self, run: etree._Element, cursor: int) -> tuple[list[ParagraphNode], int]:
        nodes: list[ParagraphNode] = []

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
                cursor += 1

            elif child.tag == f"{{{W_NS}}}footnoteReference":
                node = ParagraphNode(
                    kind="footnote_reference",
                    text="",
                    element=child,
                    run_element=run,
                    start=cursor,
                    end=cursor,
                    meta={"footnote_id": child.get(f"{{{W_NS}}}id")},
                )
                nodes.append(node)

            elif child.tag == f"{{{W_NS}}}commentReference":
                node = ParagraphNode(
                    kind="comment_reference",
                    text="",
                    element=child,
                    run_element=run,
                    start=cursor,
                    end=cursor,
                    meta={"comment_id": child.get(f"{{{W_NS}}}id")},
                )
                nodes.append(node)

            elif child.tag == f"{{{W_NS}}}drawing":
                node = ParagraphNode(
                    kind="drawing",
                    text="",
                    element=child,
                    run_element=run,
                    start=cursor,
                    end=cursor,
                )
                nodes.append(node)

            elif child.tag == f"{{{W_NS}}}fldChar":
                node = ParagraphNode(
                    kind="field_char",
                    text="",
                    element=child,
                    run_element=run,
                    start=cursor,
                    end=cursor,
                    meta={"fldCharType": child.get(f"{{{W_NS}}}fldCharType")},
                )
                nodes.append(node)

            elif child.tag == f"{{{W_NS}}}instrText":
                text = child.text or ""
                node = ParagraphNode(
                    kind="instr_text",
                    text=text,
                    element=child,
                    run_element=run,
                    start=cursor,
                    end=cursor,
                )
                nodes.append(node)

            elif child.tag == f"{{{W_NS}}}rPr":
                continue

            else:
                node = ParagraphNode(
                    kind="other",
                    text="",
                    element=child,
                    run_element=run,
                    start=cursor,
                    end=cursor,
                    meta={"tag": child.tag},
                )
                nodes.append(node)

        return nodes, cursor