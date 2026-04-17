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
# Datamodeller för footnotes.xml
# ============================================================================

@dataclass
class FootnoteNode:
    """
    En nod i en fotnot. Synlig text och specialnoder hålls isär.
    """
    kind: Literal[
        "text",
        "tab",
        "linebreak",
        "footnote_ref",
        "comment_reference",
        "field_char",
        "instr_text",
        "other"
    ]
    text: str
    element: etree._Element
    run_element: Optional[etree._Element]
    paragraph_element: Optional[etree._Element]
    start: int
    end: int
    meta: dict = field(default_factory=dict)


@dataclass
class FootnoteModel:
    """
    Modell av en fotnot med synlig text, nodsekvens och referens till fotnotens XML.
    """
    footnote_id: str
    footnote_element: etree._Element
    paragraphs: list[etree._Element]
    nodes: list[FootnoteNode]
    visible_text: str


# ============================================================================
# Adapter
# ============================================================================

class FootnotesPartAdapter:
    """
    Adapter för word/footnotes.xml.

    Första versionens ansvar:
    - läsa footnotes.xml via DocxPackage
    - hitta fotnot via footnote_id
    - bygga FootnoteModel
    - lokalisera textankare utan att behandla footnoteRef som vanlig text
    - ge senare renderers en säker nodnivå att arbeta på

    Den här versionen modifierar inte XML ännu.
    """

    def __init__(self, package: DocxPackage, logger):
        self.package = package
        self.logger = logger
        self.tree = self.package.read_footnotes_tree(create_if_missing=False)
        self.root = self.tree.getroot()

    # ------------------------------------------------------------------
    # Publika metoder
    # ------------------------------------------------------------------

    def refresh(self) -> None:
        self.tree = self.package.read_footnotes_tree(create_if_missing=False)
        self.root = self.tree.getroot()

    def get_footnote_model_for_plan(self, plan: ChangePlan) -> FootnoteModel:
        if plan.target.element_type != "footnote":
            raise ValueError(
                f"FootnotesPartAdapter requires element_type=footnote, got {plan.target.element_type}"
            )

        footnote_id = plan.target.footnote_id
        if not footnote_id:
            raise ValueError("ChangePlan for footnote is missing footnote_id")

        footnote_element = self._find_footnote_by_id(footnote_id)
        if footnote_element is None:
            raise ValueError(f"Footnote id={footnote_id} not found in footnotes.xml")

        return self._build_footnote_model(footnote_element, footnote_id)

    def locate_plan_nodes(self, plan: ChangePlan) -> dict:
        """
        Returnerar nodintervall för planens ankare inom fotnotens synliga text.
        """
        model = self.get_footnote_model_for_plan(plan)
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
                f"No overlapping text nodes found for footnote anchor in footnote_id={model.footnote_id}"
            )

        return {
            "footnote_model": model,
            "anchor_start": anchor_start,
            "anchor_end": anchor_end,
            "overlapping_nodes": overlapping_nodes,
            "first_node": overlapping_nodes[0],
            "last_node": overlapping_nodes[-1],
        }

    def debug_plan_location(self, plan: ChangePlan) -> dict:
        located = self.locate_plan_nodes(plan)
        model = located["footnote_model"]

        return {
            "footnote_id": model.footnote_id,
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
            "has_footnote_ref": any(n.kind == "footnote_ref" for n in model.nodes),
        }

    def get_leading_footnote_ref_node(self, model: FootnoteModel) -> Optional[FootnoteNode]:
        """
        Returnerar den första footnoteRef-noden om den finns.
        Denna ska senare bevaras orörd vid rendering.
        """
        for node in model.nodes:
            if node.kind == "footnote_ref":
                return node
        return None

    # ------------------------------------------------------------------
    # Hitta fotnot
    # ------------------------------------------------------------------

    def _find_footnote_by_id(self, footnote_id: str) -> Optional[etree._Element]:
        xpath = f".//w:footnote[@w:id='{footnote_id}']"
        matches = self.root.xpath(xpath, namespaces=NSMAP)
        if not matches:
            return None
        return matches[0]

    # ------------------------------------------------------------------
    # Bygg modell
    # ------------------------------------------------------------------

    def _build_footnote_model(self, footnote_element: etree._Element, footnote_id: str) -> FootnoteModel:
        paragraphs = footnote_element.findall(f"./{{{W_NS}}}p")

        nodes: list[FootnoteNode] = []
        visible_parts: list[str] = []
        cursor = 0

        for p_index, paragraph in enumerate(paragraphs):
            paragraph_nodes, cursor = self._extract_paragraph_nodes(paragraph, cursor)
            nodes.extend(paragraph_nodes)

            for node in paragraph_nodes:
                if node.kind in {"text", "tab", "linebreak"}:
                    visible_parts.append(node.text)

            # lägg in synlig radbrytning mellan stycken i modellen
            if p_index < len(paragraphs) - 1:
                nodes.append(FootnoteNode(
                    kind="linebreak",
                    text="\n",
                    element=paragraph,
                    run_element=None,
                    paragraph_element=paragraph,
                    start=cursor,
                    end=cursor + 1,
                    meta={"synthetic": True, "paragraph_boundary": True},
                ))
                visible_parts.append("\n")
                cursor += 1

        visible_text = "".join(visible_parts)

        return FootnoteModel(
            footnote_id=footnote_id,
            footnote_element=footnote_element,
            paragraphs=paragraphs,
            nodes=nodes,
            visible_text=visible_text,
        )

    def _extract_paragraph_nodes(self, paragraph: etree._Element, cursor: int) -> tuple[list[FootnoteNode], int]:
        nodes: list[FootnoteNode] = []

        for run in paragraph.findall(f"./{{{W_NS}}}r"):
            run_nodes, cursor = self._extract_run_nodes(run, paragraph, cursor)
            nodes.extend(run_nodes)

        return nodes, cursor

    def _extract_run_nodes(
        self,
        run: etree._Element,
        paragraph: etree._Element,
        cursor: int,
    ) -> tuple[list[FootnoteNode], int]:
        nodes: list[FootnoteNode] = []

        for child in run:
            if child.tag == f"{{{W_NS}}}t":
                text = child.text or ""
                node = FootnoteNode(
                    kind="text",
                    text=text,
                    element=child,
                    run_element=run,
                    paragraph_element=paragraph,
                    start=cursor,
                    end=cursor + len(text),
                )
                nodes.append(node)
                cursor += len(text)

            elif child.tag == f"{{{W_NS}}}tab":
                node = FootnoteNode(
                    kind="tab",
                    text="\t",
                    element=child,
                    run_element=run,
                    paragraph_element=paragraph,
                    start=cursor,
                    end=cursor + 1,
                )
                nodes.append(node)
                cursor += 1

            elif child.tag in {f"{{{W_NS}}}br", f"{{{W_NS}}}cr"}:
                node = FootnoteNode(
                    kind="linebreak",
                    text="\n",
                    element=child,
                    run_element=run,
                    paragraph_element=paragraph,
                    start=cursor,
                    end=cursor + 1,
                )
                nodes.append(node)
                cursor += 1

            elif child.tag == f"{{{W_NS}}}footnoteRef":
                node = FootnoteNode(
                    kind="footnote_ref",
                    text="",
                    element=child,
                    run_element=run,
                    paragraph_element=paragraph,
                    start=cursor,
                    end=cursor,
                )
                nodes.append(node)

            elif child.tag == f"{{{W_NS}}}commentReference":
                node = FootnoteNode(
                    kind="comment_reference",
                    text="",
                    element=child,
                    run_element=run,
                    paragraph_element=paragraph,
                    start=cursor,
                    end=cursor,
                    meta={"comment_id": child.get(f"{{{W_NS}}}id")},
                )
                nodes.append(node)

            elif child.tag == f"{{{W_NS}}}fldChar":
                node = FootnoteNode(
                    kind="field_char",
                    text="",
                    element=child,
                    run_element=run,
                    paragraph_element=paragraph,
                    start=cursor,
                    end=cursor,
                    meta={"fldCharType": child.get(f"{{{W_NS}}}fldCharType")},
                )
                nodes.append(node)

            elif child.tag == f"{{{W_NS}}}instrText":
                text = child.text or ""
                node = FootnoteNode(
                    kind="instr_text",
                    text=text,
                    element=child,
                    run_element=run,
                    paragraph_element=paragraph,
                    start=cursor,
                    end=cursor,
                )
                nodes.append(node)

            elif child.tag == f"{{{W_NS}}}rPr":
                continue

            else:
                node = FootnoteNode(
                    kind="other",
                    text="",
                    element=child,
                    run_element=run,
                    paragraph_element=paragraph,
                    start=cursor,
                    end=cursor,
                    meta={"tag": child.tag},
                )
                nodes.append(node)

        return nodes, cursor