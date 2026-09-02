import posixpath
from typing import Optional

from lxml import etree

try:
    from app.src.JBGDocxPackage import NSMAP, DocxPackage
    from app.src.JBGChangePlanner import ChangePlan
    from app.src.JBGDocumentPartAdapter import ParagraphModel, DocumentPartAdapter
except ModuleNotFoundError:
    from JBGDocxPackage import NSMAP, DocxPackage
    from JBGChangePlanner import ChangePlan
    from JBGDocumentPartAdapter import ParagraphModel, DocumentPartAdapter


class HeaderFooterPartAdapter(DocumentPartAdapter):
    """Locate paragraph models in the real header/footer OOXML part.

    The extractor stores both ``part_name`` and ``container_path``.  This is
    deliberately independent of section numbering because Word may share a
    story part between sections and relationship targets need not be named in
    section order.
    """

    def __init__(self, package: DocxPackage, logger):
        self.package = package
        self.logger = logger
        self._trees: dict[str, etree._ElementTree] = {}

    def refresh(self, part_name: Optional[str] = None) -> None:
        if part_name is None:
            self._trees.clear()
        else:
            self._trees.pop(part_name, None)

    def get_tree(self, part_name: str) -> etree._ElementTree:
        self._validate_part_name(part_name)
        tree = self._trees.get(part_name)
        if tree is None:
            tree = self.package.read_xml_tree(part_name)
            self._trees[part_name] = tree
        return tree

    def write_tree(self, part_name: str) -> None:
        self.package.write_xml_tree(part_name, self.get_tree(part_name))

    def get_paragraph_model_for_plan(self, plan: ChangePlan) -> ParagraphModel:
        if plan.target.element_type not in {"header", "footer"}:
            raise ValueError(
                "HeaderFooterPartAdapter supports only header/footer, got "
                f"{plan.target.element_type}"
            )

        part_name = plan.target.part_name
        container_path = plan.target.container_path
        if not container_path:
            raise ValueError(
                f"Missing container_path for {plan.target.element_id} in {part_name}"
            )

        tree = self.get_tree(part_name)
        expected_root = "hdr" if plan.target.element_type == "header" else "ftr"
        if tree.getroot().tag != f"{{{NSMAP['w']}}}{expected_root}":
            raise ValueError(
                f"Part {part_name} is not a {plan.target.element_type} OOXML story"
            )
        matches = tree.xpath(container_path, namespaces=NSMAP)
        if len(matches) != 1:
            raise ValueError(
                f"Expected one paragraph for {plan.target.element_id} at "
                f"{container_path}, found {len(matches)}"
            )

        paragraph = matches[0]
        return self._build_paragraph_model(
            paragraph,
            target_path=f"{part_name}:{container_path}",
        )

    def locate_plan_nodes(self, plan: ChangePlan) -> dict:
        model = self.get_paragraph_model_for_plan(plan)
        overlapping_nodes = [
            node
            for node in model.nodes
            if node.kind in {"text", "tab", "linebreak"}
            and max(node.start, plan.anchor.start) < min(node.end, plan.anchor.end)
        ]
        if not overlapping_nodes:
            raise ValueError(
                f"No overlapping text nodes found for anchor in {plan.target.element_id}"
            )

        return {
            "paragraph_model": model,
            "anchor_start": plan.anchor.start,
            "anchor_end": plan.anchor.end,
            "overlapping_nodes": overlapping_nodes,
            "first_node": overlapping_nodes[0],
            "last_node": overlapping_nodes[-1],
        }

    @staticmethod
    def _validate_part_name(part_name: str) -> None:
        normalized = posixpath.normpath(part_name.replace("\\", "/"))
        if (
            part_name.startswith("/")
            or normalized != part_name
            or not part_name.startswith("word/")
            or not part_name.endswith(".xml")
        ):
            raise ValueError(f"Unsafe or invalid OOXML story part: {part_name}")
