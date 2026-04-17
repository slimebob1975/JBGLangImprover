from copy import deepcopy
from dataclasses import dataclass
from typing import Optional

from lxml import etree

try:
    from app.src.JBGDocxPackage import W_NS, DocxPackage
except ModuleNotFoundError:
    from JBGDocxPackage import W_NS, DocxPackage

try:
    from app.src.JBGDocumentPartAdapter import DocumentPartAdapter
except ModuleNotFoundError:
    from JBGDocumentPartAdapter import DocumentPartAdapter

try:
    from app.src.JBGFootnotesPartAdapter import FootnotesPartAdapter
except ModuleNotFoundError:
    from JBGFootnotesPartAdapter import FootnotesPartAdapter

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


class SimpleMarkupRenderer:
    """
    Enkel visuell markup-renderer.

    Första versionens ansvar:
    - applicera ChangePlan i document.xml eller footnotes.xml
    - använda adapters för att hitta rätt paragraph/footnote
    - ersätta endast berörd textregion
    - bevara specialnoder som footnoteReference/footnoteRef bättre än legacy
    - skriva röd/struken gammal text + grön ny text

    Begränsningar i första versionen:
    - inga kommentarer
    - inga tracked changes
    - inga headers/footers
    - ingen konfliktupplösning (konflikter bör filtreras innan rendering)
    """

    def __init__(self, package: DocxPackage, logger):
        self.package = package
        self.logger = logger
        self.document_adapter = DocumentPartAdapter(package, logger)

        self.footnotes_adapter = None
        if self.package.part_exists("word/footnotes.xml"):
            self.footnotes_adapter = FootnotesPartAdapter(package, logger)

    # ------------------------------------------------------------------
    # Publikt API
    # ------------------------------------------------------------------

    def apply_plan(self, plan: ChangePlan) -> RenderResult:
        try:
            if plan.target.element_type == "footnote":
                if self.footnotes_adapter is None:
                    return RenderResult(
                        plan=plan,
                        applied=False,
                        message="Document has no footnotes.xml part",
                    )

                self._apply_footnote_plan(plan)
                self.package.write_footnotes_tree(self.footnotes_adapter.tree)
                return RenderResult(plan=plan, applied=True, message="Applied footnote markup")

            if plan.target.element_type in {"paragraph", "table_cell", "textbox"}:
                self._apply_document_plan(plan)
                self.package.write_document_tree(self.document_adapter.tree)
                return RenderResult(plan=plan, applied=True, message="Applied document markup")

            return RenderResult(
                plan=plan,
                applied=False,
                message=f"Unsupported element_type for SimpleMarkupRenderer: {plan.target.element_type}",
            )

        except Exception as ex:
            return RenderResult(plan=plan, applied=False, message=str(ex))

    def apply_plans(self, plans: list[ChangePlan]) -> list[RenderResult]:
        results = []

        for plan in plans:
            if "overlapping_change_conflict" in plan.notes:
                results.append(RenderResult(
                    plan=plan,
                    applied=False,
                    message="Skipped due to overlapping_change_conflict",
                ))
                continue

            result = self.apply_plan(plan)
            results.append(result)

            # refresh adapters after each successful mutation
        if result.applied:
            self.document_adapter.refresh()
            if self.footnotes_adapter is not None:
                self.footnotes_adapter.refresh()

        return results

    # ------------------------------------------------------------------
    # Document.xml
    # ------------------------------------------------------------------

    def _apply_document_plan(self, plan: ChangePlan) -> None:
        located = self.document_adapter.locate_plan_nodes(plan)
        model = located["paragraph_model"]

        self._rewrite_container_with_markup(
            container_element=model.paragraph_element,
            visible_text=model.visible_text,
            anchor_start=plan.anchor.start,
            anchor_end=plan.anchor.end,
            old_text=plan.old_text,
            new_text=plan.new_text,
            preserve_special_zero_width=True,
        )

    # ------------------------------------------------------------------
    # Footnotes.xml
    # ------------------------------------------------------------------

    def _apply_footnote_plan(self, plan: ChangePlan) -> None:
        located = self.footnotes_adapter.locate_plan_nodes(plan)
        model = located["footnote_model"]

        # Första versionen: skriv om första stycket som bär den synliga textregionen.
        # Vi bevarar footnoteRef-run explicit om den finns.
        paragraph = self._find_best_footnote_paragraph_for_anchor(model, plan)

        paragraph_visible_text = self._get_visible_text_from_paragraph(paragraph)
        paragraph_anchor = self._map_global_anchor_to_paragraph(model, paragraph, plan.anchor.start, plan.anchor.end)

        self._rewrite_container_with_markup(
            container_element=paragraph,
            visible_text=paragraph_visible_text,
            anchor_start=paragraph_anchor["start"],
            anchor_end=paragraph_anchor["end"],
            old_text=plan.old_text,
            new_text=plan.new_text,
            preserve_special_zero_width=True,
        )

    def _find_best_footnote_paragraph_for_anchor(self, model, plan: ChangePlan):
        """
        Första version:
        - om fotnoten bara har ett stycke, använd det
        - annars välj det stycke vars synliga text omfattar anchor.start
        """
        if len(model.paragraphs) == 1:
            return model.paragraphs[0]

        cursor = 0
        for p in model.paragraphs:
            p_text = self._get_visible_text_from_paragraph(p)
            start = cursor
            end = cursor + len(p_text)

            if start <= plan.anchor.start <= end:
                return p

            cursor = end + 1  # +1 för syntetisk paragraph-boundary linebreak i modellen

        return model.paragraphs[0]

    def _map_global_anchor_to_paragraph(self, model, paragraph, global_start: int, global_end: int) -> dict:
        cursor = 0

        for p in model.paragraphs:
            p_text = self._get_visible_text_from_paragraph(p)
            start = cursor
            end = cursor + len(p_text)

            if p is paragraph:
                local_start = max(0, global_start - start)
                local_end = max(0, min(len(p_text), global_end - start))
                return {"start": local_start, "end": local_end}

            cursor = end + 1

        raise ValueError("Could not map global anchor to footnote paragraph")

    # ------------------------------------------------------------------
    # Omskrivning av en container (paragraph-liknande)
    # ------------------------------------------------------------------

    def _rewrite_container_with_markup(
        self,
        container_element: etree._Element,
        visible_text: str,
        anchor_start: int,
        anchor_end: int,
        old_text: str,
        new_text: str,
        preserve_special_zero_width: bool = True,
    ) -> None:
        if anchor_start < 0 or anchor_end < anchor_start or anchor_end > len(visible_text):
            raise ValueError("Invalid anchor range for container rewrite")

        matched = visible_text[anchor_start:anchor_end]
        if matched != old_text:
            # tolerera inte tyst mismatch här
            raise ValueError(
                f"Anchor text mismatch. Expected old_text={old_text!r}, found={matched!r}"
            )

        prefix = visible_text[:anchor_start]
        suffix = visible_text[anchor_end:]

        # samla specialruns som ska bevaras
        preserved_runs_before = []
        preserved_runs_after = []

        if preserve_special_zero_width:
            preserved_runs_before, preserved_runs_after = self._collect_preserved_special_runs(container_element)

        # ta bort alla runs i containern
        runs = container_element.findall(f"./{{{W_NS}}}r")
        for run in runs:
            container_element.remove(run)

        # bygg om minimalt
        if preserved_runs_before:
            for run in preserved_runs_before:
                container_element.append(run)

        if prefix:
            container_element.append(self._make_plain_run(prefix))

        if old_text:
            container_element.append(self._make_deleted_markup_run(old_text))

        if new_text:
            container_element.append(self._make_inserted_markup_run(new_text))

        if suffix:
            container_element.append(self._make_plain_run(suffix))

        if preserved_runs_after:
            for run in preserved_runs_after:
                container_element.append(run)

    def _collect_preserved_special_runs(self, container_element: etree._Element) -> tuple[list[etree._Element], list[etree._Element]]:
        """
        Första versionens strategi:
        - bevara runs som endast innehåller zero-width-specialinnehåll
        - särskilt viktigt för footnoteReference i document.xml och footnoteRef i footnotes.xml
        - lägg dem före/efter text beroende på deras ursprungliga position
        """
        all_runs = container_element.findall(f"./{{{W_NS}}}r")
        if not all_runs:
            return [], []

        preserved_before = []
        preserved_after = []

        text_seen = False

        for run in all_runs:
            visible_len = self._visible_length_of_run(run)
            has_special = self._run_contains_preservable_special(run)

            if not has_special or visible_len > 0:
                if visible_len > 0:
                    text_seen = True
                continue

            run_copy = deepcopy(run)
            if text_seen:
                preserved_after.append(run_copy)
            else:
                preserved_before.append(run_copy)

        return preserved_before, preserved_after

    def _visible_length_of_run(self, run: etree._Element) -> int:
        length = 0
        for child in run:
            if child.tag == f"{{{W_NS}}}t":
                length += len(child.text or "")
            elif child.tag == f"{{{W_NS}}}tab":
                length += 1
            elif child.tag in {f"{{{W_NS}}}br", f"{{{W_NS}}}cr"}:
                length += 1
        return length

    def _run_contains_preservable_special(self, run: etree._Element) -> bool:
        for child in run:
            if child.tag in {
                f"{{{W_NS}}}footnoteReference",
                f"{{{W_NS}}}footnoteRef",
                f"{{{W_NS}}}commentReference",
                f"{{{W_NS}}}fldChar",
                f"{{{W_NS}}}drawing",
            }:
                return True
        return False

    # ------------------------------------------------------------------
    # Textutvinning per paragraph
    # ------------------------------------------------------------------

    def _get_visible_text_from_paragraph(self, paragraph: etree._Element) -> str:
        parts = []

        for run in paragraph.findall(f"./{{{W_NS}}}r"):
            for child in run:
                if child.tag == f"{{{W_NS}}}t":
                    parts.append(child.text or "")
                elif child.tag == f"{{{W_NS}}}tab":
                    parts.append("\t")
                elif child.tag in {f"{{{W_NS}}}br", f"{{{W_NS}}}cr"}:
                    parts.append("\n")

        return "".join(parts)

    # ------------------------------------------------------------------
    # Run-fabriker
    # ------------------------------------------------------------------

    def _make_plain_run(self, text: str) -> etree._Element:
        run = etree.Element(f"{{{W_NS}}}r")
        t = etree.SubElement(run, f"{{{W_NS}}}t")
        t.set(f"{{{XML_NS}}}space", "preserve")
        t.text = text
        return run

    def _make_deleted_markup_run(self, text: str) -> etree._Element:
        run = etree.Element(f"{{{W_NS}}}r")

        rpr = etree.SubElement(run, f"{{{W_NS}}}rPr")
        color = etree.SubElement(rpr, f"{{{W_NS}}}color")
        color.set(f"{{{W_NS}}}val", "FF0000")
        etree.SubElement(rpr, f"{{{W_NS}}}strike")

        t = etree.SubElement(run, f"{{{W_NS}}}t")
        t.set(f"{{{XML_NS}}}space", "preserve")
        t.text = text
        return run

    def _make_inserted_markup_run(self, text: str) -> etree._Element:
        run = etree.Element(f"{{{W_NS}}}r")

        rpr = etree.SubElement(run, f"{{{W_NS}}}rPr")
        color = etree.SubElement(rpr, f"{{{W_NS}}}color")
        color.set(f"{{{W_NS}}}val", "008000")

        t = etree.SubElement(run, f"{{{W_NS}}}t")
        t.set(f"{{{XML_NS}}}space", "preserve")
        t.text = text
        return run