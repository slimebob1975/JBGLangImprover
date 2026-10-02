"""
Renderar globala iakttagelser som Word-kommentarer (Fas 3/4).

Varje iakttagelse blir en kommentar kring hela det stycke där läsaren bör göra
något (första element_id). Övriga ställen anges i kommentaren med avsnittets
rubrik och början av texten, eftersom Word-kommentarer inte kan länka.

- Stycken och tabellceller förankras direkt.
- Textrutor förankras i det stycke i brödtexten som bär textrutan.
- Fotnoter, sidhuvuden och sidfötter används aldrig som ankare (kommentarer
  i fotnoter visas inte korrekt i Word, se BACKLOG).
"""

import re
from dataclasses import dataclass
from types import SimpleNamespace
from typing import Optional

try:
    from app.src.JBGCommentsRenderer import CommentsRenderer
    from app.src.JBGDocumentPartAdapter import DocumentPartAdapter
    from app.src.JBGReadabilityMetrics import is_caption
except ModuleNotFoundError:
    from JBGCommentsRenderer import CommentsRenderer
    from JBGDocumentPartAdapter import DocumentPartAdapter
    from JBGReadabilityMetrics import is_caption


# Egen kommentarförfattare, så att de globala kommentarerna kan visas eller
# döljas för sig i Word (Granska > Visa markering > Specifika personer).
GLOBAL_COMMENT_AUTHOR = "JBG Klarspråkningstjänst (global granskning)"
SNIPPET_WORDS = 8

# Inledning till raden om övriga ställen, per kategori.
RELATED_LABELS = {
    "repetition": "Hänger ihop med",
    "inconsistency": "Jämför med",
    "error": "Se även",
    "disposition": "Se även",
    "heading": "Se även",
    "conclusion": "Jämför med",
}
_TEXTBOX_HOST_RE = re.compile(r"^/document/body/paragraph\[(\d+)\]/textbox")


@dataclass
class GlobalCommentResult:
    finding: object
    applied: bool
    message: str
    comment_id: Optional[int] = None


class GlobalFindingsRenderer:
    def __init__(self, package, logger, structure: dict):
        self.package = package
        self.logger = logger
        self.elements = {e["element_id"]: e for e in structure.get("elements", [])}
        self.ordered = sorted(
            (e for e in structure.get("elements", []) if e.get("doc_order") is not None),
            key=lambda e: e["doc_order"],
        )

    # ------------------------------------------------------------------
    # Publikt API
    # ------------------------------------------------------------------

    def apply(self, findings: list) -> list[GlobalCommentResult]:
        if not findings:
            return []
        adapter = DocumentPartAdapter(self.package, self.logger)
        comments = CommentsRenderer(self.package, self.logger, author=GLOBAL_COMMENT_AUTHOR)
        results: list[GlobalCommentResult] = []
        changed = False

        for finding in findings:
            try:
                paragraph = locate_paragraph(adapter, self.elements, finding.element_ids[0])
                comment_id = comments.add_paragraph_comment(paragraph, self.comment_text(finding))
                changed = True
                results.append(GlobalCommentResult(finding, True, "Comment applied", comment_id))
            except Exception as ex:
                self.logger.warning(
                    f"Global finding comment skipped for {finding.element_ids[:1]}: {ex}"
                )
                results.append(GlobalCommentResult(finding, False, str(ex)))

        if changed:
            self.package.write_document_tree(adapter.tree)
        applied = len([r for r in results if r.applied])
        self.logger.info(f"Global findings rendered as comments: {applied} of {len(results)}")
        return results

    # ------------------------------------------------------------------
    # Kommentartext
    # ------------------------------------------------------------------

    def comment_text(self, finding) -> str:
        # Etiketten inleder första raden: "Onödig upprepning. Samma resultat …"
        lines = [f"{finding.label}. {finding.description}"]
        if finding.proposal:
            lines.append(f"Förslag: {finding.proposal}")
        related_quote = getattr(finding, "related_quote", "") or ""
        related = [
            self._describe_location(element_id, quote=related_quote if index == 0 else "")
            for index, element_id in enumerate(finding.element_ids[1:])
        ]
        related = [r for r in related if r]
        if related:
            label = RELATED_LABELS.get(finding.category, "Hänger ihop med")
            lines.append(f"{label}: " + "; ".join(related) + ".")
        proposed_order = getattr(finding, "proposed_order", None) or []
        if proposed_order:
            lines.append("Möjlig ordning: " + ", ".join(f"”{h}”" for h in proposed_order) + ".")
        return "\n".join(lines)

    def _describe_location(self, element_id: str, quote: str = "") -> Optional[str]:
        element = self.elements.get(element_id)
        if element is None:
            return None
        if quote:
            snippet = quote   # det verifierade citatet visar själva motsägelsen
        else:
            words = (element.get("text") or "").split()
            snippet = " ".join(words[:SNIPPET_WORDS]) + (" …" if len(words) > SNIPPET_WORDS else "")

        if (element.get("heading_level") or 0) >= 1:
            return f"rubriken ”{snippet}”"
        heading = self._nearest_heading(element)
        if heading:
            return f"avsnittet ”{heading}” (”{snippet}”)"
        return f"”{snippet}”"

    def _nearest_heading(self, element: dict) -> Optional[str]:
        order = element.get("doc_order")
        if order is None:
            return None
        heading = None
        for candidate in self.ordered:
            if candidate["doc_order"] >= order:
                break
            if (
                (candidate.get("heading_level") or 0) >= 1
                and (candidate.get("text") or "").strip()
                and not is_caption(candidate)
            ):
                heading = " ".join(candidate["text"].split())
        return heading



def locate_paragraph(adapter: DocumentPartAdapter, elements: dict, element_id: str):
    """
    w:p-elementet för ett element i strukturen, för kommentarer kring hela
    stycket. Stycken och tabellceller förankras direkt; textrutor i stycket
    som bär dem. Fotnoter, sidhuvuden och sidfötter stöds inte.
    """
    element = elements.get(element_id)
    if element is None:
        raise ValueError(f"Unknown element_id: {element_id}")

    element_type = element.get("type")
    if element_type == "textbox":
        match = _TEXTBOX_HOST_RE.match(element.get("container_path") or "")
        if not match:
            raise ValueError(f"Cannot find host paragraph for {element_id}")
        element_type, element_id = "paragraph", f"paragraph_{match.group(1)}"

    plan = SimpleNamespace(target=SimpleNamespace(element_type=element_type, element_id=element_id))
    if element_type == "paragraph":
        return adapter._find_main_document_paragraph(plan)
    if element_type == "table_cell":
        return adapter._find_table_cell_paragraph(plan)
    raise ValueError(f"Comments are not anchored in {element_type} elements")
