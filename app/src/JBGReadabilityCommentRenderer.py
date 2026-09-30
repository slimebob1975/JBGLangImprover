"""
LIX som kommentar vid dokumentets titel.

Används när LIX har beräknats men avsnittet "Om klarspråkningen" inte läggs
till. Titeln väljs så här, för alla typer av dokument:
1. första stycket med rubriknivå 0 (formatmallen Rubrik/Title) och text
2. annars första stycket i brödtexten med text (i läsordning)
Tabellceller (t.ex. faktarutor) och textrutor används inte som titel.
"""

from dataclasses import dataclass
from typing import Optional

try:
    from app.src.JBGCommentsRenderer import CommentsRenderer
    from app.src.JBGDocumentPartAdapter import DocumentPartAdapter
    from app.src.JBGGlobalFindingsRenderer import locate_paragraph
except ModuleNotFoundError:
    from JBGCommentsRenderer import CommentsRenderer
    from JBGDocumentPartAdapter import DocumentPartAdapter
    from JBGGlobalFindingsRenderer import locate_paragraph


READABILITY_COMMENT_AUTHOR = "JBG Klarspråkningstjänst"


@dataclass
class ReadabilityCommentResult:
    applied: bool
    message: str
    element_id: Optional[str] = None


def find_title_element(structure: dict) -> Optional[dict]:
    body = sorted(
        (
            e for e in structure.get("elements", [])
            if e.get("type") == "paragraph"
            and e.get("doc_order") is not None
            and (e.get("text") or "").strip()
        ),
        key=lambda e: e["doc_order"],
    )
    titles = [e for e in body if e.get("heading_level") == 0]
    if titles:
        return titles[0]
    return body[0] if body else None


def _fmt(value: Optional[float]) -> str:
    return "–" if value is None else f"{float(value):.1f}".replace(".", ",")


def build_readability_comment(readability: dict) -> str:
    before = readability.get("before") or {}
    after = readability.get("after") or {}
    before_lix, after_lix = before.get("lix"), after.get("lix")

    first = f"Läsbarhet (LIX): {_fmt(before_lix)}"
    if before.get("band"):
        first += f" ({before['band'].lower()})"
    first += f" före granskningen och {_fmt(after_lix)}"
    if after.get("band"):
        first += f" ({after['band'].lower()})"
    first += " om alla förslag godtas"
    if before_lix is not None and after_lix is not None:
        delta = round(round(after_lix, 1) - round(before_lix, 1), 1)
        if delta:
            sign = "+" if delta > 0 else "\u2212"
            first += f", en förändring med {sign}{abs(delta):.1f}".replace(".", ",")
        else:
            first += ", oförändrat"
    first += "."

    return "\n".join([
        first,
        "LIX bygger på meningarnas längd och andelen ord med fler än sex bokstäver. "
        "Rubriker, sidhuvuden, sidfötter och fotnoter räknas inte med.",
        "Riktvärden: under 30 mycket lättläst, 30–40 lättläst, 40–50 medelsvår, "
        "50–60 svår och över 60 mycket svår.",
    ])


class ReadabilityCommentRenderer:
    def __init__(self, package, logger, structure: dict):
        self.package = package
        self.logger = logger
        self.structure = structure
        self.elements = {e["element_id"]: e for e in structure.get("elements", [])}

    def apply(self, readability: dict) -> ReadabilityCommentResult:
        title = find_title_element(self.structure)
        if title is None:
            return ReadabilityCommentResult(False, "No title paragraph found")
        adapter = DocumentPartAdapter(self.package, self.logger)
        paragraph = locate_paragraph(adapter, self.elements, title["element_id"])
        comments = CommentsRenderer(self.package, self.logger, author=READABILITY_COMMENT_AUTHOR)
        comments.add_paragraph_comment(paragraph, build_readability_comment(readability))
        self.package.write_document_tree(adapter.tree)
        self.logger.info(f"LIX comment added at the title ({title['element_id']})")
        return ReadabilityCommentResult(True, "Comment applied", title["element_id"])
