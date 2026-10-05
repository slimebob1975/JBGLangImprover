"""
Klassning av innehåll som inte ska granskas som vanlig text.

1. Dold text (HiddenTextResolver)
   Ett stycke vars text är helt dold (t.ex. en mallinstruktion som
   "Ta ej bort denna avsnittsbrytning!!") syns inte för läsaren och ska varken
   granskas, räknas i LIX eller få kommentarer. Dold-egenskapen (w:vanish,
   w:specVanish) löses upp som i Word, från det mest specifika till det mest
   allmänna: direkt formatering på texten, textens teckenformat, styckets
   formatmall och sist dokumentets standardformat.

   Delvis dolda stycken räknas inte som dolda. Att ta bort enstaka dolda ord
   skulle flytta positionerna som de lokala förslagen förankras mot.

2. Automatiskt genererat innehåll (generated_kind_*, GeneratedFieldTracker)
   Innehållsförteckningar, figur- och tabellförteckningar, register och
   källförteckningar skapas av Word och skrivs över när de uppdateras. De
   känns igen i tre former:
   - en innehållskontroll (w:sdt) med galleriet "Table of Contents" m.fl.
   - ett fält (TOC, INDEX, BIBLIOGRAPHY) som kan sträcka sig över flera stycken
   - stycken i Words inbyggda förteckningsformat (toc 1-9, table of figures …)
"""

import re
from typing import Optional

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"

# Etikett för platshållaren som den globala granskningen ser.
GENERATED_LABELS = {
    "toc": "Innehållsförteckning",
    "figures": "Figur- eller tabellförteckning",
    "index": "Register",
    "bibliography": "Källförteckning",
}

_SDT_GALLERY_KINDS = {
    "table of contents": "toc",
    "table of figures": "figures",
    "bibliographies": "bibliography",
    "bibliography": "bibliography",
}

# Inbyggda formatnamn lagras på engelska i styles.xml, oavsett språkversion.
_GENERATED_STYLE_RE = re.compile(
    r"^(?:(?P<toc>toc \d)|(?P<figures>table of figures)|(?P<index>index \d|index heading)"
    r"|(?P<bibliography>bibliography))$",
    re.IGNORECASE,
)

_FALSE_VALUES = {"false", "0", "off"}


def generated_label(kind: str) -> str:
    return f"[{GENERATED_LABELS.get(kind, 'Automatiskt innehåll')}, skapas automatiskt]"


# ============================================================================
# Dold text
# ============================================================================

class HiddenTextResolver:
    def __init__(self, styles_root: Optional[etree._Element]):
        self.paragraph_styles: dict[str, dict] = {}
        self.character_styles: dict[str, dict] = {}
        self.default_paragraph_style: Optional[str] = None
        self.document_default: Optional[bool] = None
        if styles_root is None:
            return

        defaults = styles_root.find(f"{W}docDefaults/{W}rPrDefault/{W}rPr")
        self.document_default = self._vanish(defaults)

        for style in styles_root.findall(f"{W}style"):
            style_id = style.get(f"{W}styleId")
            if not style_id:
                continue
            based_on = style.find(f"{W}basedOn")
            info = {
                "based_on": based_on.get(f"{W}val") if based_on is not None else None,
                "vanish": self._vanish(style.find(f"{W}rPr")),
            }
            style_type = style.get(f"{W}type")
            if style_type == "paragraph":
                self.paragraph_styles[style_id] = info
                if style.get(f"{W}default") in {"1", "true", "on"}:
                    self.default_paragraph_style = style_id
            elif style_type == "character":
                self.character_styles[style_id] = info

    # ------------------------------------------------------------------

    @staticmethod
    def _vanish(rpr: Optional[etree._Element]) -> Optional[bool]:
        """True/False om dold-egenskapen är satt, None om den inte är satt."""
        if rpr is None:
            return None
        for tag in ("vanish", "specVanish"):
            element = rpr.find(f"{W}{tag}")
            if element is not None:
                return (element.get(f"{W}val") or "true").lower() not in _FALSE_VALUES
        return None

    @staticmethod
    def _from_chain(styles: dict, style_id: Optional[str]) -> Optional[bool]:
        seen = set()
        while style_id and style_id in styles and style_id not in seen:
            seen.add(style_id)
            value = styles[style_id]["vanish"]
            if value is not None:
                return value
            style_id = styles[style_id]["based_on"]
        return None

    def is_run_hidden(self, run: etree._Element, paragraph_style: Optional[str]) -> bool:
        rpr = run.find(f"{W}rPr")
        direct = self._vanish(rpr)
        if direct is not None:
            return direct
        if rpr is not None:
            r_style = rpr.find(f"{W}rStyle")
            if r_style is not None:
                from_character = self._from_chain(self.character_styles, r_style.get(f"{W}val"))
                if from_character is not None:
                    return from_character
        from_paragraph = self._from_chain(self.paragraph_styles, paragraph_style)
        if from_paragraph is not None:
            return from_paragraph
        return bool(self.document_default)

    def is_paragraph_fully_hidden(self, paragraph: etree._Element) -> bool:
        """
        True när stycket har synlig text och all den texten är dold.
        Tomma stycken och stycken med någon synlig text räknas inte som dolda.
        """
        p_style = paragraph.find(f"{W}pPr/{W}pStyle")
        paragraph_style = p_style.get(f"{W}val") if p_style is not None else self.default_paragraph_style

        has_text = False
        for run in paragraph.iterdescendants(f"{W}r"):
            if next(run.iterancestors(f"{W}p"), None) is not paragraph:
                continue   # körning i en textruta inne i stycket
            carries_text = any(
                (child.tag == f"{W}t" and (child.text or "")) or child.tag in (f"{W}tab", f"{W}br", f"{W}cr")
                for child in run
            )
            if not carries_text:
                continue
            has_text = True
            if not self.is_run_hidden(run, paragraph_style):
                return False
        return has_text


# ============================================================================
# Automatiskt genererat innehåll
# ============================================================================

def generated_kind_for_sdt(sdt: etree._Element) -> Optional[str]:
    gallery = sdt.find(f"{W}sdtPr/{W}docPartObj/{W}docPartGallery")
    if gallery is None:
        return None
    return _SDT_GALLERY_KINDS.get((gallery.get(f"{W}val") or "").strip().lower())


def generated_kind_for_style(style_name: Optional[str]) -> Optional[str]:
    match = _GENERATED_STYLE_RE.match((style_name or "").strip())
    if not match:
        return None
    return next(kind for kind, value in match.groupdict().items() if value)


def generated_kind_for_field(instruction: str) -> Optional[str]:
    tokens = instruction.strip().split()
    if not tokens:
        return None
    name = tokens[0].upper()
    if name == "TOC":
        # \c (bildtexter av en viss etikett) och \a ger figur- och tabellförteckningar
        return "figures" if re.search(r"\\[ca]\b", instruction) else "toc"
    if name == "INDEX":
        return "index"
    if name == "BIBLIOGRAPHY":
        return "bibliography"
    return None


class GeneratedFieldTracker:
    """
    Följer fält genom styckena i läsordning. Ett fält kan börja i ett stycke
    och sluta i ett annat, och en innehållsförteckning innehåller egna fält
    (PAGEREF, HYPERLINK), så fälten hålls på en stack.
    """

    def __init__(self):
        self.stack: list[dict] = []

    def _active_kind(self) -> Optional[str]:
        for field in self.stack:
            if field["kind"] and field["separated"]:
                return field["kind"]
        return None

    def paragraph_kind(self, paragraph: etree._Element) -> Optional[str]:
        """Typen av genererat innehåll som stycket tillhör, eller None."""
        kind = None
        for element in paragraph.iter():
            tag = element.tag
            if tag == f"{W}fldSimple":
                simple = generated_kind_for_field(element.get(f"{W}instr") or "")
                if simple:
                    kind = kind or simple
            elif tag == f"{W}fldChar":
                field_type = element.get(f"{W}fldCharType")
                if field_type == "begin":
                    self.stack.append({"instr": "", "kind": None, "separated": False})
                elif field_type == "separate" and self.stack:
                    top = self.stack[-1]
                    top["separated"] = True
                    top["kind"] = generated_kind_for_field(top["instr"])
                elif field_type == "end" and self.stack:
                    self.stack.pop()
            elif tag == f"{W}instrText" and self.stack:
                self.stack[-1]["instr"] += element.text or ""
            elif tag == f"{W}t" and (element.text or "").strip():
                kind = kind or self._active_kind()
        return kind
