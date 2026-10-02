"""
Avsnittet "Om klarspråkningen" sist i resultatdokumentet (G5).

Tre regler gäller för alla dokument, oavsett mall:
1. Sist: avsnittet läggs efter allt befintligt innehåll i brödtexten, direkt
   före body-sectPr, och börjar på en ny sida.
2. Spårat: i läget "spåra ändringar" är hela avsnittet en spårad infogning.
3. Onumrerad rubrik: rubriken får alltid en direkt numreringsspärr
   (w:numPr/w:numId = 0), oavsett om stilen eller numbering.xml numrerar den.

Två delar:
- build_about_section_blocks(): bygger innehållet (ren text) från en
  RunSummary-dict. Ingen XML, lätt att testa.
- AboutSectionRenderer: infogar blocken sist i word/document.xml, efter
  allt befintligt innehåll och före body-sectPr.

Avsnittet omges av det dolda bokmärket ABOUT_BOOKMARK_NAME. Extraktorn
hoppar över allt inom bokmärket, och vid en ny körning ersätts ett
befintligt avsnitt i stället för att ett till läggs till.

I läget "spåra ändringar" infogas avsnittet som en spårad infogning enligt
samma modell som Word själv använder när man skriver nya stycken sist i ett
dokument: det tidigare sista styckets styckemarkering och alla nya stycken
utom det sista markeras som infogade, och det sista nya stycket får det
tidigare sista styckets styckeformat. Avvisas infogningen blir dokumentet
därför identiskt med originalet.
"""

import copy
import re
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Any, Literal, Optional

from lxml import etree

try:
    from app.src.JBGDocxPackage import DocxPackage
    from app.src.JBGRevisionIds import max_revision_id
except ModuleNotFoundError:
    from JBGDocxPackage import DocxPackage
    from JBGRevisionIds import max_revision_id


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NSMAP = {"w": W_NS}
W = f"{{{W_NS}}}"

# Dolt bokmärke (namn som börjar med "_" visas inte i Words bokmärkeslista).
ABOUT_BOOKMARK_NAME = "_JBG_OmKlarsprakningen"
ABOUT_HEADING = "Om klarspråkningen"

# Kategorinamn i plural för raden om global granskning.
GLOBAL_CATEGORY_NAMES_PLURAL = {
    "repetition": "upprepningar",
    "inconsistency": "inkonsekvenser",
    "error": "troliga fel",
    "disposition": "förslag om disposition",
    "heading": "förslag om rubriker",
    "conclusion": "slutsatser som behöver stöd",
}

SWEDISH_MONTHS = (
    "januari", "februari", "mars", "april", "maj", "juni",
    "juli", "augusti", "september", "oktober", "november", "december",
)


# Ordning för barn i w:pPr enligt schemat (de som den här modulen berör).
_PPR_TAIL_TAGS = (f"{W}sectPr", f"{W}pPrChange")


# ============================================================================
# Innehåll
# ============================================================================

@dataclass
class AboutBlock:
    """
    Ett block i avsnittet:
    - "heading": avsnittets rubrik
    - "paragraph": stycke med körningar (text, fet); keep_with_next håller
      en tabellrubrik ihop med tabellen
    - "table": rows[0] är rubrikrad; numeric_columns högerställs
    """
    kind: Literal["heading", "paragraph", "table"]
    runs: list[tuple[str, bool]] = field(default_factory=list)  # (text, fet)
    rows: list[list[str]] = field(default_factory=list)
    numeric_columns: tuple[int, ...] = ()
    keep_with_next: bool = False

    @property
    def text(self) -> str:
        return "".join(text for text, _ in self.runs)


def _fmt_int(value: Any) -> str:
    try:
        return f"{int(value):,}".replace(",", "\u00a0")
    except (TypeError, ValueError):
        return "0"


def _fmt_lix(value: Optional[float], band: Optional[str]) -> str:
    if value is None:
        return "kunde inte beräknas"
    number = f"{float(value):.1f}".replace(".", ",")
    return f"{number} ({band.lower()})" if band else number


def _fmt_date(iso_timestamp: Optional[str]) -> str:
    try:
        moment = datetime.fromisoformat(iso_timestamp) if iso_timestamp else datetime.now(timezone.utc)
    except ValueError:
        moment = datetime.now(timezone.utc)
    try:
        from zoneinfo import ZoneInfo
        moment = moment.astimezone(ZoneInfo("Europe/Stockholm"))
    except Exception:
        moment = moment.astimezone()  # tzdata saknas (t.ex. Windows): lokal tid
    return f"{moment.day} {SWEDISH_MONTHS[moment.month - 1]} {moment.year}"


def _item(label: str, value: str) -> AboutBlock:
    return AboutBlock("paragraph", [(f"{label}: ", True), (value, False)])


def _fmt_delta(before: Optional[float], after: Optional[float]) -> str:
    """Förändring beräknad från de avrundade värdena, med typografiskt minustecken."""
    if before is None or after is None:
        return "–"
    delta = round(round(after, 1) - round(before, 1), 1)
    if delta == 0:
        return "0,0"
    sign = "+" if delta > 0 else "\u2212"
    return f"{sign}{abs(delta):.1f}".replace(".", ",")


def _fmt_lix_value(value: Optional[float]) -> str:
    return "–" if value is None else f"{float(value):.1f}".replace(".", ",")


def _table_title(text: str) -> AboutBlock:
    return AboutBlock("paragraph", [(text, True)], keep_with_next=True)


PHASE_NAMES = (("local", "Lokal granskning"), ("global", "Global granskning"))

LIX_EXPLANATION = (
    "LIX (läsbarhetsindex) bygger på meningarnas längd och andelen ord med fler "
    "än sex bokstäver. Rubriker, sidhuvuden, sidfötter och fotnoter räknas inte med. "
    "Riktvärden: under 30 mycket lättläst, 30–40 lättläst, 40–50 medelsvår, "
    "50–60 svår och över 60 mycket svår. LIX säger inget om hur begriplig texten "
    "är i övrigt, och en tydligare formulering kan ibland ge ett högre värde."
)


def _usage_table(usage: dict) -> AboutBlock:
    by_phase = usage.get("by_phase") or {}
    total = usage.get("total") or {}
    phases = [(key, name) for key, name in PHASE_NAMES if key in by_phase]
    columns = [by_phase[key] for key, _ in phases]
    header = [""] + [name for _, name in phases]
    if len(phases) != 1:
        columns.append(total)
        header.append("Totalt")

    def row(label: str, field_name: str) -> list[str]:
        return [label] + [_fmt_int(column.get(field_name)) for column in columns]

    rows = [header, row("Anrop till språkmodellen", "calls")]
    if total.get("failed_calls"):
        rows.append(row("varav misslyckade", "failed_calls"))
    rows.append(row("Tokens skickade", "prompt_tokens"))
    if total.get("cached_prompt_tokens"):
        # Återanvända (cachade) tokens debiteras till ett lägre pris, så utan
        # den här raden överskattas kostnaden.
        rows.append(row("varav återanvända (lägre kostnad)", "cached_prompt_tokens"))
    rows.append(row("Tokens mottagna", "completion_tokens"))
    if total.get("reasoning_tokens"):
        rows.append(row("varav för resonemang", "reasoning_tokens"))
    return AboutBlock("table", rows=rows, numeric_columns=tuple(range(1, len(header))))


def _suggestions_table(summary: dict) -> AboutBlock:
    local = summary.get("local_suggestions") or {}
    rows = [
        ["Lokala förslag", "Antal"],
        ["Förslag från språkmodellen", _fmt_int(local.get("raw"))],
        ["Förslag som klarade kontrollerna", _fmt_int(local.get("accepted"))],
        ["Förslag som förts in i dokumentet", _fmt_int(local.get("applied"))],
    ]
    if local.get("failed"):
        rows.append(["Förslag som inte kunde föras in", _fmt_int(local.get("failed"))])
    return AboutBlock("table", rows=rows, numeric_columns=(1,))


def _global_blocks(summary: dict) -> list[AboutBlock]:
    global_findings = summary.get("global_findings") or {}
    accepted = global_findings.get("accepted") or 0
    title = _table_title("Iakttagelser från den globala granskningen")
    if global_findings.get("errors") and not accepted:
        return [title, AboutBlock("paragraph", [("Den globala granskningen kunde inte genomföras.", False)])]
    rows = [["Kategori", "Antal"]]
    for category, count in (global_findings.get("by_category") or {}).items():
        name = GLOBAL_CATEGORY_NAMES_PLURAL.get(category, category)
        rows.append([name[:1].upper() + name[1:], _fmt_int(count)])
    rows.append(["Totalt", _fmt_int(accepted)])
    blocks = [title, AboutBlock("table", rows=rows, numeric_columns=(1,))]
    if accepted:
        blocks.append(AboutBlock("paragraph", [("Iakttagelserna finns som kommentarer i dokumentet.", False)]))
    return blocks


def _readability_blocks(readability: dict) -> list[AboutBlock]:
    before = readability.get("before") or {}
    after = readability.get("after") or {}
    rows = [
        ["Del av dokumentet", "Före", "Om alla förslag godtas", "Förändring"],
        [
            "Hela dokumentet",
            _fmt_lix(before.get("lix"), before.get("band")),
            _fmt_lix(after.get("lix"), after.get("band")),
            _fmt_delta(before.get("lix"), after.get("lix")),
        ],
    ]
    for section in readability.get("sections") or []:
        section_before = (section.get("before") or {}).get("lix")
        section_after = (section.get("after") or {}).get("lix")
        rows.append([
            section.get("heading") or "Före första rubriken",
            _fmt_lix_value(section_before),
            _fmt_lix_value(section_after),
            _fmt_delta(section_before, section_after),
        ])
    return [
        _table_title("Läsbarhet (LIX)"),
        AboutBlock("table", rows=rows, numeric_columns=(1, 2, 3)),
        AboutBlock("paragraph", [(LIX_EXPLANATION, False)]),
    ]


def build_about_section_blocks(summary: dict, tracked: bool) -> list[AboutBlock]:
    """
    Bygger avsnittets innehåll från RunSummary.to_dict() (med usage ifylld).
    Uppgifter om körningen står som korta rader; siffror står i tabeller.
    Avsnittet slutar alltid med ett stycke (där bokmärket slutar).
    """
    blocks = [AboutBlock("heading", [(ABOUT_HEADING, False)])]
    blocks.append(AboutBlock("paragraph", [(
        "Det här dokumentet har granskats med en klarspråkningstjänst. "
        "Förslagen till ändringar har tagits fram av en AI-språkmodell.",
        False,
    )]))

    # --- Körningen
    blocks.append(_item("Datum", _fmt_date(summary.get("started_at"))))
    blocks.append(_item("Språkmodell", str(summary.get("model") or "okänd")))
    if summary.get("docx_mode") == "tracked":
        presentation = "Spåra ändringar"
        if summary.get("include_motivations"):
            presentation += " med kommentarer"
    else:
        presentation = "Enkel färgmarkering"
    blocks.append(_item("Visning av förslagen", presentation))
    prompt_customized = summary.get("prompt_customized")
    if prompt_customized is not None:
        blocks.append(_item(
            "Promptinstruktion",
            "anpassad för den här granskningen" if prompt_customized else "standard",
        ))

    # --- Tabeller
    blocks.append(_table_title("Anrop och tokens"))
    blocks.append(_usage_table(summary.get("usage") or {}))
    blocks.append(_table_title("Förslag"))
    blocks.append(_suggestions_table(summary))
    if summary.get("global_review"):
        blocks.extend(_global_blocks(summary))
    if summary.get("readability"):
        blocks.extend(_readability_blocks(summary["readability"]))

    # --- Avslutning
    closing = "AI kan göra misstag, så granska förslagen innan du godtar dem."
    if tracked:
        closing += (
            " Avsnittet är infogat som en spårad ändring. Du tar bort det genom "
            "att markera hela avsnittet och avvisa ändringen."
        )
    else:
        closing += " Du kan ta bort avsnittet när granskningen är klar."
    blocks.append(AboutBlock("paragraph", [(closing, False)]))
    return blocks


# ============================================================================
# Rendering
# ============================================================================

@dataclass
class AboutSectionResult:
    applied: bool
    message: str
    heading_style_id: Optional[str] = None
    replaced_existing: bool = False


class AboutSectionRenderer:
    SOURCE_HEADING_RE = re.compile(
        r"^\s*(källor|källförteckning|referenser|referenslista|bilaga\b|bilagor|appendix|references)",
        re.IGNORECASE,
    )
    CAPTION_TOKENS = ("caption", "beskrivning", "tabellrubrik", "figurrubrik", "diagramrubrik")

    def __init__(self, package: DocxPackage, logger, author: str = "JBG Klarspråkningstjänst"):
        self.package = package
        self.logger = logger
        self.author = author
        self._next_id: Optional[int] = None

    # ------------------------------------------------------------------
    # Publikt API
    # ------------------------------------------------------------------

    def apply(self, blocks: list[AboutBlock], tracked: bool) -> AboutSectionResult:
        tree = self.package.read_document_tree()
        body = tree.getroot().find(f"{W}body")
        if body is None:
            return AboutSectionResult(False, "Document has no w:body")

        replaced = self._remove_existing_section(body)
        heading_style = self._resolve_heading_style(body)
        timestamp = datetime.now(timezone.utc).replace(microsecond=0).isoformat()

        if not blocks or blocks[-1].kind != "paragraph":
            raise ValueError("The section must end with a paragraph (the bookmark ends there)")
        elements = [
            self._build_table(block) if block.kind == "table" else self._build_paragraph(block, heading_style)
            for block in blocks
        ]
        paragraphs = elements  # stycken och tabeller i ordning (namnet behålls för läsbarhet nedan)

        final_sectpr = body.find(f"{W}sectPr")
        previous_last = self._last_block(body)

        if tracked:
            self._next_id = max_revision_id(self.package) + 1
            for element in elements:
                if element.tag == f"{W}tbl":
                    self._mark_table_inserted(element, timestamp)
                else:
                    self._wrap_runs_in_insertion(element, timestamp)

            closing = etree.Element(f"{W}p")
            if self._can_extend_last_paragraph(previous_last):
                # Wordmodellen: tidigare sista stycket får infogad styckemarkering,
                # nya stycken likaså, och ett avslutande stycke bär det tidigare
                # sista styckets format. Avvisas infogningen blir dokumentet
                # identiskt med originalet.
                previous_ppr = previous_last.find(f"{W}pPr")
                if previous_ppr is not None:
                    closing.append(copy.deepcopy(previous_ppr))
                self._mark_paragraph_inserted(previous_last, timestamp)
            # Övriga fall (dokumentet slutar med tabell, innehållskontroll,
            # avsnittsbrytning eller är tomt): alla nya stycken markeras och ett
            # tomt avslutande stycke blir kvar vid avvisning. Word kräver ändå
            # ett stycke sist i dokumentet, så resultatet motsvarar originalet.
            for element in elements:
                if element.tag == f"{W}p":
                    self._mark_paragraph_inserted(element, timestamp)
            paragraphs.append(closing)

        self._add_bookmark(paragraphs[0], paragraphs[len(blocks) - 1])

        for paragraph in paragraphs:
            if final_sectpr is not None:
                final_sectpr.addprevious(paragraph)
            else:
                body.append(paragraph)

        self.package.write_document_tree(tree)
        tables = len([b for b in blocks if b.kind == "table"])
        message = (
            f"Inserted {len(blocks) - tables} paragraphs and {tables} tables "
            f"({'tracked' if tracked else 'plain'})"
        )
        if replaced:
            message += "; replaced an existing section"
        self.logger.info(f"About section: {message}; heading style: {heading_style or 'bold fallback'}")
        return AboutSectionResult(True, message, heading_style, replaced)

    # ------------------------------------------------------------------
    # Befintligt avsnitt
    # ------------------------------------------------------------------

    @staticmethod
    def find_section_blocks(body: etree._Element) -> list[etree._Element]:
        """
        Avsnittets block (w:p, w:tbl, w:sdt) i brödtexten, från stycket med
        bokmärkets början till stycket med dess slut. Tom lista om inget avsnitt.
        """
        blocks = [child for child in body if child.tag in (f"{W}p", f"{W}tbl", f"{W}sdt")]
        start_index = end_index = None
        bookmark_id = None
        for index, block in enumerate(blocks):
            if start_index is None and block.tag == f"{W}p":
                for start in block.iter(f"{W}bookmarkStart"):
                    if start.get(f"{W}name") == ABOUT_BOOKMARK_NAME:
                        start_index, bookmark_id = index, start.get(f"{W}id")
                        break
            if start_index is not None:
                for end in block.iter(f"{W}bookmarkEnd"):
                    if end.get(f"{W}id") == bookmark_id:
                        end_index = index
                        break
                if end_index is not None:
                    break
        if start_index is None:
            return []
        return blocks[start_index:(end_index if end_index is not None else len(blocks) - 1) + 1]

    @classmethod
    def find_section_indices(cls, body: etree._Element) -> tuple[set[int], set[int]]:
        """
        1-baserade index för avsnittets stycken och tabeller, med samma
        numrering som doc.paragraphs och doc.tables (brödtextens w:p och w:tbl).
        """
        # Listan hålls vid liv under hela jämförelsen: lxml skapar tillfälliga
        # Python-objekt för noderna, så id() kan återanvändas om de släpps.
        # Elementen jämförs därför direkt (identitet), inte via id().
        section_blocks = cls.find_section_blocks(body)
        section = set(section_blocks)
        paragraphs: set[int] = set()
        tables: set[int] = set()
        p_count = t_count = 0
        for child in body:
            if child.tag == f"{W}p":
                p_count += 1
                if child in section:
                    paragraphs.add(p_count)
            elif child.tag == f"{W}tbl":
                t_count += 1
                if child in section:
                    tables.add(t_count)
        return paragraphs, tables

    def _remove_existing_section(self, body: etree._Element) -> bool:
        section = self.find_section_blocks(body)
        if not section:
            return False
        following = []
        sibling = section[-1].getnext()
        while sibling is not None and sibling.tag == f"{W}p":
            following.append(sibling)
            sibling = sibling.getnext()
        for block in section:
            body.remove(block)
        # Tomma avslutande stycken från förra körningen tas också bort.
        for paragraph in following:
            if self._is_empty_paragraph(paragraph):
                body.remove(paragraph)
            else:
                break
        self.logger.warning("Existing 'Om klarspråkningen' section found and replaced")
        return True

    @staticmethod
    def _is_empty_paragraph(paragraph: etree._Element) -> bool:
        has_text = any((t.text or "").strip() for t in paragraph.iter(f"{W}t"))
        has_objects = any(
            True for tag in ("drawing", "pict", "object", "fldChar", "instrText")
            for _ in paragraph.iter(f"{W}{tag}")
        )
        return not has_text and not has_objects

    # ------------------------------------------------------------------
    # Rubrikstil
    # ------------------------------------------------------------------

    def _resolve_heading_style(self, body: etree._Element) -> Optional[str]:
        """
        Rubrikstil på nivå 1, i prioritetsordning. Onumrerade stilar föredras
        eftersom de ser ut som avsnitt av typen Källor/Bilaga; en numrerad stil
        används bara om ingen onumrerad finns, och numreringen spärras då
        direkt i stycket (se _build_paragraph).
        1. stilen på en rubrik som Källor/Bilaga/Referenser i dokumentet
        2. första nivå 1-rubrik i brödtexten (ej bildtext)
        3. inbyggd Heading1/Rubrik1 om den finns
        Annars None, vilket ger fet text utan rubrikstil.
        """
        styles = self._style_index()
        if not styles:
            return None
        numbered_by_list = self._styles_numbered_by_numbering_part()

        def info_for(style_id: Optional[str]) -> Optional[dict]:
            info = self._resolved(styles, style_id, numbered_by_list)
            if info is None or info["outline"] != 0:
                return None
            name = f"{style_id} {info['name'] or ''}".lower()
            if any(token in name for token in self.CAPTION_TOKENS):
                return None
            return info

        # (prioritet, onumrerad först, dokumentordning) -> style_id
        candidates: list[tuple[int, int, int, str]] = []
        for order, paragraph in enumerate(body.findall(f"{W}p")):
            style_el = paragraph.find(f"{W}pPr/{W}pStyle")
            style_id = style_el.get(f"{W}val") if style_el is not None else None
            info = info_for(style_id)
            if info is None:
                continue
            text = "".join(t.text or "" for t in paragraph.iter(f"{W}t"))
            if not text.strip():
                continue
            priority = 0 if self.SOURCE_HEADING_RE.match(text) else 1
            candidates.append((int(info["numbered"]), priority, order, style_id))

        for order, candidate in enumerate(("Heading1", "Rubrik1")):
            info = info_for(candidate)
            if info is not None:
                candidates.append((int(info["numbered"]), 2, order, candidate))

        return min(candidates)[3] if candidates else None

    def _styles_numbered_by_numbering_part(self) -> set[str]:
        """Stilar som numreras via numbering.xml (w:lvl/w:pStyle), inte via styles.xml."""
        if not self.package.part_exists("word/numbering.xml"):
            return set()
        try:
            root = self.package.read_xml_root("word/numbering.xml")
        except Exception as ex:
            self.logger.warning(f"About section: could not read numbering.xml: {ex}")
            return set()
        return {
            style.get(f"{W}val") for style in root.iter(f"{W}pStyle")
            if style.getparent() is not None and style.getparent().tag == f"{W}lvl"
            and style.get(f"{W}val")
        }

    def _style_index(self) -> dict[str, dict]:
        try:
            root = self.package.read_styles_tree().getroot()
        except Exception as ex:
            self.logger.warning(f"About section: could not read styles.xml: {ex}")
            return {}
        index = {}
        for style in root.findall(f"{W}style"):
            if style.get(f"{W}type") != "paragraph" or not style.get(f"{W}styleId"):
                continue
            name = style.find(f"{W}name")
            based_on = style.find(f"{W}basedOn")
            outline = style.find(f"{W}pPr/{W}outlineLvl")
            num_id = style.find(f"{W}pPr/{W}numPr/{W}numId")
            index[style.get(f"{W}styleId")] = {
                "name": name.get(f"{W}val") if name is not None else None,
                "based_on": based_on.get(f"{W}val") if based_on is not None else None,
                "outline": int(outline.get(f"{W}val")) if outline is not None else None,
                "num_id": num_id.get(f"{W}val") if num_id is not None else None,
            }
        return index

    @staticmethod
    def _resolved(styles: dict, style_id: Optional[str], numbered_by_list: frozenset = frozenset()) -> Optional[dict]:
        if not style_id or style_id not in styles:
            return None
        outline = num_id = None
        linked = False
        current, seen = style_id, set()
        while current and current in styles and current not in seen:
            seen.add(current)
            info = styles[current]
            if outline is None:
                outline = info["outline"]
            if num_id is None:
                num_id = info["num_id"]
            linked = linked or current in numbered_by_list
            current = info["based_on"]
        return {
            "name": styles[style_id]["name"],
            "outline": outline,
            "numbered": (num_id not in (None, "0")) or (linked and num_id != "0"),
        }

    # ------------------------------------------------------------------
    # XML-byggare
    # ------------------------------------------------------------------

    def _build_paragraph(self, block: AboutBlock, heading_style: Optional[str]) -> etree._Element:
        paragraph = etree.Element(f"{W}p")
        is_heading = block.kind == "heading"

        if block.keep_with_next and not is_heading:
            # Tabellrubrik: hålls ihop med tabellen och får luft ovanför
            ppr = etree.SubElement(paragraph, f"{W}pPr")
            etree.SubElement(ppr, f"{W}keepNext")
            etree.SubElement(ppr, f"{W}spacing", {f"{W}before": "240"})

        if is_heading:
            ppr = etree.SubElement(paragraph, f"{W}pPr")
            if heading_style:
                etree.SubElement(ppr, f"{W}pStyle").set(f"{W}val", heading_style)
            etree.SubElement(ppr, f"{W}keepNext")
            etree.SubElement(ppr, f"{W}pageBreakBefore")
            # Regel 3: rubriken är alltid onumrerad, oavsett stil och
            # numbering.xml. numId 0 stänger av numreringen för stycket.
            num_pr = etree.SubElement(ppr, f"{W}numPr")
            etree.SubElement(num_pr, f"{W}numId").set(f"{W}val", "0")

        for text, bold in block.runs:
            run = etree.SubElement(paragraph, f"{W}r")
            if bold or (is_heading and not heading_style):
                rpr = etree.SubElement(run, f"{W}rPr")
                etree.SubElement(rpr, f"{W}b")
                if is_heading:
                    etree.SubElement(rpr, f"{W}sz").set(f"{W}val", "32")
            t = etree.SubElement(run, f"{W}t")
            t.set(f"{{{XML_NS}}}space", "preserve")
            t.text = text
        return paragraph

    # Tabellens totala bredd i twips (A4 med normala marginaler) och den första
    # kolumnens andel beroende på antal kolumner; övriga kolumner delar resten.
    TABLE_WIDTH_TWIPS = 9000
    FIRST_COLUMN_SHARE = {1: 1.0, 2: 0.75, 3: 0.55}
    FIRST_COLUMN_SHARE_MANY = 0.40

    def _build_table(self, block: AboutBlock) -> etree._Element:
        """
        Enkel tabell med tunna ramar och fet rubrikrad som upprepas vid
        sidbrytning. Inga tabellformat används, så utseendet blir detsamma i
        alla mallar.
        """
        columns = max(len(row) for row in block.rows)
        share = self.FIRST_COLUMN_SHARE.get(columns, self.FIRST_COLUMN_SHARE_MANY)
        first = int(self.TABLE_WIDTH_TWIPS * share)
        rest = (self.TABLE_WIDTH_TWIPS - first) // max(columns - 1, 1)
        widths = [first] + [rest] * (columns - 1)

        table = etree.Element(f"{W}tbl")
        tbl_pr = etree.SubElement(table, f"{W}tblPr")
        etree.SubElement(tbl_pr, f"{W}tblW", {f"{W}w": "5000", f"{W}type": "pct"})
        borders = etree.SubElement(tbl_pr, f"{W}tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            etree.SubElement(borders, f"{W}{side}", {
                f"{W}val": "single", f"{W}sz": "4", f"{W}space": "0", f"{W}color": "auto",
            })
        margins = etree.SubElement(tbl_pr, f"{W}tblCellMar")
        for side in ("left", "right"):
            etree.SubElement(margins, f"{W}{side}", {f"{W}w": "80", f"{W}type": "dxa"})
        etree.SubElement(tbl_pr, f"{W}tblLook", {
            f"{W}val": "04A0", f"{W}firstRow": "1", f"{W}lastRow": "0",
            f"{W}firstColumn": "1", f"{W}lastColumn": "0", f"{W}noHBand": "0", f"{W}noVBand": "1",
        })
        grid = etree.SubElement(table, f"{W}tblGrid")
        for width in widths:
            etree.SubElement(grid, f"{W}gridCol", {f"{W}w": str(width)})

        for row_index, row in enumerate(block.rows):
            tr = etree.SubElement(table, f"{W}tr")
            tr_pr = etree.SubElement(tr, f"{W}trPr")
            etree.SubElement(tr_pr, f"{W}cantSplit")
            if row_index == 0:
                etree.SubElement(tr_pr, f"{W}tblHeader")
            for column in range(columns):
                text = row[column] if column < len(row) else ""
                tc = etree.SubElement(tr, f"{W}tc")
                tc_pr = etree.SubElement(tc, f"{W}tcPr")
                etree.SubElement(tc_pr, f"{W}tcW", {f"{W}w": str(widths[column]), f"{W}type": "dxa"})
                paragraph = etree.SubElement(tc, f"{W}p")
                p_pr = etree.SubElement(paragraph, f"{W}pPr")
                etree.SubElement(p_pr, f"{W}spacing", {f"{W}before": "20", f"{W}after": "20"})
                if column in block.numeric_columns:
                    etree.SubElement(p_pr, f"{W}jc", {f"{W}val": "right"})
                run = etree.SubElement(paragraph, f"{W}r")
                if row_index == 0:
                    etree.SubElement(etree.SubElement(run, f"{W}rPr"), f"{W}b")
                t = etree.SubElement(run, f"{W}t")
                t.set(f"{{{XML_NS}}}space", "preserve")
                t.text = text
        return table

    def _mark_table_inserted(self, table: etree._Element, timestamp: str) -> None:
        """Spårad infogning av en hel tabell: varje rad, stycke och körning."""
        for tr in table.findall(f"{W}tr"):
            tr_pr = tr.find(f"{W}trPr")
            if tr_pr is None:
                tr_pr = etree.Element(f"{W}trPr")
                tr.insert(0, tr_pr)
            tr_pr.append(self._revision_element("ins", timestamp))
            for paragraph in tr.iter(f"{W}p"):
                self._wrap_runs_in_insertion(paragraph, timestamp)
                self._mark_paragraph_inserted(paragraph, timestamp)

    def _add_bookmark(self, first: etree._Element, last: etree._Element) -> None:
        tree_root = self.package.read_document_tree().getroot()
        ids = [
            int(b.get(f"{W}id")) for b in tree_root.iter(f"{W}bookmarkStart")
            if (b.get(f"{W}id") or "").isdigit()
        ]
        bookmark_id = str(max(ids, default=0) + 1)

        start = etree.Element(f"{W}bookmarkStart")
        start.set(f"{W}id", bookmark_id)
        start.set(f"{W}name", ABOUT_BOOKMARK_NAME)
        end = etree.Element(f"{W}bookmarkEnd")
        end.set(f"{W}id", bookmark_id)

        # I spårat läge ligger bokmärket inne i w:ins, så att det försvinner
        # tillsammans med texten om infogningen avvisas. Annars kunde ett tomt
        # bokmärke bli kvar i baksidans sista stycke och få nästa körning att
        # hoppa över det.
        first_ins = first.find(f"{W}ins")
        if first_ins is not None:
            first_ins.insert(0, start)
        else:
            ppr = first.find(f"{W}pPr")
            if ppr is not None:
                ppr.addnext(start)
            else:
                first.insert(0, start)

        last_ins = last.findall(f"{W}ins")
        (last_ins[-1] if last_ins else last).append(end)

    def _wrap_runs_in_insertion(self, paragraph: etree._Element, timestamp: str) -> None:
        runs = paragraph.findall(f"{W}r")
        if not runs:
            return
        ins = self._revision_element("ins", timestamp)
        runs[0].addprevious(ins)
        for run in runs:
            ins.append(run)

    def _mark_paragraph_inserted(self, paragraph: etree._Element, timestamp: str) -> None:
        ppr = paragraph.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.Element(f"{W}pPr")
            paragraph.insert(0, ppr)
        rpr = ppr.find(f"{W}rPr")
        if rpr is None:
            rpr = etree.Element(f"{W}rPr")
            tail = next((child for child in ppr if child.tag in _PPR_TAIL_TAGS), None)
            if tail is not None:
                tail.addprevious(rpr)
            else:
                ppr.append(rpr)
        if rpr.find(f"{W}ins") is None:
            rpr.insert(0, self._revision_element("ins", timestamp))

    def _revision_element(self, tag: str, timestamp: str) -> etree._Element:
        element = etree.Element(f"{W}{tag}")
        element.set(f"{W}id", str(self._next_id))
        self._next_id += 1
        element.set(f"{W}author", self.author)
        element.set(f"{W}date", timestamp)
        return element

    @staticmethod
    def _can_extend_last_paragraph(block: Optional[etree._Element]) -> bool:
        """Sista blocket är ett vanligt stycke med styckeformat som kan kopieras."""
        if block is None or block.tag != f"{W}p":
            return False
        ppr = block.find(f"{W}pPr")
        if ppr is None:
            return True
        # Ett stycke som avslutar ett Word-avsnitt får inte sin sectPr kopierad.
        return ppr.find(f"{W}sectPr") is None and ppr.find(f"{W}pPrChange") is None

    @staticmethod
    def _last_block(body: etree._Element) -> Optional[etree._Element]:
        blocks = [child for child in body if child.tag in (f"{W}p", f"{W}tbl", f"{W}sdt")]
        return blocks[-1] if blocks else None
