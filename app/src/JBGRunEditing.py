"""
Gemensamma byggstenar för att ändra text i ett stycke, för båda renderarna
(spåra ändringar och enkel färgmarkering).

Tre regler gäller för alla ändringar:

1. Positioner räknas i hela textbitens (w:r) synliga text, där en tabb är
   "\\t" och en radbrytning "\\n". En textbit kan innehålla flera delar
   (text, radbrytning, text); det som inte ändras ska stå kvar.
2. Ny text byggs med riktiga element: "\\n" blir w:br och "\\t" blir w:tab,
   så att radbrytningar och tabbar inte blir vanliga tecken i texten.
3. En ändring kan sträcka sig över textbitar i olika behållare, t.ex. direkt
   i stycket och inuti en hyperlänk. Ändringen delas då per behållare, och
   varje del ändras i sin egen behållare, så att länken finns kvar.
"""

import difflib
from dataclasses import dataclass, field
from typing import Optional

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_NS = "http://www.w3.org/XML/1998/namespace"
W = f"{{{W_NS}}}"

# Behållare vars textbitar kan ändras: stycket självt och hyperlänkar i det.
EDITABLE_CONTAINERS = {f"{W}p", f"{W}hyperlink"}


# ============================================================================
# Synlig text i en textbit
# ============================================================================

def visible_text_from_run_child(child: etree._Element) -> str:
    if child.tag == f"{W}t":
        return child.text or ""
    if child.tag == f"{W}tab":
        return "\t"
    if child.tag in (f"{W}br", f"{W}cr"):
        return "\n"
    return ""


def run_visible_text(run: etree._Element) -> str:
    return "".join(visible_text_from_run_child(child) for child in run)


def local_offset_in_run(node, global_offset: int) -> int:
    """Styckets position -> position i textbitens hela synliga text."""
    run = node.run_element
    if run is None:
        raise ValueError("Cannot calculate a local offset without run_element")
    if global_offset < node.start or global_offset > node.end:
        raise ValueError("Global offset is outside the boundary node")
    local_offset = 0
    for child in run:
        if child is node.element:
            return local_offset + (global_offset - node.start)
        local_offset += len(visible_text_from_run_child(child))
    raise ValueError("Boundary node is not a child of its run")


def append_visible_text(run: etree._Element, text: str, text_tag: str = "t") -> None:
    """Lägger till text i en textbit; "\\n" blir w:br och "\\t" blir w:tab."""
    buffer = []

    def flush():
        if buffer:
            element = etree.SubElement(run, f"{W}{text_tag}")
            element.set(f"{{{XML_NS}}}space", "preserve")
            element.text = "".join(buffer)
            buffer.clear()

    for ch in text:
        if ch == "\n":
            flush()
            etree.SubElement(run, f"{W}br")
        elif ch == "\t":
            flush()
            etree.SubElement(run, f"{W}tab")
        else:
            buffer.append(ch)
    flush()


# ============================================================================
# Textbitar i läsordning och per behållare
# ============================================================================

def _nearest_paragraph(element: etree._Element) -> Optional[etree._Element]:
    return next(element.iterancestors(f"{W}p"), None)


def runs_in_order(paragraph: etree._Element) -> list[etree._Element]:
    """Styckets egna textbitar, även i hyperlänkar (inte i textrutor)."""
    return [
        run for run in paragraph.iter(f"{W}r")
        if run.getparent().tag in EDITABLE_CONTAINERS and _nearest_paragraph(run) is paragraph
    ]


def runs_between(paragraph: etree._Element, first_run: etree._Element, last_run: etree._Element) -> list[etree._Element]:
    runs = runs_in_order(paragraph)
    if first_run not in runs or last_run not in runs:
        return []
    start, end = runs.index(first_run), runs.index(last_run)
    return runs[start:end + 1] if start <= end else []


def split_by_container(runs: list[etree._Element]) -> list[list[etree._Element]]:
    """Följande textbitar med samma förälder bildar en del."""
    segments: list[list[etree._Element]] = []
    for run in runs:
        if segments and segments[-1][-1].getparent() is run.getparent() \
                and segments[-1][-1].getnext() is run:
            segments[-1].append(run)
        else:
            segments.append([run])
    return segments


def map_boundaries(old: str, new: str, boundaries: list[int]) -> list[int]:
    """
    Var i den nya texten varje gräns i den gamla texten hamnar. Gränser i
    oförändrad text följer texten exakt; en gräns inne i en ändrad del läggs
    efter den delens nya text, så att en ändring hamnar i den behållare där
    den börjar.
    """
    opcodes = difflib.SequenceMatcher(None, old, new, autojunk=False).get_opcodes()
    mapped = []
    for boundary in boundaries:
        position = len(new)
        for tag, i1, i2, j1, j2 in opcodes:
            if i1 <= boundary < i2:
                position = j1 + (boundary - i1) if tag == "equal" else j2
                break
        mapped.append(position)
    # Gränserna måste vara stigande
    for index in range(1, len(mapped)):
        mapped[index] = max(mapped[index], mapped[index - 1])
    return mapped


# ============================================================================
# Plan för en ändring
# ============================================================================

@dataclass
class SpanSegment:
    container: etree._Element
    runs: list
    old: str
    new: str = ""
    before: str = ""          # text som behålls före ändringen (bara första delen)
    after: str = ""           # text som behålls efter ändringen (bara sista delen)
    first_run: Optional[etree._Element] = None
    last_run: Optional[etree._Element] = None


@dataclass
class SpanPlan:
    segments: list[SpanSegment] = field(default_factory=list)
    actual_old: str = ""


def plan_span_edit(paragraph, first_node, last_node, anchor_start: int, anchor_end: int, new_text: str) -> SpanPlan:
    """
    Delar upp en ändring av styckets text [anchor_start, anchor_end) i delar
    per behållare, med den text som ska behållas, tas bort och läggas till.
    """
    first_run, last_run = first_node.run_element, last_node.run_element
    if first_run is None or last_run is None:
        raise ValueError("Resolved text boundary node(s) without run_element")

    first_local = local_offset_in_run(first_node, anchor_start)
    last_local = local_offset_in_run(last_node, anchor_end)
    first_text, last_text = run_visible_text(first_run), run_visible_text(last_run)
    if not 0 <= first_local <= len(first_text):
        raise ValueError("Invalid first_local_start")
    if not 0 <= last_local <= len(last_text):
        raise ValueError("Invalid last_local_end")

    changed = [first_run] if first_run is last_run else runs_between(paragraph, first_run, last_run)
    if not changed:
        raise ValueError("Could not resolve changed run span")
    unsupported = [r for r in changed if r.getparent().tag not in EDITABLE_CONTAINERS]
    if unsupported:
        raise ValueError("Changed span contains runs in an unsupported container")

    segments = []
    for runs in split_by_container(changed):
        parts = []
        for run in runs:
            text = run_visible_text(run)
            start = first_local if run is first_run else 0
            end = last_local if run is last_run else len(text)
            parts.append(text[start:end])
        segments.append(SpanSegment(container=runs[0].getparent(), runs=runs, old="".join(parts),
                                    first_run=runs[0], last_run=runs[-1]))

    segments[0].before = first_text[:first_local]
    segments[-1].after = last_text[last_local:]
    actual_old = "".join(segment.old for segment in segments)

    if len(segments) == 1:
        segments[0].new = new_text
    else:
        cumulative, boundaries = 0, []
        for segment in segments[:-1]:
            cumulative += len(segment.old)
            boundaries.append(cumulative)
        cuts = [0] + map_boundaries(actual_old, new_text, boundaries) + [len(new_text)]
        for index, segment in enumerate(segments):
            segment.new = new_text[cuts[index]:cuts[index + 1]]

    return SpanPlan(segments=segments, actual_old=actual_old)


def replace_runs(segment: SpanSegment, new_elements: list[etree._Element]) -> None:
    """Ersätter delens textbitar med nya element på samma plats i behållaren."""
    container = segment.container
    index = container.index(segment.runs[0])
    for run in segment.runs:
        container.remove(run)
    for offset, element in enumerate(new_elements):
        container.insert(index + offset, element)
