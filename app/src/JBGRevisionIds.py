"""
Id för spårade ändringar (w:ins, w:del m.fl.) ska vara unika i hela dokumentet.
Nya ändringar numreras därför från det högsta id som redan finns i någon
del av dokumentet, så att de inte krockar med ändringar som fanns i originalet.
"""

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"

REVISION_TAGS = {
    f"{W}{name}" for name in (
        "ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange",
        "sectPrChange", "tblPrChange", "trPrChange", "tcPrChange",
        "tblGridChange", "numberingChange", "cellIns", "cellDel", "cellMerge",
    )
}


def max_revision_id(package) -> int:
    """Högsta id bland spårade ändringar i alla word/*.xml-delar (0 om inga finns)."""
    highest = 0
    for part_name in package.list_parts():
        if not (part_name.startswith("word/") and part_name.endswith(".xml")):
            continue
        try:
            root = package.read_xml_root(part_name)
        except Exception:
            continue
        for element in root.iter():
            if element.tag in REVISION_TAGS:
                value = element.get(f"{W}id") or ""
                if value.isdigit():
                    highest = max(highest, int(value))
    return highest
