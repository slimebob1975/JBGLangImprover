import os
import shutil
import zipfile
from tempfile import mkdtemp
from typing import Optional
from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OD_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

NSMAP = {
    "w": W_NS,
    "rel": REL_NS,
}


class DocxPackage:
    """
    Ren OOXML-wrapper för .docx-paket.

    Ansvar:
    - extrahera .docx till temporär katalog
    - läsa XML-delar som lxml-träd
    - skriva tillbaka XML-delar
    - skapa saknade delar vid behov
    - packa ihop till .docx igen
    - städa temporära filer
    """

    def __init__(self, docx_path: str, logger):
        self.docx_path = os.path.abspath(docx_path)
        self.logger = logger
        self.temp_dir: Optional[str] = None
        self.is_open = False

    # ------------------------------------------------------------------
    # Livscykel
    # ------------------------------------------------------------------

    def open(self) -> None:
        if self.is_open:
            return

        if not os.path.exists(self.docx_path):
            raise FileNotFoundError(f"DOCX file not found: {self.docx_path}")

        if not zipfile.is_zipfile(self.docx_path):
            raise ValueError(f"File is not a valid DOCX/ZIP archive: {self.docx_path}")

        self.temp_dir = mkdtemp(prefix="jbg_docxpkg_")
        with zipfile.ZipFile(self.docx_path, "r") as zin:
            zin.extractall(self.temp_dir)

        self.is_open = True
        self.logger.info(f"Opened DOCX package into temp dir: {self.temp_dir}")

    def close(self, cleanup: bool = True) -> None:
        if cleanup and self.temp_dir and os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)

        self.temp_dir = None
        self.is_open = False

    def __enter__(self):
        self.open()
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close(cleanup=True)

    # ------------------------------------------------------------------
    # Basala paths
    # ------------------------------------------------------------------

    def get_abs_part_path(self, part_name: str) -> str:
        self._ensure_open()
        part_name = part_name.lstrip("/")
        return os.path.join(self.temp_dir, part_name)

    def part_exists(self, part_name: str) -> bool:
        return os.path.exists(self.get_abs_part_path(part_name))

    # ------------------------------------------------------------------
    # XML-läsning / skrivning
    # ------------------------------------------------------------------

    def read_xml_tree(self, part_name: str, create_if_missing: bool = False,
                      root_tag: Optional[str] = None,
                      nsmap: Optional[dict] = None) -> etree._ElementTree:
        """
        Läser en XML-del som ElementTree.
        Om create_if_missing=True kan en minimal rot skapas.
        """
        abs_path = self.get_abs_part_path(part_name)

        if not os.path.exists(abs_path):
            if not create_if_missing:
                raise FileNotFoundError(f"Missing XML part: {part_name}")

            if not root_tag:
                raise ValueError("root_tag is required when create_if_missing=True")

            os.makedirs(os.path.dirname(abs_path), exist_ok=True)
            root = etree.Element(root_tag, nsmap=nsmap)
            tree = etree.ElementTree(root)
            tree.write(abs_path, pretty_print=True, xml_declaration=True, encoding="UTF-8")
            self.logger.info(f"Created missing XML part: {part_name}")

        parser = etree.XMLParser(remove_blank_text=False, ns_clean=True, recover=True)
        return etree.parse(abs_path, parser)

    def write_xml_tree(self, part_name: str, tree: etree._ElementTree) -> None:
        abs_path = self.get_abs_part_path(part_name)
        os.makedirs(os.path.dirname(abs_path), exist_ok=True)
        tree.write(abs_path, pretty_print=True, xml_declaration=True, encoding="UTF-8")
        self.logger.info(f"Wrote XML part: {part_name}")

    def read_xml_root(self, part_name: str, create_if_missing: bool = False,
                      root_tag: Optional[str] = None,
                      nsmap: Optional[dict] = None) -> etree._Element:
        return self.read_xml_tree(
            part_name=part_name,
            create_if_missing=create_if_missing,
            root_tag=root_tag,
            nsmap=nsmap,
        ).getroot()

    # ------------------------------------------------------------------
    # Byte- / text-access
    # ------------------------------------------------------------------

    def read_bytes(self, part_name: str) -> bytes:
        abs_path = self.get_abs_part_path(part_name)
        if not os.path.exists(abs_path):
            raise FileNotFoundError(f"Missing part: {part_name}")
        with open(abs_path, "rb") as f:
            return f.read()

    def write_bytes(self, part_name: str, data: bytes) -> None:
        abs_path = self.get_abs_part_path(part_name)
        os.makedirs(os.path.dirname(abs_path), exist_ok=True)
        with open(abs_path, "wb") as f:
            f.write(data)
        self.logger.info(f"Wrote binary part: {part_name}")

    def read_text(self, part_name: str, encoding: str = "utf-8") -> str:
        return self.read_bytes(part_name).decode(encoding)

    def write_text(self, part_name: str, text: str, encoding: str = "utf-8") -> None:
        self.write_bytes(part_name, text.encode(encoding))

    # ------------------------------------------------------------------
    # Part-helpers
    # ------------------------------------------------------------------

    def list_parts(self) -> list[str]:
        self._ensure_open()
        parts = []
        for root_dir, _, files in os.walk(self.temp_dir):
            for file in files:
                full_path = os.path.join(root_dir, file)
                rel_path = os.path.relpath(full_path, self.temp_dir).replace("\\", "/")
                parts.append(rel_path)
        return sorted(parts)

    def ensure_part(self, part_name: str, data: bytes = b"") -> None:
        abs_path = self.get_abs_part_path(part_name)
        if not os.path.exists(abs_path):
            os.makedirs(os.path.dirname(abs_path), exist_ok=True)
            with open(abs_path, "wb") as f:
                f.write(data)
            self.logger.info(f"Ensured missing part: {part_name}")

    def remove_part(self, part_name: str) -> None:
        abs_path = self.get_abs_part_path(part_name)
        if os.path.exists(abs_path):
            os.remove(abs_path)
            self.logger.info(f"Removed part: {part_name}")

    # ------------------------------------------------------------------
    # Standard Word-delar
    # ------------------------------------------------------------------

    def read_document_tree(self) -> etree._ElementTree:
        return self.read_xml_tree("word/document.xml")

    def write_document_tree(self, tree: etree._ElementTree) -> None:
        self.write_xml_tree("word/document.xml", tree)

    def read_footnotes_tree(self, create_if_missing: bool = False) -> etree._ElementTree:
        return self.read_xml_tree(
            "word/footnotes.xml",
            create_if_missing=create_if_missing,
            root_tag=f"{{{W_NS}}}footnotes",
            nsmap={"w": W_NS},
        )

    def write_footnotes_tree(self, tree: etree._ElementTree) -> None:
        self.write_xml_tree("word/footnotes.xml", tree)

    def read_comments_tree(self, create_if_missing: bool = False) -> etree._ElementTree:
        return self.read_xml_tree(
            "word/comments.xml",
            create_if_missing=create_if_missing,
            root_tag=f"{{{W_NS}}}comments",
            nsmap={"w": W_NS},
        )

    def write_comments_tree(self, tree: etree._ElementTree) -> None:
        self.write_xml_tree("word/comments.xml", tree)

    def read_document_rels_tree(self, create_if_missing: bool = False) -> etree._ElementTree:
        return self.read_xml_tree(
            "word/_rels/document.xml.rels",
            create_if_missing=create_if_missing,
            root_tag=f"{{{REL_NS}}}Relationships",
            nsmap={None: REL_NS},
        )

    def write_document_rels_tree(self, tree: etree._ElementTree) -> None:
        self.write_xml_tree("word/_rels/document.xml.rels", tree)

    def read_content_types_tree(self) -> etree._ElementTree:
        return self.read_xml_tree("[Content_Types].xml")

    def write_content_types_tree(self, tree: etree._ElementTree) -> None:
        self.write_xml_tree("[Content_Types].xml", tree)

    def read_settings_tree(self, create_if_missing: bool = False) -> etree._ElementTree:
        return self.read_xml_tree(
            "word/settings.xml",
            create_if_missing=create_if_missing,
            root_tag=f"{{{W_NS}}}settings",
            nsmap={"w": W_NS},
        )

    def write_settings_tree(self, tree: etree._ElementTree) -> None:
        self.write_xml_tree("word/settings.xml", tree)

    def read_styles_tree(self, create_if_missing: bool = False) -> etree._ElementTree:
        return self.read_xml_tree(
            "word/styles.xml",
            create_if_missing=create_if_missing,
            root_tag=f"{{{W_NS}}}styles",
            nsmap={"w": W_NS},
        )

    def write_styles_tree(self, tree: etree._ElementTree) -> None:
        self.write_xml_tree("word/styles.xml", tree)

    # ------------------------------------------------------------------
    # Relationship helpers
    # ------------------------------------------------------------------

    def ensure_office_relationship(
        self,
        rel_type_suffix: str,
        target: str,
        rel_id: Optional[str] = None,
    ) -> str:
        """
        Säkerställer att word/_rels/document.xml.rels innehåller en relation.
        rel_type_suffix ex: 'comments', 'footnotes', 'styles', 'settings'
        """
        tree = self.read_document_rels_tree(create_if_missing=True)
        root = tree.getroot()

        wanted_type = f"{OD_REL_NS}/{rel_type_suffix}"

        for rel in root.findall(f"{{{REL_NS}}}Relationship"):
            if rel.get("Type") == wanted_type and rel.get("Target") == target:
                return rel.get("Id")

        existing_ids = {
            rel.get("Id") for rel in root.findall(f"{{{REL_NS}}}Relationship")
        }

        if rel_id is None:
            i = 1
            while f"rId{i}" in existing_ids:
                i += 1
            rel_id = f"rId{i}"

        new_rel = etree.SubElement(root, f"{{{REL_NS}}}Relationship")
        new_rel.set("Id", rel_id)
        new_rel.set("Type", wanted_type)
        new_rel.set("Target", target)

        self.write_document_rels_tree(tree)
        self.logger.info(f"Added relationship: {wanted_type} -> {target}")
        return rel_id

    # ------------------------------------------------------------------
    # Content types helpers
    # ------------------------------------------------------------------

    def ensure_content_type_override(self, part_name: str, content_type: str) -> None:
        tree = self.read_content_types_tree()
        root = tree.getroot()

        overrides = root.findall(f"{{{CT_NS}}}Override")
        wanted_part_name = part_name if part_name.startswith("/") else f"/{part_name}"

        for ov in overrides:
            if ov.get("PartName") == wanted_part_name:
                return

        new_ov = etree.SubElement(root, f"{{{CT_NS}}}Override")
        new_ov.set("PartName", wanted_part_name)
        new_ov.set("ContentType", content_type)

        self.write_content_types_tree(tree)
        self.logger.info(f"Added content type override: {wanted_part_name}")

    # ------------------------------------------------------------------
    # Spara tillbaka .docx
    # ------------------------------------------------------------------

    def save(self, output_path: Optional[str] = None) -> str:
        self._ensure_open()

        if output_path is None:
            base, ext = os.path.splitext(self.docx_path)
            output_path = f"{base}_rebuilt{ext}"

        output_path = os.path.abspath(output_path)

        with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for root_dir, _, files in os.walk(self.temp_dir):
                for file in files:
                    full_path = os.path.join(root_dir, file)
                    arc_name = os.path.relpath(full_path, self.temp_dir).replace("\\", "/")
                    zout.write(full_path, arc_name)

        self.logger.info(f"Saved DOCX package to: {output_path}")
        return output_path

    # ------------------------------------------------------------------
    # Interna hjälpare
    # ------------------------------------------------------------------

    def _ensure_open(self) -> None:
        if not self.is_open or not self.temp_dir:
            raise RuntimeError("DocxPackage is not open. Call open() first.")