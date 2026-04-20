from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Optional

from lxml import etree

try:
    from app.src.JBGDocxPackage import W_NS, DocxPackage
except ModuleNotFoundError:
    from JBGDocxPackage import W_NS, DocxPackage


XML_NS = "http://www.w3.org/XML/1998/namespace"


@dataclass
class CommentRenderResult:
    plan: object
    applied: bool
    message: str
    comment_id: Optional[int] = None


class CommentsRenderer:
    def __init__(self, package: DocxPackage, logger, author: str = "JBG Klarspråkningstjänst", initials: str = "JBG"):
        self.package = package
        self.logger = logger
        self.author = author
        self.initials = initials

        self.comments_tree = self.package.read_comments_tree(create_if_missing=True)
        self.comments_root = self.comments_tree.getroot()

        self.package.ensure_office_relationship("comments", "comments.xml")
        self.package.ensure_content_type_override(
            "word/comments.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        )

        self.comment_id_counter = self._get_next_comment_id()

    # ------------------------------------------------------------------
    # Publikt API
    # ------------------------------------------------------------------

    def apply_comments_for_results(self, render_results) -> list[CommentRenderResult]:
        results: list[CommentRenderResult] = []

        for render_result in render_results:
            plan = render_result.plan

            if not render_result.applied:
                results.append(CommentRenderResult(
                    plan=plan,
                    applied=False,
                    message="Skipped because tracked change was not applied",
                    comment_id=None,
                ))
                continue

            motivation = getattr(plan, "motivation", None)
            if not motivation:
                results.append(CommentRenderResult(
                    plan=plan,
                    applied=False,
                    message="No motivation on plan",
                    comment_id=None,
                ))
                continue

            anchor_part_name = getattr(render_result, "anchor_part_name", None)
            anchor_kind = getattr(render_result, "anchor_kind", None)
            anchor_revision_id = getattr(render_result, "anchor_revision_id", None)

            if not anchor_revision_id:
                results.append(CommentRenderResult(
                    plan=plan,
                    applied=False,
                    message="RenderResult has no anchor_revision_id",
                    comment_id=None,
                ))
                continue

            if anchor_part_name != "word/document.xml":
                results.append(CommentRenderResult(
                    plan=plan,
                    applied=False,
                    message=f"CommentsRenderer v3 supports only word/document.xml anchors, got {anchor_part_name}",
                    comment_id=None,
                ))
                continue

            try:
                anchor_element = self._find_anchor_by_revision_id(
                    part_name=anchor_part_name,
                    anchor_kind=anchor_kind,
                    revision_id=anchor_revision_id,
                )
                if anchor_element is None:
                    raise ValueError(
                        f"Could not find anchor {anchor_kind} with revision id={anchor_revision_id}"
                    )

                comment_id = self._attach_comment_to_anchor(anchor_element, motivation)
                self.package.write_comments_tree(self.comments_tree)
                self.package.write_document_tree(anchor_element.getroottree())

                results.append(CommentRenderResult(
                    plan=plan,
                    applied=True,
                    message="Comment applied",
                    comment_id=comment_id,
                ))
            except Exception as ex:
                results.append(CommentRenderResult(
                    plan=plan,
                    applied=False,
                    message=str(ex),
                    comment_id=None,
                ))

        return results

    # ------------------------------------------------------------------
    # Hitta anchor i aktuellt träd
    # ------------------------------------------------------------------

    def _find_anchor_by_revision_id(
        self,
        part_name: str,
        anchor_kind: str,
        revision_id: str,
    ) -> Optional[etree._Element]:
        if part_name != "word/document.xml":
            return None

        tree = self.package.read_document_tree()
        root = tree.getroot()

        tag = "ins" if anchor_kind == "ins" else "del"
        xpath = f".//w:{tag}[@w:id='{revision_id}']"
        matches = root.xpath(xpath, namespaces={"w": W_NS})

        if not matches:
            return None

        return matches[0]

    # ------------------------------------------------------------------
    # Huvudlogik
    # ------------------------------------------------------------------

    def _attach_comment_to_anchor(
        self,
        anchor_element: etree._Element,
        motivation: str,
    ) -> int:
        paragraph = self._find_ancestor_paragraph(anchor_element)
        if paragraph is None:
            raise ValueError("Anchor element is not inside a paragraph")

        parent = anchor_element.getparent()
        if parent is None:
            raise ValueError("Anchor element has no parent")

        comment_id = self._create_comment(motivation)

        start = etree.Element(f"{{{W_NS}}}commentRangeStart")
        start.set(f"{{{W_NS}}}id", str(comment_id))

        end = etree.Element(f"{{{W_NS}}}commentRangeEnd")
        end.set(f"{{{W_NS}}}id", str(comment_id))

        ref_run = etree.Element(f"{{{W_NS}}}r")
        ref = etree.SubElement(ref_run, f"{{{W_NS}}}commentReference")
        ref.set(f"{{{W_NS}}}id", str(comment_id))

        anchor_index = parent.index(anchor_element)
        parent.insert(anchor_index, start)
        parent.insert(anchor_index + 2, end)

        paragraph_children = list(paragraph)
        if parent in paragraph_children:
            paragraph_index = paragraph_children.index(parent)
            paragraph.insert(paragraph_index + 1, ref_run)
        else:
            paragraph.append(ref_run)

        return comment_id

    def _find_ancestor_paragraph(self, element: etree._Element) -> Optional[etree._Element]:
        current = element
        while current is not None:
            if current.tag == f"{{{W_NS}}}p":
                return current
            current = current.getparent()
        return None

    # ------------------------------------------------------------------
    # Kommentarobjekt
    # ------------------------------------------------------------------

    def _create_comment(self, text: str) -> int:
        comment_id = self.comment_id_counter
        self.comment_id_counter += 1

        comment = etree.SubElement(self.comments_root, f"{{{W_NS}}}comment")
        comment.set(f"{{{W_NS}}}id", str(comment_id))
        comment.set(f"{{{W_NS}}}author", self.author)
        comment.set(f"{{{W_NS}}}initials", self.initials)
        comment.set(f"{{{W_NS}}}date", self._current_word_timestamp())

        p = etree.SubElement(comment, f"{{{W_NS}}}p")
        r = etree.SubElement(p, f"{{{W_NS}}}r")
        t = etree.SubElement(r, f"{{{W_NS}}}t")
        t.set(f"{{{XML_NS}}}space", "preserve")
        t.text = text

        return comment_id

    def _get_next_comment_id(self) -> int:
        max_id = -1
        for comment in self.comments_root.findall(f"{{{W_NS}}}comment"):
            try:
                max_id = max(max_id, int(comment.get(f"{{{W_NS}}}id")))
            except Exception:
                continue
        return max_id + 1

    def _current_word_timestamp(self) -> str:
        return datetime.now(timezone.utc).replace(microsecond=0).isoformat()