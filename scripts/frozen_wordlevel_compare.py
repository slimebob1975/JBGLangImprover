"""Reproduce identical, frozen DOCX proposals across two repository revisions.

No AI calls, network, writes to user documents, or imports from a second revision
into this process. The baseline runs in its own temporary checkout via git archive.
"""
import argparse
import io
import json
import logging
import shutil
import subprocess
import sys
import tarfile
import tempfile
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{" + W_NS + "}"
NS = {"w": W_NS}


def fixture_path():
    return ROOT / "tests" / "fixtures" / "wordlevel_frozen_v1.json"


def view_text(paragraph, accept):
    chunks = []
    for el in paragraph:
        if el.tag not in (W + "r", W + ("ins" if accept else "del")):
            continue
        for child in el.iter():
            if child.tag in (W + "t", W + "delText"):
                chunks.append(child.text or "")
    return "".join(chunks)


def reference_locations(paragraph, accept):
    offset, locations = 0, []
    for el in paragraph:
        if el.tag == W + ("del" if accept else "ins"):
            continue
        for child in el.iter():
            if child.tag in (W + "t", W + "delText"):
                offset += len(child.text or "")
            elif child.tag == W + "footnoteReference":
                locations.append((child.get(W + "id"), offset))
    return locations


def revision_counts(paragraph):
    counts = {}
    for tag in ("ins", "del"):
        wrappers = paragraph.xpath("./w:" + tag, namespaces=NS)
        counts[tag] = len(wrappers)
        counts[tag + "_chars"] = sum(
            len(el.text or "") for wrapper in wrappers for el in wrapper.iter()
            if el.tag in (W + "t", W + "delText")
        )
    return counts


def run_case(case, temp, logger):
    from app.src.JBGChangePlanner import ChangePlanner
    from app.src.JBGCommentsRenderer import CommentsRenderer
    from app.src.JBGDocxPackage import DocxPackage
    from app.src.JBGDocumentStructureExtractor import DocumentStructureExtractor
    from app.src.JBGLangImprovSuggestorAI import SuggestedChange
    from app.src.JBGTokenDiffEngine import TokenDiffEngine
    from app.src.JBGTrackedChangesRenderer import TrackedChangesRenderer

    source, destination = temp / "src.docx", temp / "result.docx"
    doc = Document()
    p = doc.add_paragraph()
    for spec in case["runs"]:
        run = p.add_run(spec.get("text", ""))
        if spec.get("bold"):
            run.bold = True
        if "footnote_id" in spec:
            etree.SubElement(run._r, W + "footnoteReference").set(W + "id", spec["footnote_id"])
    doc.save(source)
    expected_old = "".join(item.get("text", "") for item in case["runs"])
    if expected_old != case["old"]:
        raise AssertionError(case["id"] + ": fixture old text mismatch")
    structure = DocumentStructureExtractor(str(source), logger).extract()
    matches = [x for x in structure["elements"] if x.get("type") == "paragraph" and x.get("text") == case["old"]]
    if len(matches) != 1:
        raise AssertionError(case["id"] + ": source paragraph not unique")
    change = SuggestedChange(element_type="paragraph", element_id=matches[0]["element_id"],
                             footnote_id=None, old=case["old"], new=case["new"],
                             motivation="Frozen deterministic replay", match_status="exact")
    plans = ChangePlanner(structure, TokenDiffEngine(), logger).build_plans([change])
    if len(plans) != 1:
        raise AssertionError(case["id"] + ": planner did not return exactly one proposal")
    with DocxPackage(str(source), logger) as pkg:
        results = TrackedChangesRenderer(pkg, logger, word_level_changes=True).apply_plans(plans)
        if len(results) != 1:
            raise AssertionError(case["id"] + ": no render result")
        result = results[0]
        if not result.applied:
            raise AssertionError(case["id"] + ": rendering failed: " + result.message)
        comments = CommentsRenderer(pkg, logger).apply_comments_for_results(results)
        if len(comments) != 1 or not comments[0].applied:
            raise AssertionError(case["id"] + ": comment was not attached")
        pkg.save(str(destination))
    with ZipFile(destination) as zip_file:
        root = etree.fromstring(zip_file.read("word/document.xml"))
        comment_root = etree.fromstring(zip_file.read("word/comments.xml"))
    paragraph = root.xpath(".//w:body/w:p", namespaces=NS)[0]
    if view_text(paragraph, False) != case["old"] or view_text(paragraph, True) != case["new"]:
        raise AssertionError(case["id"] + ": accepted/rejected text mismatch")
    original_reference_offsets = []
    position = 0
    for item in case["runs"]:
        if "footnote_id" in item:
            original_reference_offsets.append((item["footnote_id"], position))
        position += len(item.get("text", ""))
    if reference_locations(paragraph, False) != original_reference_offsets:
        raise AssertionError(case["id"] + ": original footnote moved")
    if len(reference_locations(paragraph, True)) != len(original_reference_offsets):
        raise AssertionError(case["id"] + ": accepted footnote missing")
    if len(comment_root.xpath("./w:comment", namespaces=NS)) != 1:
        raise AssertionError(case["id"] + ": comment count mismatch")
    for marker in ("commentRangeStart", "commentRangeEnd", "commentReference"):
        if len(paragraph.xpath(".//w:" + marker, namespaces=NS)) != 1:
            raise AssertionError(case["id"] + ": comment markers missing")
    if paragraph.xpath(".//w:ins//w:footnoteReference | .//w:del//w:footnoteReference", namespaces=NS):
        raise AssertionError(case["id"] + ": reference inside revision")
    revisions = paragraph.xpath(".//w:ins | .//w:del", namespaces=NS)
    ids = [r.get(W + "id") for r in revisions]
    if len(ids) != len(set(ids)):
        raise AssertionError(case["id"] + ": duplicate revision IDs")
    metrics = revision_counts(paragraph)
    return {
        "id": case["id"], "status": result.word_level_status,
        "fallback_reason": result.word_level_fallback_reason,
        "revisions": metrics, "reference_offsets_rejected": original_reference_offsets,
        "reference_offsets_accepted": reference_locations(paragraph, True),
        "invariants_passed": True
    }


def probe(fixtures):
    logger = logging.getLogger("frozen-wordlevel")
    logger.addHandler(logging.NullHandler())
    cases = fixtures["cases"]
    output = []
    with tempfile.TemporaryDirectory(prefix="frozen-wordlevel-") as tmp:
        for case in cases:
            case_dir = Path(tmp) / case["id"]
            case_dir.mkdir()
            output.append(run_case(case, case_dir, logger))
    return output


def baseline_probe(ref, fixture_file):
    with tempfile.TemporaryDirectory(prefix="frozen-baseline-") as tmp:
        baseline = Path(tmp)
        archive = subprocess.run(["git", "archive", "--format=tar", ref], cwd=ROOT,
                                 stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
        # Git archive is trusted from the selected local Git repository only.
        with tarfile.open(fileobj=io.BytesIO(archive.stdout)) as tar:
            tar.extractall(baseline, filter="data")
        target = baseline / "scripts" / "frozen_wordlevel_compare.py"
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(Path(__file__), target)
        fixture_dest = baseline / "tests" / "fixtures" / fixture_file.name
        fixture_dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(fixture_file, fixture_dest)
        runner = subprocess.run([sys.executable, str(target), "--probe-only"],
                                cwd=baseline, capture_output=True, text=True)
        if runner.returncode:
            raise RuntimeError("Baseline probe failed:\n" + runner.stderr[-4000:])
        return json.loads(runner.stdout)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline-ref", default=None, help="Local committed Git ref, e.g. HEAD")
    parser.add_argument("--probe-only", action="store_true")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    fixture_file = fixture_path()
    fixtures = json.loads(fixture_file.read_text(encoding="utf-8"))
    current = probe(fixtures)
    if args.probe_only:
        report = current
    else:
        if not args.baseline_ref:
            parser.error("--baseline-ref is required unless --probe-only")
        previous = baseline_probe(args.baseline_ref, fixture_file)
        by_id = {x["id"]: x for x in previous}
        report = {"schema": fixtures["schema"], "baseline_ref": args.baseline_ref,
                  "note": "Frozen synthetic cases, not the original AI proposals; compare revisions with identical inputs.",
                  "cases": [{"id": item["id"], "baseline": by_id[item["id"]], "candidate": item,
                             "status_changed": (by_id[item["id"]]["status"] != item["status"]),
                             "delta_revised_chars": (
                                 item["revisions"]["ins_chars"] + item["revisions"]["del_chars"]
                                 - by_id[item["id"]]["revisions"]["ins_chars"]
                                 - by_id[item["id"]]["revisions"]["del_chars"]
                             )} for item in current]}
    encoded = json.dumps(report, ensure_ascii=False, indent=2) + "\n"
    if args.output:
        args.output.write_text(encoded, encoding="utf-8")
        print("Report saved:", args.output)
    else:
        print(encoded, end="")


if __name__ == "__main__":
    main()
