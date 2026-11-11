import io
import json
import unicodedata
from collections import Counter
from copy import deepcopy
from uuid import uuid4

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from lxml import etree

import translation_lib as tl
from review_revisions import replace_paragraph_with_revision


HEADERS = ("Mã", "Nội dung cần dịch", "Tiếng Việt tham chiếu", "Bản dịch tham chiếu", "Bản dịch hoàn thiện")
SOURCE_LABELS = {"dictionary": "Từ điển", "prior": "Báo cáo kỳ trước"}
MANIFEST_NAMESPACE = "urn:translation-tool:ai-review:1"
MANIFEST_PATH = "/customXml/translationToolReview.xml"
LANGUAGES = {"E", "Hs", "Ht"}


def match_percentage(score):
    percentage = 100 if score >= 1 else min(99.99, round(score * 100, 2))
    return f"{percentage:g}%"


def reference_label(reference):
    return f"{SOURCE_LABELS[reference['source']]} ({match_percentage(reference['score'])})"


def reference_text(references, field):
    return "\n".join(f"{reference_label(reference)}\n{reference[field]}" for reference in references)


def _normalize(text):
    return " ".join(unicodedata.normalize("NFC", str(text)).split())


def _visible_text(paragraph):
    separators = {qn("w:tab"): "\t", qn("w:br"): "\n", qn("w:cr"): "\n"}
    return "".join(element.text or "" if element.tag == qn("w:t") else separators.get(element.tag, "") for element in paragraph._p.iter())


def _manifest_part(document):
    matches = []
    for relationship in document.part.rels.values():
        if not relationship.is_external and relationship.reltype == RT.CUSTOM_XML:
            part = relationship.target_part
            try:
                root = etree.fromstring(part.blob, parser=etree.XMLParser(resolve_entities=False, no_network=True))
            except etree.XMLSyntaxError:
                continue
            if root.tag == "{" + MANIFEST_NAMESPACE + "}manifest":
                matches.append(part)
    if len(matches) > 1:
        raise ValueError("This report contains duplicate AI review mappings. Use the original matching report.")
    return matches[0] if matches else None


def _write_manifest(document, manifest):
    root = etree.Element("{" + MANIFEST_NAMESPACE + "}manifest")
    root.text = json.dumps(manifest, ensure_ascii=False)
    blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8")
    part = _manifest_part(document)
    if part is None:
        package = document.part.package
        path = PackURI(MANIFEST_PATH)
        if any(existing.partname == path for existing in package.parts):
            path = package.next_partname("/customXml/translationToolReview%d.xml")
        part = Part(path, "application/xml", blob, package)
        document.part.relate_to(part, RT.CUSTOM_XML)
    else:
        part._blob = blob


def _read_manifest(document):
    part = _manifest_part(document)
    if part is None:
        raise ValueError("This report has no AI review mapping. Process it again and download the new five-column worksheet and matching Comments + Track Changes report.")
    root = etree.fromstring(part.blob, parser=etree.XMLParser(resolve_entities=False, no_network=True))
    if root.tag != "{" + MANIFEST_NAMESPACE + "}manifest":
        raise ValueError("The report contains an invalid AI review mapping.")
    manifest = json.loads(root.text or "{}")
    if manifest.get("version") != 1 or manifest.get("target_col") not in LANGUAGES:
        raise ValueError("The report contains an unsupported AI review mapping.")
    return manifest


def _anchors(document, comment_id):
    starts = [element for element in document.element.iter(qn("w:commentRangeStart")) if element.get(qn("w:id")) == comment_id]
    ends = [element for element in document.element.iter(qn("w:commentRangeEnd")) if element.get(qn("w:id")) == comment_id]
    if len(starts) != 1 or len(ends) != 1:
        return None
    paragraph = starts[0].getparent()
    if paragraph is not ends[0].getparent() or paragraph.tag != qn("w:p"):
        return None
    return Paragraph(paragraph, document._body)


def _resolve_anchor(document, item):
    anchor_name = item.get("anchor_name")
    starts = [element for element in document.element.iter(qn("w:bookmarkStart"))
              if anchor_name and element.get(qn("w:name")) == anchor_name]
    if len(starts) != 1:
        return None, None
    paragraph_element = starts[0].getparent()
    ends = [element for element in document.element.iter(qn("w:bookmarkEnd"))
            if element.get(qn("w:id")) == starts[0].get(qn("w:id"))]
    if len(ends) != 1 or ends[0].getparent() is not paragraph_element or paragraph_element.tag != qn("w:p"):
        return None, None
    comments = {str(comment.comment_id): comment for comment in document.comments}
    matches = []
    for marker in paragraph_element.iter(qn("w:commentRangeStart")):
        comment_id = marker.get(qn("w:id"))
        comment = comments.get(comment_id)
        paragraph = _anchors(document, comment_id)
        if (paragraph is not None and paragraph._p is paragraph_element and comment is not None
                and _normalize(comment.text) == item.get("comment_text") and comment.author == item.get("comment_author")):
            matches.append((paragraph, comment_id))
    return matches[0] if len(matches) == 1 else (None, None)


def bind_review_manifest(document, items, target_col):
    if not items:
        return
    if target_col not in LANGUAGES:
        raise ValueError("Unsupported translation language.")
    batch = uuid4().hex[:8].upper()
    existing_ids = [int(element.get(qn("w:id"))) for element in document.element.iter()
                    if (element.get(qn("w:id")) or "").isdigit()]
    bookmark_id = max(existing_ids, default=0) + 1
    comments = {str(comment.comment_id): comment for comment in document.comments}
    records = []
    for index, item in enumerate(items, 1):
        item["review_id"] = f"{batch}-C{index:03d}"
        item["batch"] = batch
        paragraph = _anchors(document, str(item["comment_id"]))
        anchor_name = f"TT_AI_{batch}_C{index:03d}"
        if paragraph is not None:
            start, end = OxmlElement("w:bookmarkStart"), OxmlElement("w:bookmarkEnd")
            start.set(qn("w:name"), anchor_name)
            start.set(qn("w:id"), str(bookmark_id))
            end.set(qn("w:id"), str(bookmark_id))
            paragraph._p.insert(1 if paragraph._p.pPr is not None else 0, start)
            paragraph._p.append(end)
            bookmark_id += 1
        comment = comments.get(str(item["comment_id"]))
        records.append({"review_id": item["review_id"], "comment_id": str(item["comment_id"]),
                        "anchor_name": anchor_name, "comment_text": _normalize(comment.text) if comment is not None else None,
                        "comment_author": comment.author if comment is not None else None,
                        "original": item["original"], "references": item["references"],
                        "report_text": _visible_text(paragraph) if paragraph is not None else None})
    _write_manifest(document, {"version": 1, "batch": batch, "target_col": target_col, "items": records, "applied_ids": []})


def _read_document(source):
    if isinstance(source, bytes):
        source = io.BytesIO(source)
    elif hasattr(source, "seek"):
        source.seek(0)
    return Document(source)


def _worksheet_rows(document):
    rows = []
    for table in document.tables:
        if len(table.rows) != 2 or any(len(row.cells) != 5 for row in table.rows):
            raise ValueError("Use the new five-column AI worksheet, with one header and one content row per table. Do not merge or split its tables.")
        if any(len({cell._tc for cell in row.cells}) != 5 for row in table.rows):
            raise ValueError("The worksheet contains merged cells. Restore the original table structure.")
        if [_normalize(cell.text).casefold() for cell in table.rows[0].cells] != [_normalize(header).casefold() for header in HEADERS]:
            raise ValueError("The worksheet headers have changed or this is not an AI translation worksheet.")
        cells = table.rows[1].cells
        if any(cell.tables for cell in cells):
            raise ValueError("Nested tables are not supported in the AI worksheet.")
        rows.append({"review_id": cells[0].text.strip(), "source_values": [cell.text for cell in cells[1:4]],
                     "translation": cells[4].text,
                     "has_revisions": bool(cells[4]._tc.xpath(".//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo | .//w:rPrChange")),
                     "has_objects": any(element.tag == qn(f"w:{tag}") for element in cells[4]._tc.iter()
                                        for tag in ("fldChar", "instrText", "fldSimple", "hyperlink", "drawing", "object", "sdt"))})
    if not rows:
        raise ValueError("The AI worksheet contains no review tables.")
    return rows


def _format_signature(run):
    properties = deepcopy(run._r.rPr)
    if properties is None:
        return ()
    ignored = {qn(f"w:{tag}") for tag in ("rFonts", "lang", "noProof")}
    return tuple(sorted(etree.tostring(child, method="c14n") for child in properties
                        if child.tag not in ignored and not (child.tag == qn("w:highlight") and child.get(qn("w:val")) == "yellow")))


def _bookmark_reason(paragraph, document, anchor_name):
    starts = list(paragraph._p.iter(qn("w:bookmarkStart")))
    ends = list(paragraph._p.iter(qn("w:bookmarkEnd")))
    if anchor_name is not None and sum(start.get(qn("w:name")) == anchor_name for start in starts) != 1:
        return "The AI review bookmark is missing or ambiguous; manual review is required."
    if not starts and not ends:
        return ""
    content_runs = [run for run in paragraph.runs if run.text]
    if not content_runs:
        return "The source paragraph has no editable text."
    first_position = paragraph._p.index(content_runs[0]._r)
    last_position = paragraph._p.index(content_runs[-1]._r)
    all_starts = list(document.element.iter(qn("w:bookmarkStart")))
    all_ends = list(document.element.iter(qn("w:bookmarkEnd")))
    for bookmark_id in {element.get(qn("w:id")) for element in starts + ends}:
        matching_starts = [element for element in all_starts if element.get(qn("w:id")) == bookmark_id]
        matching_ends = [element for element in all_ends if element.get(qn("w:id")) == bookmark_id]
        if not bookmark_id or len(matching_starts) != 1 or len(matching_ends) != 1:
            return "Contains an incomplete or duplicate bookmark; manual review is required."
        start, end = matching_starts[0], matching_ends[0]
        name = start.get(qn("w:name"))
        if (not name or sum(element.get(qn("w:name")) == name for element in all_starts) != 1
                or any(start.get(qn(f"w:{attribute}")) is not None for attribute in ("colFirst", "colLast"))):
            return "Contains an ambiguous or table-column bookmark; manual review is required."
        if start.getparent() is not paragraph._p or end.getparent() is not paragraph._p:
            return "Contains a bookmark spanning paragraph boundaries; manual review is required."
        start_position, end_position = paragraph._p.index(start), paragraph._p.index(end)
        if start_position >= end_position:
            return "Contains an invalid bookmark range; manual review is required."
        if name == anchor_name:
            continue
        if (end_position < first_position or start_position > last_position
                or (start_position < first_position and end_position > last_position)):
            continue
        return "Contains a partial-text bookmark that cannot be mapped safely to the translation; manual review is required."
    return ""


def _unsafe_reason(paragraph, comment_id, document, anchor_name):
    if tl.has_fields(paragraph):
        return "Contains an Excel link or Word field; manual review is required."
    revision_tags = ("ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange")
    if any(any(True for _ in paragraph._p.iter(qn(f"w:{tag}"))) for tag in revision_tags):
        return "Track Changes already existed in the source report before AI import. This AI result will not be applied; existing revisions and comments are kept unchanged."
    complex_tags = ("hyperlink", "drawing", "pict", "object", "sdt")
    if any(any(True for _ in paragraph._p.iter(qn(f"w:{tag}"))) for tag in complex_tags):
        return "Contains a link, content control or embedded object; manual review is required."
    if reason := _bookmark_reason(paragraph, document, anchor_name):
        return reason
    allowed_children = {qn(f"w:{tag}") for tag in ("pPr", "r", "commentRangeStart", "commentRangeEnd", "bookmarkStart", "bookmarkEnd", "proofErr")}
    allowed_run_children = {qn(f"w:{tag}") for tag in ("rPr", "t", "tab", "br", "cr", "commentReference")}
    for child in paragraph._p:
        if child.tag not in allowed_children:
            return "Contains unsupported paragraph content; manual review is required."
        if child.tag == qn("w:r") and any(element.tag not in allowed_run_children for element in child):
            return "Contains unsupported inline content; manual review is required."
    for tag in ("commentRangeStart", "commentRangeEnd", "commentReference"):
        if any(element.get(qn("w:id")) != comment_id for element in paragraph._p.iter(qn(f"w:{tag}"))):
            return "Another comment overlaps this paragraph; manual review is required."
    active_comments, active_bookmarks = set(), set()
    for element in document.element.iter():
        if element is paragraph._p:
            if active_comments - {comment_id}:
                return "Another comment spans this paragraph; manual review is required."
            if active_bookmarks:
                return "Another bookmark spans this paragraph; manual review is required."
            break
        if element.tag == qn("w:commentRangeStart"):
            active_comments.add(element.get(qn("w:id")))
        elif element.tag == qn("w:commentRangeEnd"):
            active_comments.discard(element.get(qn("w:id")))
        elif element.tag == qn("w:bookmarkStart"):
            active_bookmarks.add(element.get(qn("w:id")))
        elif element.tag == qn("w:bookmarkEnd"):
            active_bookmarks.discard(element.get(qn("w:id")))
    if not paragraph.runs or not any(run.text for run in paragraph.runs):
        return "The source paragraph has no editable text."
    if any(element.get(qn("w:type"), "textWrapping") != "textWrapping" for element in paragraph._p.iter(qn("w:br"))):
        return "Contains a page or column break; manual review is required."
    if len({_format_signature(run) for run in paragraph.runs if run.text.strip()}) > 1:
        return "Contains mixed text formatting that cannot be mapped safely to a translated sentence; manual review is required."
    return ""


def _inspect(report_source, worksheet_source):
    report = _read_document(report_source)
    manifest = _read_manifest(report)
    worksheet = _read_document(worksheet_source)
    supplied = _worksheet_rows(worksheet)
    expected = {item["review_id"]: item for item in manifest["items"]}
    if any(row["review_id"] not in expected for row in supplied):
        raise ValueError("One or more review IDs do not belong to this report. Upload the matching Comments + Track Changes report and AI worksheet.")
    identifier = worksheet.core_properties.identifier
    if identifier.startswith("TranslationTool:") and identifier != f"TranslationTool:{manifest['batch']}:{manifest['target_col']}":
        raise ValueError("The worksheet belongs to another report or translation language.")
    counts = Counter(row["review_id"] for row in supplied)
    comments = {str(comment.comment_id) for comment in report.comments}
    applied = set(manifest.get("applied_ids", []))
    results = []
    for row in supplied:
        item = expected[row["review_id"]]
        result = {"review_id": row["review_id"], "original": item["original"], "translation": row["translation"], "status": "ready", "reason": ""}
        expected_values = [item["original"], reference_text(item["references"], "original"), reference_text(item["references"], "translation")]
        paragraph, comment_id = _resolve_anchor(report, item)
        result["comment_id"] = comment_id
        result["source_has_revisions"] = paragraph is not None and bool(paragraph._p.xpath(".//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo | .//w:rPrChange | .//w:pPrChange"))
        result["source_current_text"] = _visible_text(paragraph) if paragraph is not None else ""
        if counts[row["review_id"]] > 1:
            result.update(status="blocked", reason="Duplicate review ID; none of its rows will be applied.")
        elif row["review_id"] in applied:
            result.update(status="applied", reason="Already applied to this report.")
        elif [_normalize(value) for value in row["source_values"]] != [_normalize(value) for value in expected_values]:
            result.update(status="blocked", reason="The source or reference columns were changed by AI.")
        elif row["has_revisions"]:
            result.update(status="blocked", reason="Accept or reject revisions in the AI result cell before importing it.")
        elif row["has_objects"]:
            result.update(status="blocked", reason="The AI result cell must contain plain translation text, not fields, links or objects.")
        elif not row["translation"].strip():
            result.update(status="blank", reason="No completed translation; the source and comment will be kept.")
        elif paragraph is None or comment_id not in comments:
            result.update(status="blocked", reason="The original comment was changed, or its anchor is missing or ambiguous.")
        elif _normalize(_visible_text(paragraph)) != _normalize(item["report_text"]):
            result.update(status="blocked", reason="The report text has changed since the worksheet was exported.")
        elif reason := _unsafe_reason(paragraph, comment_id, report, item["anchor_name"]):
            if (result["source_has_revisions"] and any(reference["source"] == "prior" and reference["score"] >= 0.999999 for reference in item["references"])
                    and any(element.get(qn("w:author")) == "Translation Tool" for element in paragraph._p.iter(qn("w:ins")))):
                reason = ("Already translated during initial processing from a 100% prior-report match. "
                          "These Track Changes were present in the source report BEFORE AI import. "
                          "This AI result is NOT applied; the existing translation, revisions and comment are kept unchanged.")
            result.update(status="blocked", reason=reason)
        elif _normalize(row["translation"]) == _normalize(paragraph.text):
            result.update(status="unchanged", reason="The AI result is identical to the current report text.")
        results.append(result)
    for review_id, item in expected.items():
        if review_id not in counts and review_id not in applied:
            results.append({"review_id": review_id, "original": item["original"], "translation": "", "status": "missing", "reason": "This table is missing from the AI worksheet; the source and comment will be kept."})
    return report, manifest, results


def preview_ai_translations(report_source, worksheet_source):
    _, manifest, rows = _inspect(report_source, worksheet_source)
    return {"target_col": manifest["target_col"], "rows": rows, "counts": dict(Counter(row["status"] for row in rows))}


def _replace_with_revision(document, paragraph, comment_id, translation, target_col, revision_id, anchor_name):
    replace_paragraph_with_revision(paragraph, translation, target_col, revision_id, "Translation Tool AI")

    for tag in ("commentRangeStart", "commentRangeEnd", "commentReference"):
        for element in list(paragraph._p.iter(qn(f"w:{tag}"))):
            if element.get(qn("w:id")) == comment_id:
                parent = element.getparent()
                parent.remove(element)
                if parent.tag == qn("w:r") and not any(child.tag != qn("w:rPr") for child in parent):
                    parent.getparent().remove(parent)
    anchor = next(element for element in paragraph._p.iter(qn("w:bookmarkStart")) if element.get(qn("w:name")) == anchor_name)
    anchor_id = anchor.get(qn("w:id"))
    for element in list(paragraph._p):
        if element.tag in (qn("w:bookmarkStart"), qn("w:bookmarkEnd")) and element.get(qn("w:id")) == anchor_id:
            paragraph._p.remove(element)
    for comment in list(document.comments):
        if str(comment.comment_id) == comment_id:
            comment._element.getparent().remove(comment._element)


def apply_ai_translations(report_source, worksheet_source, selected_ids=None):
    report, manifest, rows = _inspect(report_source, worksheet_source)
    ready = {row["review_id"]: row for row in rows if row["status"] == "ready"}
    selected = set(ready) if selected_ids is None else set(selected_ids)
    if not selected or not selected.issubset(ready):
        raise ValueError("Select at least one ready translation. Blocked, blank or missing rows cannot be applied.")
    identifiers = [int(element.get(qn("w:id"))) for part in report.part.package.parts if hasattr(part, "element")
                   for element in part.element.iter() if (element.get(qn("w:id")) or "").isdigit()]
    revision_id = max(identifiers, default=0) + 1
    items = {item["review_id"]: item for item in manifest["items"]}
    for review_id in sorted(selected):
        item = items[review_id]
        paragraph, comment_id = _resolve_anchor(report, item)
        _replace_with_revision(report, paragraph, comment_id, ready[review_id]["translation"], manifest["target_col"], revision_id, item["anchor_name"])
        revision_id += 2
    manifest["applied_ids"] = sorted(set(manifest.get("applied_ids", [])) | selected)
    _write_manifest(report, manifest)
    settings = report.settings.element
    tracking = settings.find(qn("w:trackRevisions"))
    if tracking is None:
        tracking = OxmlElement("w:trackRevisions")
        settings.append(tracking)
    tracking.set(qn("w:val"), "true")
    output = io.BytesIO()
    report.save(output)
    output.seek(0)
    return output, {"applied": len(selected), "applied_ids": sorted(selected),
                    "preserved_ids": sorted(set(items) - selected),
                    "preexisting_revision_ids": sorted({row["review_id"] for row in rows if row.get("source_has_revisions") and row["review_id"] not in selected}),
                    "rows": rows}
