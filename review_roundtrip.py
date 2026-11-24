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
from docx.table import _Cell
from docx.text.paragraph import Paragraph
from lxml import etree

import translation_lib as tl
from review_revisions import effective_highlight, replace_paragraph_with_revision


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
        raise ValueError("Báo cáo này có dữ liệu liên kết AI bị trùng. Hãy dùng đúng báo cáo gốc đi kèm.")
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
        raise ValueError("Báo cáo này chưa có dữ liệu liên kết AI. Hãy xử lý lại báo cáo và tải bảng AI_Form 5 cột mới cùng báo cáo Comments + Track Changes đi kèm.")
    root = etree.fromstring(part.blob, parser=etree.XMLParser(resolve_entities=False, no_network=True))
    if root.tag != "{" + MANIFEST_NAMESPACE + "}manifest":
        raise ValueError("Dữ liệu liên kết AI trong báo cáo không hợp lệ.")
    manifest = json.loads(root.text or "{}")
    if manifest.get("version") != 1 or manifest.get("target_col") not in LANGUAGES:
        raise ValueError("Dữ liệu liên kết AI trong báo cáo không được hỗ trợ.")
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
        raise ValueError("Ngôn ngữ dịch không được hỗ trợ.")
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


YELLOW_FILLS = {"FFFF00", "FFFF66", "FFF200", "FFEB3B", "FFFF99"}


def _is_yellow(run, paragraph):
    if effective_highlight(run, paragraph) == "yellow":
        return True
    shading = run._r.rPr.find(qn("w:shd")) if run._r.rPr is not None else None
    return shading is not None and (shading.get(qn("w:fill")) or "").upper() in YELLOW_FILLS


def _yellow_ranges(cell):
    """Character ranges of cell.text that the AI marked with a yellow highlight or yellow run shading."""
    ranges, offset = [], 0
    for index, paragraph in enumerate(cell.paragraphs):
        if index:
            offset += 1
        for run in paragraph.runs:
            length = len(run.text)
            if length and _is_yellow(run, paragraph):
                if ranges and ranges[-1][1] == offset:
                    ranges[-1] = (ranges[-1][0], offset + length)
                else:
                    ranges.append((offset, offset + length))
            offset += length
    return ranges


REVISION_XPATH = ".//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo | .//w:rPrChange | .//w:pPrChange"


def _accepted_cell(cell):
    """A copy of the cell with all Track Changes accepted (the final text a reviewer sees with All Markup off)."""
    tc = deepcopy(cell._tc)
    for element in tc.xpath(".//w:del | .//w:moveFrom | .//w:rPrChange | .//w:pPrChange | .//w:moveFromRangeStart | .//w:moveFromRangeEnd"
                            " | .//w:moveToRangeStart | .//w:moveToRangeEnd"):
        element.getparent().remove(element)
    for wrapper in tc.xpath(".//w:ins | .//w:moveTo"):
        parent = wrapper.getparent()
        if parent.tag == qn("w:rPr") or parent.tag == qn("w:pPr") or wrapper.getparent().tag == qn("w:trPr"):
            parent.remove(wrapper)  # a paragraph-mark or row insertion mark: nothing to keep
            continue
        index = parent.index(wrapper)
        for child in list(wrapper):
            parent.insert(index, child)
            index += 1
        parent.remove(wrapper)
    for paragraph in tc.xpath(".//w:p"):
        properties = paragraph.find(qn("w:pPr"))
        marks = properties.find(qn("w:rPr")) if properties is not None else None
        if marks is not None:
            for mark in marks.xpath("./w:del | ./w:ins"):
                marks.remove(mark)
    return _Cell(tc, cell._parent)


def _worksheet_rows(document):
    rows = []
    for table in document.tables:
        if len(table.rows) != 2 or any(len(row.cells) != 5 for row in table.rows):
            raise ValueError("Hãy dùng bảng AI_Form 5 cột mới, mỗi bảng gồm một dòng tiêu đề và một dòng nội dung. Không gộp hoặc tách bảng.")
        if any(len({cell._tc for cell in row.cells}) != 5 for row in table.rows):
            raise ValueError("Bảng AI có ô bị gộp. Hãy khôi phục cấu trúc bảng ban đầu.")
        if [_normalize(cell.text).casefold() for cell in table.rows[0].cells] != [_normalize(header).casefold() for header in HEADERS]:
            raise ValueError("Tiêu đề cột đã bị sửa hoặc đây không phải bảng AI_Form.")
        cells = table.rows[1].cells
        if any(cell.tables for cell in cells):
            raise ValueError("Bảng AI không được chứa bảng lồng bên trong.")
        result_cell = cells[4]
        accepted = bool(result_cell._tc.xpath(REVISION_XPATH))
        if accepted:
            result_cell = _accepted_cell(result_cell)
        rows.append({"review_id": cells[0].text.strip(), "source_values": [cell.text for cell in cells[1:4]],
                     "translation": result_cell.text, "highlights": _yellow_ranges(result_cell),
                     "has_revisions": False, "accepted_revisions": accepted,
                     "has_objects": any(element.tag == qn(f"w:{tag}") for element in cells[4]._tc.iter()
                                        for tag in ("fldChar", "instrText", "fldSimple", "hyperlink", "drawing", "object", "sdt"))})
    if not rows:
        raise ValueError("Bảng AI không có mục nào để nhập.")
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
        return "Bookmark liên kết AI bị mất hoặc không rõ ràng. Cần xử lý thủ công."
    if not starts and not ends:
        return ""
    content_runs = [run for run in paragraph.runs if run.text]
    if not content_runs:
        return "Đoạn nguồn không có chữ để thay."
    first_position = paragraph._p.index(content_runs[0]._r)
    last_position = paragraph._p.index(content_runs[-1]._r)
    all_starts = list(document.element.iter(qn("w:bookmarkStart")))
    all_ends = list(document.element.iter(qn("w:bookmarkEnd")))
    for bookmark_id in {element.get(qn("w:id")) for element in starts + ends}:
        matching_starts = [element for element in all_starts if element.get(qn("w:id")) == bookmark_id]
        matching_ends = [element for element in all_ends if element.get(qn("w:id")) == bookmark_id]
        if not bookmark_id or len(matching_starts) != 1 or len(matching_ends) != 1:
            return "Đoạn có bookmark thiếu hoặc bị trùng. Cần xử lý thủ công."
        start, end = matching_starts[0], matching_ends[0]
        name = start.get(qn("w:name"))
        if (not name or sum(element.get(qn("w:name")) == name for element in all_starts) != 1
                or any(start.get(qn(f"w:{attribute}")) is not None for attribute in ("colFirst", "colLast"))):
            return "Đoạn có bookmark không rõ ràng hoặc bookmark theo cột bảng. Cần xử lý thủ công."
        if start.getparent() is not paragraph._p or end.getparent() is not paragraph._p:
            return "Đoạn có bookmark trải qua nhiều đoạn. Cần xử lý thủ công."
        start_position, end_position = paragraph._p.index(start), paragraph._p.index(end)
        if start_position >= end_position:
            return "Đoạn có phạm vi bookmark không hợp lệ. Cần xử lý thủ công."
        if name == anchor_name:
            continue
        if (end_position < first_position or start_position > last_position
                or (start_position < first_position and end_position > last_position)):
            continue
        return "Đoạn có bookmark bao một phần câu, không thể chuyển an toàn sang bản dịch. Cần xử lý thủ công."
    return ""


def _unsafe_reason(paragraph, comment_id, document, anchor_name):
    if tl.has_fields(paragraph):
        return "Đoạn có liên kết Excel hoặc field Word. Cần xử lý thủ công."
    revision_tags = ("ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange")
    if any(any(True for _ in paragraph._p.iter(qn(f"w:{tag}"))) for tag in revision_tags):
        return "Báo cáo nguồn đã có Track Changes ở đoạn này từ trước khi nhập AI. Bản dịch AI không được áp dụng; Track Changes và comment hiện có được giữ nguyên."
    complex_tags = ("hyperlink", "drawing", "pict", "object", "sdt")
    if any(any(True for _ in paragraph._p.iter(qn(f"w:{tag}"))) for tag in complex_tags):
        return "Đoạn có liên kết, content control hoặc đối tượng nhúng. Cần xử lý thủ công."
    if reason := _bookmark_reason(paragraph, document, anchor_name):
        return reason
    allowed_children = {qn(f"w:{tag}") for tag in ("pPr", "r", "commentRangeStart", "commentRangeEnd", "bookmarkStart", "bookmarkEnd", "proofErr")}
    allowed_run_children = {qn(f"w:{tag}") for tag in ("rPr", "t", "tab", "br", "cr", "commentReference")}
    for child in paragraph._p:
        if child.tag not in allowed_children:
            return "Đoạn có nội dung không được hỗ trợ. Cần xử lý thủ công."
        if child.tag == qn("w:r") and any(element.tag not in allowed_run_children for element in child):
            return "Đoạn có thành phần trong dòng không được hỗ trợ. Cần xử lý thủ công."
    for tag in ("commentRangeStart", "commentRangeEnd", "commentReference"):
        if any(element.get(qn("w:id")) != comment_id for element in paragraph._p.iter(qn(f"w:{tag}"))):
            return "Có comment khác chồng lên đoạn này. Cần xử lý thủ công."
    active_comments, active_bookmarks = set(), set()
    for element in document.element.iter():
        if element is paragraph._p:
            if active_comments - {comment_id}:
                return "Có comment khác trải qua đoạn này. Cần xử lý thủ công."
            if active_bookmarks:
                return "Có bookmark khác trải qua đoạn này. Cần xử lý thủ công."
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
        return "Đoạn nguồn không có chữ để thay."
    if any(element.get(qn("w:type"), "textWrapping") != "textWrapping" for element in paragraph._p.iter(qn("w:br"))):
        return "Đoạn có ngắt trang hoặc ngắt cột. Cần xử lý thủ công."
    if len({_format_signature(run) for run in paragraph.runs if run.text.strip()}) > 1:
        return "Đoạn có định dạng chữ hỗn hợp, không thể chuyển an toàn sang câu dịch. Cần xử lý thủ công."
    return ""


def _expected_values(item):
    """The first three content columns of an AI_Form row, as exported for this manifest item."""
    return [item["original"], reference_text(item["references"], "original"), reference_text(item["references"], "translation")]


def apply_overrides(manifest, rows, overrides):
    """Replace translations with the user's in-app edits ({review_id: text}); edited rows go through the same checks.

    Highlighted phrases of the original translation stay highlighted wherever they still occur in the edited text.
    An edit for a code the supplied rows lack (for example a batch that failed) adds a row for it.
    Edited text is plain text typed in the app, so the cell's own Track Changes or objects no longer apply.
    """
    if not overrides:
        return rows
    expected = {item["review_id"]: item for item in manifest["items"]}
    applied = set(manifest.get("applied_ids", []))
    updated, present = [], set()
    for row in rows:
        row = dict(row)
        if row["review_id"] in overrides:
            present.add(row["review_id"])
            text = str(overrides[row["review_id"]] or "")
            phrases = [row["translation"][start:end] for start, end in row.get("highlights", [])]
            row.update(translation=text, highlights=highlight_ranges(text, phrases), has_revisions=False, has_objects=False, edited=True)
        updated.append(row)
    for review_id, text in overrides.items():
        if review_id not in present and review_id in expected and review_id not in applied:
            updated.append({"review_id": review_id, "source_values": _expected_values(expected[review_id]),
                            "translation": str(text or ""), "highlights": [], "has_revisions": False, "has_objects": False, "edited": True})
    return updated


def _inspect(report_source, worksheet_source, overrides=None):
    report = _read_document(report_source)
    manifest = _read_manifest(report)
    worksheet = _read_document(worksheet_source)
    supplied = apply_overrides(manifest, _worksheet_rows(worksheet), overrides)
    return _validate(report, manifest, supplied, worksheet.core_properties.identifier)


def _validate(report, manifest, supplied, identifier=""):
    """Check supplied translation rows against the report manifest; shared by the AI_Form file and API flows."""
    expected = {item["review_id"]: item for item in manifest["items"]}
    if any(row["review_id"] not in expected for row in supplied):
        raise ValueError("Có mã không thuộc báo cáo này. Hãy tải đúng cặp báo cáo Comments + Track Changes và bảng AI_Form.")
    if identifier.startswith("TranslationTool:") and identifier != f"TranslationTool:{manifest['batch']}:{manifest['target_col']}":
        raise ValueError("Bảng AI này thuộc báo cáo khác hoặc ngôn ngữ dịch khác.")
    counts = Counter(row["review_id"] for row in supplied)
    comments = {str(comment.comment_id) for comment in report.comments}
    applied = set(manifest.get("applied_ids", []))
    results = []
    for row in supplied:
        item = expected[row["review_id"]]
        result = {"review_id": row["review_id"], "original": item["original"], "translation": row["translation"],
                  "highlights": row["highlights"], "status": "ready", "reason": "", "edited": row.get("edited", False)}
        if row.get("accepted_revisions") and not row.get("edited"):
            result["reason"] = "Đã tự lấy nội dung cuối cùng của Track Changes trong ô bản dịch (như Accept All)."
        expected_values = _expected_values(item)
        paragraph, comment_id = _resolve_anchor(report, item)
        result["comment_id"] = comment_id
        result["source_has_revisions"] = paragraph is not None and bool(paragraph._p.xpath(".//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo | .//w:rPrChange | .//w:pPrChange"))
        result["source_current_text"] = _visible_text(paragraph) if paragraph is not None else ""
        if counts[row["review_id"]] > 1:
            result.update(status="blocked", reason="Mã bị trùng; không áp dụng dòng nào của mã này.")
        elif row["review_id"] in applied:
            result.update(status="applied", reason="Đã áp dụng vào báo cáo này.")
        elif [_normalize(value) for value in row["source_values"]] != [_normalize(value) for value in expected_values]:
            result.update(status="blocked", reason="AI đã sửa cột nội dung gốc hoặc cột tham chiếu.")
        elif row["has_revisions"]:
            result.update(status="blocked", reason="Hãy Accept hoặc Reject các Track Changes trong ô kết quả AI trước khi nhập.")
        elif row["has_objects"]:
            result.update(status="blocked", reason="Ô kết quả AI chỉ được chứa chữ bản dịch, không chứa field, liên kết hay đối tượng.")
        elif not row["translation"].strip():
            result.update(status="blank", reason="AI chưa điền bản dịch; giữ nguyên nội dung gốc và comment.")
        elif paragraph is None or comment_id not in comments:
            result.update(status="blocked", reason="Comment gốc đã bị sửa, hoặc vị trí neo bị mất/không rõ ràng.")
        elif _normalize(_visible_text(paragraph)) != _normalize(item["report_text"]):
            result.update(status="blocked", reason="Nội dung báo cáo đã thay đổi kể từ khi xuất bảng AI.")
        elif reason := _unsafe_reason(paragraph, comment_id, report, item["anchor_name"]):
            if (result["source_has_revisions"] and any(reference["source"] == "prior" and reference["score"] >= 0.999999 for reference in item["references"])
                    and any(element.get(qn("w:author")) == "Translation Tool" for element in paragraph._p.iter(qn("w:ins")))):
                reason = ("Đã được dịch ở bước xử lý ban đầu do khớp 100% với báo cáo năm trước. "
                          "Track Changes này có sẵn trong báo cáo nguồn TRƯỚC khi nhập AI. "
                          "Bản dịch AI KHÔNG được áp dụng; bản dịch, Track Changes và comment hiện có được giữ nguyên.")
            result.update(status="blocked", reason=reason)
        elif _normalize(row["translation"]) == _normalize(paragraph.text):
            result.update(status="unchanged", reason="Bản dịch AI giống hệt nội dung hiện tại của báo cáo.")
        results.append(result)
    for review_id, item in expected.items():
        if review_id not in counts and review_id not in applied:
            results.append({"review_id": review_id, "original": item["original"], "translation": "", "status": "missing", "reason": "Bảng AI thiếu mục này; giữ nguyên nội dung gốc và comment."})
    return report, manifest, results


def _preview(manifest, rows):
    order = {item["review_id"]: index for index, item in enumerate(manifest["items"])}
    rows = sorted(rows, key=lambda row: order.get(row["review_id"], len(order)))
    return {"target_col": manifest["target_col"], "rows": rows, "counts": dict(Counter(row["status"] for row in rows))}


def preview_ai_translations(report_source, worksheet_source, overrides=None):
    _, manifest, rows = _inspect(report_source, worksheet_source, overrides)
    return _preview(manifest, rows)


def read_review_manifest(report_source):
    """Items still awaiting AI translation, limited to the AI_Form content columns (no report text, comments or metadata)."""
    from review_annotations import needs_ai_translation

    manifest = _read_manifest(_read_document(report_source))
    applied = set(manifest.get("applied_ids", []))
    items = [{"review_id": item["review_id"], "original": item["original"],
              "references": [{field: reference[field] for field in ("source", "score", "original", "translation")}
                             for reference in item["references"]]}
             for item in manifest["items"] if item["review_id"] not in applied and needs_ai_translation(item)]
    return {"batch": manifest["batch"], "target_col": manifest["target_col"], "items": items}


def highlight_ranges(translation, phrases):
    """Character ranges of every occurrence of the given phrases, merged and sorted."""
    spans = []
    for phrase in phrases:
        if not isinstance(phrase, str) or not phrase:
            continue
        start = translation.find(phrase)
        while start != -1:
            spans.append((start, start + len(phrase)))
            start = translation.find(phrase, start + len(phrase))
    merged = []
    for start, end in sorted(spans):
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    return merged


def rows_from_ai_results(manifest, results):
    """Turn API results [{review_id, translation, highlight: [phrases]}] into rows in the AI_Form worksheet shape."""
    expected = {item["review_id"]: item for item in manifest["items"]}
    rows = []
    for result in results:
        if not isinstance(result, dict) or not isinstance(result.get("review_id"), str) or not isinstance(result.get("translation", ""), str):
            raise ValueError("Kết quả AI không đúng định dạng.")
        translation = result.get("translation", "")
        item = expected.get(result["review_id"].strip())
        rows.append({"review_id": result["review_id"].strip(),
                     "source_values": _expected_values(item) if item is not None else [],
                     "translation": translation, "highlights": highlight_ranges(translation, result.get("highlight") or []),
                     "has_revisions": False, "has_objects": False})
    return rows


def _inspect_results(report_source, results, overrides=None):
    report = _read_document(report_source)
    manifest = _read_manifest(report)
    return _validate(report, manifest, apply_overrides(manifest, rows_from_ai_results(manifest, results), overrides))


def preview_ai_results(report_source, results, overrides=None):
    _, manifest, rows = _inspect_results(report_source, results, overrides)
    return _preview(manifest, rows)


def _replace_with_revision(document, paragraph, comment_id, translation, target_col, revision_id, anchor_name, highlights=None):
    replace_paragraph_with_revision(paragraph, translation, target_col, revision_id, "Translation Tool AI", highlights)

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


def apply_ai_translations(report_source, worksheet_source, selected_ids=None, overrides=None):
    return _apply(*_inspect(report_source, worksheet_source, overrides), selected_ids)


def apply_ai_results(report_source, results, selected_ids=None, overrides=None):
    return _apply(*_inspect_results(report_source, results, overrides), selected_ids)


def _apply(report, manifest, rows, selected_ids=None):
    ready = {row["review_id"]: row for row in rows if row["status"] == "ready"}
    selected = set(ready) if selected_ids is None else set(selected_ids)
    if not selected or not selected.issubset(ready):
        raise ValueError("Hãy chọn ít nhất một bản dịch sẵn sàng. Không thể áp dụng các dòng bị chặn, trống hoặc thiếu.")
    identifiers = [int(element.get(qn("w:id"))) for part in report.part.package.parts if hasattr(part, "element")
                   for element in part.element.iter() if (element.get(qn("w:id")) or "").isdigit()]
    revision_id = max(identifiers, default=0) + 1
    items = {item["review_id"]: item for item in manifest["items"]}
    for review_id in sorted(selected):
        item = items[review_id]
        paragraph, comment_id = _resolve_anchor(report, item)
        _replace_with_revision(report, paragraph, comment_id, ready[review_id]["translation"], manifest["target_col"], revision_id, item["anchor_name"],
                               ready[review_id]["highlights"])
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
