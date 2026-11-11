import difflib
from copy import deepcopy
from docx.oxml.ns import qn
from docx.shared import RGBColor
from docx.text.paragraph import Paragraph
import translation_lib as tl
from review_revisions import effective_highlight, replace_paragraph_with_revision
from review_roundtrip import match_percentage


def _match(text, mapping, cutoff=None):
    cleaned = tl.clean_text(text)
    if not cleaned or not mapping:
        return None
    if cutoff is None:
        for key, value in mapping.items():
            if tl.clean_text(key) == cleaned:
                return key, value, 1.0
        return None
    keys = list(mapping.keys())
    found = difflib.get_close_matches(cleaned, keys, n=1, cutoff=cutoff)
    if not found:
        return None
    key = found[0]
    score = difflib.SequenceMatcher(None, cleaned, tl.clean_text(key)).ratio()
    return key, mapping[key], score


def _first_review_run(paragraph):
    for run in paragraph.runs:
        if effective_highlight(run, paragraph) == "yellow":
            return run
    return paragraph.runs[0] if paragraph.runs else None


def _has_yellow_highlight(paragraph):
    return _first_review_run(paragraph) is not None and any(
        effective_highlight(run, paragraph) == "yellow" for run in paragraph.runs
    )


def _add_revision(paragraph, deleted_text, inserted_text, author, revision_id, target_col="E"):
    return replace_paragraph_with_revision(paragraph, inserted_text, target_col, revision_id, author)


def _copy_comment_run_format(source_run, target_run, color=None, bold=None):
    source_properties = source_run._element.find(qn("w:rPr"))
    if source_properties is not None:
        copied_properties = deepcopy(source_properties)
        highlight = copied_properties.find(qn("w:highlight"))
        if highlight is not None:
            copied_properties.remove(highlight)
        target_run._element.insert(0, copied_properties)
    if color is not None:
        target_run.font.color.rgb = color
    if bold is not None:
        target_run.font.bold = bold


def _add_comment_section(comment, label, value, source_run, color):
    paragraph = comment.paragraphs[0] if len(comment.paragraphs) == 1 and not comment.paragraphs[0].text else comment.add_paragraph()
    label_run = paragraph.add_run(label)
    _copy_comment_run_format(source_run, label_run, color=color, bold=True)
    paragraph.add_run("\n")
    value_run = paragraph.add_run(str(value))
    _copy_comment_run_format(source_run, value_run, color=color, bold=False)


def _add_formatted_comment(doc, paragraph, original, dictionary_match, prior_match, author):
    anchor_run = _first_review_run(paragraph)
    if anchor_run is None:
        return None
    source_run = anchor_run
    comment = doc.add_comment([run for run in paragraph.runs if run.text], text="", author=author, initials="TT")
    _add_comment_section(comment, "[Original]", original, source_run, RGBColor(0, 51, 204))
    if dictionary_match:
        key, value, score = dictionary_match
        _add_comment_section(comment, "[Dict VN]", key, source_run, RGBColor(0, 51, 204))
        _add_comment_section(comment, "[Suggest]", value, source_run, RGBColor(0, 51, 204))
        _add_comment_section(comment, "[Match]", match_percentage(score), source_run, RGBColor(0, 51, 204))
    if prior_match:
        _, value, score = prior_match
        prior_vn, prior_translation = value
        _add_comment_section(comment, "[Prior VN]", prior_vn, source_run, RGBColor(204, 0, 0))
        _add_comment_section(comment, "[Suggest]", prior_translation, source_run, RGBColor(204, 0, 0))
        _add_comment_section(comment, "[Match]", match_percentage(score), source_run, RGBColor(204, 0, 0))
    if not dictionary_match and not prior_match:
        _add_comment_section(comment, "[AI translation needed]", "Chưa có tham chiếu phù hợp. Dịch đầy đủ nội dung nguồn hiện tại.", source_run, RGBColor(204, 0, 0))
    return comment


def needs_ai_translation(item):
    references = item.get("references", [])
    if any(reference["source"] == "prior" and reference["score"] >= 1.0 for reference in references):
        return False
    return not references or any(reference["score"] < 1.0 for reference in references)


def _make_review_item(paragraph, original, dictionary_match, prior_match):
    references = []
    if dictionary_match:
        source, translation, score = dictionary_match
        references.append({"source": "dictionary", "original": source, "translation": translation, "score": score})
    if prior_match:
        _, (source, translation), score = prior_match
        references.append({"source": "prior", "original": source, "translation": translation, "score": score})
    if not needs_ai_translation({"references": references}):
        return None

    source_run = _first_review_run(paragraph)
    formatting = {}
    for property_name in ("size", "bold", "italic", "underline"):
        value = getattr(source_run.font, property_name)
        for initial_style in (source_run.style, paragraph.style):
            style = initial_style
            while value is None and style is not None:
                value = getattr(style.font, property_name)
                style = style.base_style
        formatting[property_name] = value.pt if property_name == "size" and value is not None else value
    formatting["alignment"] = paragraph.alignment
    style = paragraph.style
    while formatting["alignment"] is None and style is not None:
        formatting["alignment"] = style.paragraph_format.alignment
        style = style.base_style
    return {"original": original, "references": references, "formatting": formatting}


def add_comments_and_track_changes(doc, translation_map, py_dict=None, original_texts=None, author="Translation Tool", review_items=None, target_col="E"):
    from review_roundtrip import _unsafe_reason

    comments = 0
    revisions = 0
    identifiers = [int(element.get(qn("w:id"))) for element in doc.element.iter()
                   if (element.get(qn("w:id")) or "").isdigit()]
    revision_id = max(identifiers, default=0) + 1
    for paragraph_element in doc.element.body.xpath(".//w:p[not(ancestor::w:txbxContent)]"):
        paragraph = Paragraph(paragraph_element, doc._body)
        current = paragraph.text.strip()
        original = (original_texts or {}).get(paragraph._element, current)
        if not current or not paragraph.runs or not tl.contains_vietnamese(original):
            continue
        if current.startswith("["):
            continue

        if not _has_yellow_highlight(paragraph):
            continue
        dictionary_match = _match(original, translation_map) or _match(original, translation_map, 0.6)
        prior_match = _match(original, py_dict) or _match(original, py_dict, 0.7)
        manual_reason = None
        if prior_match and prior_match[2] >= 1.0:
            manual_reason = _unsafe_reason(paragraph, None, doc, None)
            prior_translation = str(prior_match[1][1] or "").strip()
            if not manual_reason and prior_translation and _add_revision(paragraph, original, prior_translation, author, revision_id, target_col):
                revisions += 1
                revision_id += 2
                continue
            manual_reason = manual_reason or ("Bản dịch kỳ trước trống; cần kiểm tra thủ công." if not prior_translation
                                               else "Không thể tạo Track Changes an toàn; cần kiểm tra thủ công.")
        comment = _add_formatted_comment(doc, paragraph, original, dictionary_match, prior_match, author)
        if comment is not None:
            comments += 1
            if manual_reason:
                _add_comment_section(comment, "[Manual review]", manual_reason, _first_review_run(paragraph), RGBColor(204, 0, 0))
                continue
            if review_items is not None:
                review_item = _make_review_item(paragraph, original, dictionary_match, prior_match)
                if review_item is not None:
                    review_item["comment_id"] = str(comment.comment_id)
                    review_items.append(review_item)

    return {"comments": comments, "revisions": revisions}

