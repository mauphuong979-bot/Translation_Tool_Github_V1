import difflib
from copy import deepcopy
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import RGBColor
import translation_lib as tl


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
        if run.font.highlight_color == WD_COLOR_INDEX.YELLOW:
            return run
    return paragraph.runs[0] if paragraph.runs else None


def _has_yellow_highlight(paragraph):
    return _first_review_run(paragraph) is not None and any(
        run.font.highlight_color == WD_COLOR_INDEX.YELLOW for run in paragraph.runs
    )


def _add_revision(paragraph, deleted_text, inserted_text, author, revision_id):
    runs = list(paragraph.runs)
    if not runs or not inserted_text:
        return False
    paragraph_element = paragraph._element
    first_run = runs[0]._element
    original_runs = [run._element for run in runs]

    deleted = OxmlElement("w:del")
    deleted.set(qn("w:author"), author)
    deleted.set(qn("w:date"), "2026-11-08T00:00:00Z")
    deleted.set(qn("w:id"), str(revision_id))
    old_run = OxmlElement("w:r")
    old_properties = first_run.find(qn("w:rPr"))
    if old_properties is not None:
        old_run.append(deepcopy(old_properties))
    old_text = OxmlElement("w:delText")
    old_text.set(qn("xml:space"), "preserve")
    old_text.text = deleted_text
    old_run.append(old_text)
    deleted.append(old_run)

    inserted = OxmlElement("w:ins")
    inserted.set(qn("w:author"), author)
    inserted.set(qn("w:date"), "2026-11-08T00:00:00Z")
    inserted.set(qn("w:id"), str(revision_id + 1))
    new_run = OxmlElement("w:r")
    new_properties = first_run.find(qn("w:rPr"))
    if new_properties is not None:
        new_run.append(deepcopy(new_properties))
    new_text = OxmlElement("w:t")
    new_text.set(qn("xml:space"), "preserve")
    new_text.text = inserted_text
    new_run.append(new_text)
    inserted.append(new_run)

    for run in original_runs:
        paragraph_element.remove(run)
    paragraph_element.append(deleted)
    paragraph_element.append(inserted)
    return True


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
    comment = doc.add_comment(anchor_run, text="", author=author, initials="TT")
    _add_comment_section(comment, "[Original]", original, source_run, RGBColor(0, 51, 204))
    if dictionary_match:
        key, value, score = dictionary_match
        _add_comment_section(comment, "[Dict VN]", key, source_run, RGBColor(0, 51, 204))
        _add_comment_section(comment, "[Suggest]", value, source_run, RGBColor(0, 51, 204))
        _add_comment_section(comment, "[Match]", f"{score:.0%}", source_run, RGBColor(0, 51, 204))
    if prior_match:
        _, value, score = prior_match
        prior_vn, prior_translation = value
        _add_comment_section(comment, "[Prior VN]", prior_vn, source_run, RGBColor(204, 0, 0))
        _add_comment_section(comment, "[Suggest]", prior_translation, source_run, RGBColor(204, 0, 0))
        _add_comment_section(comment, "[Match]", f"{score:.0%}", source_run, RGBColor(204, 0, 0))
    return comment


def _make_review_item(paragraph, original, dictionary_match, prior_match):
    references = []
    if dictionary_match:
        source, translation, score = dictionary_match
        references.append({"source": "dictionary", "original": source, "translation": translation, "score": score})
    if prior_match:
        _, (source, translation), score = prior_match
        references.append({"source": "prior", "original": source, "translation": translation, "score": score})
    if not any(reference["score"] < 1.0 for reference in references):
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


def add_comments_and_track_changes(doc, translation_map, py_dict=None, original_texts=None, author="Translation Tool", review_items=None):
    comments = 0
    revisions = 0
    revision_id = 9000
    for paragraph in tl._get_all_paragraphs(doc):
        if paragraph.part is not doc.part:
            continue
        current = paragraph.text.strip()
        original = (original_texts or {}).get(paragraph._element, current)
        if not current or not paragraph.runs or not tl.contains_vietnamese(original):
            continue
        if current.startswith("["):
            continue

        # Dictionary replacements are already resolved; only visible yellow text is reviewed.
        if not _has_yellow_highlight(paragraph):
            continue
        dictionary_match = _match(original, translation_map) or _match(original, translation_map, 0.6)
        prior_match = _match(original, py_dict) or _match(original, py_dict, 0.7)
        if not dictionary_match and not prior_match:
            continue

        if _add_formatted_comment(doc, paragraph, original, dictionary_match, prior_match, author) is not None:
            comments += 1
            if review_items is not None:
                review_item = _make_review_item(paragraph, original, dictionary_match, prior_match)
                if review_item is not None:
                    review_items.append(review_item)

        # Only unresolved yellow text with an exact prior-report match becomes a revision.
        if prior_match and prior_match[2] >= 0.999999:
            prior_translation = str(prior_match[1][1])
            if _add_revision(paragraph, original, prior_translation, author, revision_id):
                revisions += 1
                revision_id += 2
    return {"comments": comments, "revisions": revisions}

