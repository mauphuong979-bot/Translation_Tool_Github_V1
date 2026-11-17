import re
from copy import deepcopy
from datetime import datetime, timezone

from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.run import Run

import translation_lib as tl


def _font_properties(run, paragraph):
    yield run._r.rPr
    for initial_style in (run.style, paragraph.style):
        style = initial_style
        while style is not None:
            yield style.element.rPr
            style = style.base_style
    defaults = paragraph.part.document.styles.element.find(qn("w:docDefaults"))
    if defaults is not None:
        properties = defaults.find(qn("w:rPrDefault"))
        if properties is not None:
            yield properties.find(qn("w:rPr"))


def effective_highlight(run, paragraph):
    for properties in _font_properties(run, paragraph):
        highlight = properties.find(qn("w:highlight")) if properties is not None else None
        if highlight is not None:
            return highlight.get(qn("w:val"))
    return None


def _effective_size(run, paragraph):
    if run.font.size is not None:
        return run.font.size.pt
    for initial_style in (run.style, paragraph.style):
        style = initial_style
        while style is not None:
            if style.font.size is not None:
                return style.font.size.pt
            style = style.base_style
    return 10


def _clear_yellow(run, paragraph):
    properties = run._r.get_or_add_rPr()
    highlight = properties.find(qn("w:highlight"))
    if highlight is not None and highlight.get(qn("w:val")) == "yellow":
        properties.remove(highlight)
    if effective_highlight(run, paragraph) == "yellow":
        highlight = OxmlElement("w:highlight")
        highlight.set(qn("w:val"), "none")
        properties.append(highlight)


def _segments(translation, target_col, highlights):
    """Split the translation into (text, highlighted) pieces at script and highlight boundaries."""
    cuts = {0, len(translation)}
    if target_col in ("Hs", "Ht"):
        for match in re.finditer(r"[一-鿿㐀-䶿豈-﫿]+", translation):
            cuts.update(match.span())
    for start, end in highlights or ():
        cuts.update((max(0, min(start, len(translation))), max(0, min(end, len(translation)))))
    points = sorted(cuts)
    for start, end in zip(points, points[1:]):
        if start < end:
            yield translation[start:end], any(a <= start and end <= b for a, b in highlights or ())


def replace_paragraph_with_revision(paragraph, translation, target_col, revision_id, author, highlights=None):
    content_runs = [run for run in paragraph.runs if run.text]
    if not content_runs or not translation:
        return False
    source_run = next((run for run in content_runs if effective_highlight(run, paragraph) == "yellow"), content_runs[0])
    timestamp = datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")
    deleted, inserted = OxmlElement("w:del"), OxmlElement("w:ins")
    for offset, element in enumerate((deleted, inserted)):
        element.set(qn("w:id"), str(revision_id + offset))
        element.set(qn("w:author"), author)
        element.set(qn("w:date"), timestamp)
    for run in content_runs:
        old_run = deepcopy(run._r)
        for text in old_run.iter(qn("w:t")):
            text.tag = qn("w:delText")
        for reference in list(old_run.iter(qn("w:commentReference"))):
            reference.getparent().remove(reference)
        deleted.append(old_run)

    for text, highlighted in _segments(translation, target_col, highlights):
        element = OxmlElement("w:r")
        if source_run._r.rPr is not None:
            element.append(deepcopy(source_run._r.rPr))
        new_run = Run(element, paragraph)
        new_run.text = text
        _clear_yellow(new_run, paragraph)
        if highlighted:
            properties = element.get_or_add_rPr()
            for old in properties.findall(qn("w:highlight")):
                properties.remove(old)
            new_run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        tl._set_run_fonts_refined(new_run, "Times New Roman", "DFKai-SB" if target_col in ("Hs", "Ht") else "Times New Roman", _effective_size(source_run, paragraph))
        properties = element.get_or_add_rPr()
        fonts = properties.get_or_add_rFonts()
        for attribute in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            fonts.attrib.pop(qn(f"w:{attribute}"), None)
        language = properties.find(qn("w:lang"))
        if language is None:
            language = OxmlElement("w:lang")
            properties.append(language)
        language.set(qn("w:val"), "en-US")
        if target_col in ("Hs", "Ht"):
            language.set(qn("w:eastAsia"), "zh-CN" if target_col == "Hs" else "zh-TW")
        inserted.append(element)

    position = paragraph._p.index(content_runs[0]._r)
    for run in content_runs:
        paragraph._p.remove(run._r)
    paragraph._p.insert(position, deleted)
    paragraph._p.insert(position + 1, inserted)
    return True
