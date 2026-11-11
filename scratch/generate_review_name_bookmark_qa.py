from pathlib import Path
import sys

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from review_annotations import add_comments_and_track_changes
from review_export import build_ai_review_document
from review_roundtrip import apply_ai_translations, bind_review_manifest, preview_ai_translations
from test_review_names_bookmarks import add_bookmark
from test_review_roundtrip import saved


directory = Path(__file__).resolve().parent / "review_name_bookmark_qa"
directory.mkdir(exist_ok=True)
report = Document()
title_style = report.styles["Title"]
title_style.font.color.rgb = RGBColor(0, 0, 0)
border = title_style.element.get_or_add_pPr().find(qn("w:pBdr"))
if border is not None:
    border.getparent().remove(border)
report.add_paragraph("Kiểm thử tên riêng và thẻ mẫu", "Title")
sources = (
    "Công ty TNHH ABC trình bày báo cáo cho giai đoạn 【自年月日至年月日止階段】.",
    "Nguyễn Văn A chịu trách nhiệm lập báo cáo tài chính.",
    "Công ty TNHH ABC có vốn điều lệ là 1.234,56 đồng.",
)
translations = (
    "Công ty TNHH ABC 呈報財務報表 【自年月日至年月日止階段】。",
    "Nguyễn Văn A 負責編製財務報表。",
    "Công ty TNHH ABC 之註冊資本為 VND 1,234.56。",
)
for index, source in enumerate(sources):
    paragraph = report.add_paragraph(source)
    paragraph.runs[0].font.size = Pt(10)
    paragraph.runs[0].font.highlight_color = WD_COLOR_INDEX.YELLOW
    add_bookmark(paragraph, "before" if index == 0 else "whole", str(700 + index), f"QA_Bookmark_{index}")
items = []
counts = add_comments_and_track_changes(report, {}, {sources[2]: (sources[2], translations[2])}, review_items=items, target_col="Ht")
assert counts == {"comments": 2, "revisions": 1}, counts
bind_review_manifest(report, items, "Ht")
report.save(directory / "01_Report.docx")
worksheet_data = build_ai_review_document(items, "Ht")
(directory / "02_Worksheet.docx").write_bytes(worksheet_data.getvalue())
completed = Document(worksheet_data)
for table, translation in zip(completed.tables, translations):
    table.cell(1, 4).text = translation
assert preview_ai_translations(saved(report), saved(completed))["counts"] == {"ready": 2}
output, summary = apply_ai_translations(saved(report), saved(completed))
assert summary["applied"] == 2
(directory / "04_Tracked_Report.docx").write_bytes(output.getvalue())
print(directory)
