from pathlib import Path
import sys

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import Pt

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from review_annotations import add_comments_and_track_changes
from review_export import build_ai_review_document
from review_roundtrip import apply_ai_translations, bind_review_manifest, preview_ai_translations


directory = Path(__file__).resolve().parent / "review_enhancement_qa"
directory.mkdir(exist_ok=True)
report = Document()
report.add_heading("Kiểm thử hoàn thiện bản dịch", 0)
exact = "Công ty có vốn điều lệ là 1.234,56 đồng tại ngày 31/12/2025."
unmatched = "Nội dung mới chưa có bản dịch và cần bổ sung đầy đủ."
translation = "本公司之新增資料為 VND 1,234.56。"
yellow_style = report.styles.add_style("Yellow source", WD_STYLE_TYPE.CHARACTER)
yellow_style.font.highlight_color = WD_COLOR_INDEX.YELLOW
for index, text in enumerate((exact, unmatched, unmatched)):
    paragraph = report.add_paragraph(text)
    run = paragraph.runs[0]
    run.font.name = "Times New Roman"
    run.font.size = Pt(10)
    if index == 1:
        run.style = yellow_style
    else:
        run.font.highlight_color = WD_COLOR_INDEX.YELLOW
items = []
counts = add_comments_and_track_changes(report, {}, {exact: (exact, translation)}, review_items=items, target_col="Ht")
assert counts == {"comments": 2, "revisions": 1}, counts
assert len(items) == 2
bind_review_manifest(report, items, "Ht")
report_path = directory / "01_Report.docx"
report.save(report_path)
worksheet = build_ai_review_document(items, "Ht")
(directory / "02_Worksheet.docx").write_bytes(worksheet.getvalue())
completed = Document(worksheet)
for table in completed.tables:
    table.cell(1, 4).text = translation
completed_path = directory / "03_AI_Completed.docx"
completed.save(completed_path)
assert preview_ai_translations(report_path, completed_path)["counts"] == {"ready": 2}
output, summary = apply_ai_translations(report_path, completed_path)
(directory / "04_Tracked_Report.docx").write_bytes(output.getvalue())
assert summary["applied"] == 2
print(directory)
