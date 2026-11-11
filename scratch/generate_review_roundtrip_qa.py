import io
from pathlib import Path
import sys

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import Pt

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from review_annotations import add_comments_and_track_changes
from review_export import build_ai_review_document
from review_roundtrip import apply_ai_translations, bind_review_manifest, preview_ai_translations


directory = Path(__file__).resolve().parent / "review_roundtrip_qa"
directory.mkdir(exist_ok=True)
report = Document()
report.add_heading("Kiểm thử nhập bản dịch AI", 0)
source = "Công ty có vốn điều lệ là 1.234,56 đồng tại ngày 31/12/2025."
reference = "Công ty có vốn điều lệ là 1.000,00 đồng tại ngày 31/12/2024."
translated = "本公司於2024/12/31日之註冊資本為 VND 1,000.00。"
for size in (10, 10, 7):
    paragraph = report.add_table(rows=1, cols=1).cell(0, 0).paragraphs[0] if size == 7 else report.add_paragraph()
    run = paragraph.add_run(source)
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    run.bold = True
items = []
add_comments_and_track_changes(report, {reference: translated}, review_items=items)
bind_review_manifest(report, items, "Ht")
source_path = directory / "01_Review_Source.docx"
report.save(source_path)
blank = build_ai_review_document(items, "Ht")
(directory / "02_AI_Worksheet.docx").write_bytes(blank.getvalue())
worksheet = Document(blank)
for table in worksheet.tables:
    table.cell(1, 4).text = "本公司於2025/12/31日之註冊資本為 VND 1,234.56。"
completed = directory / "03_AI_Completed.docx"
worksheet.save(completed)
print(preview_ai_translations(source_path, completed)["counts"])
output, summary = apply_ai_translations(source_path, completed)
(directory / "04_Report_AI_Reviewed.docx").write_bytes(output.getvalue())
print(summary["applied"], "translations applied")
print(directory)
