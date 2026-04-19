
from docx import Document
import os

doc_path = r'd:\AI\Python\11_Project\Translation_Tool\Translation_Tool_Antigravity_V2\test.docx'
doc = Document(doc_path)

print("--- Run Inspection of Table 1 Cell(1, 1) ---")
# Table 1 Cell(1, 1) was 'CHỈ TIÊU'
cell = doc.tables[0].cell(0, 0)
print(f"Cell text: '{cell.text}'")
for i, para in enumerate(cell.paragraphs):
    print(f"  Para {i}:")
    for j, run in enumerate(para.runs):
        print(f"    Run {j}: '{run.text}'")

print("\n--- Run Inspection of Table 1 Cell(1, 2) ---")
# Table 1 Cell(1, 2) was 'Mã số'
cell = doc.tables[0].cell(0, 1)
print(f"Cell text: '{cell.text}'")
for i, para in enumerate(cell.paragraphs):
    print(f"  Para {i}:")
    for j, run in enumerate(para.runs):
        print(f"    Run {j}: '{run.text}'")

print("\n--- Run Inspection of Table 1 Cell(13, 1) ---")
# '6.	Doanh thu hoạt động tài chính'
cell = doc.tables[0].cell(12, 0)
print(f"Cell text: '{cell.text}'")
for i, para in enumerate(cell.paragraphs):
    print(f"  Para {i}:")
    for j, run in enumerate(para.runs):
        print(f"    Run {j}: '{run.text}' (repr: {repr(run.text)})")

print("\n--- Run Inspection of Table 2 Cell(2, 5) ---")
# 'Người lập biểu kiêm kế toán trưởng'
cell = doc.tables[1].cell(1, 4) # 1-indexed (2, 5) -> (1, 4)
print(f"Cell text: '{cell.text}'")
for i, para in enumerate(cell.paragraphs):
    print(f"  Para {i}:")
    for j, run in enumerate(para.runs):
        print(f"    Run {j}: '{run.text}'")
