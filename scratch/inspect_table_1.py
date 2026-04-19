
from docx import Document
import os
import sys

# Reconfigure stdout to use UTF-8
if sys.stdout.encoding != 'utf-8':
    try:
        import codecs
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'replace')
    except:
        pass

doc_path = r'd:\AI\Python\11_Project\Translation_Tool\Translation_Tool_Antigravity_V2\test_translated.docx'
if os.path.exists(doc_path):
    doc = Document(doc_path)
    with open('scratch/inspect_table_1.txt', 'w', encoding='utf-8') as f:
        t = doc.tables[0]
        for i, row in enumerate(t.rows):
            f.write(f"Row {i+1}: {repr(row.cells[0].text)}\n")
else:
    print("File not found")
