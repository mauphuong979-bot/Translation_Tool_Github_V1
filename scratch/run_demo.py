
import os
import sys
import pandas as pd
from docx import Document
import unicodedata

# Add parent dir to path to find translation_lib
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import translation_lib as tl

def main():
    doc_path = r'd:\AI\Python\11_Project\Translation_Tool\Translation_Tool_Antigravity_V2\test.docx'
    dict_path = r'd:\AI\Python\11_Project\Translation_Tool\Translation_Tool_Antigravity_V2\Dictionary_v3.xlsx'
    output_path = r'd:\AI\Python\11_Project\Translation_Tool\Translation_Tool_Antigravity_V2\test_translated.docx'
    
    if not os.path.exists(doc_path):
        return

    # 1. Load Document
    doc = Document(doc_path)
    
    # 2. Mock Metadata
    metadata = {
        "name_vn": "CÔNG TY TNHH ABC",
        "name_trans": "ABC COMPANY LIMITED",
        "year_end": "31/12/2025",
        "report_date": "20/03/2026",
        "period_in": "01/01/2025 - 31/12/2025"
    }
    
    # 3. Load and fill Dictionary
    df_dict = tl.load_and_fill_v3_dictionary(metadata)
    if df_dict is None:
        return
        
    translation_map = dict(zip(df_dict['Vietnamese'], df_dict['E']))
    
    with open('scratch/demo_debug.txt', 'w', encoding='utf-8') as f:
        f.write("--- Demo Debug Log ---\n")
        test_key = "Doanh thu hoạt động tài chính"
        f.write(f"Looking for: '{test_key}'\n")
        f.write(f"Dict Value: '{translation_map.get(test_key)}'\n")
        
        # Check if any key *contains* it
        f.write("\nKeys containing the term:\n")
        for k in translation_map:
            if test_key in k:
                f.write(f"  - '{k}': '{translation_map[k]}'\n")

    # 4. Perform Translation
    count = tl.replace_text_in_document(doc, translation_map, target_col="E", metadata=metadata)
    
    doc.save(output_path)
    print(f"Done. Replacements: {count}")

if __name__ == "__main__":
    main()
