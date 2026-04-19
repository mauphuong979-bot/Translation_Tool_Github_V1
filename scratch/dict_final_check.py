
import os
import sys
import pandas as pd
from docx import Document

# Add parent dir to path to find translation_lib
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import translation_lib as tl

def main():
    metadata = {
        "name_vn": "CÔNG TY TNHH ABC",
        "name_trans": "ABC COMPANY LIMITED",
        "year_end": "31/12/2025",
        "report_date": "20/03/2026",
        "period_in": "01/01/2025 - 31/12/2025"
    }
    
    df_dict = tl.load_and_fill_v3_dictionary(metadata)
    
    target_key = "Doanh thu hoạt động tài chính"
    cleaned_target = tl.clean_text(target_key)
    
    matches = df_dict[df_dict['Vietnamese'] == cleaned_target]
    
    with open('scratch/dict_final_check.txt', 'w', encoding='utf-8') as f:
        f.write(f"Matches for '{cleaned_target}':\n")
        f.write(matches.to_string())
        
        # Check for ALL items that failed
        failed = ["Doanh thu hoạt động tài chính", "Chi phí tài chính", "Thu nhập khác", "Chi phí khác"]
        for item in failed:
            cleaned = tl.clean_text(item)
            f.write(f"\n\nChecking: '{cleaned}'\n")
            m = df_dict[df_dict['Vietnamese'] == cleaned]
            f.write(m.to_string())

if __name__ == "__main__":
    main()
