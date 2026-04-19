
import pandas as pd
import unicodedata
import os
import re

def clean_text(text):
    if not isinstance(text, str) or pd.isna(text):
        return ""
    text = unicodedata.normalize('NFC', str(text))
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f\xad\u200b\u200c\u200d\u2060\ufeff\xb7\u2022]', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

DICT_FILE = r'd:\AI\Python\11_Project\Translation_Tool\Translation_Tool_Antigravity_V2\Dictionary_v3.xlsx'

search_terms = [
    "Doanh thu bán hàng và cung cấp dịch vụ",
    "CHỈ TIÊU",
    "Mã số",
    "Thuyết minh",
    "Tổng Giám đốc",
    "Chi nhánh Fair Consulting Co., Ltd. Nhật Bản tại Việt Nam"
]

if os.path.exists(DICT_FILE):
    df = pd.read_excel(DICT_FILE).astype(str)
    
    with open('scratch/search_dict_results.txt', 'w', encoding='utf-8') as f:
        for term in search_terms:
            cleaned_term = clean_text(term)
            f.write(f"Searching for: '{term}' (Cleaned: '{cleaned_term}')\n")
            
            # Search by contains (cleaned)
            matches = []
            for idx, row in df.iterrows():
                vn_clean = clean_text(row['Vietnamese'])
                if cleaned_term.lower() in vn_clean.lower():
                    matches.append((row['Vietnamese'], row['E'], vn_clean))
            
            if matches:
                f.write(f"  Found {len(matches)} matches:\n")
                for vn, en, vnc in matches:
                    f.write(f"    - Dict: '{vn}' -> '{en}'\n")
                    f.write(f"      (Cleaned Dict: '{vnc}')\n")
            else:
                f.write("  No match found.\n")
            f.write("-" * 20 + "\n")
else:
    print("Dict not found")
