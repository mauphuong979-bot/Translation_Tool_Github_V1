
import pandas as pd
import unicodedata
import os
import re

def clean_text(text):
    if not isinstance(text, str) or pd.isna(text):
        return ""
    text = unicodedata.normalize('NFC', str(text))
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f\xad\u200b\u200c\u200d\u2060\ufeff\xb7\u2022\u202a-\u202e\u200e\u200f]', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

DICT_FILE = r'd:\AI\Python\11_Project\Translation_Tool\Translation_Tool_Antigravity_V2\Dictionary_v3.xlsx'

search_terms = [
    "Doanh thu hoạt động tài chính",
    "Chi phí tài chính",
    "Thu nhập khác",
    "Chi phí khác"
]

if os.path.exists(DICT_FILE):
    df = pd.read_excel(DICT_FILE).astype(str)
    with open('scratch/search_dict_results_v2.txt', 'w', encoding='utf-8') as f:
        for term in search_terms:
            cleaned_term = clean_text(term)
            f.write(f"Searching for: '{term}'\n")
            matches = []
            for idx, row in df.iterrows():
                vn_clean = clean_text(row['Vietnamese'])
                if cleaned_term.lower() in vn_clean.lower():
                    matches.append((row['Vietnamese'], row['E']))
            if matches:
                for vn, en in matches:
                    f.write(f"  Match: '{vn}' -> '{en}'\n")
            else:
                f.write("  No match found.\n")
else:
    print("Dict not found")
