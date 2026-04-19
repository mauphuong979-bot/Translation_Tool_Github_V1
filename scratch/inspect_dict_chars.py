
import pandas as pd
import unicodedata
import os
import re

def inspect_chars(text, label):
    res = [f"{label}: {repr(text)}"]
    for char in text:
        res.append(f"  Char: {repr(char)} | Code: {ord(char)}")
    return "\n".join(res)

DICT_FILE = r'd:\AI\Python\11_Project\Translation_Tool\Translation_Tool_Antigravity_V2\Dictionary_v3.xlsx'

if os.path.exists(DICT_FILE):
    df = pd.read_excel(DICT_FILE).astype(str)
    
    with open('scratch/dict_char_debug.txt', 'w', encoding='utf-8') as f:
        target = "Doanh thu hoạt động tài chính"
        for idx, row in df.iterrows():
            vn = row['Vietnamese']
            # Simple check if it looks the same
            if target in vn or vn in target or "Doanh thu hoạt động tài chính" in vn:
                f.write(inspect_chars(vn, f"Dict Row {idx}") + "\n")
                f.write("-" * 20 + "\n")
else:
    print("Dict not found")
