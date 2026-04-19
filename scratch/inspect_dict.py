
import pandas as pd
import os
import unicodedata

def clean_text(text):
    if not isinstance(text, str) or pd.isna(text):
        return ""
    text = unicodedata.normalize('NFC', str(text))
    import re
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f\xad\u200b\u200c\u200d\u2060\ufeff\xb7\u2022]', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

DICT_FILE = r'd:\AI\Python\11_Project\Translation_Tool\Translation_Tool_Antigravity_V2\Dictionary_v3.xlsx'

if os.path.exists(DICT_FILE):
    df = pd.read_excel(DICT_FILE)
    print(f"Dictionary Columns: {df.columns.tolist()}")
    print(f"Number of entries: {len(df)}")
    
    # Check for some common entries or specifically the ones the user might be concerned about
    # Since I don't know the specific terms, I'll just look at the first 20.
    print("\nFirst 20 entries (Vietnamese -> E):")
    for idx, row in df.head(20).iterrows():
        print(f"'{row['Vietnamese']}' -> '{row['E']}'")
        
    # Also check if there are duplicates with different normalization or something
    df['VN_Clean'] = df['Vietnamese'].apply(clean_text)
    dupes = df[df.duplicated(subset=['VN_Clean'], keep=False)]
    if not dupes.empty:
        print(f"\nFound {len(dupes)} potentially duplicate entries (after cleaning).")
else:
    print("Dictionary not found.")
