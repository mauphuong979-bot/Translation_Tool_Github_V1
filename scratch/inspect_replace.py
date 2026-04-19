import pandas as pd
import re
import os
import unicodedata
import translation_lib as tl

def debug_replace():
    if not os.path.exists('Dictionary_v3.xlsx'):
        print("Dictionary_v3.xlsx not found")
        return
    
    # Dummy metadata
    meta = {
        "name_vn": "CONG TY TEST",
        "name_trans": "TEST COMPANY",
        "year_end": "31/12/2025",
        "report_date": "08/04/2026",
        "period_in": "01/01/2025 - 31/12/2025",
        "period_in_2": "01/01/2024 - 31/12/2024"
    }
    
    print("Loading Dictionary_v3.xlsx...")
    df = pd.read_excel('Dictionary_v3.xlsx')
    
    print("Building sub map...")
    sub_map = tl.get_metadata_substitution_map(meta)
    
    regex_subs = {}
    for tag_with_brackets, val in sub_map.items():
        tag_name = tag_with_brackets.replace("[", "").replace("]", "").strip()
        pattern = re.compile(r"\[\s*" + re.escape(tag_name) + r"\s*\]", re.IGNORECASE)
        regex_subs[pattern] = str(val) if val is not None else ""
        print(f"  Tag: {tag_name} -> {val}")

    print("Forcing types and replacing...")
    df_new = df.astype(str)
    df_new = df_new.replace(to_replace=regex_subs, regex=True)
    
    # Check row 1 (index 0) Col A (V_NAME usually)
    print("\nSample Inspection (Row 0):")
    print(df_new.iloc[0, 0:2])
    
    # Check for any remaining [
    print("\nRemaining tags?")
    found_remaining = False
    for col in df_new.columns:
        res = df_new[df_new[col].str.contains(r'\[v_name\]', case=False, regex=True)]
        if len(res) > 0:
            print(f"Found unreplaced [v_name] in column {col}:")
            print(res[col].head())
            found_remaining = True
            
    if not found_remaining:
        print("Success! No unreplaced [v_name] found in Row 0 area.")

    df_new.to_excel('debug_resolved.xlsx', index=False)
    print("\nSaved debug_resolved.xlsx")

if __name__ == "__main__":
    debug_replace()
