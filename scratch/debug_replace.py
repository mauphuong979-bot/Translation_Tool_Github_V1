import pandas as pd
import re
import os
import unicodedata

def clean_text(text):
    if not isinstance(text, str): return ""
    text = unicodedata.normalize('NFC', str(text))
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f\xad\u200b\u200c\u200d\u2060\ufeff\xb7\u2022]', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def parse_period_to_tags(period_str, date1_prefix, date2_prefix):
    tags = {}
    dates = re.findall(r"(\d{1,2})/(\d{1,2})/(\d{4})", period_str)
    if len(dates) >= 1:
        d, m, y = dates[0]
        tags.update({f"[{date1_prefix}_day]": f"{int(d):02d}", f"[{date1_prefix}_month]": f"{int(m):02d}", f"[{date1_prefix}_year]": y})
    else:
        tags.update({f"[{date1_prefix}_day]": "", f"[{date1_prefix}_month]": "", f"[{date1_prefix}_year]": ""})
    if len(dates) >= 2:
        d, m, y = dates[1]
        tags.update({f"[{date2_prefix}_day]": f"{int(d):02d}", f"[{date2_prefix}_month]": f"{int(m):02d}", f"[{date2_prefix}_year]": y})
    else:
        tags.update({f"[{date2_prefix}_day]": "", f"[{date2_prefix}_month]": "", f"[{date2_prefix}_year]": ""})
    return tags

def test_full_replacement():
    meta = {
        "period_in": "Từ 01/01/2025 đến 31/12/2025",
        "name_vn": "ABC Co"
    }
    
    # Simulate sub_map
    subs = {"[v_name]": meta["name_vn"]}
    subs.update(parse_period_to_tags(meta["period_in"], "p1", "p2"))
    
    print("Substitution Map:")
    for k, v in subs.items():
        print(f"  {k} -> '{v}'")
        
    # Simulate DataFrame
    data = {
        "Vietnamese": ["Công ty [v_name]", "Giai đoạn [p1_day]/[p1_month]/[p1_year]"],
        "E": ["Company [v_name]", "Period [p1_day]/[p1_month]/[p1_year]"]
    }
    df = pd.DataFrame(data)
    
    # Perform replacement
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].apply(lambda x: clean_text(x))
            for tag, val in subs.items():
                print(f"Replacing {tag} with {val} in column {col}")
                df[col] = df[col].str.replace(tag, str(val), regex=False)
    
    print("\nResulting DataFrame:")
    print(df)

if __name__ == "__main__":
    test_full_replacement()
