import pandas as pd
import re
import os

def inspect_tags():
    if not os.path.exists('Dictionary_v3.xlsx'):
        print("Dictionary_v3.xlsx not found")
        return
    
    df = pd.read_excel('Dictionary_v3.xlsx')
    all_tags = set()
    
    # Search for all [anything] patterns
    for col in df.columns:
        for val in df[col]:
            if pd.notnull(val):
                found = re.findall(r'\[.*?\]', str(val))
                all_tags.update(found)
    
    print("Tags found in Dictionary_v3.xlsx:")
    for tag in sorted(list(all_tags)):
        print(f"  - {tag}")

if __name__ == "__main__":
    inspect_tags()
