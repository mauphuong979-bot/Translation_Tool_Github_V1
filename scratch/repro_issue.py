
import sys
import os
from docx import Document
import unicodedata

# Add root to path
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
import translation_lib as tl

def main():
    doc_path = r'd:\AI\Python\11_Project\Translation_Tool\Translation_Tool_Antigravity_V2\test.docx'
    doc = Document(doc_path)
    
    # Target phrase in Table 1 Cell(13,1)
    cell = doc.tables[0].cell(12, 0)
    
    # Prepare map
    translation_map = {
        "Doanh thu hoạt động tài chính": "Financial income"
    }
    
    # Run replacement
    # tl.replace_text_in_document(doc, translation_map, target_col="E")
    
    # Let's do it step by step to see where it fails
    cleaned_body_text = tl.clean_text(cell.text)
    prepared_list = tl.prepare_translation_list(translation_map, case_threshold=25)
    
    new_text, changed = tl.apply_translations_to_text(cleaned_body_text, prepared_list)
    
    with open('scratch/repro_log.txt', 'w', encoding='utf-8') as f:
        f.write(f"Original Text: {repr(cell.text)}\n")
        f.write(f"Cleaned Text: {repr(cleaned_body_text)}\n")
        f.write(f"Prepared List: {repr(prepared_list)}\n")
        f.write(f"New Text: {repr(new_text)}\n")
        f.write(f"Changed? {changed}\n")
        
        # Test the dict lookup
        f.write(f"Is 'Doanh thu hoạt động tài chính' in Cleaned Text? {'Doanh thu hoạt động tài chính' in cleaned_body_text}\n")
        
        # Test Case-insensitive match if it's a regex
        is_regex, pattern, val = prepared_list[0]
        if is_regex:
            f.write(f"Regex Match? {bool(pattern.search(cleaned_body_text))}\n")

if __name__ == "__main__":
    main()
