import pandas as pd
import os
import re
import json
import unicodedata

# Use absolute path for Streamlit Cloud compatibility
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DICTIONARY_FILE = os.path.join(BASE_DIR, "dictionary.json")

def clean_text(text):
    """
    Normalizes text to NFC form and strips common invisible characters/whitespace.
    Highly recommended for Vietnamese Unicode stability.
    """
    if not isinstance(text, str):
        return text
    # Normalize to NFC (Normalization Form C)
    text = unicodedata.normalize('NFC', text)
    # Remove some common invisible characters like zero-width space
    text = text.replace('\u200b', '').replace('\ufeff', '')
    return text.strip()

def load_dictionary():
    """
    Loads the dictionary.json file and returns it as a DataFrame.
    """
    if os.path.exists(DICTIONARY_FILE):
        try:
            # Read JSON and return as DataFrame
            df = pd.read_json(DICTIONARY_FILE, encoding='utf-8')
            if not df.empty:
                 # Normalize all columns that might contain Vietnamese
                 for col in df.columns:
                     if df[col].dtype == object:
                         df[col] = df[col].apply(clean_text)
                 # Remove rows where Vietnamese is null
                 df = df.dropna(subset=['Vietnamese'])
            return df
        except Exception as e:
            print(f"Error loading dictionary: {e}")
            return None
    return None

def save_dictionary(df):
    """
    Saves a DataFrame to dictionary.json.
    """
    try:
        # Convert to records and save as JSON
        data = df.to_dict(orient='records')
        with open(DICTIONARY_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        return True
    except Exception as e:
        print(f"Error saving dictionary: {e}")
        return False

def get_translation_map(df, target_lang):
    """
    Creates a dictionary of Vietnamese -> Target Language.
    """
    if df is None or target_lang not in df.columns:
        return {}
    
    # Create map, dropping rows with null translations in the target column
    subset = df[['Vietnamese', target_lang]].dropna()
    return dict(zip(subset['Vietnamese'], subset[target_lang]))

def find_missing_terms(text, translation_map):
    """
    Optional: Find terms that might be missing (not used currently based on user request).
    """
    pass

def replace_text_in_paragraph(paragraph, translation_map):
    """
    Replaces text in a paragraph while attempting to preserve some simple run formatting.
    Note: Multi-run phrases are difficult to replace without losing specific run formatting.
    This implementation replaces the entire text if a match is found to ensure translation accuracy.
    """
    inline = paragraph.runs
    full_text = "".join(run.text for run in inline)
    
    # Normalize document text to ensure matching with dictionary
    full_text = clean_text(full_text)
    
    changed = False
    new_text = full_text
    
    # Sort keys by length descending to match longest phrases first
    sorted_keys = sorted(translation_map.keys(), key=len, reverse=True)
    
    for key in sorted_keys:
        if key in new_text:
            new_text = new_text.replace(key, translation_map[key])
            changed = True
            
    if changed:
        # If text changed, we simplify by replacing runs. 
        # This preserves paragraph-level formatting but might reset run-level formatting
        # within the translated text if it was mixed.
        if len(inline) > 0:
            # Clear all runs and add a new one with the replaced text
            p_text = new_text
            for run in inline:
                run.text = ""
            inline[0].text = p_text
        else:
            paragraph.add_run(new_text)
    return changed

def replace_text_in_document(doc, translation_map):
    """
    Performs global search and replace in paragraphs and tables.
    """
    count = 0
    # Process paragraphs
    for p in doc.paragraphs:
        if replace_text_in_paragraph(p, translation_map):
            count += 1
            
    # Process tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if replace_text_in_paragraph(p, translation_map):
                        count += 1
    return count
