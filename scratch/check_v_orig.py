
import docx
import sys
import io

# Set stdout to utf-8
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

try:
    doc = docx.Document("fs_C&C_2025_V.docx")
    print("Searching for Form Indicator text in original DOCX...")
    for i, p in enumerate(doc.paragraphs):
        # Broad search for potential indicators
        t = p.text.upper()
        if ("MẪU" in t or "MÃ" in t or "FORM" in t) and "DN" in t:
            print(f"Para[{i}]: '{p.text}'")
            for r in p.runs:
                print(f"  Run: '{r.text}'")
except Exception as e:
    print(f"Error: {e}")
