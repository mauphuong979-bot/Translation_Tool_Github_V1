
import re
import sys

def clean_text(text):
    import unicodedata
    text = unicodedata.normalize('NFC', str(text))
    # Original regex:
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f\xad\u200b\u200c\u200d\u2060\ufeff\xb7\u2022\u202a-\u202e\u200e\u200f]', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

text_with_tab = "6.\tDoanh thu hoạt động tài chính"
cleaned = clean_text(text_with_tab)

print(f"Original: {repr(text_with_tab)}")
print(f"Cleaned : {repr(cleaned)}")

key = "Doanh thu hoạt động tài chính"
print(f"In? {key in cleaned}")

# Test the regex pattern
pattern = re.compile(re.escape(key), re.IGNORECASE)
print(f"Regex Match? {bool(pattern.search(cleaned))}")
