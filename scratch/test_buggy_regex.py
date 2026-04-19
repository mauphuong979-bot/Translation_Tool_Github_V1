
import re
import unicodedata

def clean_text(text):
    text = unicodedata.normalize('NFC', str(text))
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f\xad\u200b\u200c\u200d\u2060\ufeff\xb7\u2022\u202a-\u202e\u200e\u200f]', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    # The problematic line:
    text = re.sub(r'(^|\s)(\d+)\.\s+', r'\1\2. ', text)
    return text.strip()

t1 = "6.\tDoanh thu"
c1 = clean_text(t1)
print(f"6.TAB: {repr(c1)}")

t2 = "6. Doanh thu"
c2 = clean_text(t2)
print(f"6.SPACE: {repr(c2)}")

t3 = "6.Doanh thu"
c3 = clean_text(t3)
print(f"6.NOSPC: {repr(c3)}")
