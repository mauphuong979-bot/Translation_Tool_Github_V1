
import re
key = "Doanh thu"
escaped = re.escape(key)
print(f"Escaped: {repr(escaped)}")

# Let's check regex with escaped spaces
text = "6. Doanh thu"
pattern = re.compile(escaped, re.IGNORECASE)
print(f"Match: {bool(pattern.search(text))}")
