
import re

def swap_vn_to_en_number_separators(text):
    if not text:
        return text
    pattern = re.compile(r'(?<!\d)(?:\d{1,3}(?:\.\d{3})+(?:,\d+)?|\d+,\d+)(?!\d)')
    
    def replace_func(match):
        val = match.group(0)
        return val.replace('.', 'TEMP_DOT').replace(',', '.').replace('TEMP_DOT', ',')
        
    return pattern.sub(replace_func, text)

# Test cases
test_cases = [
    "91.504.195",
    "50.000.000",
    "1.021.000.000",
    "3.588.883.200",
    " 4.882.149.339 ",
    "1.234,56",
    "6,78"
]

for tc in test_cases:
    print(f"'{tc}' -> '{swap_vn_to_en_number_separators(tc)}'")
