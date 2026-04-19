import re
from collections import Counter
import unicodedata

def format_vn_date_to_digit(vn_date_str):
    if not vn_date_str:
        return vn_date_str
    vn_date_str = unicodedata.normalize('NFC', vn_date_str)
    match = re.search(r"(\d{1,2})\s+th\u00e1ng\s+(\d{1,2})\s+n\u0103m\s+(\d{4})", vn_date_str)
    if match:
        day, month, year = match.groups()
        return f"{int(day):02d}/{int(month):02d}/{year}"
    return vn_date_str

def extract_metadata_test(full_content):
    full_content = unicodedata.normalize('NFC', full_content)
    metadata = {
        "year_end": None,
        "period_out": None,
        "period_in": None,
    }
    
    # 2. Year-end date
    ye_pattern = re.compile(r"kết thúc ngày\s+(?:ngày\s+)?(\d{1,2}\s+tháng\s+\d{1,2}\s+năm\s+\d{4})", re.IGNORECASE)
    ye_match = ye_pattern.search(full_content)
    if ye_match:
        metadata["year_end"] = format_vn_date_to_digit(ye_match.group(1).strip())
    else:
        # Fallback 1: Common in Auditor's opinion paragraphs
        ye_pattern_2 = re.compile(r"tại ngày\s+(\d{1,2}\s+tháng\s+\d{1,2}\s+năm\s+\d{4})", re.IGNORECASE)
        ye_match_2 = ye_pattern_2.search(full_content)
        if ye_match_2:
            metadata["year_end"] = format_vn_date_to_digit(ye_match_2.group(1).strip())
            
    # 4. Period (out of table)
    period_pattern = re.compile(
        r"(t\u1eeb ng\u00e0y\s+\d{1,2}\s+th\u00e1ng\s+\d{1,2}\s+n\u0103m\s+\d{4}\s+\u0111\u1ebfn ng\u00e0y\s+\d{1,2}\s+th\u00e1ng\s+\d{1,2}\s+n\u0103m\s+\d{4})",
        re.IGNORECASE
    )
    all_periods = period_pattern.findall(full_content)
    if all_periods:
        most_common = Counter(all_periods).most_common(1)[0][0]
        metadata["period_out"] = most_common.strip()

    # 5. Periods (in table) - Short format
    short_period_pattern = re.compile(
        r"([Tt]\u1eeb\s+\d{1,2}/\d{1,2}/\d{4}\s+[\u0111\u00d0]\u1ebfn\s+\d{1,2}/\d{1,2}/\d{4})",
        re.IGNORECASE
    )
    all_short_periods = short_period_pattern.findall(full_content)
    if all_short_periods:
        counts = Counter(all_short_periods).most_common(2)
        metadata["period_in"] = counts[0][0].strip()

    # --- FINAL FALLBACK: Derive Year-end from Period if still missing ---
    if not metadata["year_end"]:
        if metadata["period_out"]:
            p_dates = re.findall(r"(\d{1,2})\s+th\u00e1ng\s+(\d{1,2})\s+n\u0103m\s+(\d{4})", metadata["period_out"])
            if len(p_dates) >= 2:
                d, m, y = p_dates[-1]
                metadata["year_end"] = f"{int(d):02d}/{int(m):02d}/{y}"
        if not metadata["year_end"] and metadata["period_in"]:
            s_dates = re.findall(r"(\d{1,2})/(\d{1,2})/(\d{4})", metadata["period_in"])
            if len(s_dates) >= 2:
                d, m, y = s_dates[-1]
                metadata["year_end"] = f"{int(d):02d}/{int(m):02d}/{y}"
    
    return metadata

# Test Case 1: "tại ngày" fallback
content1 = "tình hình tài chính của Công ty tại ngày 31 tháng 12 năm 2025"
res1 = extract_metadata_test(content1)
print(f"Test 1 (tai ngay): {res1['year_end']}")

# Test Case 2: Derive from period_out
content2 = "từ ngày 23 tháng 12 năm 2024 đến ngày 31 tháng 12 năm 2025"
res2 = extract_metadata_test(content2)
print(f"Test 2 (period_out): {res2['year_end']}")

# Test Case 3: Derive from period_in
content3 = "Từ 01/01/2025 đến 31/12/2025"
res3 = extract_metadata_test(content3)
print(f"Test 3 (period_in): {res3['year_end']}")
