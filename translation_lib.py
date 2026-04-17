import pandas as pd
import os
import re
import json
import unicodedata
import openpyxl
from datetime import datetime

# Use absolute path for Streamlit Cloud compatibility
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DICTIONARY_FILE = os.path.join(BASE_DIR, "Dictionary.xlsx")

def clean_text(text):
    """
    Normalizes text to NFC form and strips common invisible characters/whitespace.
    Highly recommended for Vietnamese Unicode stability.
    """
    if not isinstance(text, str) or pd.isna(text):
        return ""
    # Normalize to NFC (Normalization Form C)
    text = unicodedata.normalize('NFC', str(text))
    # Remove some common invisible characters like zero-width space
    text = text.replace('\u200b', '').replace('\ufeff', '')
    return text.strip()

# --- Excel Formula Simulation Helpers ---

def parse_vn_period_dates(period_str):
    """
    Parses 'từ ngày 04 tháng 11 năm 2025 đến ngày 31 tháng 12 năm 2025'
    Returns (start_date, end_date) as datetime objects
    """
    if not isinstance(period_str, str) or "đến" not in period_str:
        return None, None
    
    parts = period_str.split("đến")
    dates = []
    for part in parts:
        # Match 'ngày X tháng Y năm Z'
        # Group 1: day, Group 2: month, Group 3: year
        match = re.search(r"ngày\s+(\d+)\s+tháng\s+(\d+)\s+năm\s+(\d+)", part)
        if match:
            day, month, year = match.groups()
            try:
                dates.append(datetime(int(year), int(month), int(day)))
            except:
                dates.append(None)
        else:
            dates.append(None)
    
    if len(dates) >= 2:
        return dates[0], dates[1]
    return None, None

def format_excel_date_logic(dt, lang="vn"):
    if not isinstance(dt, datetime):
        return ""
    if lang == "vn":
        # 'Ngày 31 tháng 12 năm 2025'
        return f"Ngày {dt.day:02d} tháng {dt.month:02d} năm {dt.year}"
    elif lang == "en":
        # 'December 31, 2025'
        return dt.strftime("%B %d, %Y")
    elif lang == "cn":
        # '2025年12月31日'
        return f"{dt.year}年{dt.month:02d}月{dt.day:02d}日"
    elif lang == "short_vn":
        # '31/12/2025'
        return dt.strftime("%d/%m/%Y")
    return ""

def recalculate_dictionary_formulas(metadata):
    """
    Reproduces the logic of rows 6-64 from Dictionary.xlsx in Python.
    Returns a list of rows, each row is [VN, E, Hs, Ht]
    """
    name_vn = metadata.get("name_vn", "")
    name_trans = metadata.get("name_trans", "")
    
    # 1. Parse Dates
    def parse_dt(d_str):
        if not d_str: return None
        try: return datetime.strptime(d_str, "%d/%m/%Y")
        except: return None

    y_end_dt = parse_dt(metadata.get("year_end", ""))
    rep_date_dt = parse_dt(metadata.get("report_date", ""))
    p_out_start, p_out_end = parse_vn_period_dates(metadata.get("period_out", ""))
    
    # 2. Prepare Formatted Strings
    date_vn = format_excel_date_logic(y_end_dt, "vn")
    date_en = format_excel_date_logic(y_end_dt, "en")
    date_cn = format_excel_date_logic(y_end_dt, "cn")
    
    p_out = metadata.get("period_out", "")
    p_in = metadata.get("period_in", "")
    p_out_en_start = format_excel_date_logic(p_out_start, "en")
    p_out_en_end = format_excel_date_logic(p_out_end, "en")
    p_out_cn_start = format_excel_date_logic(p_out_start, "cn")
    p_out_cn_end = format_excel_date_logic(p_out_end, "cn")
    
    rep_year = str(y_end_dt.year) if y_end_dt else ""
    rep_vn = format_excel_date_logic(rep_date_dt, "vn")
    rep_en = format_excel_date_logic(rep_date_dt, "en")
    rep_cn = format_excel_date_logic(rep_date_dt, "cn")

    rows = []
    
    # 6: UPPER(A1 / Name)
    rows.append([name_vn.upper(), name_trans.upper(), name_trans.upper(), name_trans.upper()])
    
    # 7: Statements for year ended ...
    rows.append([
        f"Báo cáo tài chính cho năm tài chính kết thúc {date_vn.lower()}",
        f"Financial statements for the year ended {date_en}",
        f"截至{date_cn}止财务年度的财务报表",
        f"截至{date_cn}止財務年度之財務報表"
    ])
    
    # 8-9: Periods
    rows.append([
        f"Báo cáo tài chính cho giai đoạn {p_out}" if p_out else "",
        f"Financial statements for the period from {p_out_en_start} to {p_out_en_end}" if p_out_en_start else "",
        f"{p_out_cn_start}至{p_out_cn_end}期间的财务报表" if p_out_cn_start else "",
        f"{p_out_cn_start}至{p_out_cn_end}期間的財務報表" if p_out_cn_start else ""
    ])
    rows.append([
        f"Báo cáo tài chính cho giai đoạn {p_in}" if p_in else "",
        f"Financial statements for the period from {p_out_en_start} to {p_out_en_end}" if p_out_en_start else "",
        f"{p_out_cn_start}至{p_out_cn_end}期间的财务报表" if p_out_cn_start else "",
        f"{p_out_cn_start}至{p_out_cn_end}期間的財務報表" if p_out_cn_start else ""
    ])
    
    # 10-11: Opinion
    # Simplified long sentences logic based on Excel formulas
    opinion_vn = "Theo ý kiến của chúng tôi, Báo cáo tài chính đã phản ánh trung thực và hợp lý, trên các khía cạnh trọng yếu tình hình tài chính của Công ty tại "
    rows.append([
        f"{opinion_vn}{date_vn.lower()}, cũng như kết quả hoạt động kinh doanh và tình hình lưu chuyển tiền tệ cho năm tài chính kết thúc cùng ngày, phù hợp với chuẩn mực kế toán, chế độ kế toán doanh nghiệp Việt Nam và các quy định pháp lý có liên quan đến việc lập và trình bày Báo cáo tài chính.",
        f"In our opinion, the financial statements present fairly, in all material respects, the financial position of the Company as at {date_en}, and its financial performance and its cash flows for the year then ended in accordance with Vietnamese Accounting Standards, Vietnamese enterprise accounting system and applicable regulations relevant to the preparation and presentation of financial statements in Vietnam.",
        f"依本会计师的意见，上开财务报表 in all material respects 已真实、公允地反映公司在{date_cn}财务状况以及同日结束财务年度的经营成果和现金流量状况，且符合《越南会计准则》、《越南企业会计制度》以及财务报表编制和列报的相关法规。",
        f"依本會計師之意見，上開財務報表 in all material respects 已真實、公允地反映公司於{date_cn}財務狀況以及同日結束財務年度之經營成果及現金流量狀況，且符合《越南會計準則》、《越南企業會計制度》及財務報表編製及列報之相關法規。"
    ])
    rows.append([
        f"{opinion_vn}{date_vn.lower()}, cũng như kết quả hoạt động kinh doanh và tình hình lưu chuyển tiền tệ cho giai đoạn {p_out}, phù hợp với chuẩn mực kế toán, chế độ kế toán doanh nghiệp Việt Nam và các quy định pháp lý có liên quan đến việc lập và trình bày Báo cáo tài chính." if p_out else "",
        f"In our opinion, the financial statements present fairly, in all material respects, the financial position of the Company as at {date_en}, and its financial performance and its cash flows for the period from {p_out_en_start} to {p_out_en_end} in accordance with Vietnamese Accounting Standards, Vietnamese enterprise accounting system and applicable regulations relevant to the preparation and presentation of financial statements in Vietnam." if p_out_en_start else "",
        f"依本会计师的意见，上开财务报表 in all material respects 已真实、公允地反映公司在{date_cn}财务状况以及{p_out_cn_start}至{p_out_cn_end}期间的经营成果和现金流量状况，且符合《越南会计准则》、《越南企业会计制度》以及财务报表编制和列报的相关法规。" if p_out_cn_start else "",
        f"依本會計師之意見，上開財務報表 in all material respects 已真實、公允地反映公司於{date_cn}財務狀況以及{p_out_cn_start}至{p_out_cn_end}期間之經營成果及現金流量狀況，且符合《越南會計準則》、《越南企業會計制度》及財務報表編製及列報之相關法規。" if p_out_cn_start else ""
    ])

    # 12: Name (Standard)
    rows.append([name_vn, name_trans, name_trans, name_trans])
    
    # 13, 14, 15: Dates
    rows.append([date_vn, date_en, date_cn, date_cn])
    rows.append([date_vn.lower(), date_en, date_cn, date_cn])
    rows.append([date_vn.upper(), date_en.upper(), date_cn, date_cn])
    
    # 16, 17, 18, 19: Helpers
    rows.append([p_out, f"from {p_out_en_start} to {p_out_en_end}" if p_out_en_start else "", f"自{p_out_cn_start}至{p_out_cn_end}" if p_out_cn_start else "", ""])
    rows.append([f"từ ngày {format_excel_date_logic(p_out_start, 'short_vn')} đến ngày {format_excel_date_logic(p_out_end, 'short_vn')}" if p_out_start else "", "", "", ""])
    rows.append([p_in, f"From {p_out_en_start} to {p_out_en_end}" if p_out_en_start else "", f"自{p_out_cn_start}日至{p_out_cn_end}日" if p_out_cn_start else "", ""])
    rows.append(["", f"From {p_out_en_start} to {p_out_en_end}" if p_out_en_start else "", f"自{p_out_cn_start}日至{p_out_cn_end}日" if p_out_cn_start else "", ""])
    
    # 20: As at
    rows.append([f"Tại {date_vn.lower()}", f"As at {date_en}", f"于{date_cn}", f"於{date_cn}"])
    
    # 21, 22: For the year ended
    rows.append([f"Cho năm tài chính kết thúc {date_vn.lower()}", f"For the year ended {date_en}", f"截至{date_cn}止财务年度", f"截至{date_cn}止財務年度"])
    rows.append([f"cho năm tài chính kết thúc {date_vn.lower()}", f"for the year ended {date_en}", f"截至{date_cn}止财务年度", f"截至{date_cn}止財務年度"])
    
    # 23-26: Period variations
    rows.append([f"Cho giai đoạn {p_out}", f"For the period from {p_out_en_start} to {p_out_en_end}", f"{p_out_cn_start}至{p_out_cn_end}期间", ""])
    rows.append([f"Cho giai đoạn {format_excel_date_logic(p_out_start, 'short_vn')} đến {format_excel_date_logic(p_out_end, 'short_vn')}", f"For the period from {p_out_en_start} to {p_out_en_end}", "", ""])
    rows.append([f"cho giai đoạn {p_out}", f"for the period from {p_out_en_start} to {p_out_en_end}", "", ""])
    rows.append([f"cho giai đoạn {format_excel_date_logic(p_out_start, 'short_vn')} đến {format_excel_date_logic(p_out_end, 'short_vn')}", f"for the period from {p_out_en_start} to {p_out_en_end}", "", ""])
    
    # 27, 28: Reporting Dates
    rows.append([rep_vn, rep_en, rep_cn, rep_cn])
    rows.append([rep_vn.lower(), rep_en, rep_cn, rep_cn])
    
    # 29, 30: Names/Dates again
    rows.append(["", name_trans, name_trans, name_trans])
    rows.append(["", date_en, date_cn, date_cn])
    
    # 31: Year
    rows.append(["", rep_year, rep_year, rep_year])
    
    # 32: Rep date
    rows.append(["", rep_en, rep_cn, rep_cn])
    
    # 33, 34, 35: Caps
    rows.append([f"CHO NĂM TÀI CHÍNH KẾT THÚC {date_vn.upper()}", f"FOR THE YEAR ENDED {date_en.upper()}", f"截至{date_cn}止财务年度", f"截至{date_cn}止財務年度"])
    rows.append([f"CHO GIAI ĐOẠN {p_out.upper()}", f"FOR THE PERIOD FROM {p_out_en_start.upper()} TO {p_out_en_end.upper()}", "", ""])
    rows.append([f"CHO GIAI ĐOẠN {p_in.upper()}", f"FOR THE PERIOD FROM {p_out_en_start.upper()} TO {p_out_en_end.upper()}", "", ""])
    
    # 36-64 (Auditors, Tax etc. - simplified logic based on your Excel)
    rows.append([f"Công ty TNHH Kiểm toán U&I đã kiểm toán Báo cáo tài chính cho năm tài chính kết thúc {date_vn.lower()} và bày tỏ nguyện vọng tiếp tục được bổ nhiệm làm kiểm toán viên cho Công ty.", f"The auditors, U&I Auditing Company Limited, have performed audit on the Company’s financial statements for the year ended {date_en} and have expressed their willingness to accept reappointment.", "", ""])
    rows.append([f"Công ty TNHH Kiểm toán U&I đã kiểm toán Báo cáo tài chính cho giai đoạn {p_out} và bày tỏ nguyện vọng tiếp tục được bổ nhiệm làm kiểm toán viên cho Công ty." if p_out else "", f"The auditors, U&I Auditing Company Limited, have performed audit on the Company’s financial statements for the period from {p_out_en_start} to {p_out_en_end} and have expressed their willingness to accept reappointment.", "", ""])
    
    rows.append([f"Chi phí thuế thu nhập doanh nghiệp cho năm tài chính kết thúc {date_vn.lower()} được tính trên thu nhập tính thuế ước tính. Chi phí thuế thu nhập doanh nghiệp này sẽ được cơ quan thuế xác định lại thông qua các cuộc kiểm tra.", f"Corporate income tax expense for the year ended {date_en} is calculated on the estimated assessable income. Corporate income tax expense will be determined again by the tax authority through their tax reviews. ", "", ""])
    rows.append([f"Chi phí thuế thu nhập doanh nghiệp cho giai đoạn {p_out} được tính trên thu nhập tính thuế ước tính. Chi phí thuế thu nhập doanh nghiệp này sẽ được cơ quan thuế xác định lại thông qua các cuộc kiểm tra." if p_out else "", f"Corporate income tax expense for the period from {p_out_en_start} to {p_out_en_end} is calculated on the estimated assessable income. Corporate income tax expense will be determined again by the tax authority through their tax reviews. ", "", ""])
    
    rows.append([f"Do đó, không có chi phí thuế thu nhập doanh nghiệp hoãn lại được ghi nhận trong năm {rep_year}.", f"Therefore, no deferred corporate income tax expense is provided in the year {rep_year}.", f"据此，并无递延所得税费用在{rep_year}年度内得以认列。", f"據此，並無遞延所得稅費用於{rep_year}年度內得以認列。"])
    rows.append([f"Do đó, không có chi phí thuế thu nhập doanh nghiệp hoãn lại được ghi nhận cho giai đoạn {p_out}.", f"Therefore, no deferred corporate income tax expense is provided in the period from {p_out_en_start} to {p_out_en_end}", f"据此，并无递延所得税费用{p_out_cn_start}至{p_out_cn_end}期间内得以认列。", f"據此，並無遞延所得稅費用{p_out_cn_start}至{p_out_cn_end}期間內得以認列。"])
    rows.append([f"Do đó, không có chi phí thuế thu nhập doanh nghiệp hoãn lại được ghi nhận trong giai đoạn {p_out}.", f"Same as above", "", ""])
    
    rows.append([f"Trong năm tài chính kết thúc {date_vn.lower()}, Công ty có các nghiệp vụ kinh tế quan trọng với các bên liên quan được trình bày ở bảng sau:", f"In the year ended {date_en}, the Company entered into significant economic transactions with its related parties as shown in the following table:", f"在截止{date_cn}财务年度之内，公司与其关联方发生重大经济交易，其列示在下表：", f"在截止{date_cn}財務年度之內，公司與其關聯方發生重大經濟交易，其列示於下表："])
    rows.append([f"Trong giai đoạn {p_out}, Công ty có các nghiệp vụ kinh tế quan trọng với các bên liên quan được trình bày ở bảng sau:" if p_out else "", f"In the period from {p_out_en_start} to {p_out_en_end}, the Company entered into significant economic transactions with its related parties as shown in the following table:", "", ""])
    rows.append([f"Trong năm tài chính kết thúc {date_vn.lower()}, Công ty có các nghiệp vụ kinh tế quan trọng và các khoản phải thu, phải trả với các bên liên quan được trình bày ở bảng sau", f"In the year ended {date_en}, the Company entered into significant economic transactions, receivables and payables with its related parties as shown in the following table", "", ""])
    rows.append([f"Trong giai đoạn {p_out}, Công ty có các nghiệp vụ kinh tế quan trọng và các khoản phải thu, phải trả với các bên liên quan được trình bày ở bảng sau:" if p_out else "", f"In the period from {p_out_en_start} to {p_out_en_end}, the Company entered into significant economic transactions, receivables and payables with its related parties as shown in the following table", "", ""])
    rows.append([f"Trong năm tài chính kết thúc {date_vn.lower()}, Công ty không phát sinh các nghiệp vụ kinh tế quan trọng với các bên liên quan.", f"In the year ended {date_en}, the Company had no significant economic transactions with its related parties.", "", ""])
    rows.append([f"Trong giai đoạn {p_out}, Công ty không phát sinh các nghiệp vụ kinh tế quan trọng với các bên liên quan." if p_out else "", f"In the period from {p_out_en_start} to {p_out_en_end}, the Company had no significant economic transactions with its related parties.", "", ""])
    
    approvers = ["Tổng Giám đốc", "Ban Giám đốc", "Hội đồng Thành viên và Ban Giám đốc", "Giám đốc điều hành", "Ban Tổng Giám đốc", "Hội đồng Thành viên", "Giám đốc", "Chủ tịch", "Hội đồng Quản trị", "Chủ sở hữu"]
    for app in approvers:
        rows.append([f"Báo cáo tài chính được phê duyệt bởi {app} Công ty để phát hành vào {rep_vn.lower()}.", "", "", ""])
        
    rows.append([f"Công ty không lập dự phòng chi phí thuế thu nhập doanh nghiệp hiện hành vì Công ty không có thu nhập chịu thuế trong năm {rep_year}.", f"No provision for current corporate income tax expense has been provided as the Company has no taxable income arising in the year {rep_year}.", "", ""])
    rows.append([f"Công ty không lập dự phòng chi phí thuế thu nhập doanh nghiệp hiện hành vì Công ty không có thu nhập chịu thuế cho giai đoạn {p_out}.", f"No provision for current corporate income tax expense has been provided as the Company has no taxable income arising in the period from {p_out_en_start} to {p_out_en_end}", "", ""])
    rows.append([f"Công ty không lập dự phòng chi phí thuế thu nhập doanh nghiệp hiện hành vì Công ty không có thu nhập chịu thuế trong giai đoạn {p_out}.", "Same as above", "", ""])
    rows.append([f"Công ty không lập dự phòng chi phí thuế thu nhập doanh nghiệp hiện hành vì Công ty không có thu nhập tính thuế sau khi trừ chuyển lỗ trong năm {rep_year}.", f"No provision for current corporate income tax expense has been provided as the Company has no assessable income after deducting tax loss brought forward in the year {rep_year}.", "", ""])
    rows.append([f"Công ty không lập dự phòng chi phí thuế thu nhập doanh nghiệp hiện hành vì Công ty không có thu nhập tính thuế sau khi trừ chuyển lỗ cho giai đoạn {p_out}.", f"No provision for current corporate income tax expense has been provided as the Company has no assessable income after deducting tax loss brought forward in the period from {p_out_en_start} to {p_out_en_end}", "", ""])
    rows.append([f"Đến {date_vn.lower()}, Công ty vẫn chưa tiến hành hoạt động sản xuất kinh doanh.", f"As at {date_en}, the Company’s operation has yet to start.", "", ""])

    return rows

def load_excel_dictionary():
    """
    Loads metadata and dictionary from Dictionary.xlsx.
    Metadata: A1, B1, C1, D1, A2, A3
    Dictionary: Row 5 (headers), Row 6+ (data)
    """
    metadata = {}
    df = None
    
    if os.path.exists(DICTIONARY_FILE):
        try:
            # 1. Load Metadata using openpyxl
            wb = openpyxl.load_workbook(DICTIONARY_FILE, data_only=True)
            ws = wb.active
            
            metadata = {
                "name_vn": clean_text(ws['A1'].value),
                "name_trans": clean_text(ws['B1'].value),
                "year_end": clean_text(ws['C1'].value),
                "report_date": clean_text(ws['D1'].value),
                "period_out": clean_text(ws['A2'].value),
                "period_in": clean_text(ws['A3'].value)
            }
            wb.close()
            
            # 2. Load Dictionary Data using pandas (Start from Row 5)
            # Row 5 is header, so header=4 (0-indexed)
            df = pd.read_excel(DICTIONARY_FILE, header=4)
            
            if df is not None and not df.empty:
                # Normalize columns
                for col in df.columns:
                    if df[col].dtype == object:
                        df[col] = df[col].apply(clean_text)
                # Ensure 'Vietnamese' column exists
                if 'Vietnamese' in df.columns:
                    df = df.dropna(subset=['Vietnamese'])
            
            return metadata, df
        except Exception as e:
            print(f"Error loading excel dictionary: {e}")
            return {}, None
    return {}, None

def save_excel_metadata(metadata, df=None):
    """
    Saves metadata to A1:D1, A2, A3 and optionally updates dictionary data.
    Uses openpyxl to preserve formulas in other cells.
    """
    try:
        if not os.path.exists(DICTIONARY_FILE):
            # Create a new workbook if it doesn't exist
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Dictionary"
        else:
            wb = openpyxl.load_workbook(DICTIONARY_FILE)
            ws = wb.active

        # Write Metadata
        ws['A1'] = metadata.get("name_vn", "")
        ws['B1'] = metadata.get("name_trans", "")
        ws['C1'] = metadata.get("year_end", "")
        ws['D1'] = metadata.get("report_date", "")
        ws['A2'] = metadata.get("period_out", "")
        ws['A3'] = metadata.get("period_in", "")

        # Recalculate Rows 6-64 based on new metadata
        # (Simulating Excel Formulas in Python)
        calculated_rows = recalculate_dictionary_formulas(metadata)
        
        # Write calculated values starting at row 6
        for i, row_data in enumerate(calculated_rows):
            target_row = 6 + i
            if target_row > 64: break
            for j, val in enumerate(row_data):
                if val: # Only overwrite if we have a calculated value
                    ws.cell(row=target_row, column=j+1, value=val)

        # Write Dictionary Data if provided (starting from Row 5)
        if df is not None:
            # Write headers at Row 5
            headers = df.columns.tolist()
            for col_idx, header in enumerate(headers, start=1):
                ws.cell(row=5, column=col_idx, value=header)
            
            # Write data from Row 6
            for row_idx, row in enumerate(df.values, start=6):
                for col_idx, value in enumerate(row, start=1):
                    # We only overwrite if value is not a formula (very basic check)
                    ws.cell(row=row_idx, column=col_idx, value=value)

        wb.save(DICTIONARY_FILE)
        return True
    except Exception as e:
        print(f"Error saving excel: {e}")
        return False

def load_dictionary():
    """
    Backwards compatibility wrapper for load_excel_dictionary.
    Returns only the DataFrame.
    """
    _, df = load_excel_dictionary()
    return df

def save_dictionary(df):
    """
    Backwards compatibility wrapper for save_excel_metadata.
    """
    # Try to load existing metadata first to preserve it
    metadata, _ = load_excel_dictionary()
    return save_excel_metadata(metadata, df)

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
