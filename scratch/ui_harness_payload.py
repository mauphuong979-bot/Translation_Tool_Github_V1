"""Dev-only page: the 'Xem dữ liệu sẽ gửi' table with long sample data and a sample API error. No API calls.
streamlit run scratch/ui_harness_payload.py --server.port 8504"""
from pathlib import Path
import sys

import streamlit as st

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import review_import_ui  # noqa: E402

st.set_page_config(layout="wide")
st.markdown(f"<style>{(ROOT / 'style.css').read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)
long_vn = ("Các thành viên Ban Giám đốc Công ty TNHH New Increase (VN) (gọi tắt là “Công ty”) đệ trình báo cáo này cùng với "
           "báo cáo tài chính đã được kiểm toán của Công ty cho năm tài chính kết thúc ngày 31 tháng 12 năm 2025.")
payload = [{"review_id": f"099FFE8E-C{index:03d}", "source": long_vn,
            "references": [{"label": "Báo cáo kỳ trước (99.5%)", "vietnamese": long_vn.replace("2025", "2024"),
                            "translation": "New Increase (VN) Company Limited（簡稱為“公司”）管理階層呈上此報告及截至2024年12月31日止財務年度之業經查核之財務報表。"}]
            + ([{"label": "Từ điển (91%)", "vietnamese": "Ban Giám đốc", "translation": "Board of Management"}] if index % 2 else []),
            **({"company_names": [{"vietnamese": "Công ty TNHH New Increase (VN)", "international": "New Increase (VN) Co., Ltd"}]} if index == 1 else {})}
           for index in range(1, 6)]
with st.expander("Xem dữ liệu sẽ gửi", expanded=True):
    review_import_ui._render_payload(payload, ["Công ty TNHH New Increase (VN)", "Ngân hàng TMCP Ngoại thương Việt Nam"])

import anthropic, httpx  # noqa: E401,E402
error = anthropic.AuthenticationError("Error code: 401 - {'type': 'error', 'error': {'type': 'authentication_error', 'message': 'Invalid bearer token'}}",
                                      response=httpx.Response(401, request=httpx.Request("POST", "https://api.anthropic.com/v1/messages")), body=None)
review_import_ui._show_api_error("Không gọi được AI", error)
