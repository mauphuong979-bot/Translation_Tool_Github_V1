"""Dev-only page that renders the company-name confirmation table with sample data, for visual checks.
Saving writes to the file in COMPANY_NAMES_FILE (default: a temp folder), never to the project's company_names.csv.
streamlit run scratch/ui_harness_company_names.py --server.port 8503"""
import os
from pathlib import Path
import sys
import tempfile

import streamlit as st

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import company_names_store  # noqa: E402
import review_import_ui  # noqa: E402

import usage_logger  # noqa: E402
usage_logger.get_gsheet_client = lambda: (None, "harness: no Google Sheet")  # never write harness names to the real Sheet
company_names_store.LOCAL_FILE = os.environ.get("COMPANY_NAMES_FILE", str(Path(tempfile.gettempdir()) / "harness_company_names.csv"))
review_import_ui.log_event = lambda *args: None  # never write harness actions to the real usage log
st.set_page_config(layout="wide")
st.markdown(f"<style>{(ROOT / 'style.css').read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)
st.caption("Lưu vào: " + company_names_store.LOCAL_FILE)
review_import_ui._render_company_names({
    "company_names": {
        "Công ty TNHH Thương mại và Dịch vụ ABC": {"international": "ABC Trading and Services Co., Ltd", "source_url": "https://masothue.com/0101234567-abc"},
        "Ngân hàng TMCP Ngoại thương Việt Nam": {"international": "Joint Stock Commercial Bank for Foreign Trade of Vietnam", "source_url": "", "saved": True},
    },
    "searched_names": ["Công ty TNHH Thương mại và Dịch vụ ABC", "Công ty Cổ phần XYZ Miền Nam"],
})
