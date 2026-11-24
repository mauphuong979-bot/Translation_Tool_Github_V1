"""Dev-only page: the AI import review table with the QA report and a completed AI_Form preloaded (no login, no API).
streamlit run scratch/ui_harness_review_table.py --server.port 8506"""
import io
from pathlib import Path
import sys
from unittest.mock import patch

import streamlit as st

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from review_import_ui import render_ai_import  # noqa: E402

QA = ROOT / "scratch/review_enhancement_qa"
st.set_page_config(layout="wide")
st.markdown(f"<style>{(ROOT / 'style.css').read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)
if "processed_outputs" not in st.session_state:
    st.session_state.processed_outputs = {"review": (QA / "01_Report.docx").read_bytes()}
    st.session_state.processed_filename = "Report.docx"


class Upload(io.BytesIO):
    def __init__(self, path):
        super().__init__(path.read_bytes())
        self.name = path.name


def uploaded(label, **kwargs):
    return Upload(QA / "03_AI_Completed.docx") if label == "Bảng AI_Form đã được AI điền" else None


with patch("review_import_ui._api_config", return_value=None), patch("review_import_ui.st.file_uploader", side_effect=uploaded):
    render_ai_import()
