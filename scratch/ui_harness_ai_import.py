"""Dev-only page that renders the AI import panel without the login, for visual checks.
streamlit run scratch/ui_harness_ai_import.py --server.port 8502"""
from pathlib import Path
import sys

import streamlit as st

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from review_import_ui import render_ai_import  # noqa: E402

st.set_page_config(layout="wide")
st.markdown(f"<style>{(ROOT / 'style.css').read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)
if "processed_outputs" not in st.session_state:
    st.session_state.processed_outputs = {"review": (ROOT / "scratch/review_enhancement_qa/01_Report.docx").read_bytes()}
    st.session_state.processed_filename = "Report.docx"
render_ai_import()
