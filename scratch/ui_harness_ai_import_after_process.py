"""Dev-only page: a "process" button that renders a spinner (like app.py) above the AI import panel, to check that
the panel stays open on the first rerun after processing.
streamlit run scratch/ui_harness_ai_import_after_process.py --server.port 8508
Add ?stable=1 to wrap the processing step in its own container (the fix used in app.py)."""
from contextlib import nullcontext
from pathlib import Path
import sys
import time

import streamlit as st

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from review_import_ui import render_ai_import  # noqa: E402

st.set_page_config(layout="wide")
st.markdown(f"<style>{(ROOT / 'style.css').read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)
clicked = st.button("Xử lý báo cáo")
with st.container() if st.query_params.get("stable") == "1" else nullcontext():
    if clicked:
        with st.spinner("Đang xử lý báo cáo..."):
            time.sleep(0.5)
            st.session_state.processed_outputs = {"review": (ROOT / "scratch/review_enhancement_qa/01_Report.docx").read_bytes()}
            st.session_state.processed_filename = "Report.docx"
        # Like st_auto_download in app.py on Streamlit Cloud: an invisible element drawn only in the processing run.
        st.components.v1.html("<script>/* auto download */</script>", height=0)
if "processed_outputs" in st.session_state:
    render_ai_import()
