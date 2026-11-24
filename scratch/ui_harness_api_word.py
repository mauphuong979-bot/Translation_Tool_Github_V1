"""Dev-only page: the API translation flow with a fake Claude client (no network, no cost), to check the Word round trip.
streamlit run scratch/ui_harness_api_word.py --server.port 8507"""
from pathlib import Path
import sys
from unittest.mock import patch

import streamlit as st

ROOT = Path(__file__).resolve().parents[1]
sys.path[:0] = [str(ROOT), str(ROOT / "scratch")]
import review_import_ui  # noqa: E402
from ai_translate_api import AIConfig  # noqa: E402
from test_ai_translate_api import FakeClient, echo  # noqa: E402

st.set_page_config(layout="wide")
st.markdown(f"<style>{(ROOT / 'style.css').read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)
if "processed_outputs" not in st.session_state:
    st.session_state.processed_outputs = {"review": (ROOT / "scratch/review_enhancement_qa/01_Report.docx").read_bytes()}
    st.session_state.processed_filename = "Report.docx"
config = AIConfig(api_key="fake", batch_size=5, max_workers=1)
client = FakeClient(lambda payload: echo(payload, lambda entry: "Bản dịch giả lập cho " + entry["review_id"]))
with patch("review_import_ui._api_config", return_value=config), patch("ai_translate_api.make_client", return_value=client), \
        patch("review_import_ui.log_event"), patch("review_import_ui._saved_names", return_value={}):
    review_import_ui.render_ai_import()
