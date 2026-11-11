import io
from pathlib import Path
import sys
from unittest.mock import patch

import streamlit as st

root = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(root))
from review_import_ui import render_ai_import

st.set_page_config(page_title="AI translation review QA", layout="wide")
st.markdown("<style>" + (root / "style.css").read_text(encoding="utf-8") + "</style>", unsafe_allow_html=True)
st.markdown("### Translation review")
directory = root / "scratch" / "review_enhancement_qa"
st.session_state.processed_outputs = {"review": (directory / "01_Report.docx").read_bytes()}
st.session_state.processed_filename = "QA_Report.docx"


class Upload(io.BytesIO):
    name = "03_AI_Completed.docx"


real_uploader = st.file_uploader


def uploaded(label, **kwargs):
    upload = real_uploader(label, **kwargs)
    if upload is not None:
        return upload
    if label == "AI-completed translation worksheet":
        return Upload((directory / "03_AI_Completed.docx").read_bytes())
    return None


with patch("review_import_ui.st.file_uploader", side_effect=uploaded):
    render_ai_import()
