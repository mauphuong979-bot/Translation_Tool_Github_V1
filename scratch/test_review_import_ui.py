from pathlib import Path
import json
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

try:
    from streamlit.testing.v1 import AppTest
except ImportError:
    AppTest = None

from docx import Document
import test_review_roundtrip as fixtures


@unittest.skipIf(AppTest is None, "Streamlit is not installed in this runtime")
class ReviewImportUITests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.report_path = Path(self.directory.name) / "report.docx"
        self.worksheet_path = Path(self.directory.name) / "ai.docx"
        report, worksheet, _ = fixtures.ReviewRoundtripTests().pair()
        self.report_data = fixtures.saved(report)
        self.report_path.write_bytes(self.report_data)
        self.worksheet_path.write_bytes(fixtures.saved(worksheet))

    def app(self, current=True, empty=False):
        source = f'''
import io
from pathlib import Path
import sys
from unittest.mock import patch
import streamlit as st
sys.path.insert(0, {str(Path(__file__).resolve().parents[1])!r})
from review_import_ui import render_ai_import
class Upload(io.BytesIO):
    def __init__(self, path):
        super().__init__(Path(path).read_bytes())
        self.name = Path(path).name
report = Upload({str(self.report_path)!r})
worksheet = Upload({str(self.worksheet_path)!r})
if {current!r}:
    st.session_state.processed_outputs = {{"review": report.getvalue()}}
    st.session_state.processed_filename = "Report.docx"
def uploaded(label, **kwargs):
    if {empty!r}:
        return None
    return worksheet if label == "Bảng AI_Form đã được AI điền" else report
with patch("review_import_ui.st.file_uploader", side_effect=uploaded):
    render_ai_import()
'''
        return AppTest.from_string(source, default_timeout=15).run()

    def test_current_report_requires_confirmation_and_keeps_download(self):
        app = self.app()
        self.assertFalse(app.exception)
        app.button(key="ai_check_translations").click().run()
        self.assertFalse(app.exception)
        self.assertEqual(len(app.dataframe), 1)
        self.assertTrue(app.button(key="ai_apply_translations").disabled)
        app.checkbox[1].check().run()
        self.assertFalse(app.button(key="ai_apply_translations").disabled)
        app.button(key="ai_apply_translations").click().run()
        self.assertFalse(app.exception)
        self.assertEqual(len(app.success), 1)
        result = app.session_state["ai_import_result"]
        self.assertEqual(result["count"], 1)
        self.assertEqual(result["name"], "Report_AI_TC.docx")
        self.assertEqual(app.session_state["processed_outputs"]["review"], self.report_data)
        import io
        self.assertFalse(Document(io.BytesIO(result["data"])).comments)
        app.run()
        self.assertEqual(len(app.get("download_button")), 1)
        self.assertEqual(app.session_state["ai_import_result"]["data"], result["data"])
        app.multiselect[0].set_value([]).run()
        self.assertFalse(app.checkbox[1].value)
        self.assertTrue(app.button(key="ai_apply_translations").disabled)
        self.assertFalse(app.get("download_button"))

    def test_uploaded_report_works_without_previous_session(self):
        app = self.app(current=False)
        app.button(key="ai_check_translations").click().run()
        self.assertFalse(app.exception)
        self.assertEqual(len(app.multiselect[0].value), 1)

    def test_legacy_uploaded_report_gets_a_short_non_stacking_output_name(self):
        self.report_path = self.report_path.with_name("Bao_cao_Ht_tool_111126 1737_review.docx")
        self.report_path.write_bytes(self.report_data)
        app = self.app(current=False)
        app.button(key="ai_check_translations").click().run()
        app.checkbox[0].check().run()
        app.button(key="ai_apply_translations").click().run()
        self.assertFalse(app.exception)
        self.assertEqual(app.session_state["ai_import_result"]["name"], "Bao_cao_Ht_111126_1737_AI_TC.docx")

    def test_changed_input_hides_stale_preview_and_invalid_ids_show_error(self):
        app = self.app(current=False)
        app.button(key="ai_check_translations").click().run()
        worksheet = Document(self.worksheet_path)
        worksheet.tables[0].cell(1, 0).text = "WRONG-ID"
        worksheet.save(self.worksheet_path)
        app.run()
        self.assertFalse(app.dataframe)
        self.assertFalse(app.get("download_button"))
        app.button(key="ai_check_translations").click().run()
        self.assertFalse(app.exception)
        self.assertIn("không thuộc báo cáo này", app.error[0].value)

    def test_missing_uploads_show_instructions_without_exception(self):
        app = self.app(current=False, empty=True)
        self.assertFalse(app.exception)
        self.assertTrue(app.info)
        self.assertFalse(app.button)

    def test_note_column_is_wider_and_full_notes_are_available(self):
        app = self.app()
        app.button(key="ai_check_translations").click().run()
        self.assertFalse(app.exception)
        columns = json.loads(app.dataframe[0].proto.columns)
        self.assertGreater(columns["Ghi chú"]["width"], columns["Nội dung gốc"]["width"])
        self.assertGreater(columns["Ghi chú"]["width"], columns["Bản dịch AI"]["width"])
        self.assertTrue(columns["Mã"]["pinned"])
        self.assertGreater(app.dataframe[0].proto.row_height, 32)
        self.assertEqual(len(app.selectbox), 1)
        self.assertTrue(any("Ghi chú (đầy đủ)" in element.value for element in app.markdown))

    def test_c009_is_shown_as_existing_not_selected_and_not_newly_applied(self):
        report, worksheet, items = fixtures.ReviewRoundtripTests().pair(count=10, prior_revision_indices=(8,))
        self.report_data = fixtures.saved(report)
        self.report_path.write_bytes(self.report_data)
        self.worksheet_path.write_bytes(fixtures.saved(worksheet))
        original_c009 = report.paragraphs[8]._p.xml
        c009 = items[8]["review_id"]
        selected = items[1]["review_id"]
        app = self.app()
        app.button(key="ai_check_translations").click().run()
        self.assertFalse(app.exception)
        self.assertNotIn(c009, app.multiselect[0].options)
        self.assertEqual(app.selectbox[0].value, 8)
        self.assertEqual(app.dataframe[0].value.iloc[8]["Trạng thái"], "Track Changes có sẵn")
        self.assertTrue(any("TRƯỚC khi nhập AI" in element.value for element in app.info))
        self.assertTrue(any("Existing prior-report translation" in element.value for element in app.text))
        app.multiselect[0].set_value([selected]).run()
        app.checkbox[1].check().run()
        app.button(key="ai_apply_translations").click().run()
        self.assertFalse(app.exception)
        result = app.session_state["ai_import_result"]
        self.assertEqual(result["applied_ids"], [selected])
        self.assertIn(c009, result["preserved_ids"])
        self.assertIn(c009, result["preexisting_revision_ids"])
        import io
        output = Document(io.BytesIO(result["data"]))
        self.assertEqual(output.paragraphs[8]._p.xml, original_c009)
        self.assertEqual(len(output.comments), 9)
        self.assertTrue(any(element.value == "Bản dịch AI CHỈ được áp dụng cho: " + selected for element in app.caption))

    def test_old_preview_is_invalidated_without_reading_old_result_fields(self):
        app = self.app()
        app.button(key="ai_check_translations").click().run()
        preview = app.session_state["ai_import_preview"]
        preview.pop("version")
        app.session_state["ai_import_preview"] = preview
        app.run()
        self.assertFalse(app.exception)
        self.assertFalse(app.dataframe)
        self.assertTrue(any("Kiểm tra bản dịch AI lần nữa" in element.value for element in app.info))


if __name__ == "__main__":
    unittest.main()
