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

    def review_state(self, app, **changes):
        """Change the review table state the way ticking boxes or editing cells in the table does."""
        state = dict(app.session_state["ai_review_state"])
        state.update(changes, version=state["version"] + 1)
        app.session_state["ai_review_state"] = state
        return app.run()

    def table(self, app):
        return app.dataframe[0].value

    def confirm(self, app):
        return next(box for box in app.checkbox if "Tôi đã rà soát" in box.label).check().run()

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
        self.review_state(app, excluded=list(self.table(app)["Mã"]))
        self.assertFalse(self.table(app)["Áp dụng"].any())
        self.assertTrue(app.button(key="ai_apply_translations").disabled)
        self.assertFalse(app.get("download_button"))

    def test_uploaded_report_works_without_previous_session(self):
        app = self.app(current=False)
        app.button(key="ai_check_translations").click().run()
        self.assertFalse(app.exception)
        self.assertEqual(list(self.table(app)["Áp dụng"]), [True])

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
        self.assertTrue(columns["Mã"]["pinned"] and columns["Áp dụng"]["pinned"])
        self.assertGreater(app.dataframe[0].proto.row_height, 32)
        self.assertEqual(list(self.table(app).columns), ["Áp dụng", "Mã", "Trạng thái", "Nội dung gốc", "Bản dịch AI (sửa được)", "Ghi chú"])
        self.assertFalse(app.selectbox)
        self.assertFalse(app.multiselect)

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
        table = self.table(app)
        self.assertFalse(table.iloc[8]["Áp dụng"])
        self.assertEqual(table.iloc[8]["Trạng thái"], "Track Changes có sẵn")
        self.assertIn("TRƯỚC khi nhập AI", table.iloc[8]["Ghi chú"])
        self.assertIn("Existing prior-report translation", table.iloc[8]["Ghi chú"])
        self.review_state(app, excluded=[review_id for review_id in table["Mã"] if review_id != selected])
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
        self.assertIn("version", app.session_state["ai_import_preview"])
        self.assertTrue(app.dataframe)

    def test_edit_translation_then_apply_again_from_the_original_report(self):
        import io
        app = self.app()
        app.button(key="ai_check_translations").click().run()
        review_id = self.table(app)["Mã"][0]
        self.confirm(app)
        app.button(key="ai_apply_translations").click().run()
        first = app.session_state["ai_import_result"]["data"]
        self.review_state(app, edits={review_id: "Edited by reviewer."})
        self.assertFalse(app.exception)
        self.assertEqual(self.table(app)["Bản dịch AI (sửa được)"][0], "Edited by reviewer.")
        self.assertEqual(self.table(app)["Trạng thái"][0], "✏️ Đã sửa · Sẵn sàng")
        self.assertFalse(app.get("download_button"))
        self.assertTrue(any("đã thay đổi sau lần áp dụng trước" in element.value for element in app.info))
        self.confirm(app)
        app.button(key="ai_apply_translations").click().run()
        self.assertFalse(app.exception)
        result = app.session_state["ai_import_result"]
        self.assertNotEqual(result["data"], first)
        self.assertEqual(result["edited_ids"], [review_id])
        output = Document(io.BytesIO(result["data"]))
        self.assertEqual(output.paragraphs[0]._p.xpath("./w:ins//w:t/text()"), ["Edited by reviewer."])
        self.assertEqual(app.session_state["processed_outputs"]["review"], self.report_data)

    def test_blank_edit_is_not_applied(self):
        app = self.app()
        app.button(key="ai_check_translations").click().run()
        review_id = self.table(app)["Mã"][0]
        self.review_state(app, edits={review_id: ""})
        self.assertFalse(app.exception)
        self.assertEqual(app.session_state["ai_import_preview"]["data"]["counts"], {"blank": 1})
        self.assertFalse(any(button.key == "ai_apply_translations" for button in app.button))


if __name__ == "__main__":
    unittest.main()
