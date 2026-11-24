"""UI tests for the "Dịch tự động bằng AI" source in review_import_ui, using a fake API client (no network, no cost).
Run from the scratch folder: python -m unittest test_review_import_api_ui.py"""
import io
import json
from pathlib import Path
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from streamlit.testing.v1 import AppTest
from docx import Document

import test_review_roundtrip as fixtures

SCRATCH = str(Path(__file__).resolve().parent)
ROOT = str(Path(__file__).resolve().parents[1])


class ReviewImportApiUITests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.folder = Path(self.directory.name)
        report, _, self.items = fixtures.ReviewRoundtripTests().pair(count=3, target="E")
        self.report_data = fixtures.saved(report)
        (self.folder / "report.docx").write_bytes(self.report_data)

    def app(self, configured=True, refuse_first_once=False, secrets=None, username="user", web=False, tick_all=False):
        source = f'''
from pathlib import Path
import sys
from unittest.mock import patch
import streamlit as st
sys.path[:0] = [{ROOT!r}, {SCRATCH!r}]
import review_import_ui
from ai_translate_api import AIConfig
from test_ai_translate_api import FakeClient, echo, lookup_reply, message
folder = Path({str(self.folder)!r})
marker = folder / "refused_once"
def respond(payload):
    if {refuse_first_once!r} and not marker.exists() and payload[0]["review_id"] == {self.items[0]["review_id"]!r}:
        marker.write_text("x")
        return message(stop_reason="refusal", text="")
    return echo(payload, lambda entry: "Charter capital VND 1,234.56.")
st.session_state.processed_outputs = {{"review": (folder / "report.docx").read_bytes()}}
st.session_state.processed_filename = "Report.docx"
st.session_state.username = {username!r}
def logged(*args):
    with open(folder / "log.txt", "a", encoding="utf-8") as handle:
        handle.write(" | ".join(map(str, args)) + "\\n")
secrets = {secrets!r}
config = AIConfig(api_key="test", batch_size=1, max_workers=1, web_search={web!r}) if {configured!r} else None
lookup = lambda request: lookup_reply({{"Công ty TNHH ABC Việt Nam": ("ABC Vietnam Co., Ltd", "https://masothue.com/abc")}})
import company_names_store
def ticked(rows, **kwargs):
    if rows and "Tên quốc tế" not in rows[0]:
        return rows
    return [dict(row, **{{"Lưu": True, "Tên quốc tế": row["Tên quốc tế"] or "Typed Name Ltd"}}) for row in rows]
patches = [patch("company_names_store.LOCAL_FILE", str(folder / "company_names.csv")),
           patch("usage_logger.get_gsheet_client", return_value=(None, "test: no Google Sheet")),
           patch("usage_logger.is_local_env", return_value=True),
           patch("review_import_ui.log_event", side_effect=logged),
           patch("ai_translate_api.make_client", return_value=FakeClient(respond, lookup))]
import io
class Upload(io.BytesIO):
    name = "AI_Form_API.docx"
def uploader(label, **kwargs):
    word = folder / "word.docx"
    return Upload(word.read_bytes()) if label.startswith("Tải lên AI_Form") and word.exists() else None
patches.append(patch("review_import_ui.st.file_uploader", side_effect=uploader))
if {tick_all!r}:
    patches.append(patch("review_import_ui.st.data_editor", side_effect=ticked))
if secrets is None:
    patches.append(patch("review_import_ui._api_config", return_value=config))
else:
    patches.append(patch("review_import_ui.st.secrets", secrets))
from contextlib import ExitStack
with ExitStack() as stack:
    for item in patches:
        stack.enter_context(item)
    company_names_store.clear_cache()
    review_import_ui.render_ai_import()
    st.session_state.api_config_seen = review_import_ui._api_config() is not None
'''
        return AppTest.from_string(source, default_timeout=20).run()

    def checkbox(self, app, text):
        return next(box for box in app.checkbox if text in box.label)

    def test_translate_review_and_apply(self):
        app = self.app()
        self.assertFalse(app.exception)
        self.assertEqual(app.radio[0].value, "🤖 Dịch tự động bằng AI")
        self.assertTrue(app.button(key="ai_translate_api").disabled)
        self.checkbox(app, "Tôi đồng ý gửi").check().run()
        app.button(key="ai_translate_api").click().run()
        self.assertFalse(app.exception)
        self.assertEqual(app.session_state["ai_import_preview"]["data"]["counts"], {"ready": 3})
        self.assertTrue(app.dataframe[0].value["Áp dụng"].all())
        self.checkbox(app, "Tôi đã rà soát").check().run()
        app.button(key="ai_apply_translations").click().run()
        self.assertFalse(app.exception)
        result = app.session_state["ai_import_result"]
        self.assertEqual((result["count"], result["name"]), (3, "Report_AI_TC.docx"))
        output = Document(io.BytesIO(result["data"]))
        self.assertFalse(output.comments)
        self.assertEqual(output.paragraphs[0]._p.xpath("./w:ins//w:t/text()"), ["Charter capital VND 1,234.56."])
        log = (self.folder / "log.txt").read_text(encoding="utf-8")
        self.assertEqual(log.count("AI API"), 1)
        self.assertIn("3 mục E", log)

    def payload_codes(self, app):
        """Turn on the JSON view of 'Xem dữ liệu sẽ gửi' and return the shown JSON blocks."""
        app.toggle(key="ai_api_payload_json").set_value(True).run()
        self.assertFalse(app.exception)
        return [json.loads(block.value) for block in app.code]

    def test_data_preview_shows_only_ai_form_content(self):
        app = self.app()
        table = next(element.value for element in app.markdown if "ai-payload" in element.value)
        for item in self.items:
            self.assertIn(item["review_id"], table)
        self.assertNotIn("<script", table)
        self.assertEqual(len(app.code), 0)
        shown = json.dumps(self.payload_codes(app)[0], ensure_ascii=False)
        self.assertEqual([entry["review_id"] for entry in json.loads(shown)], [item["review_id"] for item in self.items])
        for forbidden in ("Reference suggestion", "Translation Tool", "TT_AI_"):
            self.assertNotIn(forbidden, shown)

    def test_edited_result_is_applied_without_calling_the_api_again(self):
        app = self.app()
        self.checkbox(app, "Tôi đồng ý gửi").check().run()
        app.button(key="ai_translate_api").click().run()
        first = self.items[0]["review_id"]
        state = dict(app.session_state["ai_review_state"])
        app.session_state["ai_review_state"] = dict(state, edits={first: "Edited charter capital."}, version=state["version"] + 1)
        app.run()
        self.assertFalse(app.exception)
        self.assertFalse(any("json" in button.proto.label for button in app.get("download_button")))
        self.assertEqual(app.session_state["ai_import_preview"]["data"]["counts"], {"ready": 3})
        self.checkbox(app, "Tôi đã rà soát").check().run()
        app.button(key="ai_apply_translations").click().run()
        self.assertFalse(app.exception)
        output = Document(io.BytesIO(app.session_state["ai_import_result"]["data"]))
        self.assertEqual(output.paragraphs[0]._p.xpath("./w:ins//w:t/text()"), ["Edited charter capital."])
        self.assertEqual((self.folder / "log.txt").read_text(encoding="utf-8").count("AI API"), 1)

    def test_word_round_trip_with_track_changes(self):
        from docx.oxml import parse_xml
        app = self.app()
        self.checkbox(app, "Tôi đồng ý gửi").check().run()
        app.button(key="ai_translate_api").click().run()
        self.assertTrue(any(button.proto.label.startswith("📄 Tải AI_Form") for button in app.get("download_button")))
        worksheet = Document(io.BytesIO(app.session_state["ai_api_word_export"]["data"]))
        self.assertEqual(worksheet.tables[0].cell(1, 4).text, "Charter capital VND 1,234.56.")
        cell = worksheet.tables[1].cell(1, 4)
        namespace = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        cell.paragraphs[0]._p.append(parse_xml(f'<w:ins {namespace} w:id="990" w:author="Translator"><w:r><w:t xml:space="preserve"> Reviewed.</w:t></w:r></w:ins>'))
        worksheet.save(self.folder / "word.docx")
        app.run()
        self.assertFalse(app.exception)
        rows = {row["review_id"]: row for row in app.session_state["ai_import_preview"]["data"]["rows"]}
        second = self.items[1]["review_id"]
        self.assertEqual(rows[second]["translation"], "Charter capital VND 1,234.56. Reviewed.")
        self.assertTrue(rows[second]["word_edited"])
        self.assertFalse(rows[self.items[0]["review_id"]]["word_edited"])
        self.assertEqual(app.session_state["ai_import_preview"]["data"]["counts"], {"ready": 3})
        self.assertIn("📝 Sửa trong Word · Sẵn sàng", list(app.dataframe[0].value["Trạng thái"]))
        self.checkbox(app, "Tôi đã rà soát").check().run()
        app.button(key="ai_apply_translations").click().run()
        self.assertFalse(app.exception)
        output = Document(io.BytesIO(app.session_state["ai_import_result"]["data"]))
        self.assertEqual(output.paragraphs[1]._p.xpath("./w:ins//w:t/text()"), ["Charter capital VND 1,234.56. Reviewed."])
        self.assertEqual((self.folder / "log.txt").read_text(encoding="utf-8").count("AI API"), 1)
        (self.folder / "word.docx").unlink()
        app.run()
        self.assertFalse(app.exception)
        self.assertNotIn("word_edited", app.session_state["ai_import_preview"]["data"]["rows"][0])

    def test_effective_results_include_edits(self):
        from review_import_ui import _effective_results
        results = [{"review_id": "A-C001", "translation": "One", "highlight": ["One"]}]
        effective = _effective_results(results, {"A-C001": "Uno", "A-C002": "Dos"})
        self.assertEqual(effective, [{"review_id": "A-C001", "translation": "Uno", "highlight": ["One"]},
                                     {"review_id": "A-C002", "translation": "Dos", "highlight": []}])

    def test_failed_batch_can_be_retried(self):
        app = self.app(refuse_first_once=True)
        self.checkbox(app, "Tôi đồng ý gửi").check().run()
        app.button(key="ai_translate_api").click().run()
        self.assertFalse(app.exception)
        self.assertTrue(any("từ chối" in element.value for element in app.error))
        self.assertEqual(app.session_state["ai_import_preview"]["data"]["counts"], {"ready": 2, "missing": 1})
        app.button(key="ai_api_retry").click().run()
        self.assertFalse(app.exception)
        self.assertEqual(app.session_state["ai_import_preview"]["data"]["counts"], {"ready": 3})
        self.assertEqual((self.folder / "log.txt").read_text(encoding="utf-8").count("AI API"), 2)

    def test_company_lookup_preview_and_results(self):
        report, _, self.items = fixtures.ReviewRoundtripTests().pair(target="E", source=["Công ty TNHH ABC Việt Nam", "Công ty TNHH Trong Câu có vốn điều lệ."])
        (self.folder / "report.docx").write_bytes(fixtures.saved(report))
        app = self.app(web=True)
        self.assertFalse(app.exception)
        self.assertEqual(self.payload_codes(app)[0], ["Công ty TNHH ABC Việt Nam"])
        self.assertTrue(any("chỉ gửi tên công ty" in element.value for element in app.caption))
        self.checkbox(app, "Tôi đồng ý gửi").check().run()
        app.button(key="ai_translate_api").click().run()
        self.assertFalse(app.exception)
        found = app.session_state["ai_api_run"]["outcome"]["company_names"]
        self.assertEqual(found["Công ty TNHH ABC Việt Nam"]["international"], "ABC Vietnam Co., Ltd")
        self.assertTrue(any("1 có tên quốc tế" in element.label for element in app.expander))
        self.assertIn("tra web 1 lần", (self.folder / "log.txt").read_text(encoding="utf-8"))

    def test_confirmed_names_are_saved_and_reused(self):
        source = ["Công ty TNHH ABC Việt Nam", "Công ty Cổ phần XYZ Miền Nam"]
        report, _, self.items = fixtures.ReviewRoundtripTests().pair(target="E", source=source)
        (self.folder / "report.docx").write_bytes(fixtures.saved(report))
        app = self.app(web=True, tick_all=True)
        self.assertEqual(self.payload_codes(app)[0], ["Công ty TNHH ABC Việt Nam", "Công ty Cổ phần XYZ Miền Nam"])
        self.checkbox(app, "Tôi đồng ý gửi").check().run()
        app.button(key="ai_translate_api").click().run()
        self.assertTrue(any("1 chưa tìm thấy" in element.label for element in app.expander))
        app.button(key="ai_company_names_save").click().run()
        self.assertFalse(app.exception)
        self.assertTrue(any("Đã lưu 2 tên công ty" in element.value for element in app.success))
        saved = (self.folder / "company_names.csv").read_text(encoding="utf-8-sig")
        self.assertIn('Công ty TNHH ABC Việt Nam,"ABC Vietnam Co., Ltd",https://masothue.com/abc,user,', saved)
        self.assertIn("Công ty Cổ phần XYZ Miền Nam,Typed Name Ltd,,user,", saved)
        self.assertIn("Company Names", (self.folder / "log.txt").read_text(encoding="utf-8"))
        again = self.app(web=True)
        self.assertTrue(any("Dùng 2 tên công ty đã xác nhận" in element.value for element in again.caption))
        payload = self.payload_codes(again)[-1]
        self.assertEqual(sorted(name["international"] for entry in payload for name in entry["company_names"]),
                         ["ABC Vietnam Co., Ltd", "Typed Name Ltd"])

    def test_saving_requires_an_international_name(self):
        source = ["Công ty TNHH ABC Việt Nam", "Công ty Cổ phần XYZ Miền Nam"]
        report, _, self.items = fixtures.ReviewRoundtripTests().pair(target="E", source=source)
        (self.folder / "report.docx").write_bytes(fixtures.saved(report))
        app = self.app(web=True)
        self.checkbox(app, "Tôi đồng ý gửi").check().run()
        app.button(key="ai_translate_api").click().run()
        self.assertTrue(app.button(key="ai_company_names_save").disabled)
        self.assertFalse((self.folder / "company_names.csv").exists())

    def test_hidden_without_config(self):
        app = self.app(configured=False)
        self.assertFalse(app.exception)
        self.assertFalse(app.radio)
        self.assertFalse(app.button)  # file flow waits for the AI_Form upload

    def test_secrets_and_enabled_users(self):
        allowed = {"anthropic": {"api_key": "k", "enabled_users": ["admin"]}}
        self.assertFalse(self.app(secrets=allowed, username="user").session_state["api_config_seen"])
        self.assertTrue(self.app(secrets=allowed, username="admin").session_state["api_config_seen"])
        self.assertTrue(self.app(secrets={"anthropic": {"api_key": "k"}}).session_state["api_config_seen"])
        self.assertFalse(self.app(secrets={}).session_state["api_config_seen"])


if __name__ == "__main__":
    unittest.main()
