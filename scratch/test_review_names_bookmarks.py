from pathlib import Path
import sys
import unittest

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from review_annotations import add_comments_and_track_changes
from review_export import get_ai_review_prompt
from review_roundtrip import apply_ai_translations
import test_review_processing as processing_fixtures
import test_review_roundtrip as roundtrip_fixtures
from test_review_processing import resolved
from test_review_roundtrip import saved


MARKER = "【自年月日至年月日止階段】"
SOURCE = "Công ty trình bày báo cáo cho giai đoạn " + MARKER + "."
TRANSLATION = "公司呈報財務報表 " + MARKER + "。"


def add_bookmark(paragraph, placement, identifier="700", name="UserBookmark"):
    start, end = OxmlElement("w:bookmarkStart"), OxmlElement("w:bookmarkEnd")
    start.set(qn("w:id"), identifier)
    start.set(qn("w:name"), name)
    end.set(qn("w:id"), identifier)
    if placement == "before":
        paragraph._p.insert(0, start)
        paragraph._p.insert(1, end)
    elif placement == "after":
        paragraph._p.append(start)
        paragraph._p.append(end)
    elif placement == "whole":
        paragraph._p.insert(0, start)
        paragraph._p.append(end)
    return start, end


class ReviewNameBookmarkTests(unittest.TestCase):
    def test_all_prompts_preserve_company_and_person_names_without_overmatching_generic_company(self):
        for language in ("E", "Hs", "Ht"):
            with self.subTest(language=language):
                prompt = " ".join(get_ai_review_prompt(language))
                for required in ("Công ty TNHH ABC", "Công ty Cổ phần XYZ", "Nguyễn Văn A", "Nguyen Van A", "không tự dịch, phiên âm",
                                 "masothue.com", "thuvienphapluat.vn", "tô nền vàng", "không dấu",
                                 "đúng tên chính thức", "Phân biệt tên pháp lý", "vẫn dịch theo ngữ cảnh",
                                 "vào cột cuối thay vì để trống"):
                    self.assertIn(required, prompt)

    def test_prompts_do_not_treat_template_markers_as_a_reason_to_skip(self):
        for language in ("E", "Hs", "Ht"):
            prompt = " ".join(get_ai_review_prompt(language))
            for required in (MARKER, "không phải lỗi", "giữ nguyên chính xác thẻ", "không tự đoán ngày tháng",
                             "Không mang thẻ chỉ có trong tham chiếu", "không được làm trống ô kết quả"):
                self.assertIn(required, prompt)

    def test_literal_template_marker_can_be_applied_without_a_warning(self):
        fixture = roundtrip_fixtures.ReviewRoundtripTests()
        report, worksheet, items = fixture.pair(source=SOURCE)
        worksheet.tables[0].cell(1, 4).text = TRANSLATION
        preview = fixture.preview(report, worksheet)
        self.assertEqual(preview["counts"], {"ready": 1})
        self.assertEqual(preview["rows"][0]["reason"], "")
        output, summary = apply_ai_translations(saved(report), saved(worksheet), [items[0]["review_id"]])
        self.assertEqual(summary["applied"], 1)
        result = Document(output)
        self.assertEqual(resolved(result).paragraphs[0].text, TRANSLATION)
        self.assertEqual(resolved(result, accept=False).paragraphs[0].text, SOURCE)

    def test_safe_word_bookmarks_are_preserved_during_ai_import_and_accept_reject(self):
        for placement in ("before", "after", "whole"):
            with self.subTest(placement=placement):
                fixture = roundtrip_fixtures.ReviewRoundtripTests()
                report, worksheet, _ = fixture.pair(source=SOURCE)
                worksheet.tables[0].cell(1, 4).text = TRANSLATION
                add_bookmark(report.paragraphs[0], placement, name="OLE_LINK3")
                self.assertEqual(fixture.preview(report, worksheet)["counts"], {"ready": 1})
                output, _ = apply_ai_translations(saved(report), saved(worksheet))
                result = Document(output)
                for accept in (True, False):
                    paragraph = resolved(result, accept=accept).paragraphs[0]
                    starts = paragraph._p.xpath("./w:bookmarkStart")
                    ends = paragraph._p.xpath("./w:bookmarkEnd")
                    self.assertEqual(len(starts), 1)
                    self.assertEqual(len(ends), 1)
                    self.assertEqual(starts[0].get(qn("w:name")), "OLE_LINK3")
                    self.assertEqual(starts[0].get(qn("w:id")), ends[0].get(qn("w:id")))
                    self.assertEqual(paragraph.text, TRANSLATION if accept else SOURCE)
                    first_run, last_run = paragraph.runs[0]._r, paragraph.runs[-1]._r
                    if placement in ("before", "whole"):
                        self.assertLess(paragraph._p.index(starts[0]), paragraph._p.index(first_run))
                    if placement == "before":
                        self.assertLess(paragraph._p.index(ends[0]), paragraph._p.index(first_run))
                    if placement in ("after", "whole"):
                        self.assertGreater(paragraph._p.index(ends[0]), paragraph._p.index(last_run))
                    if placement == "after":
                        self.assertGreater(paragraph._p.index(starts[0]), paragraph._p.index(last_run))

    def test_safe_bookmarks_do_not_block_exact_prior_automatic_revision(self):
        fixture = processing_fixtures.ReviewProcessingTests()
        for placement in ("before", "after", "whole"):
            with self.subTest(placement=placement):
                report, paragraph = fixture.report()
                source = paragraph.text
                add_bookmark(paragraph, placement)
                items = []
                counts = add_comments_and_track_changes(report, {}, {source: (source, TRANSLATION)}, review_items=items, target_col="Ht")
                self.assertEqual(counts, {"comments": 0, "revisions": 1})
                self.assertEqual(items, [])
                self.assertFalse(report.comments)
                self.assertEqual(resolved(report).paragraphs[0].text, TRANSLATION)
                self.assertEqual(resolved(report, accept=False).paragraphs[0].text, source)

    def test_malformed_and_spanning_bookmarks_remain_blocked(self):
        for kind in ("duplicate_id", "duplicate_name", "reversed", "spanning", "table_columns"):
            with self.subTest(kind=kind):
                fixture = roundtrip_fixtures.ReviewRoundtripTests()
                report, worksheet, _ = fixture.pair()
                paragraph = report.paragraphs[0]
                start, end = add_bookmark(paragraph, "whole")
                if kind == "duplicate_id":
                    add_bookmark(report.add_paragraph("Other"), "whole")
                elif kind == "duplicate_name":
                    add_bookmark(report.add_paragraph("Other"), "whole", identifier="701")
                elif kind == "reversed":
                    paragraph._p.remove(start)
                    paragraph._p.append(start)
                elif kind == "spanning":
                    paragraph._p.remove(end)
                    report.add_paragraph("Other")._p.append(end)
                else:
                    start.set(qn("w:colFirst"), "0")
                    start.set(qn("w:colLast"), "1")
                self.assertEqual(fixture.preview(report, worksheet)["counts"], {"blocked": 1})
                with self.assertRaises(ValueError):
                    apply_ai_translations(saved(report), saved(worksheet))

    def test_yellow_marked_company_name_stays_yellow_in_the_report(self):
        from docx.enum.text import WD_COLOR_INDEX
        name = "Công ty TNHH ABC"
        for target, marking in (("E", "highlight"), ("Ht", "highlight"), ("E", "shading")):
            with self.subTest(target=target, marking=marking):
                fixture = roundtrip_fixtures.ReviewRoundtripTests()
                report, worksheet, _ = fixture.pair(target=target, source=f"Giao dịch với {name} trong năm.")
                cell = worksheet.tables[0].cell(1, 4)
                paragraph = cell.paragraphs[0]
                for run in list(paragraph.runs):
                    paragraph._p.remove(run._r)
                prefix = "與" if target == "Ht" else "Transactions with "
                suffix = "之交易。" if target == "Ht" else " during the year."
                paragraph.add_run(prefix)
                marked = paragraph.add_run(name)
                if marking == "highlight":
                    marked.font.highlight_color = WD_COLOR_INDEX.YELLOW
                else:
                    shading = OxmlElement("w:shd")
                    shading.set(qn("w:val"), "clear")
                    shading.set(qn("w:fill"), "FFFF00")
                    marked._r.get_or_add_rPr().append(shading)
                paragraph.add_run(suffix)
                self.assertEqual(fixture.preview(report, worksheet)["counts"], {"ready": 1})
                output, _ = apply_ai_translations(saved(report), saved(worksheet))
                result = resolved(Document(output)).paragraphs[0]
                self.assertEqual(result.text, prefix + name + suffix)
                yellow = "".join(run.text for run in result.runs if run.font.highlight_color == WD_COLOR_INDEX.YELLOW)
                self.assertEqual(yellow, name)

    def test_a_name_copied_unchanged_is_not_forced_into_a_fake_translation(self):
        for name in ("Công ty TNHH ABC", "Nguyễn Văn A"):
            with self.subTest(name=name):
                fixture = roundtrip_fixtures.ReviewRoundtripTests()
                report, worksheet, _ = fixture.pair(source=name)
                worksheet.tables[0].cell(1, 4).text = name
                self.assertEqual(fixture.preview(report, worksheet)["counts"], {"unchanged": 1})


if __name__ == "__main__":
    unittest.main()
