import io
from pathlib import Path
import sys
import unittest
from unittest.mock import patch

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from processor import process_financial_report
from review_annotations import add_comments_and_track_changes
from review_export import build_ai_review_document, get_ai_review_prompt
from review_revisions import effective_highlight
from review_roundtrip import apply_ai_translations, bind_review_manifest, preview_ai_translations


SOURCE = "Nội dung mới chưa có bản dịch và cần bổ sung đầy đủ."
TRANSLATION = "本公司之新增資料為 VND 1,234.56。"


def saved(document):
    output = io.BytesIO()
    document.save(output)
    return output.getvalue()


def resolved(document, accept=True):
    result = Document(io.BytesIO(saved(document)))
    keep_tag, remove_tag = ("ins", "del") if accept else ("del", "ins")
    for element in list(result.element.iter(qn(f"w:{remove_tag}"))):
        element.getparent().remove(element)
    for element in list(result.element.iter(qn(f"w:{keep_tag}"))):
        parent = element.getparent()
        position = parent.index(element)
        for child in list(element):
            for text in child.iter(qn("w:delText")):
                text.tag = qn("w:t")
            parent.insert(position, child)
            position += 1
        parent.remove(element)
    return result


class ReviewProcessingTests(unittest.TestCase):
    def report(self, highlight="direct", size=10):
        report = Document()
        paragraph = report.add_paragraph(SOURCE)
        run = paragraph.runs[0]
        run.font.size = Pt(size)
        if highlight == "direct":
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        elif highlight == "character":
            style = report.styles.add_style("Yellow source", WD_STYLE_TYPE.CHARACTER)
            style.font.highlight_color = WD_COLOR_INDEX.YELLOW
            run.style = style
        elif highlight == "paragraph":
            report.styles["Normal"].font.highlight_color = WD_COLOR_INDEX.YELLOW
        elif highlight == "default":
            defaults = report.styles.element.find(qn("w:docDefaults"))
            properties = defaults.find(qn("w:rPrDefault")).find(qn("w:rPr"))
            yellow = OxmlElement("w:highlight")
            yellow.set(qn("w:val"), "yellow")
            properties.append(yellow)
        return report, paragraph

    def test_unmatched_yellow_creates_comment_table_and_roundtrip_id(self):
        report, _ = self.report()
        items = []
        counts = add_comments_and_track_changes(report, {}, review_items=items, target_col="Ht")
        self.assertEqual(counts, {"comments": 1, "revisions": 0})
        self.assertIn("[AI translation needed]", next(iter(report.comments)).text)
        self.assertEqual(items[0]["references"], [])
        bind_review_manifest(report, items, "Ht")
        worksheet = Document(build_ai_review_document(items, "Ht"))
        self.assertEqual(worksheet.tables[0].cell(1, 0).text, items[0]["review_id"])
        self.assertEqual(worksheet.tables[0].cell(1, 2).text, "")
        self.assertEqual(worksheet.tables[0].cell(1, 3).text, "")
        worksheet.tables[0].cell(1, 4).text = TRANSLATION
        self.assertEqual(preview_ai_translations(saved(report), saved(worksheet))["counts"], {"ready": 1})
        output, summary = apply_ai_translations(saved(report), saved(worksheet), [items[0]["review_id"]])
        result = Document(output)
        self.assertEqual(summary["applied_ids"], [items[0]["review_id"]])
        self.assertEqual(resolved(result).paragraphs[0].text, TRANSLATION)
        self.assertFalse(result.comments)

    def test_duplicate_unmatched_text_gets_distinct_ids(self):
        report, _ = self.report()
        report.add_paragraph(SOURCE).runs[0].font.highlight_color = WD_COLOR_INDEX.YELLOW
        items = []
        add_comments_and_track_changes(report, {}, review_items=items)
        bind_review_manifest(report, items, "E")
        self.assertEqual(len(items), 2)
        self.assertEqual(len({item["review_id"] for item in items}), 2)
        self.assertEqual(len(Document(build_ai_review_document(items)).tables), 2)

    def test_merged_table_cell_is_not_exported_twice(self):
        report = Document()
        cell = report.add_table(rows=1, cols=2).cell(0, 0)
        cell.merge(report.tables[0].cell(0, 1))
        cell.text = SOURCE
        cell.paragraphs[0].runs[0].font.highlight_color = WD_COLOR_INDEX.YELLOW
        items = []
        add_comments_and_track_changes(report, {}, review_items=items)
        self.assertEqual(len(items), 1)
        self.assertEqual(len(report.comments), 1)

    def test_unmatched_yellow_inside_nested_table_gets_a_linked_comment(self):
        report = Document()
        cell = report.add_table(rows=1, cols=1).cell(0, 0)
        paragraph = cell.add_table(rows=1, cols=1).cell(0, 0).paragraphs[0]
        paragraph.add_run(SOURCE).font.highlight_color = WD_COLOR_INDEX.YELLOW
        items = []
        add_comments_and_track_changes(report, {}, review_items=items)
        bind_review_manifest(report, items, "E")
        self.assertEqual(len(items), 1)
        worksheet = Document(build_ai_review_document(items))
        worksheet.tables[0].cell(1, 4).text = "New translation"
        output, summary = apply_ai_translations(saved(report), saved(worksheet))
        self.assertEqual(summary["applied"], 1)
        result = resolved(Document(output))
        self.assertEqual(result.tables[0].cell(0, 0).tables[0].cell(0, 0).text, "New translation")

    def test_accept_removes_direct_and_inherited_yellow_for_both_paths(self):
        for highlight in ("direct", "character", "paragraph", "default"):
            for automatic in (False, True):
                with self.subTest(highlight=highlight, automatic=automatic):
                    report, _ = self.report(highlight)
                    items = []
                    add_comments_and_track_changes(report, {}, py_dict={SOURCE: (SOURCE, TRANSLATION)} if automatic else None,
                                                  review_items=items, target_col="Ht")
                    if not automatic:
                        bind_review_manifest(report, items, "Ht")
                        worksheet = Document(build_ai_review_document(items, "Ht"))
                        worksheet.tables[0].cell(1, 4).text = TRANSLATION
                        output, _ = apply_ai_translations(saved(report), saved(worksheet))
                        report = Document(output)
                    accepted = resolved(report)
                    self.assertEqual(accepted.paragraphs[0].text, TRANSLATION)
                    for run in accepted.paragraphs[0].runs:
                        if run.text:
                            self.assertNotEqual(effective_highlight(run, accepted.paragraphs[0]), "yellow")
                    rejected = resolved(report, accept=False)
                    self.assertEqual(rejected.paragraphs[0].text, SOURCE)
                    self.assertEqual(effective_highlight(rejected.paragraphs[0].runs[0], rejected.paragraphs[0]), "yellow")

    def test_exact_prior_applies_fonts_size_without_a_redundant_comment(self):
        for target in ("E", "Hs", "Ht"):
            for size in (7, 10, 12):
                with self.subTest(target=target, size=size):
                    report, paragraph = self.report(size=size)
                    counts = add_comments_and_track_changes(report, {}, {SOURCE: (SOURCE, TRANSLATION)}, target_col=target)
                    self.assertEqual(counts["revisions"], 1)
                    for properties in paragraph._p.xpath("./w:ins/w:r/w:rPr"):
                        fonts = properties.find(qn("w:rFonts"))
                        self.assertEqual(fonts.get(qn("w:ascii")), "Times New Roman")
                        self.assertEqual(fonts.get(qn("w:eastAsia")), "Times New Roman" if target == "E" else "DFKai-SB")
                        self.assertEqual(properties.find(qn("w:sz")).get(qn("w:val")), str(size * 2))
                        self.assertIsNone(properties.find(qn("w:highlight")))
                    self.assertFalse(report.comments)
                    self.assertEqual(counts["comments"], 0)
                    self.assertFalse(paragraph._p.xpath(".//w:commentRangeStart | .//w:commentRangeEnd | .//w:commentReference"))
                    self.assertFalse(resolved(report).comments)

    def test_partial_bookmark_exact_match_is_left_for_manual_review(self):
        report, paragraph = self.report()
        paragraph.runs[0].text = SOURCE[:10]
        remainder = paragraph.add_run(SOURCE[10:])
        remainder.font.size = Pt(10)
        remainder.font.highlight_color = WD_COLOR_INDEX.YELLOW
        start, end = OxmlElement("w:bookmarkStart"), OxmlElement("w:bookmarkEnd")
        start.set(qn("w:id"), "3")
        start.set(qn("w:name"), "UserBookmark")
        end.set(qn("w:id"), "3")
        paragraph._p.insert(0, start)
        paragraph._p.insert(paragraph._p.index(remainder._r), end)
        items = []
        counts = add_comments_and_track_changes(report, {}, {SOURCE: (SOURCE, TRANSLATION)}, review_items=items)
        self.assertEqual(counts, {"comments": 1, "revisions": 0})
        self.assertEqual(paragraph.text, SOURCE)
        self.assertEqual(items, [])
        self.assertIn("[Manual review]", next(iter(report.comments)).text)

    def test_nearly_exact_prior_is_not_auto_applied(self):
        report, paragraph = self.report()
        items = []
        with patch("review_annotations._match", return_value=(SOURCE, (SOURCE, TRANSLATION), 0.9999999)):
            counts = add_comments_and_track_changes(report, {}, {SOURCE: (SOURCE, TRANSLATION)}, review_items=items)
        self.assertEqual(counts["revisions"], 0)
        self.assertEqual(paragraph.text, SOURCE)
        self.assertEqual(len(items), 1)
        self.assertIn("99.99%", next(iter(report.comments)).text)
        self.assertNotIn("100%", next(iter(report.comments)).text)

    def test_empty_dictionary_still_runs_review_pipeline(self):
        report, _ = self.report()
        settings = {name: False for name in ("metadata", "unicode", "clean_v", "para_template", "dictionary", "dual_font",
                    "table_size", "date_format", "textbox", "suggestion", "suggest_py", "number_swap", "signer_accents", "always_original")}
        settings.update(highlight=True, review_annotations=True)
        items = []
        output, message = process_financial_report(io.BytesIO(saved(report)), translation_map={}, process_settings=settings, review_items=items)
        self.assertIsNotNone(output, message)
        self.assertEqual(len(items), 1)
        self.assertIn("review_id", items[0])
        self.assertEqual(len(Document(output).comments), 1)

    def test_shared_prompt_requires_translation_without_references(self):
        for target in ("E", "Hs", "Ht"):
            prompt = " ".join(get_ai_review_prompt(target))
            self.assertIn("Nếu hai cột tham chiếu để trống", prompt)
            self.assertIn("không bỏ qua bảng đó", prompt)

    def test_mixed_report_exports_only_unresolved_items_and_preserves_auto_revision(self):
        report, _ = self.report()
        prior_source = "Các khoản đầu tư tài chính dài hạn đã khớp báo cáo kỳ trước."
        fuzzy_source = "Công ty có vốn điều lệ là 1.234,56 đồng tại ngày 31/12/2025."
        for text in (prior_source, fuzzy_source):
            report.add_paragraph(text).runs[0].font.highlight_color = WD_COLOR_INDEX.YELLOW
        items = []
        counts = add_comments_and_track_changes(report,
            {"Công ty có vốn điều lệ là 1.000,00 đồng tại ngày 31/12/2024.": "Dictionary reference"},
            {prior_source: (prior_source, "Exact prior translation")}, review_items=items)
        self.assertEqual(counts, {"comments": 2, "revisions": 1})
        self.assertEqual([item["original"] for item in items], [SOURCE, fuzzy_source])
        bind_review_manifest(report, items, "E")
        original_revision = report.paragraphs[1]._p.xml
        worksheet = Document(build_ai_review_document(items))
        self.assertEqual(len(worksheet.tables), 2)
        for table in worksheet.tables:
            table.cell(1, 4).text = "AI completed translation"
        preview = preview_ai_translations(saved(report), saved(worksheet))
        self.assertEqual(preview["counts"], {"ready": 2})
        output, summary = apply_ai_translations(saved(report), saved(worksheet), [items[0]["review_id"]])
        result = Document(output)
        self.assertEqual(summary["applied_ids"], [items[0]["review_id"]])
        self.assertEqual(result.paragraphs[1]._p.xml, original_revision)
        self.assertEqual(len(result.comments), 1)

    def test_prior_automatic_processing_never_removes_existing_reviewer_comments(self):
        for overlaps in (True, False):
            with self.subTest(overlaps=overlaps):
                report, paragraph = self.report()
                anchor = paragraph if overlaps else report.add_paragraph("Reviewer content")
                reviewer_comment = report.add_comment(anchor.runs[0], "Keep this note", author="Reviewer")
                original_comment_xml = reviewer_comment._element.xml
                items = []
                counts = add_comments_and_track_changes(report, {}, {SOURCE: (SOURCE, TRANSLATION)}, review_items=items)
                self.assertEqual(report.comments.get(reviewer_comment.comment_id)._element.xml, original_comment_xml)
                self.assertEqual(counts, {"comments": 1 if overlaps else 0, "revisions": 0 if overlaps else 1})
                self.assertEqual(len(report.comments), 2 if overlaps else 1)
                self.assertEqual(items, [])
                if overlaps:
                    self.assertEqual(paragraph.text, SOURCE)
                    self.assertIn("[Manual review]", list(report.comments)[-1].text)

    def test_empty_exact_prior_translation_keeps_a_manual_review_comment(self):
        report, paragraph = self.report()
        items = []
        counts = add_comments_and_track_changes(report, {}, {SOURCE: (SOURCE, "")}, review_items=items)
        self.assertEqual(counts, {"comments": 1, "revisions": 0})
        self.assertEqual(paragraph.text, SOURCE)
        self.assertIn("Bản dịch kỳ trước trống", next(iter(report.comments)).text)
        self.assertEqual(items, [])


if __name__ == "__main__":
    unittest.main()
