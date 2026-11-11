import io
from pathlib import Path
import sys
import unittest

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml.ns import qn
from docx.shared import Pt

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from processor import process_financial_report
from review_annotations import add_comments_and_track_changes
from review_export import HEADERS, build_ai_review_document, get_ai_review_prompt


CURRENT = "Công ty có vốn điều lệ là 1.234,56 đồng tại ngày 31/12/2025."
DICTIONARY_SOURCE = "Công ty có vốn điều lệ là 1.000,00 đồng tại ngày 31/12/2024."
PRIOR_SOURCE = "Công ty có vốn điều lệ là 1.100,00 đồng tại ngày 31/12/2024."


class ReviewExportTests(unittest.TestCase):
    def make_document(self, original=CURRENT, highlighted=True):
        document = Document()
        paragraph = document.add_paragraph(original)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.runs[0]
        run.font.size = Pt(12)
        run.bold = True
        run.italic = True
        if highlighted:
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        return document, paragraph

    def collect(self, document, **kwargs):
        items = []
        counts = add_comments_and_track_changes(
            document, {DICTIONARY_SOURCE: "Dictionary reference translation"},
            py_dict={PRIOR_SOURCE: (PRIOR_SOURCE, "Prior reference translation")},
            review_items=items, **kwargs
        )
        return counts, items

    def test_collects_both_sources_from_one_comment(self):
        document, _ = self.make_document()
        counts, items = self.collect(document)
        self.assertEqual(counts, {"comments": 1, "revisions": 0})
        self.assertEqual(len(items), 1)
        self.assertEqual(items[0]["original"], CURRENT)
        self.assertEqual([reference["source"] for reference in items[0]["references"]], ["dictionary", "prior"])
        self.assertTrue(all(reference["score"] < 1 for reference in items[0]["references"]))

    def test_non_review_paragraphs_are_excluded(self):
        document, _ = self.make_document(highlighted=False)
        counts, items = self.collect(document)
        self.assertEqual(counts["comments"], 0)
        self.assertEqual(items, [])

    def test_exact_only_still_creates_revision_but_not_worksheet_item(self):
        document, _ = self.make_document()
        items = []
        counts = add_comments_and_track_changes(
            document, {}, py_dict={CURRENT: (CURRENT, "Exact prior translation")}, review_items=items
        )
        self.assertEqual(counts, {"comments": 1, "revisions": 1})
        self.assertEqual(items, [])
        self.assertIn("Exact prior translation", document._element.xml)

    def test_original_snapshot_not_processed_text_is_exported(self):
        document, paragraph = self.make_document("Nội dung đã được xử lý và định dạng.")
        _, items = self.collect(document, original_texts={paragraph._element: CURRENT})
        self.assertEqual(items[0]["original"], CURRENT)

    def test_landscape_four_columns_one_table_per_comment(self):
        document, _ = self.make_document()
        _, items = self.collect(document)
        output = build_ai_review_document(items + items)
        self.assertEqual(output.tell(), 0)
        worksheet = Document(output)
        self.assertEqual(worksheet.sections[0].orientation, WD_ORIENT.LANDSCAPE)
        self.assertGreater(worksheet.sections[0].page_width, worksheet.sections[0].page_height)
        self.assertEqual(len(worksheet.tables), 2)
        for table in worksheet.tables:
            self.assertEqual([cell.text for cell in table.rows[0].cells], list(HEADERS))
            self.assertEqual(table.cell(1, 0).text, CURRENT)
            self.assertEqual(table.cell(1, 3).text, "")
            self.assertIn(DICTIONARY_SOURCE, table.cell(1, 1).text)
            self.assertIn(PRIOR_SOURCE, table.cell(1, 1).text)
            self.assertIn("Dictionary reference translation", table.cell(1, 2).text)
            self.assertIn("Prior reference translation", table.cell(1, 2).text)
            self.assertIsNotNone(table.rows[0]._tr.trPr.find(qn("w:tblHeader")))
            self.assertIsNone(table.rows[1]._tr.trPr)
            self.assertEqual(sum(column.width.twips for column in table.columns),
                             worksheet.sections[0].page_width.twips - worksheet.sections[0].left_margin.twips - worksheet.sections[0].right_margin.twips)
        text = "\n".join(paragraph.text for paragraph in worksheet.paragraphs)
        self.assertEqual(text.count("PROMPT chung cho AI"), 1)
        self.assertNotIn("Mã rà soát", text)
        self.assertNotIn("Vị trí trong tài liệu", text)

    def test_english_result_cell_preserves_source_emphasis_and_size(self):
        document, _ = self.make_document()
        _, items = self.collect(document)
        worksheet = Document(build_ai_review_document(items, "E"))
        paragraph = worksheet.tables[0].cell(1, 3).paragraphs[0]
        run = paragraph.runs[0]
        self.assertEqual(run.font.name, "Times New Roman")
        self.assertEqual(run.font.size.pt, 12)
        self.assertTrue(run.bold)
        self.assertTrue(run.italic)
        self.assertEqual(paragraph.alignment, WD_ALIGN_PARAGRAPH.CENTER)
        self.assertIsNone(run.font.highlight_color)
        self.assertIsNotNone(paragraph._p.pPr.find(qn("w:rPr")))

    def test_chinese_result_cell_has_dual_fonts_and_correct_script(self):
        document, _ = self.make_document()
        _, items = self.collect(document)
        for target_col, language in (("Hs", "zh-CN"), ("Ht", "zh-TW")):
            with self.subTest(target_col=target_col):
                worksheet = Document(build_ai_review_document(items, target_col))
                run = worksheet.tables[0].cell(1, 3).paragraphs[0].runs[0]
                self.assertEqual(run.font.name, "Times New Roman")
                self.assertEqual(run.font.size.pt, 10)
                self.assertEqual(run._r.rPr.rFonts.get(qn("w:eastAsia")), "DFKai-SB")
                self.assertEqual(run._r.rPr.find(qn("w:lang")).get(qn("w:eastAsia")), language)
                prompt = " ".join(get_ai_review_prompt(target_col))
                self.assertIn("2025/12/31日", prompt)
                self.assertIn("Chỉ dùng chữ giản thể" if target_col == "Hs" else "Chỉ dùng chữ phồn thể", prompt)

    def test_inherited_source_formatting_is_preserved(self):
        document, paragraph = self.make_document()
        document.styles["Normal"].font.size = Pt(11.5)
        document.styles["Normal"].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        paragraph.alignment = None
        paragraph.runs[0].font.size = None
        paragraph.runs[0].bold = None
        paragraph.runs[0].style = "Strong"
        _, items = self.collect(document)
        worksheet = Document(build_ai_review_document(items))
        result = worksheet.tables[0].cell(1, 3).paragraphs[0]
        self.assertEqual(result.runs[0].font.size.pt, 11.5)
        self.assertTrue(result.runs[0].bold)
        self.assertEqual(result.alignment, WD_ALIGN_PARAGRAPH.RIGHT)

    def test_title_has_no_decorative_border(self):
        document, _ = self.make_document()
        _, items = self.collect(document)
        worksheet = Document(build_ai_review_document(items))
        self.assertIsNone(worksheet.styles["Title"]._element.pPr.find(qn("w:pBdr")))
        self.assertEqual(str(worksheet.paragraphs[0].runs[0].font.color.rgb), "000000")

    def test_near_exact_score_is_not_rounded_to_100_percent(self):
        item = {"original": CURRENT, "references": [{"source": "dictionary", "original": DICTIONARY_SOURCE,
                "translation": "Reference", "score": 0.9999999}]}
        worksheet = Document(build_ai_review_document([item]))
        self.assertIn("99.99%", worksheet.tables[0].cell(1, 1).text)
        self.assertNotIn("100%", worksheet.tables[0].cell(1, 1).text)

    def test_exact_companion_reference_is_kept_with_fuzzy_reference(self):
        document, _ = self.make_document()
        items = []
        add_comments_and_track_changes(document, {DICTIONARY_SOURCE: "Fuzzy reference"},
                                      py_dict={CURRENT: (CURRENT, "Exact prior reference")}, review_items=items)
        worksheet = Document(build_ai_review_document(items))
        self.assertEqual(len(worksheet.tables), 1)
        self.assertIn("100%", worksheet.tables[0].cell(1, 1).text)
        self.assertIn("Exact prior reference", worksheet.tables[0].cell(1, 2).text)

    def test_empty_and_exact_only_inputs_do_not_generate_empty_document(self):
        self.assertIsNone(build_ai_review_document([]))
        self.assertIsNone(build_ai_review_document([{
            "original": CURRENT, "references": [{"source": "dictionary", "original": CURRENT,
            "translation": "Exact translation", "score": 1.0}]}]))

    def test_processing_pipeline_collects_items_without_changing_return_contract(self):
        document, _ = self.make_document()
        source = io.BytesIO()
        document.save(source)
        source.seek(0)
        settings = {name: False for name in (
            "metadata", "unicode", "clean_v", "para_template", "dictionary", "dual_font", "table_size",
            "table_7pt", "date_format", "textbox", "suggestion", "suggest_py", "number_swap", "signer_accents", "always_original"
        )}
        settings.update(highlight=True, review_annotations=True)
        items = []
        output, message = process_financial_report(source, translation_map={DICTIONARY_SOURCE: "Reference translation"},
                                                  process_settings=settings, review_items=items)
        self.assertIsNotNone(output, message)
        self.assertEqual(message, "Report processed.")
        self.assertEqual(len(items), 1)
        self.assertEqual(len(Document(output).comments), 1)


if __name__ == "__main__":
    unittest.main()
