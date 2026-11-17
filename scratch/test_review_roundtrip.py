import io
from copy import deepcopy
from datetime import datetime, timezone
from pathlib import Path
import sys
import unittest
from unittest.mock import patch

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.shared import Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from review_export import build_ai_review_document
from review_roundtrip import apply_ai_translations, bind_review_manifest, preview_ai_translations


SOURCE = "Công ty có vốn điều lệ 1.234,56 đồng."
TRANSLATION = "本公司資本為 VND 1,234.56。"


def saved(document):
    buffer = io.BytesIO()
    document.save(buffer)
    return buffer.getvalue()


class ReviewRoundtripTests(unittest.TestCase):
    def pair(self, count=1, target="Ht", size=10, in_table=False, source=SOURCE, prior_revision_indices=()):
        report = Document()
        items = []
        for index in range(count):
            paragraph = report.add_table(rows=1, cols=1).cell(0, 0).paragraphs[0] if in_table else report.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            paragraph.paragraph_format.left_indent = Inches(0.2)
            run = paragraph.add_run(source)
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            run.font.size = Pt(size)
            run.bold = True
            run.italic = True
            comment = report.add_comment(run, "Reference suggestion", author="Translation Tool")
            items.append({"comment_id": str(comment.comment_id), "original": source,
                          "references": [{"source": "dictionary", "original": "Công ty có vốn điều lệ 1.000,00 đồng.",
                                          "translation": "Reference translation", "score": 0.9}],
                          "formatting": {"size": size, "bold": True, "italic": True}})
            if index in prior_revision_indices:
                from review_annotations import _add_revision
                _add_revision(paragraph, source, "Existing prior-report translation", "Translation Tool", 9000 + index * 2)
                items[-1]["references"].append({"source": "prior", "original": source,
                                                "translation": "Existing prior-report translation", "score": 1.0})
        bind_review_manifest(report, items, target)
        with patch("review_export.needs_ai_translation", return_value=True):
            worksheet = Document(build_ai_review_document(items, target))
        for table in worksheet.tables:
            table.cell(1, 4).text = TRANSLATION
        return report, worksheet, items

    def preview(self, report, worksheet):
        return preview_ai_translations(saved(report), saved(worksheet))

    def test_roundtrip_records_real_revisions_and_preserves_paragraph(self):
        for target in ("E", "Hs", "Ht"):
            with self.subTest(target=target):
                report, worksheet, _ = self.pair(target=target)
                source_bytes = saved(report)
                source_stream = io.BytesIO(source_bytes)
                original_properties = report.paragraphs[0]._p.pPr.xml
                output, summary = apply_ai_translations(source_stream, saved(worksheet))
                result = Document(output)
                self.assertEqual(summary["applied"], 1)
                self.assertEqual(source_stream.getvalue(), source_bytes)
                self.assertEqual(result.paragraphs[0]._p.pPr.xml, original_properties)
                paragraph = result.paragraphs[0]._p
                self.assertEqual("".join(paragraph.xpath("./w:del//w:delText/text()")), SOURCE)
                self.assertEqual("".join(paragraph.xpath("./w:ins//w:t/text()")), TRANSLATION)
                self.assertFalse(result.comments)
                self.assertFalse(paragraph.xpath(".//w:commentReference | .//w:commentRangeStart | .//w:commentRangeEnd"))
                self.assertEqual(result.settings.element.find(qn("w:trackRevisions")).get(qn("w:val")), "true")
                for run in paragraph.xpath("./w:ins/w:r"):
                    properties = run.find(qn("w:rPr"))
                    fonts = properties.find(qn("w:rFonts"))
                    self.assertEqual(fonts.get(qn("w:ascii")), "Times New Roman")
                    self.assertEqual(fonts.get(qn("w:hAnsi")), "Times New Roman")
                    self.assertEqual(fonts.get(qn("w:eastAsia")), "Times New Roman" if target == "E" else "DFKai-SB")
                    self.assertEqual(properties.find(qn("w:sz")).get(qn("w:val")), "20")
                    self.assertIsNotNone(properties.find(qn("w:b")))
                    self.assertIsNotNone(properties.find(qn("w:i")))
                    self.assertIsNone(properties.find(qn("w:highlight")))
                    self.assertIsNone(fonts.get(qn("w:asciiTheme")))
                    if target != "E":
                        self.assertEqual(properties.find(qn("w:lang")).get(qn("w:eastAsia")), "zh-CN" if target == "Hs" else "zh-TW")

    def test_table_font_size_is_not_forced_to_ten(self):
        for size in (7, 8.5, 12):
            with self.subTest(size=size):
                report, worksheet, _ = self.pair(size=size, in_table=True)
                output, _ = apply_ai_translations(saved(report), saved(worksheet))
                cell = Document(output).tables[0].cell(0, 0)
                self.assertTrue(cell._tc.xpath(".//w:ins"))
                self.assertTrue(all(value == str(size * 2).removesuffix(".0") for value in cell._tc.xpath(".//w:ins//w:sz/@w:val")))

    def test_selection_retains_other_comments_and_can_be_imported_again(self):
        report, worksheet, items = self.pair(count=3)
        worksheet.tables[2].cell(1, 4).text = ""
        unrelated = report.add_paragraph("Unrelated report text")
        report.add_comment(unrelated.runs[0], "Keep this comment", author="Reviewer")
        first_id = items[0]["review_id"]
        output, _ = apply_ai_translations(saved(report), saved(worksheet), [first_id])
        result = Document(output)
        self.assertEqual(len(result.comments), 3)
        preview = self.preview(result, worksheet)
        self.assertEqual(preview["counts"], {"applied": 1, "ready": 1, "blank": 1})
        output, _ = apply_ai_translations(saved(result), saved(worksheet))
        final = Document(output)
        self.assertEqual(len(final.comments), 2)
        self.assertEqual(len(final.element.xpath(".//w:ins")), 2)

    def test_ids_distinguish_identical_sources_and_reordered_tables(self):
        report, worksheet, items = self.pair(count=2)
        worksheet.tables[0].cell(1, 4).text = "First translation"
        worksheet.tables[1].cell(1, 4).text = "Second translation"
        first, second = [table._tbl for table in worksheet.tables]
        second.addnext(first)
        output, _ = apply_ai_translations(saved(report), saved(worksheet), [items[1]["review_id"]])
        result = Document(output)
        self.assertEqual(result.paragraphs[0].text, SOURCE)
        self.assertEqual(result.paragraphs[1]._p.xpath("./w:ins//w:t/text()"), ["Second translation"])

    def test_blank_missing_and_unchanged_rows_are_not_applied(self):
        report, worksheet, _ = self.pair(count=3)
        worksheet.tables[0].cell(1, 4).text = ""
        worksheet.tables[1].cell(1, 4).text = SOURCE
        table = worksheet.tables[2]._tbl
        table.getparent().remove(table)
        self.assertEqual(self.preview(report, worksheet)["counts"], {"blank": 1, "unchanged": 1, "missing": 1})
        with self.assertRaises(ValueError):
            apply_ai_translations(saved(report), saved(worksheet))

    def test_wrong_report_unknown_id_and_language_are_rejected(self):
        report, worksheet, _ = self.pair()
        other_report, _, _ = self.pair()
        with self.assertRaisesRegex(ValueError, "không thuộc báo cáo này"):
            self.preview(other_report, worksheet)
        identifier = worksheet.core_properties.identifier
        worksheet.core_properties.identifier = identifier.rsplit(":", 1)[0] + ":E"
        with self.assertRaisesRegex(ValueError, "ngôn ngữ dịch khác"):
            self.preview(report, worksheet)
        worksheet.core_properties.identifier = identifier
        worksheet.tables[0].cell(1, 0).text = "UNKNOWN-C001"
        with self.assertRaisesRegex(ValueError, "không thuộc báo cáo này"):
            self.preview(report, worksheet)

    def test_duplicate_ids_block_all_copies(self):
        report, worksheet, _ = self.pair(count=2)
        worksheet.tables[1].cell(1, 0).text = worksheet.tables[0].cell(1, 0).text
        preview = self.preview(report, worksheet)
        self.assertEqual(preview["counts"], {"blocked": 2, "missing": 1})
        self.assertTrue(all("Mã bị trùng" in row["reason"] for row in preview["rows"][:2]))

    def test_changed_source_or_reference_is_blocked(self):
        for column in (1, 2, 3):
            with self.subTest(column=column):
                report, worksheet, _ = self.pair()
                worksheet.tables[0].cell(1, column).text += " Modified by AI"
                self.assertEqual(self.preview(report, worksheet)["counts"], {"blocked": 1})

    def test_modified_report_and_missing_anchor_are_blocked(self):
        report, worksheet, _ = self.pair()
        report.paragraphs[0].runs[0].text += " Changed report"
        self.assertIn("Nội dung báo cáo đã thay đổi", self.preview(report, worksheet)["rows"][0]["reason"])
        report, worksheet, _ = self.pair()
        anchor = report.element.xpath(".//w:commentRangeEnd")[0]
        anchor.getparent().remove(anchor)
        self.assertIn("vị trí neo", self.preview(report, worksheet)["rows"][0]["reason"])

    def test_fields_objects_and_existing_revisions_are_blocked(self):
        for tag in ("fldChar", "instrText", "drawing", "object", "bookmarkStart", "ins", "del", "rPrChange"):
            with self.subTest(tag=tag):
                report, worksheet, _ = self.pair()
                node = OxmlElement(f"w:{tag}")
                if tag in ("fldChar", "instrText", "drawing", "object"):
                    report.paragraphs[0].runs[0]._r.append(node)
                elif tag == "rPrChange":
                    report.paragraphs[0].runs[0]._r.get_or_add_rPr().append(node)
                else:
                    report.paragraphs[0]._p.append(node)
                self.assertEqual(self.preview(report, worksheet)["counts"], {"blocked": 1})

    def test_overlapping_and_spanning_comments_are_blocked(self):
        report, worksheet, _ = self.pair()
        report.add_comment(report.paragraphs[0].runs[0], "Second comment", author="Reviewer")
        self.assertIn("chồng lên đoạn này", self.preview(report, worksheet)["rows"][0]["reason"])
        report, worksheet, _ = self.pair()
        first = report.add_paragraph("Before")
        last = report.add_paragraph("After")
        report.element.body.insert(0, first._p)
        report.add_comment([first.runs[0], last.runs[0]], "Spanning comment", author="Reviewer")
        self.assertIn("trải qua", self.preview(report, worksheet)["rows"][0]["reason"])

    def test_mixed_emphasis_requires_manual_review(self):
        report, worksheet, _ = self.pair()
        first = report.paragraphs[0].runs[0]
        first.text = SOURCE[:10]
        second = report.paragraphs[0].add_run(SOURCE[10:])
        second.bold = False
        self.assertIn("định dạng chữ hỗn hợp", self.preview(report, worksheet)["rows"][0]["reason"])

    def test_multi_run_deletions_and_line_breaks_survive_accept_reject(self):
        source = "Công ty\tcó vốn\n1.234,56 đồng."
        report, worksheet, _ = self.pair(source=source)
        first = report.paragraphs[0].runs[0]
        first.text = source[:7]
        second = report.paragraphs[0].add_run(source[7:])
        second._r.insert(0, deepcopy(first._r.rPr))
        translation = "\t本公司\tVND\n1,234.56。\n"
        worksheet.tables[0].cell(1, 4).text = translation
        output, _ = apply_ai_translations(saved(report), saved(worksheet))
        result = Document(output)
        accepted, rejected = deepcopy(result.paragraphs[0]._p), deepcopy(result.paragraphs[0]._p)
        for paragraph, keep, discard in ((accepted, "ins", "del"), (rejected, "del", "ins")):
            for node in list(paragraph.findall(qn(f"w:{discard}"))):
                paragraph.remove(node)
            for node in list(paragraph.findall(qn(f"w:{keep}"))):
                position = paragraph.index(node)
                for offset, child in enumerate(list(node)):
                    paragraph.insert(position + offset, child)
                paragraph.remove(node)
            for node in paragraph.iter(qn("w:delText")):
                node.tag = qn("w:t")
        from docx.text.paragraph import Paragraph
        self.assertEqual(Paragraph(parse_xml(accepted.xml), result._body).text, translation)
        self.assertEqual(Paragraph(parse_xml(rejected.xml), result._body).text, source)

    def test_worksheet_revisions_and_objects_are_blocked(self):
        for tag in ("ins", "del", "fldSimple", "object"):
            with self.subTest(tag=tag):
                report, worksheet, _ = self.pair()
                worksheet.tables[0].cell(1, 4).paragraphs[0]._p.append(OxmlElement(f"w:{tag}"))
                self.assertEqual(self.preview(report, worksheet)["counts"], {"blocked": 1})

    def test_legacy_reports_and_merged_worksheets_are_rejected(self):
        report, worksheet, _ = self.pair()
        with self.assertRaisesRegex(ValueError, "chưa có dữ liệu liên kết AI"):
            self.preview(Document(), worksheet)
        worksheet.tables[0].cell(1, 0).merge(worksheet.tables[0].cell(1, 1))
        with self.assertRaisesRegex(ValueError, "ô bị gộp"):
            self.preview(report, worksheet)

    def test_revision_ids_are_unique_and_timestamp_is_current(self):
        report, worksheet, _ = self.pair(count=2)
        header_revision = OxmlElement("w:ins")
        header_revision.set(qn("w:id"), "20000")
        report.sections[0].header.paragraphs[0]._p.append(header_revision)
        output, _ = apply_ai_translations(saved(report), saved(worksheet))
        revisions = Document(output).element.xpath(".//w:ins | .//w:del")
        self.assertEqual(len({node.get(qn("w:id")) for node in revisions}), 4)
        self.assertTrue(all(int(node.get(qn("w:id"))) > 20000 for node in revisions))
        for revision in revisions:
            self.assertEqual(revision.get(qn("w:author")), "Translation Tool AI")
            date = datetime.fromisoformat(revision.get(qn("w:date")).replace("Z", "+00:00"))
            self.assertLess(abs((datetime.now(timezone.utc) - date).total_seconds()), 60)

    def test_explicit_invalid_selection_never_produces_output(self):
        report, worksheet, _ = self.pair()
        for selection in ([], ["UNKNOWN"]):
            with self.subTest(selection=selection), self.assertRaises(ValueError):
                apply_ai_translations(saved(report), saved(worksheet), selection)

    def test_mapping_survives_word_custom_xml_part_rename(self):
        report, worksheet, _ = self.pair()
        for relation in report.part.rels.values():
            if relation.reltype == RT.CUSTOM_XML and "translationToolReview" in str(relation.target_part.partname):
                relation.target_part._partname = PackURI("/customXml/item2.xml")
        self.assertEqual(self.preview(report, worksheet)["counts"], {"ready": 1})
        output, _ = apply_ai_translations(saved(report), saved(worksheet))
        self.assertFalse(Document(output).comments)

    def test_renumbered_comments_still_map_to_the_correct_identical_paragraph(self):
        report, worksheet, items = self.pair(count=3)
        for index, table in enumerate(worksheet.tables, 1):
            table.cell(1, 4).text = f"Translation {index}"
        output, _ = apply_ai_translations(saved(report), saved(worksheet), [items[0]["review_id"]])
        report = Document(output)
        for tag in ("commentRangeStart", "commentRangeEnd", "commentReference"):
            for element in report.element.iter(qn(f"w:{tag}")):
                element.set(qn("w:id"), str(int(element.get(qn("w:id"))) - 1))
        for comment in report.comments:
            element = comment._element
            element.set(qn("w:id"), str(int(element.get(qn("w:id"))) - 1))
        self.assertEqual(self.preview(report, worksheet)["counts"], {"applied": 1, "ready": 2})
        output, _ = apply_ai_translations(saved(report), saved(worksheet), [items[2]["review_id"]])
        result = Document(output)
        self.assertEqual(result.paragraphs[1].text, SOURCE)
        self.assertEqual(result.paragraphs[2]._p.xpath("./w:ins//w:t/text()"), ["Translation 3"])
        self.assertEqual(len(result.comments), 1)

    def test_duplicated_bookmark_identity_and_edited_comments_are_blocked(self):
        report, worksheet, _ = self.pair()
        paragraph = report.paragraphs[0]._p
        report.element.body.append(deepcopy(paragraph))
        self.assertEqual(self.preview(report, worksheet)["counts"], {"blocked": 1})
        report, worksheet, _ = self.pair()
        report.comments.get(0).paragraphs[0].add_run(" Reviewer note")
        self.assertEqual(self.preview(report, worksheet)["counts"], {"blocked": 1})

    def test_c009_prior_revision_is_not_a_new_ai_import(self):
        report, worksheet, items = self.pair(count=10, prior_revision_indices=(8, 9))
        start, end = OxmlElement("w:bookmarkStart"), OxmlElement("w:bookmarkEnd")
        start.set(qn("w:id"), "20000")
        start.set(qn("w:name"), "ExistingUserBookmark")
        end.set(qn("w:id"), "20000")
        report.paragraphs[0]._p.insert(0, start)
        report.paragraphs[0]._p.append(end)
        preview = self.preview(report, worksheet)
        self.assertEqual(preview["counts"], {"blocked": 2, "ready": 8})
        c009 = items[8]["review_id"]
        row = next(row for row in preview["rows"] if row["review_id"] == c009)
        self.assertTrue(row["source_has_revisions"])
        self.assertIn("khớp 100% với báo cáo năm trước", row["reason"])
        self.assertIn("TRƯỚC khi nhập AI", row["reason"])
        self.assertEqual(row["source_current_text"], "Existing prior-report translation")
        selected_ids = [items[1]["review_id"], items[6]["review_id"]]
        original_paragraphs = [paragraph._p.xml for paragraph in report.paragraphs]
        output, summary = apply_ai_translations(saved(report), saved(worksheet), selected_ids)
        result = Document(output)
        self.assertEqual(summary["applied_ids"], sorted(selected_ids))
        self.assertIn(c009, summary["preserved_ids"])
        self.assertIn(c009, summary["preexisting_revision_ids"])
        for index, paragraph in enumerate(result.paragraphs):
            if items[index]["review_id"] not in selected_ids:
                self.assertEqual(paragraph._p.xml, original_paragraphs[index])
        self.assertEqual(len(result.comments), 8)
        self.assertEqual(len(result.element.xpath('.//w:ins[@w:author="Translation Tool AI"]')), 2)
        self.assertEqual(len(result.element.xpath('.//w:ins[@w:author="Translation Tool"]')), 2)
        with self.assertRaises(ValueError):
            apply_ai_translations(saved(report), saved(worksheet), [c009])


if __name__ == "__main__":
    unittest.main()
