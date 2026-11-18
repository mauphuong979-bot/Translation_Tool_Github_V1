"""Tests for the JSON (API) path of the AI review roundtrip: read_review_manifest, rows_from_ai_results,
preview_ai_results and apply_ai_results. Run from the scratch folder: python -m unittest test_review_api_results.py"""
import io
import re
import unittest
import zipfile

from docx import Document

import test_review_roundtrip as fixtures
from review_roundtrip import (_highlight_ranges, apply_ai_results, apply_ai_translations, preview_ai_results,
                              preview_ai_translations, read_review_manifest)

saved = fixtures.saved
TRANSLATION = fixtures.TRANSLATION


def normalized_parts(data):
    archive = zipfile.ZipFile(io.BytesIO(data))
    return {name: re.sub(rb'w:date="[^"]*"', b'w:date=""', archive.read(name)) for name in sorted(archive.namelist())}


class ApiResultsTests(unittest.TestCase):
    def setUp(self):
        self.builder = fixtures.ReviewRoundtripTests()

    def pair(self, **kwargs):
        return self.builder.pair(**kwargs)

    def results(self, items, translation=TRANSLATION, **extra):
        return [{"review_id": item["review_id"], "translation": translation, **extra} for item in items]

    def test_api_path_matches_worksheet_path_exactly(self):
        for target in ("E", "Hs", "Ht"):
            with self.subTest(target=target):
                report, worksheet, items = self.pair(count=2, target=target)
                report_bytes = saved(report)
                via_file, file_summary = apply_ai_translations(report_bytes, saved(worksheet))
                via_api, api_summary = apply_ai_results(report_bytes, self.results(items))
                self.assertEqual(normalized_parts(via_api.getvalue()), normalized_parts(via_file.getvalue()))
                self.assertEqual(api_summary["applied_ids"], file_summary["applied_ids"])
                self.assertEqual(preview_ai_results(report_bytes, self.results(items))["counts"],
                                 preview_ai_translations(report_bytes, saved(worksheet))["counts"])

    def test_manifest_items_match_ai_form_and_carry_only_ai_form_fields(self):
        report, worksheet, items = self.pair(count=3, prior_revision_indices=(1,))
        manifest = read_review_manifest(saved(report))
        expected_ids = [items[0]["review_id"], items[2]["review_id"]]
        self.assertEqual([item["review_id"] for item in manifest["items"]], expected_ids)
        for item in manifest["items"]:
            self.assertEqual(set(item), {"review_id", "original", "references"})
            for reference in item["references"]:
                self.assertEqual(set(reference), {"source", "score", "original", "translation"})
        self.assertEqual(manifest["target_col"], "Ht")

    def test_manifest_excludes_already_applied_items(self):
        report, _, items = self.pair(count=2)
        output, _ = apply_ai_results(saved(report), self.results(items[:1]), [items[0]["review_id"]])
        manifest = read_review_manifest(output.getvalue())
        self.assertEqual([item["review_id"] for item in manifest["items"]], [items[1]["review_id"]])

    def test_highlight_phrases_become_yellow_inserted_runs(self):
        report, _, items = self.pair()
        translation = "本公司 ABC Company Limited 資本"
        output, _ = apply_ai_results(saved(report), self.results(items, translation, highlight=["ABC Company Limited"]))
        paragraph = Document(output).paragraphs[0]._p
        yellow = "".join(text for run in paragraph.xpath("./w:ins/w:r") if run.xpath("./w:rPr/w:highlight[@w:val='yellow']")
                         for text in run.xpath("./w:t/text()"))
        self.assertEqual(yellow, "ABC Company Limited")
        self.assertEqual("".join(paragraph.xpath("./w:ins//w:t/text()")), translation)

    def test_blank_missing_duplicate_and_unknown_ids(self):
        report, _, items = self.pair(count=3)
        results = [{"review_id": items[0]["review_id"], "translation": ""},
                   {"review_id": items[1]["review_id"], "translation": "A"},
                   {"review_id": items[1]["review_id"], "translation": "B"}]
        self.assertEqual(preview_ai_results(saved(report), results)["counts"], {"blank": 1, "blocked": 2, "missing": 1})
        with self.assertRaises(ValueError):
            preview_ai_results(saved(report), [{"review_id": "UNKNOWN-C001", "translation": "x"}])
        with self.assertRaises(ValueError):
            apply_ai_results(saved(report), results)

    def test_malformed_results_are_rejected(self):
        report, _, items = self.pair()
        for bad in (["text"], [{"translation": "x"}], [{"review_id": items[0]["review_id"], "translation": 5}]):
            with self.subTest(bad=bad), self.assertRaises(ValueError):
                preview_ai_results(saved(report), bad)

    def test_highlight_ranges_merge_and_ignore_invalid(self):
        self.assertEqual(_highlight_ranges("ab ab abc", ["ab", "b a", "", None, "zz"]), [(0, 8)])
        self.assertEqual(_highlight_ranges("x ABC y ABC", ["ABC"]), [(2, 5), (8, 11)])


if __name__ == "__main__":
    unittest.main()
