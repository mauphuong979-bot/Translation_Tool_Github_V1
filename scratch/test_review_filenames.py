from pathlib import Path
import sys
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from review_filenames import review_output_filename


class ReviewFilenameTests(unittest.TestCase):
    def test_each_workflow_stage_has_a_short_recognizable_name(self):
        for kind, suffix in (("original", "_original"), ("clean", "_clean"), ("review", "_TC"), ("ai_result", "_AI_TC")):
            with self.subTest(kind=kind):
                self.assertEqual(review_output_filename("Bao_cao_Ht_111126_1737.docx", kind),
                                 f"Bao_cao_Ht_111126_1737{suffix}.docx")

    def test_ai_worksheet_uses_a_short_fixed_name(self):
        for filename in ("Bao_cao_Ht_111126_1737.docx", "annual_review_2025.docx", "_review.docx"):
            with self.subTest(filename=filename):
                self.assertEqual(review_output_filename(filename, "ai_review"), "AI_Form.docx")

    def test_legacy_workflow_suffixes_and_tool_timestamp_are_not_stacked(self):
        for suffix in ("_review", "_review_AI_reviewed", "_AI_translation_worksheet", "_TC", "_AI_Form", "_AI_TC", "_REVIEW_AI_REVIEWED"):
            with self.subTest(suffix=suffix):
                self.assertEqual(review_output_filename(f"Bao_cao_Ht_tool_111126 1737{suffix}.docx", "ai_result"),
                                 "Bao_cao_Ht_111126_1737_AI_TC.docx")

    def test_paths_and_uppercase_extensions_keep_only_the_original_report_name(self):
        for filename in ("C:\\Reports\\Báo cáo_review.DOCX", "/tmp/Báo cáo_TC.docx", "Báo cáo.docx"):
            with self.subTest(filename=filename):
                self.assertEqual(review_output_filename(filename, "ai_result"), "Báo cáo_AI_TC.docx")

    def test_business_names_are_not_truncated_or_rewritten(self):
        self.assertEqual(review_output_filename("Company_tool_111126_1737.docx", "review"),
                         "Company_tool_111126_1737_TC.docx")
        self.assertEqual(review_output_filename("annual_review_2025.docx", "ai_result"),
                         "annual_review_2025_AI_TC.docx")

    def test_empty_base_after_suffix_removal_has_a_safe_fallback(self):
        self.assertEqual(review_output_filename("_review.docx", "ai_result"), "Report_AI_TC.docx")


if __name__ == "__main__":
    unittest.main()
