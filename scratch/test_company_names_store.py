"""Tests for company_names_store (local CSV and Google Sheet paths, with fakes; no network).
Run from the scratch folder: python -m unittest test_company_names_store.py"""
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import gspread

import company_names_store as store

ABC = "Công ty TNHH Thương mại và Dịch vụ ABC"


class FakeWorksheet:
    def __init__(self):
        self.rows = []

    def append_row(self, row):
        self.rows.append(list(row))

    def append_rows(self, rows, value_input_option=None):
        self.rows.extend(list(row) for row in rows)

    def get_all_records(self):
        header, *body = self.rows
        return [dict(zip(header, row)) for row in body]


class FakeSpreadsheet:
    def __init__(self):
        self.sheets = {}

    def worksheet(self, title):
        if title not in self.sheets:
            raise gspread.WorksheetNotFound(title)
        return self.sheets[title]

    def add_worksheet(self, title, rows, cols):
        self.sheets[title] = FakeWorksheet()
        return self.sheets[title]


class FakeGClient:
    def __init__(self):
        self.spreadsheet = FakeSpreadsheet()

    def open_by_url(self, url):
        return self.spreadsheet


class LocalStoreTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.file = Path(directory.name) / "company_names.csv"
        for item in (patch.object(store, "LOCAL_FILE", str(self.file)), patch("usage_logger.is_local_env", return_value=True)):
            item.start()
            self.addCleanup(item.stop)
        store.clear_cache()
        self.addCleanup(store.clear_cache)

    def test_save_load_latest_wins_and_empty_removes(self):
        self.assertEqual(store.load_company_names(use_cache=False), {})
        saved = store.save_company_names([{"vietnamese": "  Công ty   TNHH   ABC ", "international": "ABC Co., Ltd", "source": "https://masothue.com/a"},
                                          {"vietnamese": ABC, "international": " ABC Trading  Services "},
                                          {"vietnamese": "Công ty TNHH Trống", "international": "  "}], "user1")
        self.assertEqual(saved, 2)
        names = store.load_company_names(use_cache=False)
        self.assertEqual(names["Công ty TNHH ABC"]["international"], "ABC Co., Ltd")
        self.assertEqual(names["Công ty TNHH ABC"]["saved_by"], "user1")
        self.assertEqual(names[ABC]["international"], "ABC Trading Services")
        store.save_company_names([{"vietnamese": "Công ty TNHH ABC", "international": "ABC Company Limited"}], "user2")
        self.assertEqual(store.company_name_map(use_cache=False)["Công ty TNHH ABC"], "ABC Company Limited")
        with open(self.file, "a", encoding="utf-8", newline="") as handle:
            handle.write("Công ty TNHH ABC,,,admin,2026-11-18\n")
        self.assertNotIn("Công ty TNHH ABC", store.load_company_names(use_cache=False))
        self.assertEqual(self.file.read_text(encoding="utf-8-sig").splitlines()[0], ",".join(store.HEADERS))

    def test_dictionary_rows_and_cache_cleared_on_save(self):
        self.assertEqual(store.company_name_map(), {})
        store.save_company_names([{"vietnamese": ABC, "international": "ABC Ltd"}], "user1")
        self.assertEqual(store.company_name_map(), {ABC: "ABC Ltd"})
        self.assertEqual(store.dictionary_rows(), [{"Vietnamese": ABC, "E": "ABC Ltd", "Hs": "ABC Ltd", "Ht": "ABC Ltd"}])
        self.assertEqual(store.save_company_names([], "user1"), 0)


class SheetStoreTests(unittest.TestCase):
    def setUp(self):
        self.client = FakeGClient()
        for item in (patch("usage_logger.is_local_env", return_value=False),
                     patch("usage_logger.get_gsheet_client", return_value=(self.client, "https://sheet"))):
            item.start()
            self.addCleanup(item.stop)
        store.clear_cache()
        self.addCleanup(store.clear_cache)

    def test_worksheet_is_created_and_rows_appended(self):
        self.assertEqual(store.load_company_names(use_cache=False), {})
        self.assertNotIn(store.SHEET_TITLE, self.client.spreadsheet.sheets)
        store.save_company_names([{"vietnamese": ABC, "international": "ABC Ltd", "source": "https://masothue.com/a"}], "user1")
        sheet = self.client.spreadsheet.sheets[store.SHEET_TITLE]
        self.assertEqual(sheet.rows[0], store.HEADERS)
        self.assertEqual(sheet.rows[1][:4], [ABC, "ABC Ltd", "https://masothue.com/a", "user1"])
        self.assertEqual(store.company_name_map(use_cache=False), {ABC: "ABC Ltd"})

    def test_unavailable_sheet_raises(self):
        with patch("usage_logger.get_gsheet_client", return_value=(None, "Secrets not found")):
            with self.assertRaises(RuntimeError):
                store.load_company_names(use_cache=False)


if __name__ == "__main__":
    unittest.main()
