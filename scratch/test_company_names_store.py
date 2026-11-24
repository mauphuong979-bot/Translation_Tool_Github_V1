"""Tests for company_names_store (Google Sheet and local CSV fallback, with fakes; no network).
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
OLD_HEADERS = ["Vietnamese", "International", "Source", "Saved by", "Saved at"]


class FakeWorksheet:
    id = 42

    def __init__(self, cols=9):
        self.rows, self.col_count, self.formatted = [], cols, False

    def update(self, a1, rows, value_input_option=None):
        assert a1 == "A1"
        if self.rows:
            self.rows[0] = list(rows[0])
        else:
            self.rows.append(list(rows[0]))

    def row_values(self, number):
        return list(self.rows[number - 1]) if len(self.rows) >= number else []

    def add_cols(self, count):
        self.col_count += count

    def append_rows(self, rows, value_input_option=None):
        self.rows.extend(list(row) for row in rows)

    def get_all_values(self):
        return [list(row) for row in self.rows]

    def freeze(self, rows=None):
        self.formatted = True

    def format(self, cells, style):
        pass


class FakeSpreadsheet:
    url = "https://docs.google.com/spreadsheets/d/x"

    def __init__(self):
        self.sheets, self.requests = {}, []

    def worksheet(self, title):
        if title not in self.sheets:
            raise gspread.WorksheetNotFound(title)
        return self.sheets[title]

    def add_worksheet(self, title, rows, cols):
        self.sheets[title] = FakeWorksheet(cols)
        return self.sheets[title]

    def batch_update(self, body):
        self.requests.extend(body["requests"])


class FakeGClient:
    def __init__(self):
        self.spreadsheet = FakeSpreadsheet()

    def open_by_url(self, url):
        return self.spreadsheet


class TempFile(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.folder = Path(directory.name)
        self.file = self.folder / "company_names.csv"
        item = patch.object(store, "LOCAL_FILE", str(self.file))
        item.start()
        self.addCleanup(item.stop)
        store.clear_cache()
        self.addCleanup(store.clear_cache)


class LocalFallbackTests(TempFile):
    def setUp(self):
        super().setUp()
        for item in (patch("usage_logger.is_local_env", return_value=True),
                     patch("usage_logger.get_gsheet_client", return_value=(None, "Secrets not found"))):
            item.start()
            self.addCleanup(item.stop)

    def test_save_load_latest_wins_and_empty_removes(self):
        self.assertEqual(store.load_company_names(use_cache=False), {})
        self.assertEqual(store.storage_status()["backend"], "csv")
        saved = store.save_company_names([{"vietnamese": "  Công ty   TNHH   ABC ", "international": "ABC Co., Ltd", "source": "https://masothue.com/a"},
                                          {"vietnamese": ABC, "international": " ABC Trading  Services "},
                                          {"vietnamese": "Công ty TNHH Trống", "international": "  "}], "user1")
        self.assertEqual(saved, 2)
        names = store.load_company_names(use_cache=False)
        self.assertEqual(names["Công ty TNHH ABC"]["international"], "ABC Co., Ltd")
        self.assertEqual((names["Công ty TNHH ABC"]["saved_by"], names["Công ty TNHH ABC"]["status"]), ("user1", store.STATUS_PENDING))
        self.assertEqual(names[ABC]["international"], "ABC Trading Services")
        with patch.object(store, "_timestamp", return_value="2099-01-01 00:00:00"):
            store.save_company_names([{"vietnamese": "Công ty TNHH ABC", "international": "ABC Company Limited"}], "user2")
        self.assertEqual(store.company_name_map(use_cache=False)["Công ty TNHH ABC"], "ABC Company Limited")
        with open(self.file, "a", encoding="utf-8", newline="") as handle:
            handle.write("Công ty TNHH ABC,,,admin,2099-01-02 00:00:00,,,,\n")
        self.assertNotIn("Công ty TNHH ABC", store.load_company_names(use_cache=False))
        self.assertEqual(self.file.read_text(encoding="utf-8-sig").splitlines()[0], ",".join(store.HEADERS))

    def test_old_five_column_file_is_upgraded(self):
        self.file.write_text(",".join(OLD_HEADERS) + "\n" + f"{ABC},ABC Ltd,,admin,2026-11-24 17:33:31\n", encoding="utf-8-sig")
        self.assertEqual(store.company_name_map(use_cache=False), {ABC: "ABC Ltd"})
        store.save_company_names([{"vietnamese": "Công ty TNHH XYZ", "international": "XYZ Ltd"}], "user1")
        self.assertEqual(self.file.read_text(encoding="utf-8-sig").splitlines()[0], ",".join(store.HEADERS))
        self.assertEqual(sorted(store.company_name_map(use_cache=False)), sorted([ABC, "Công ty TNHH XYZ"]))

    def test_dictionary_rows_and_cache_cleared_on_save(self):
        self.assertEqual(store.company_name_map(), {})
        store.save_company_names([{"vietnamese": ABC, "international": "ABC Ltd"}], "user1")
        self.assertEqual(store.company_name_map(), {ABC: "ABC Ltd"})
        self.assertEqual(store.dictionary_rows(), [{"Vietnamese": ABC, "E": "ABC Ltd", "Hs": "ABC Ltd", "Ht": "ABC Ltd"}])
        self.assertEqual(store.save_company_names([], "user1"), 0)


class SheetStoreTests(TempFile):
    def setUp(self):
        super().setUp()
        self.client = FakeGClient()
        for item in (patch("usage_logger.is_local_env", return_value=True),
                     patch("usage_logger.get_gsheet_client", return_value=(self.client, "https://sheet"))):
            item.start()
            self.addCleanup(item.stop)

    def sheet(self):
        return self.client.spreadsheet.sheets[store.SHEET_TITLE]

    def test_local_run_uses_the_sheet_when_configured(self):
        self.assertEqual(store.load_company_names(use_cache=False), {})
        self.assertNotIn(store.SHEET_TITLE, self.client.spreadsheet.sheets)
        store.save_company_names([{"vietnamese": ABC, "international": "ABC Ltd", "source": "https://masothue.com/a"}], "user1")
        self.assertFalse(self.file.exists())
        self.assertEqual(self.sheet().rows[0], store.HEADERS)
        self.assertEqual(self.sheet().rows[1][:4] + self.sheet().rows[1][5:6], [ABC, "ABC Ltd", "https://masothue.com/a", "user1", store.STATUS_PENDING])
        self.assertTrue(self.sheet().formatted)
        self.assertEqual(self.client.spreadsheet.requests[0]["setDataValidation"]["range"]["startColumnIndex"], 5)
        self.assertEqual(store.company_name_map(use_cache=False), {ABC: "ABC Ltd"})
        self.assertEqual(store.storage_status()["url"], "https://docs.google.com/spreadsheets/d/x#gid=42")

    def test_old_header_is_extended_and_review_columns_are_read(self):
        old = self.client.spreadsheet.add_worksheet(store.SHEET_TITLE, 1000, 5)
        old.rows = [list(OLD_HEADERS), [ABC, "ABC Ltd", "", "admin", "2026-11-24 10:00:00"]]
        self.assertEqual(store.load_company_names(use_cache=False)[ABC]["status"], store.STATUS_PENDING)
        store.save_company_names([{"vietnamese": "Công ty TNHH XYZ", "international": "XYZ Ltd"}], "user1")
        self.assertEqual((old.rows[0], old.col_count), (store.HEADERS, 9))
        old.rows[1] += [store.STATUS_APPROVED, "reviewer", "Đúng theo ĐKKD"]
        entry = store.load_company_names(use_cache=False)[ABC]
        self.assertEqual((entry["status"], entry["reviewed_by"], entry["note"]), (store.STATUS_APPROVED, "reviewer", "Đúng theo ĐKKD"))

    def test_rejected_names_are_not_used_and_sorting_the_sheet_does_not_matter(self):
        sheet = self.client.spreadsheet.add_worksheet(store.SHEET_TITLE, 1000, 9)
        sheet.rows = [list(store.HEADERS),
                      [ABC, "ABC New Ltd", "", "u", "2026-11-25 09:00:00", "", "", "", ""],
                      [ABC, "ABC Old Ltd", "", "u", "2026-11-24 09:00:00", "", "", "", ""],
                      ["Công ty TNHH Sai", "Wrong Ltd", "", "u", "2026-11-24 09:00:00", store.STATUS_REJECTED, "r", "", ""]]
        self.assertEqual(store.company_name_map(use_cache=False), {ABC: "ABC New Ltd"})

    def test_changed_header_is_refused(self):
        sheet = self.client.spreadsheet.add_worksheet(store.SHEET_TITLE, 1000, 9)
        sheet.rows = [["Tên", "Tên quốc tế"]]
        with self.assertRaisesRegex(ValueError, "tiêu đề"):
            store.save_company_names([{"vietnamese": ABC, "international": "ABC Ltd"}], "user1")

    def test_local_csv_is_uploaded_once_and_kept_as_backup(self):
        self.file.write_text(",".join(OLD_HEADERS) + "\n" + f"{ABC},ABC Ltd,,admin,2026-11-24 17:33:31\n", encoding="utf-8-sig")
        self.assertEqual(store.storage_status()["local_rows"], 1)
        self.assertEqual(store.upload_local_to_sheet(), 1)
        self.assertFalse(self.file.exists())
        self.assertEqual(len(list(self.folder.glob("company_names_uploaded_*.csv"))), 1)
        self.assertEqual(self.sheet().rows[1], [ABC, "ABC Ltd", "", "admin", "2026-11-24 17:33:31", store.STATUS_PENDING, "", "", ""])
        self.file.write_text(",".join(OLD_HEADERS) + "\n" + f"{ABC},ABC Ltd,,admin,2026-11-24 17:33:31\n", encoding="utf-8-sig")
        self.assertEqual(store.upload_local_to_sheet(), 0)
        self.assertEqual(len(self.sheet().rows), 2)

    def test_unavailable_sheet_raises_on_the_cloud(self):
        with patch("usage_logger.is_local_env", return_value=False), \
                patch("usage_logger.get_gsheet_client", return_value=(None, "Secrets not found")):
            with self.assertRaises(RuntimeError):
                store.load_company_names(use_cache=False)


if __name__ == "__main__":
    unittest.main()
