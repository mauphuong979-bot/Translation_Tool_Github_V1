"""Remember-me tokens and user changes against an in-memory spreadsheet. Run: python -m unittest scratch.test_remember_login"""
import os
import sys
import unittest
from unittest.mock import patch

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path[:0] = [ROOT, os.path.dirname(os.path.abspath(__file__))]

import session_store  # noqa: E402
import usage_logger  # noqa: E402
import user_store  # noqa: E402
from fake_gsheet import FakeSpreadsheet  # noqa: E402


class RememberLoginTest(unittest.TestCase):
    def setUp(self):
        self.sheet = FakeSpreadsheet()
        users = [{"username": "admin", "password": "a1", "role": "admin", "auto_fill": False},
                 {"username": "tester", "password": "t1", "role": "user", "auto_fill": False}]
        for item in (patch.object(usage_logger, "open_spreadsheet", return_value=self.sheet),
                     patch.object(user_store, "_local_users", return_value=users),
                     patch.object(user_store, "_sync_local_backup")):
            item.start()
            self.addCleanup(item.stop)
        user_store._cached_users.clear()

    def test_token_round_trip_and_hash_only(self):
        token = session_store.create("tester")
        self.assertEqual(session_store.username_for(token), "tester")
        self.assertIsNone(session_store.username_for("wrong"))
        self.assertNotIn(token, str(self.sheet.tabs["Sessions"].cells))
        session_store.revoke(token)
        self.assertIsNone(session_store.username_for(token))

    def test_expired_token_rejected_and_cleaned(self):
        old = session_store.create("tester")
        self.sheet.tabs["Sessions"].cells[1][2] = "2000-01-01 00:00:00"
        self.assertIsNone(session_store.username_for(old))
        session_store.create("admin")
        self.assertEqual([row[1] for row in self.sheet.tabs["Sessions"].cells[1:]], ["admin"])

    def test_password_changes_sign_out_everywhere(self):
        token = session_store.create("tester")
        self.assertFalse(user_store.change_password("tester", "bad", "t2")[0])
        self.assertEqual(session_store.username_for(token), "tester")
        self.assertTrue(user_store.change_password("tester", "t1", "t2")[0])
        self.assertIsNone(session_store.username_for(token))

        token = session_store.create("tester")
        self.assertTrue(user_store.update_user_data("tester", "tester", "t2", "admin", False)[0])
        self.assertEqual(session_store.username_for(token), "tester")  # role only: stays signed in
        self.assertTrue(user_store.update_user_data("tester", "tester", "t3", "user", False)[0])
        self.assertIsNone(session_store.username_for(token))

        token = session_store.create("tester")
        self.assertTrue(user_store.remove_user("tester")[0])
        self.assertIsNone(session_store.username_for(token))

    def test_revoke_user_keeps_others(self):
        mine, other = session_store.create("tester"), session_store.create("admin")
        session_store.revoke_user("tester")
        self.assertIsNone(session_store.username_for(mine))
        self.assertEqual(session_store.username_for(other), "admin")

    def test_sheet_down_means_no_auto_login(self):
        token = session_store.create("tester")
        with patch.object(usage_logger, "open_spreadsheet", side_effect=RuntimeError("offline")):
            self.assertIsNone(session_store.username_for(token))


if __name__ == "__main__":
    unittest.main()
