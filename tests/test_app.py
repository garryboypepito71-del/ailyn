import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

from app_logic import FULL_DAY_RATES, calculate_labor_pay
from receipt_logic import parse_scanned_receipt
from report_security import escape_report_text, report_download_name
from word_puzzle import check_guess, puzzle_for_level
from word_puzzle import check_math_answer, math_problem
from financial_close import close_month, is_month_closed, month_summary
import storage


class AppRulesTests(unittest.TestCase):
    def test_known_payroll_rates_and_fractional_pay(self):
        self.assertEqual(FULL_DAY_RATES["Laborer"], 500.0)
        self.assertEqual(calculate_labor_pay(1.5, "Laborer"), (750.0, 500.0, 250.0))

    def test_receipt_parser_extracts_total_quantity_and_delivery(self):
        result = parse_scanned_receipt("Receipt\nCement\nQty: 2\nDelivery: 50\nTotal PHP 1,050.00")
        self.assertEqual(result["name"], "Cement")
        self.assertEqual(result["qty"], 2)
        self.assertEqual(result["price"], 525.0)
        self.assertEqual(result["delivery"], 50.0)

    def test_report_text_is_escaped(self):
        self.assertEqual(escape_report_text('<script>alert("x")</script>'), "&lt;script&gt;alert(&quot;x&quot;)&lt;/script&gt;")
        self.assertEqual(report_download_name("bad/title"), '"bad_title_Receipt.png"')

    def test_word_puzzle_levels_accept_normalized_answers(self):
        self.assertEqual(puzzle_for_level(1)["answer"], "BUDGET")
        self.assertTrue(check_guess(1, " bud get "))
        self.assertFalse(check_guess(1, "MATERIAL"))

    def test_puzzle_random_seed_and_math_levels(self):
        self.assertNotEqual(puzzle_for_level(1, 1)["scramble"], puzzle_for_level(1, 2)["scramble"])
        question, answer = math_problem(8, 44)
        self.assertTrue(question)
        self.assertTrue(check_math_answer(answer, answer))
        self.assertFalse(check_math_answer(answer, answer + 1))

    def test_monthly_close_captures_summary_and_prevents_reclose(self):
        state = {
            "records": [{"type": "material", "amount": 100, "date": "Sep 10, 2026"}],
            "labor_records": [{"net": 50, "date": "Sep 10, 2026"}],
            "payroll_expenses": [],
        }
        summary = month_summary(state, "2026-09")
        self.assertEqual(summary["spent"], 150)
        closes = close_month({}, "2026-09", summary, "2026-09-30T12:00:00+08:00")
        self.assertTrue(is_month_closed(closes, "2026-09"))
        with self.assertRaises(ValueError):
            close_month(closes, "2026-09", summary, "2026-09-30T12:01:00+08:00")


class StorageTests(unittest.TestCase):
    def test_complete_backup_round_trip_preserves_files_and_state(self):
        with tempfile.TemporaryDirectory() as directory:
            data_dir = Path(directory) / "data"
            db_file = data_dir / "ailyn_state.db"
            backup_dir = data_dir / "backups"
            with patch.object(storage, "DATA_DIR", data_dir), patch.object(storage, "DB_FILE", str(db_file)), patch.object(
                storage, "BACKUP_DIR", str(backup_dir)
            ), patch.object(storage, "PHOTO_DIR", data_dir / "scanner_photos"):
                storage.save_state({"budget": 1000, "records": [{"amount": 250}]})
                photo = storage.PHOTO_DIR / "receipt.jpg"
                photo.parent.mkdir(parents=True)
                photo.write_bytes(b"photo")
                backup = storage.create_backup()
                self.assertTrue(zipfile.is_zipfile(backup))
                photo.unlink()
                restored = storage.restore_backup(backup)
                self.assertEqual(restored["budget"], 1000)
                self.assertEqual(photo.read_bytes(), b"photo")

    def test_restore_rejects_unsafe_zip_paths(self):
        with tempfile.TemporaryDirectory() as directory:
            backup = Path(directory) / "unsafe.zip"
            with zipfile.ZipFile(backup, "w") as archive:
                archive.writestr("../outside.txt", "bad")
            with self.assertRaises(ValueError):
                storage.restore_backup(backup)

    def test_restore_rejects_zip_without_application_database(self):
        with tempfile.TemporaryDirectory() as directory:
            backup = Path(directory) / "missing-db.zip"
            with zipfile.ZipFile(backup, "w") as archive:
                archive.writestr("notes.txt", "not a database")
            with self.assertRaises(ValueError):
                storage.restore_backup(backup)


if __name__ == "__main__":
    unittest.main()