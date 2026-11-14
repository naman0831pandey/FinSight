"""
Unit Tests for Personal Finance Excel Import & Insights Feature
"""

import os
import unittest
import tempfile
import openpyxl
from datetime import date
from db_manager import DatabaseManager
from tracker_oop import BudgetTracker, Transaction
import excel_import_service
from excel_import_service import ExcelImportError

class TestExcelImportService(unittest.TestCase):
    def setUp(self):
        # Create a temporary sqlite database
        self.temp_db_fd, self.temp_db_path = tempfile.mkstemp(suffix=".db")
        self.db = DatabaseManager(db_name=self.temp_db_path)
        self.tracker = BudgetTracker(db_manager=self.db)

        # Create a temporary directory for excel files
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        # Cleanup
        try:
            os.close(self.temp_db_fd)
            if os.path.exists(self.temp_db_path):
                os.remove(self.temp_db_path)
            for f in os.listdir(self.temp_dir):
                os.remove(os.path.join(self.temp_dir, f))
            os.rmdir(self.temp_dir)
        except Exception:
            pass

    def test_column_auto_detection(self):
        headers1 = ["Txn Date", "Transaction Type", "Debit Amount", "Category", "Remarks / Description"]
        mapping1 = excel_import_service.detect_column_mapping(headers1)
        self.assertEqual(mapping1["date"], "Txn Date")
        self.assertEqual(mapping1["type"], "Transaction Type")
        self.assertEqual(mapping1["amount"], "Debit Amount")
        self.assertEqual(mapping1["category"], "Category")
        self.assertEqual(mapping1["description"], "Remarks / Description")

        headers2 = ["Date", "Expense", "Food Category", "Notes"]
        mapping2 = excel_import_service.detect_column_mapping(headers2)
        self.assertEqual(mapping2["date"], "Date")
        self.assertEqual(mapping2["amount"], "Expense")
        self.assertEqual(mapping2["category"], "Food Category")
        self.assertEqual(mapping2["description"], "Notes")

    def test_parsers(self):
        # Date parser
        d1 = excel_import_service.parse_date_value("2026-11-10")
        self.assertEqual(d1, "2026-11-10")

        d2 = excel_import_service.parse_date_value("15/08/2026")
        self.assertEqual(d2, "2026-08-15")

        d3 = excel_import_service.parse_date_value("invalid-date")
        self.assertIsNone(d3)

        # Amount parser
        a1 = excel_import_service.parse_amount_value("₹ 1,500.50")
        self.assertEqual(a1, 1500.50)

        a2 = excel_import_service.parse_amount_value("-500")
        # abs value is taken if negative, or returns None if 0
        self.assertEqual(a2, 500.0)

        a3 = excel_import_service.parse_amount_value("zero")
        self.assertIsNone(a3)

        a4 = excel_import_service.parse_amount_value("0")
        self.assertIsNone(a4)

        # Type parser
        self.assertEqual(excel_import_service.parse_type_value("DR"), "Expense")
        self.assertEqual(excel_import_service.parse_type_value("CR"), "Income")
        self.assertEqual(excel_import_service.parse_type_value("credit"), "Income")
        self.assertEqual(excel_import_service.parse_type_value("Deposit"), "Income")
        self.assertEqual(excel_import_service.parse_type_value("Withdrawal"), "Expense")

    def test_excel_file_reading_and_validation(self):
        # Create an .xlsx workbook
        xlsx_path = os.path.join(self.temp_dir, "sample_statement.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Transactions"

        ws.append(["Date", "Type", "Amount", "Category", "Description"])
        ws.append(["2026-10-01", "Expense", 450.0, "Food", "Groceries at Supermarket"])
        ws.append(["2026-10-02", "Income", 50000.0, "Salary", "Monthly Paycheck"])
        ws.append(["2026-10-03", "Expense", "invalid_amt", "Shopping", "Clothing"]) # Error: invalid amount
        ws.append(["99-99-9999", "Expense", 1200.0, "Bills", "Electricity"]) # Error: invalid date
        ws.append(["2026-10-01", "Expense", 450.0, "Food", "Groceries at Supermarket"]) # In-file duplicate
        wb.save(xlsx_path)

        # Read sheet
        sheets = excel_import_service.get_sheet_names(xlsx_path)
        self.assertIn("Transactions", sheets)

        headers, rows = excel_import_service.read_excel_sheet(xlsx_path, "Transactions")
        self.assertEqual(len(headers), 5)
        self.assertEqual(len(rows), 5)

        mapping = {
            "date": "Date",
            "type": "Type",
            "amount": "Amount",
            "category": "Category",
            "description": "Description"
        }

        # Add an existing transaction into DB to test DB duplicate detection
        self.tracker.add_transaction(
            type_="Expense",
            category="Food",
            amount=450.0,
            date="2026-10-01",
            notes="Existing Manual Grocery",
            profile="Personal"
        )
        existing_txs = self.tracker.get_transactions()
        self.assertEqual(len(existing_txs), 1)
        self.assertFalse(existing_txs[0].is_imported)
        self.assertEqual(existing_txs[0].source, "Manual")

        # Validate rows
        results = excel_import_service.validate_and_prepare_rows(headers, rows, mapping, existing_txs)
        self.assertEqual(results["total_rows"], 5)
        self.assertEqual(results["error_count"], 2) # invalid_amt and invalid date
        self.assertEqual(results["duplicate_count"], 2) # 1 matches DB, 1 duplicate in-file
        self.assertEqual(results["valid_count"], 1) # Salary row

        # Execute Batch Import of valid rows
        batch = excel_import_service.execute_import_batch(
            db=self.db,
            valid_rows=results["valid_rows"],
            filename="sample_statement.xlsx",
            profile="Personal",
            skipped_count=results["error_count"] + results["duplicate_count"]
        )

        batch_id = batch["batch_id"]
        self.assertEqual(batch["imported_count"], 1)
        self.assertEqual(batch["total_amount"], 50000.0)

        # Check transactions in DB
        all_txs = self.tracker.get_transactions()
        self.assertEqual(len(all_txs), 2) # 1 manual + 1 imported

        imported_tx = next(t for t in all_txs if t.is_imported)
        manual_tx = next(t for t in all_txs if not t.is_imported)

        self.assertTrue(imported_tx.source.startswith("Imported"))
        self.assertTrue(imported_tx.is_imported)
        self.assertEqual(imported_tx.import_batch_id, batch_id)
        self.assertEqual(imported_tx.amount, 50000.0)
        self.assertEqual(manual_tx.source, "Manual")
        self.assertFalse(manual_tx.is_imported)
        self.assertIsNone(manual_tx.import_batch_id)

        # Check Excel imports table
        imports = self.tracker.get_excel_imports(profile="Personal")
        self.assertEqual(len(imports), 1)
        self.assertEqual(imports[0][0], batch_id)
        self.assertEqual(imports[0][1], "sample_statement.xlsx")

        # Check Scope calculation
        scope = excel_import_service.get_personal_insights_scope(self.db)
        self.assertEqual(scope["total_transactions"], 2)
        self.assertEqual(scope["manual_count"], 1)
        self.assertEqual(scope["imported_count"], 1)
        self.assertEqual(scope["active_batches"], 1)
        self.assertEqual(scope["earliest_date"], "2026-10-01")
        self.assertEqual(scope["latest_date"], "2026-10-02")

        # Test Rollback / Removal of Batch
        removed = excel_import_service.remove_import_batch(self.db, batch_id)
        self.assertEqual(removed, 1)

        # Verify manual record is strictly preserved and imported is gone
        remaining_txs = self.tracker.get_transactions()
        self.assertEqual(len(remaining_txs), 1)
        self.assertEqual(remaining_txs[0].source, "Manual")
        self.assertFalse(remaining_txs[0].is_imported)

        # Verify imports log is removed
        remaining_imports = self.tracker.get_excel_imports(profile="Personal")
        self.assertEqual(len(remaining_imports), 0)

        # Scope after removal
        scope_after = excel_import_service.get_personal_insights_scope(self.db)
        self.assertEqual(scope_after["total_transactions"], 1)
        self.assertEqual(scope_after["manual_count"], 1)
        self.assertEqual(scope_after["imported_count"], 0)
        self.assertEqual(scope_after["active_batches"], 0)

    def test_validation_result_keys_compatibility(self):
        headers = ["Date", "type", "amount", "category", "description"]
        rows = [
            ["2026-10-03 00:00:00", "UPI", 520, "Shopping", "T-shirt"],
            ["2026-10-04 00:00:00", "Debit Card", 270, "Entertainment", "Movie ticket"],
            ["invalid-date", "UPI", 360, "Food", "Dinner"],
        ]
        mapping = {"date": "Date", "type": "type", "amount": "amount", "category": "category", "description": "description"}
        results = excel_import_service.validate_and_prepare_rows(headers, rows, mapping, [])
        
        # Verify both key styles are present
        self.assertIn("invalid_rows", results)
        self.assertIn("error_rows", results)
        self.assertIn("valid_rows", results)
        self.assertIn("duplicate_rows", results)
        self.assertEqual(len(results["invalid_rows"]), 1)
        self.assertEqual(len(results["valid_rows"]), 2)

        # Verify row dict keys
        err_row = results["invalid_rows"][0]
        self.assertIn("row_num", err_row)
        self.assertIn("row_number", err_row)
        self.assertIn("errors", err_row)
        self.assertIn("error", err_row)

    def test_unsupported_format_rejection(self):
        # Non-existent or non-xlsx file
        bad_file = os.path.join(self.temp_dir, "legacy_file.xls")
        with open(bad_file, "w") as f:
            f.write("dummy legacy content")

        with self.assertRaises(ExcelImportError):
            excel_import_service.get_sheet_names(bad_file)


if __name__ == "__main__":
    unittest.main()
