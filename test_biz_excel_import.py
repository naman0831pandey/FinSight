"""
Unit Tests for Business ERP Excel Import & Insights Feature
"""

import os
import unittest
import tempfile
import openpyxl
from datetime import date
from db_manager import DatabaseManager
from tracker_oop import BudgetTracker
import excel_business_import_service as biz_service
from excel_business_import_service import (
    ENTITY_SALES,
    ENTITY_EXPENSES,
    ENTITY_INVENTORY,
    ExcelImportError
)

class TestBusinessExcelImport(unittest.TestCase):
    def setUp(self):
        self.temp_db_fd, self.temp_db_path = tempfile.mkstemp(suffix=".db")
        self.db = DatabaseManager(db_name=self.temp_db_path)
        self.tracker = BudgetTracker(db_manager=self.db)
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        try:
            os.close(self.temp_db_fd)
            if os.path.exists(self.temp_db_path):
                os.remove(self.temp_db_path)
            for f in os.listdir(self.temp_dir):
                os.remove(os.path.join(self.temp_dir, f))
            os.rmdir(self.temp_dir)
        except Exception:
            pass

    def test_column_auto_detection_per_entity(self):
        # 1. Sales headers
        sales_headers = ["Invoice Date", "Customer Name", "Invoice No.", "Sale Amount", "Payment Mode"]
        sales_map = biz_service.detect_biz_column_mapping(sales_headers, ENTITY_SALES)
        self.assertEqual(sales_map["date"], "Invoice Date")
        self.assertEqual(sales_map["customer"], "Customer Name")
        self.assertEqual(sales_map["invoice"], "Invoice No.")
        self.assertEqual(sales_map["amount"], "Sale Amount")
        self.assertEqual(sales_map["payment_method"], "Payment Mode")

        # 2. Expenses headers
        exp_headers = ["Date", "Expense Head", "Total Amount", "Vendor / Notes"]
        exp_map = biz_service.detect_biz_column_mapping(exp_headers, ENTITY_EXPENSES)
        self.assertEqual(exp_map["date"], "Date")
        self.assertEqual(exp_map["category"], "Expense Head")
        self.assertEqual(exp_map["amount"], "Total Amount")
        self.assertEqual(exp_map["notes"], "Vendor / Notes")

        # 3. Inventory headers
        inv_headers = ["Product Name", "In Stock Quantity", "Units Sold"]
        inv_map = biz_service.detect_biz_column_mapping(inv_headers, ENTITY_INVENTORY)
        self.assertEqual(inv_map["item_name"], "Product Name")
        self.assertEqual(inv_map["stock"], "In Stock Quantity")
        self.assertEqual(inv_map["sold"], "Units Sold")

    def test_sales_import_validation_and_batch_commit(self):
        # Add an existing manual sale to test duplicate detection
        self.db.add_business_sale(
            date="2026-11-01",
            customer_name="Acme Corp",
            invoice_no="INV-1001",
            amount=15000.0,
            payment_method="UPI"
        )
        existing_sales = self.db.get_all_business_sales(include_meta=True)
        self.assertEqual(len(existing_sales), 1)

        headers = ["Date", "Customer", "Invoice", "Amount", "Method"]
        rows = [
            ["2026-11-02", "Globex Inc", "INV-1002", 25000.0, "Card"], # Valid
            ["2026-11-03", "Wayne Enterprises", "INV-1003", 45000.0, "Net Banking"], # Valid
            ["invalid-date", "Stark Ind", "INV-1004", 5000.0, "UPI"], # Error: invalid date
            ["2026-11-04", "", "INV-1005", 2000.0, "Cash"], # Error: missing customer
            ["2026-11-05", "Cyberdyne", "INV-1006", -100.0, "UPI"], # Error: negative amount
            ["2026-11-01", "Acme Corp", "INV-1001", 15000.0, "UPI"], # Duplicate matching existing DB
            ["2026-11-02", "Globex Inc", "INV-1002", 25000.0, "Card"], # Duplicate in-file
        ]
        mapping = {"date": "Date", "customer": "Customer", "invoice": "Invoice", "amount": "Amount", "payment_method": "Method"}

        results = biz_service.validate_and_prepare_biz_rows(headers, rows, ENTITY_SALES, mapping, existing_sales)
        self.assertEqual(results["total_rows"], 7)
        self.assertEqual(results["valid_count"], 2)
        self.assertEqual(results["error_count"], 3)
        self.assertEqual(results["duplicate_count"], 2)

        # Commit valid batch
        batch = biz_service.execute_biz_import_batch(
            db=self.db,
            valid_rows=results["valid_rows"],
            entity_type=ENTITY_SALES,
            filename="sales_november.xlsx",
            skipped_count=results["error_count"] + results["duplicate_count"]
        )
        batch_id = batch["batch_id"]
        self.assertEqual(batch["imported_count"], 2)
        self.assertEqual(batch["total_volume"], 70000.0)

        # Verify DB records
        all_sales = self.db.get_all_business_sales(include_meta=True)
        self.assertEqual(len(all_sales), 3) # 1 manual + 2 imported

        imported_records = [s for s in all_sales if s[6] == batch_id]
        manual_records = [s for s in all_sales if s[6] is None]
        self.assertEqual(len(imported_records), 2)
        self.assertEqual(len(manual_records), 1)
        self.assertEqual(manual_records[0][7], "Manual")
        self.assertTrue(imported_records[0][7].startswith("Imported"))

        # Rollback batch
        deleted = biz_service.remove_biz_import_batch(self.db, batch_id)
        self.assertEqual(deleted, 2)

        remaining_sales = self.db.get_all_business_sales(include_meta=True)
        self.assertEqual(len(remaining_sales), 1)
        self.assertEqual(remaining_sales[0][2], "Acme Corp") # Manual record preserved!

    def test_expenses_import_validation_and_rollback(self):
        # Add existing manual expense
        self.db.add_business_expense("2026-11-01", "Office Rent", 20000.0, "Monthly lease")
        existing_exp = self.db.get_all_business_expenses(include_meta=True)

        headers = ["Date", "Category", "Amount", "Notes"]
        rows = [
            ["2026-11-02", "Electricity Bill", 4500.0, "Power supply"],
            ["2026-11-03", "Internet", 1500.0, "Fiber broadband"],
            ["2026-11-01", "Office Rent", 20000.0, "Monthly lease"], # Duplicate
        ]
        mapping = {"date": "Date", "category": "Category", "amount": "Amount", "notes": "Notes"}

        results = biz_service.validate_and_prepare_biz_rows(headers, rows, ENTITY_EXPENSES, mapping, existing_exp)
        self.assertEqual(results["valid_count"], 2)
        self.assertEqual(results["duplicate_count"], 1)

        batch = biz_service.execute_biz_import_batch(
            db=self.db,
            valid_rows=results["valid_rows"],
            entity_type=ENTITY_EXPENSES,
            filename="expenses_nov.xlsx"
        )
        batch_id = batch["batch_id"]

        all_exp = self.db.get_all_business_expenses(include_meta=True)
        self.assertEqual(len(all_exp), 3)

        # Rollback
        deleted = biz_service.remove_biz_import_batch(self.db, batch_id)
        self.assertEqual(deleted, 2)
        self.assertEqual(len(self.db.get_all_business_expenses()), 1)

    def test_inventory_import_validation_and_rollback(self):
        # Add existing product
        self.db.add_inventory_item("Gaming Mouse", 50, 10)

        headers = ["Item Name", "Stock Count", "Units Sold"]
        rows = [
            ["Mechanical Keyboard", 100, 5], # Valid
            ["Gaming Mouse", 20, 0], # Duplicate warning (item already exists)
            ["", 10, 0], # Error: missing name
            ["USB Cable", -5, 0], # Error: negative stock
            ["Wireless Headset", 5, 10], # Error: sold > stock
        ]
        mapping = {"item_name": "Item Name", "stock": "Stock Count", "sold": "Units Sold"}

        results = biz_service.validate_and_prepare_biz_rows(headers, rows, ENTITY_INVENTORY, mapping, self.db.get_all_inventory(include_meta=True))
        self.assertEqual(results["valid_count"], 1)
        self.assertEqual(results["duplicate_count"], 1) # Gaming Mouse duplicate warning
        self.assertEqual(results["error_count"], 3)

        batch = biz_service.execute_biz_import_batch(
            db=self.db,
            valid_rows=results["valid_rows"],
            entity_type=ENTITY_INVENTORY,
            filename="inventory_stock.xlsx"
        )
        batch_id = batch["batch_id"]

        items = self.db.get_all_inventory(include_meta=True)
        self.assertEqual(len(items), 2) # Gaming Mouse + Mechanical Keyboard

        # Rollback
        deleted = biz_service.remove_biz_import_batch(self.db, batch_id)
        self.assertEqual(deleted, 1)
        remaining = self.db.get_all_inventory()
        self.assertEqual(len(remaining), 1)
        self.assertEqual(remaining[0][1], "Gaming Mouse")

    def test_business_scope_calculation(self):
        # Insert 1 manual sale and 1 imported sale
        self.db.add_business_sale("2026-11-01", "Client A", "INV-1", 1000.0, "UPI")
        biz_service.execute_biz_import_batch(
            db=self.db,
            valid_rows=[{"date": "2026-11-05", "customer_name": "Client B", "invoice_no": "INV-2", "amount": 2000.0, "payment_method": "Card"}],
            entity_type=ENTITY_SALES,
            filename="sales.xlsx"
        )

        scope = biz_service.get_business_insights_scope(self.db)
        self.assertEqual(scope["sales"]["total_count"], 2)
        self.assertEqual(scope["sales"]["manual_count"], 1)
        self.assertEqual(scope["sales"]["imported_count"], 1)
        self.assertEqual(scope["sales"]["total_amount"], 3000.0)
        self.assertEqual(scope["sales"]["imported_amount"], 2000.0)
        self.assertEqual(scope["active_batches"], 1)

    def test_unsupported_format_rejection(self):
        bad_file = os.path.join(self.temp_dir, "legacy_file.xls")
        with open(bad_file, "w") as f:
            f.write("legacy binary data")

        with self.assertRaises(ExcelImportError):
            biz_service.get_sheet_names(bad_file)


if __name__ == "__main__":
    unittest.main()
