# -*- coding: utf-8 -*-
"""
excel_import_service.py
=======================
Local Excel (.xlsx) Import, Validation, and Insights Service for FINSIGHT.

Features:
- 100% Local processing without external API calls or network transfer.
- Reads .xlsx spreadsheets (multi-sheet support, header extraction, row parsing).
- Intelligent column mapping auto-detection (date, type, amount, category, description).
- Strict row-level validation (date parsing, positive amounts, type inference).
- Duplicate detection against database and within file.
- Batch import tracking with unique batch IDs and rollback capability.
- Personal Finance dataset scope and insights generation.
"""

import os
import sys
import re
import uuid
from datetime import datetime, date, timedelta
from typing import List, Dict, Any, Tuple, Optional

# Automatically ensure local .venv site-packages is in sys.path
_venv_site = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".venv", "Lib", "site-packages")
if os.path.exists(_venv_site) and _venv_site not in sys.path:
    sys.path.insert(0, _venv_site)

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


# Supported format is strictly modern .xlsx workbooks
SUPPORTED_EXTENSIONS = [".xlsx"]
UNSUPPORTED_FORMATS_DESC = (
    "Unsupported formats include: legacy binary Excel (.xls / Excel 97-2003), "
    "Excel binary (.xlsb), password-protected/encrypted workbooks, "
    "and non-Excel files (.csv, .numbers, .tsv)."
)


class ExcelImportError(Exception):
    """Custom exception for Excel import failures."""
    pass


def is_supported_file(filepath: str) -> bool:
    """Checks whether the file extension is supported."""
    ext = os.path.splitext(filepath)[1].lower()
    return ext in SUPPORTED_EXTENSIONS


def get_sheet_names(filepath: str) -> List[str]:
    """Retrieves all sheet names from an .xlsx workbook."""
    if not os.path.exists(filepath):
        raise ExcelImportError(f"File not found: {filepath}")
        
    ext = os.path.splitext(filepath)[1].lower()
    if ext not in SUPPORTED_EXTENSIONS:
        raise ExcelImportError(f"Unsupported spreadsheet format '{ext}'. Only .xlsx files are supported. {UNSUPPORTED_FORMATS_DESC}")

    if OPENPYXL_AVAILABLE:
        try:
            wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
            names = wb.sheetnames
            wb.close()
            return names
        except Exception as e:
            if "encrypted" in str(e).lower() or "password" in str(e).lower():
                raise ExcelImportError("This workbook is password-protected or encrypted. Please remove encryption and save as standard .xlsx.")
            raise ExcelImportError(f"Failed to read workbook sheets: {str(e)}")
    else:
        # Standard library zipfile fallback to read sheet names
        import zipfile
        import xml.etree.ElementTree as ET
        try:
            with zipfile.ZipFile(filepath, 'r') as z:
                wb_xml = z.read("xl/workbook.xml")
                root = ET.fromstring(wb_xml)
                ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                sheets = root.findall('.//main:sheet', ns)
                names = [s.get('name') for s in sheets if s.get('name')]
                return names or ["Sheet1"]
        except Exception as e:
            raise ExcelImportError(f"Unable to read Excel workbook: {str(e)}")


def read_excel_sheet(filepath: str, sheet_name: Optional[str] = None, max_preview_rows: Optional[int] = None) -> Tuple[List[str], List[List[Any]]]:
    """
    Reads headers and data rows from a specific sheet in an .xlsx workbook.
    Returns: (headers, rows)
    """
    if not os.path.exists(filepath):
        raise ExcelImportError(f"File not found: {filepath}")
        
    ext = os.path.splitext(filepath)[1].lower()
    if ext not in SUPPORTED_EXTENSIONS:
        raise ExcelImportError(f"Unsupported format '{ext}'. Only .xlsx is supported.")

    if not OPENPYXL_AVAILABLE:
        raise ExcelImportError("The 'openpyxl' library is required to process .xlsx spreadsheets.")

    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
    except Exception as e:
        if "password" in str(e).lower() or "encrypted" in str(e).lower():
            raise ExcelImportError("Encrypted or password-protected workbooks are not supported. Please unprotect the file.")
        raise ExcelImportError(f"Failed to open .xlsx file: {str(e)}")

    try:
        if sheet_name and sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.active

        raw_rows = []
        for r in ws.iter_rows(values_only=True):
            # Skip completely empty rows
            if any(cell is not None and str(cell).strip() != "" for cell in r):
                raw_rows.append(list(r))
    finally:
        wb.close()

    if not raw_rows:
        return [], []

    # First non-empty row is treated as the column header
    header_row = raw_rows[0]
    headers = [str(c).strip() if c is not None else f"Column_{idx+1}" for idx, c in enumerate(header_row)]

    data_rows = raw_rows[1:]
    if max_preview_rows is not None and max_preview_rows > 0:
        data_rows = data_rows[:max_preview_rows]

    return headers, data_rows


def detect_column_mapping(headers: List[str]) -> Dict[str, Optional[str]]:
    """
    Auto-detects mapping from spreadsheet headers to transaction fields:
    - date
    - type
    - amount
    - category
    - description (notes)
    """
    mapping: Dict[str, Optional[str]] = {
        "date": None,
        "type": None,
        "amount": None,
        "category": None,
        "description": None
    }

    # Normalize headers for matching
    norm_headers = [(h, re.sub(r'[^a-z0-9]', '', h.lower())) for h in headers]

    # 1. Date Detection
    date_keywords = ["date", "txndate", "transactiondate", "valuedate", "postingdate", "time", "transdate"]
    for orig, norm in norm_headers:
        if any(kw == norm or norm.startswith(kw) or norm.endswith(kw) for kw in date_keywords):
            mapping["date"] = orig
            break

    # 2. Type Detection
    type_keywords = ["type", "transactiontype", "txntype", "crdr", "drcr", "creditdebit", "debitcredit", "entrytype"]
    for orig, norm in norm_headers:
        if any(kw == norm or norm.startswith(kw) or norm.endswith(kw) for kw in type_keywords):
            mapping["type"] = orig
            break

    # 3. Amount Detection
    amount_keywords = ["amount", "amt", "txnamount", "transactionamount", "value", "price", "netamount", "total", "spent", "debit", "withdrawal", "expense", "cost"]
    for orig, norm in norm_headers:
        if any(kw == norm or norm.startswith(kw) or norm.endswith(kw) for kw in amount_keywords):
            mapping["amount"] = orig
            break

    # 4. Category Detection
    category_keywords = ["category", "cat", "expensecategory", "head", "tag", "subcategory", "group", "class"]
    for orig, norm in norm_headers:
        if any(kw == norm or norm.startswith(kw) or norm.endswith(kw) for kw in category_keywords):
            mapping["category"] = orig
            break

    # 5. Description / Notes Detection
    desc_keywords = ["description", "desc", "notes", "narration", "remarks", "particulars", "details", "payee", "merchant", "item", "title"]
    for orig, norm in norm_headers:
        if any(kw == norm or norm.startswith(kw) or norm.endswith(kw) for kw in desc_keywords):
            mapping["description"] = orig
            break

    return mapping


def parse_date_value(val: Any) -> Optional[str]:
    """
    Parses various date formats and Excel datetime types into standard 'YYYY-MM-DD'.
    """
    if val is None:
        return None

    # If already datetime or date object
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")

    val_str = str(val).strip()
    if not val_str:
        return None

    # Check if Excel serial integer (e.g. 45600)
    try:
        if val_str.isdigit() and len(val_str) in (4, 5):
            serial = int(val_str)
            # Excel base date is 1899-12-30 due to leap year bug
            base_date = datetime(1899, 12, 30)
            converted = base_date + timedelta(days=serial)
            return converted.strftime("%Y-%m-%d")
    except Exception:
        pass

    # Try common string date formats
    date_formats = [
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%d-%m-%Y",
        "%Y/%m/%d",
        "%d.%m.%Y",
        "%Y.%m.%d",
        "%d %b %Y",
        "%d %B %Y",
        "%b %d, %Y",
        "%B %d, %Y",
        "%Y%m%d",
        "%Y-%m-%d %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M:%S",
    ]

    for fmt in date_formats:
        try:
            parsed = datetime.strptime(val_str, fmt)
            return parsed.strftime("%Y-%m-%d")
        except ValueError:
            continue

    return None


def parse_amount_value(val: Any) -> Optional[float]:
    """
    Parses currency or numeric strings into positive float amounts.
    """
    if val is None:
        return None

    if isinstance(val, (int, float)):
        amt = abs(float(val))
        return amt if amt > 0 else None

    val_str = str(val).strip()
    if not val_str:
        return None

    # Clean out currency symbols and commas
    cleaned = re.sub(r'[₹$€£Rs,\s]', '', val_str, flags=re.IGNORECASE)
    
    # Handle accounting brackets: (1500.00) -> 1500.00
    if cleaned.startswith("(") and cleaned.endswith(")"):
        cleaned = cleaned[1:-1]
        
    try:
        amt = abs(float(cleaned))
        return amt if amt > 0 else None
    except ValueError:
        return None


def parse_type_value(val: Any, raw_amount_str: str = "", default_type: str = "Expense") -> str:
    """
    Determines whether a transaction is Income or Expense.
    """
    if val is not None:
        s = str(val).strip().lower()
        if any(w in s for w in ["inc", "income", "credit", "cr", "deposit", "dep", "received"]):
            return "Income"
        if any(w in s for w in ["exp", "expense", "debit", "dr", "payment", "pay", "withdrawal", "spent"]):
            return "Expense"

    # Check negative amount sign in raw amount
    if raw_amount_str and str(raw_amount_str).strip().startswith(("-", "(")):
        return "Expense"

    return default_type if default_type in ["Income", "Expense"] else "Expense"


def validate_and_prepare_rows(
    headers: List[str],
    rows: List[List[Any]],
    mapping: Dict[str, Optional[str]],
    existing_transactions: Optional[List[Dict[str, Any]]] = None,
    default_category: str = "Others",
    default_type: str = "Expense"
) -> Dict[str, Any]:
    """
    Validates spreadsheet rows against mapped fields, detects row-level errors
    and duplicate entries, and prepares clean records ready for import.
    """
    header_indices = {h: idx for idx, h in enumerate(headers)}

    date_col = mapping.get("date")
    type_col = mapping.get("type")
    amount_col = mapping.get("amount")
    cat_col = mapping.get("category")
    desc_col = mapping.get("description")

    valid_rows: List[Dict[str, Any]] = []
    error_rows: List[Dict[str, Any]] = []
    duplicate_rows: List[Dict[str, Any]] = []

    # Build existing transaction lookup for duplicate checking:
    # 1. Exact signature: (date, type, amount, category, notes)
    # 2. Core signature: (date, type, amount, category)
    db_exact_signatures = set()
    db_core_signatures = set()

    if existing_transactions:
        for t in existing_transactions:
            if isinstance(t, dict):
                date_k = str(t.get("date", ""))
                type_k = str(t.get("type", "")).strip().capitalize()
                amt_k = round(float(t.get("amount", 0)), 2)
                cat_k = str(t.get("category", "")).strip().lower()
                note_k = str(t.get("notes", "")).strip().lower()
            else:
                date_k = str(getattr(t, "date", ""))
                type_k = str(getattr(t, "type", "")).strip().capitalize()
                amt_k = round(float(getattr(t, "amount", 0)), 2)
                cat_k = str(getattr(t, "category", "")).strip().lower()
                note_k = str(getattr(t, "notes", "") or "").strip().lower()

            db_exact_signatures.add((date_k, type_k, amt_k, cat_k, note_k))
            db_core_signatures.add((date_k, type_k, amt_k, cat_k))

    # Also track duplicates within the spreadsheet file itself
    file_exact_signatures = set()
    file_core_signatures = set()

    for row_idx, r in enumerate(rows):
        excel_row_num = row_idx + 2  # 1-indexed, accounting for header row
        
        # Extract mapped values
        raw_date = r[header_indices[date_col]] if date_col in header_indices and header_indices[date_col] < len(r) else None
        raw_type = r[header_indices[type_col]] if type_col in header_indices and header_indices[type_col] < len(r) else None
        raw_amount = r[header_indices[amount_col]] if amount_col in header_indices and header_indices[amount_col] < len(r) else None
        raw_cat = r[header_indices[cat_col]] if cat_col in header_indices and header_indices[cat_col] < len(r) else None
        raw_desc = r[header_indices[desc_col]] if desc_col in header_indices and header_indices[desc_col] < len(r) else None

        row_errors = []

        # Validate Date
        parsed_date = parse_date_value(raw_date)
        if not parsed_date:
            row_errors.append(f"Invalid date '{raw_date}'. Must be valid date (e.g. YYYY-MM-DD or DD/MM/YYYY).")

        # Validate Amount
        parsed_amount = parse_amount_value(raw_amount)
        if parsed_amount is None:
            row_errors.append(f"Invalid amount '{raw_amount}'. Must be a positive number greater than 0.")

        # If any validation errors, record error row and skip
        if row_errors:
            error_rows.append({
                "row_number": excel_row_num,
                "row_num": excel_row_num,
                "error": " | ".join(row_errors),
                "errors": row_errors,
                "date": str(raw_date or "-"),
                "amount": str(raw_amount or "-"),
                "type": str(raw_type or "-"),
                "category": str(raw_cat or "-"),
                "description": str(raw_desc or "-"),
                "notes": str(raw_desc or "-"),
                "raw_values": {
                    "date": str(raw_date),
                    "amount": str(raw_amount),
                    "type": str(raw_type),
                    "category": str(raw_cat),
                    "description": str(raw_desc)
                }
            })
            continue

        # Validated fields
        txn_type = parse_type_value(raw_type, raw_amount_str=str(raw_amount), default_type=default_type)
        category = str(raw_cat).strip() if raw_cat is not None and str(raw_cat).strip() else default_category
        notes = str(raw_desc).strip() if raw_desc is not None else ""

        # Duplicate checking
        sig_exact = (parsed_date, txn_type, round(parsed_amount, 2), category.lower(), notes.lower())
        sig_core = (parsed_date, txn_type, round(parsed_amount, 2), category.lower())
        is_dup = False
        dup_reason = ""

        if sig_exact in db_exact_signatures:
            is_dup = True
            dup_reason = "Matches existing record in Personal Finance ledger (Exact match)"
        elif sig_core in db_core_signatures:
            is_dup = True
            dup_reason = f"Suspected duplicate: matches existing {txn_type} on {parsed_date} for ₹{parsed_amount:,.2f} ({category})"
        elif sig_exact in file_exact_signatures or sig_core in file_core_signatures:
            is_dup = True
            dup_reason = "Duplicate record within this spreadsheet file"

        file_exact_signatures.add(sig_exact)
        file_core_signatures.add(sig_core)

        record = {
            "row_number": excel_row_num,
            "row_num": excel_row_num,
            "date": parsed_date,
            "type": txn_type,
            "amount": parsed_amount,
            "category": category,
            "notes": notes,
            "description": notes,
            "is_duplicate": is_dup,
            "duplicate_reason": dup_reason,
            "warning": dup_reason
        }

        if is_dup:
            duplicate_rows.append(record)
        else:
            valid_rows.append(record)

    # Compute summary metrics
    all_dates = [r["date"] for r in valid_rows + duplicate_rows if "date" in r]
    date_min = min(all_dates) if all_dates else None
    date_max = max(all_dates) if all_dates else None

    total_valid_amt = sum(r["amount"] for r in valid_rows)
    valid_expenses = sum(r["amount"] for r in valid_rows if r["type"] == "Expense")
    valid_income = sum(r["amount"] for r in valid_rows if r["type"] == "Income")

    return {
        "total_rows": len(rows),
        "valid_count": len(valid_rows),
        "duplicate_count": len(duplicate_rows),
        "error_count": len(error_rows),
        "valid_rows": valid_rows,
        "duplicate_rows": duplicate_rows,
        "error_rows": error_rows,
        "invalid_rows": error_rows,  # Compatibility alias
        "date_min": date_min,
        "date_max": date_max,
        "total_amount": total_valid_amt,
        "total_expenses": valid_expenses,
        "total_income": valid_income
    }


def execute_import_batch(
    db_manager=None,
    rows_to_import: Optional[List[Dict[str, Any]]] = None,
    filename: str = "",
    profile: str = "Personal",
    total_processed: int = 0,
    skipped_count: int = 0,
    db=None,
    valid_rows=None
) -> Dict[str, Any]:
    """
    Executes atomic database import of validated transaction records,
    assigning a unique batch ID and marking the source.
    """
    active_db = db or db_manager
    if not active_db:
        raise ExcelImportError("Database manager instance is required.")

    records = valid_rows if valid_rows is not None else rows_to_import
    if not records:
        raise ExcelImportError("No valid rows to import.")

    timestamp_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    batch_suffix = uuid.uuid4().hex[:6]
    batch_id = f"IMP_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{batch_suffix}"
    source_label = f"Imported: {os.path.basename(filename)}"

    tuples_to_insert = []
    dates = []
    total_amt = 0.0

    for r in records:
        d = r["date"]
        t = r["type"]
        a = float(r["amount"])
        c = r["category"]
        n = r.get("notes") or r.get("description") or ""
        
        dates.append(d)
        total_amt += a
        tuples_to_insert.append((t, a, c, d, n, profile, batch_id, source_label))

    date_min = min(dates) if dates else timestamp_str[:10]
    date_max = max(dates) if dates else timestamp_str[:10]

    # Insert transactions and record batch
    active_db.add_transactions_batch(tuples_to_insert)
    active_db.record_excel_import(
        batch_id=batch_id,
        filename=os.path.basename(filename),
        import_date=timestamp_str,
        total_rows=total_processed or len(records),
        valid_rows=len(records),
        skipped_rows=skipped_count,
        date_min=date_min,
        date_max=date_max,
        total_amount=total_amt,
        profile=profile,
        notes=f"Successfully imported {len(records)} transactions from {os.path.basename(filename)}."
    )

    return {
        "batch_id": batch_id,
        "imported_count": len(records),
        "date_min": date_min,
        "date_max": date_max,
        "total_amount": total_amt,
        "import_date": timestamp_str
    }


def remove_import_batch(db_manager=None, batch_id: str = "", db=None) -> int:
    """
    Safely rolls back and removes an entire import batch from Personal Finance,
    preserving all existing manual transactions and other imports.
    """
    active_db = db or db_manager
    if not active_db:
        raise ExcelImportError("Database manager instance is required.")
    return active_db.remove_excel_import(batch_id)


def get_personal_insights_scope(db_manager=None, profile: str = "Personal", db=None) -> Dict[str, Any]:
    """
    Computes scope metrics for Personal Finance insights:
    distinguishing manual vs imported records, active import batches, and date range.
    """
    active_db = db or db_manager
    if not active_db:
        raise ExcelImportError("Database manager instance is required.")
    conn = active_db.get_connection()
    cursor = conn.cursor()

    try:
        # Total counts by source
        cursor.execute("""
            SELECT 
                COUNT(*) as total_count,
                SUM(CASE WHEN import_batch_id IS NOT NULL THEN 1 ELSE 0 END) as imported_count,
                SUM(CASE WHEN import_batch_id IS NULL THEN 1 ELSE 0 END) as manual_count,
                MIN(date) as earliest_date,
                MAX(date) as latest_date,
                SUM(CASE WHEN type = 'Expense' THEN amount ELSE 0 END) as total_expenses,
                SUM(CASE WHEN type = 'Income' THEN amount ELSE 0 END) as total_income
            FROM transactions
            WHERE profile = ?
        """, (profile,))
        row = cursor.fetchone()

        total_count = row[0] or 0
        imported_count = row[1] or 0
        manual_count = row[2] or 0
        earliest_date = row[3] or "N/A"
        latest_date = row[4] or "N/A"
        total_expenses = row[5] or 0.0
        total_income = row[6] or 0.0

        # Import batches
        cursor.execute("""
            SELECT batch_id, filename, import_date, valid_rows, date_min, date_max, total_amount
            FROM excel_imports
            WHERE profile = ?
            ORDER BY import_date DESC
        """, (profile,))
        batches = cursor.fetchall()

        # Category breakdown for imported transactions specifically
        cursor.execute("""
            SELECT category, SUM(amount) as cat_total, COUNT(*) as cat_count
            FROM transactions
            WHERE profile = ? AND import_batch_id IS NOT NULL AND type = 'Expense'
            GROUP BY category
            ORDER BY cat_total DESC
            LIMIT 5
        """, (profile,))
        top_imported_categories = cursor.fetchall()

        return {
            "total_count": total_count,
            "total_transactions": total_count,
            "imported_count": imported_count,
            "manual_count": manual_count,
            "earliest_date": earliest_date if earliest_date != "N/A" else None,
            "latest_date": latest_date if latest_date != "N/A" else None,
            "total_expenses": total_expenses,
            "total_income": total_income,
            "batches_count": len(batches),
            "active_batches": len(batches),
            "batches": batches,
            "top_imported_categories": top_imported_categories
        }
    finally:
        conn.close()
