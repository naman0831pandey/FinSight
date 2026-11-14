"""
FINSIGHT - Business ERP Excel Import Service (excel_business_import_service.py)
-------------------------------------------------------------------------------
Provides 100% local, offline Excel (.xlsx) ingestion and validation engine
strictly scoped to the Business ERP portal:
- Multi-entity support: Sales, Expenses, Inventory
- Column auto-detection per entity type
- Row-level validation and error reporting
- Duplicate detection (in-file & database matching)
- Atomic batch commits with unique batch ID and source tagging
- Safe rollback / removal of imported batches
- Business insights and dataset scope metrics calculation
"""

import os
import re
import uuid
from datetime import datetime, date
from typing import List, Dict, Any, Optional, Tuple

import excel_import_service
from excel_import_service import (
    ExcelImportError,
    read_excel_sheet,
    get_sheet_names,
    parse_date_value,
    parse_amount_value,
    UNSUPPORTED_FORMATS_DESC,
    SUPPORTED_EXTENSIONS
)

ENTITY_SALES = "Sales"
ENTITY_EXPENSES = "Expenses"
ENTITY_INVENTORY = "Inventory"

VALID_ENTITIES = [ENTITY_SALES, ENTITY_EXPENSES, ENTITY_INVENTORY]

BIZ_EXPENSE_CATEGORIES = [
    "Employee Salary", "Electricity Bill", "Internet", "Office Rent",
    "Marketing", "Transportation", "Raw Material", "Miscellaneous"
]

BIZ_PAYMENT_METHODS = ["UPI", "Cash", "Card", "Net Banking"]


def get_entity_fields(entity_type: str) -> List[Tuple[str, str, bool, str]]:
    """
    Returns list of (display_label, field_key, is_required, description)
    for the selected business entity.
    """
    if entity_type == ENTITY_SALES:
        return [
            ("Date *", "date", True, "Sale date (YYYY-MM-DD or DD/MM/YYYY)"),
            ("Customer Name *", "customer", True, "Client or customer name"),
            ("Amount (₹) *", "amount", True, "Sale value / bill amount (> 0)"),
            ("Invoice No.", "invoice", False, "Invoice or bill reference number"),
            ("Payment Method", "payment_method", False, "UPI, Cash, Card, Net Banking")
        ]
    elif entity_type == ENTITY_EXPENSES:
        return [
            ("Date *", "date", True, "Expense date (YYYY-MM-DD or DD/MM/YYYY)"),
            ("Category", "category", False, "Business category (Rent, Salary, etc.)"),
            ("Amount (₹) *", "amount", True, "Expense amount (> 0)"),
            ("Notes / Payee", "notes", False, "Vendor, narration or details")
        ]
    elif entity_type == ENTITY_INVENTORY:
        return [
            ("Item Name *", "item_name", True, "Product name, model or SKU"),
            ("Stock Units *", "stock", True, "Initial or received stock count (>= 0)"),
            ("Units Sold", "sold", False, "Units sold (>= 0, default 0)")
        ]
    else:
        raise ExcelImportError(f"Unknown business entity type '{entity_type}'")


def detect_biz_column_mapping(headers: List[str], entity_type: str) -> Dict[str, Optional[str]]:
    """
    Auto-detects mapping between spreadsheet headers and the chosen entity fields.
    """
    mapping: Dict[str, Optional[str]] = {}
    norm_headers = [(h, re.sub(r'[^a-z0-9]', '', h.lower())) for h in headers]

    if entity_type == ENTITY_SALES:
        mapping = {"date": None, "customer": None, "amount": None, "invoice": None, "payment_method": None}

        date_kws = ["date", "saledate", "txndate", "invoicedate", "time"]
        cust_kws = ["customer", "customername", "client", "buyer", "party", "name", "billedto"]
        amt_kws = ["amount", "saleamount", "total", "netamount", "value", "price", "revenue"]
        inv_kws = ["invoice", "invoiceno", "invno", "billno", "bill", "receipt", "reference"]
        pay_kws = ["paymentmethod", "paymentmode", "method", "mode", "paymode", "type"]

        for orig, norm in norm_headers:
            if not mapping["date"] and any(kw == norm or norm.startswith(kw) or norm.endswith(kw) for kw in date_kws):
                mapping["date"] = orig
            elif not mapping["customer"] and any(kw == norm or norm.startswith(kw) or norm.endswith(kw) for kw in cust_kws):
                mapping["customer"] = orig
            elif not mapping["amount"] and any(kw == norm or norm.startswith(kw) or norm.endswith(kw) for kw in amt_kws):
                mapping["amount"] = orig
            elif not mapping["invoice"] and any(kw == norm or norm.startswith(kw) or norm.endswith(kw) for kw in inv_kws):
                mapping["invoice"] = orig
            elif not mapping["payment_method"] and any(kw == norm or norm.startswith(kw) or norm.endswith(kw) for kw in pay_kws):
                mapping["payment_method"] = orig

    elif entity_type == ENTITY_EXPENSES:
        mapping = {"date": None, "category": None, "amount": None, "notes": None}

        date_kws = ["date", "expdate", "expensedate", "time"]
        cat_kws = ["category", "expcategory", "head", "expensehead", "type", "group"]
        amt_kws = ["amount", "expenseamount", "total", "cost", "spent", "debit", "value"]
        note_kws = ["notes", "description", "desc", "remarks", "vendor", "payee", "particulars", "memo"]

        for orig, norm in norm_headers:
            if not mapping["date"] and any(kw == norm or norm.startswith(kw) or norm.endswith(kw) for kw in date_kws):
                mapping["date"] = orig
            elif not mapping["category"] and any(kw == norm or norm.startswith(kw) or norm.endswith(kw) for kw in cat_kws):
                mapping["category"] = orig
            elif not mapping["amount"] and any(kw == norm or norm.startswith(kw) or norm.endswith(kw) for kw in amt_kws):
                mapping["amount"] = orig
            elif not mapping["notes"] and any(kw == norm or norm.startswith(kw) or norm.endswith(kw) for kw in note_kws):
                mapping["notes"] = orig

    elif entity_type == ENTITY_INVENTORY:
        mapping = {"item_name": None, "stock": None, "sold": None}

        name_kws = ["item", "itemname", "product", "productname", "name", "sku", "model", "article", "description"]
        stock_kws = ["stock", "qty", "quantity", "units", "inventory", "stockcount", "instock", "count", "initialstock"]
        sold_kws = ["sold", "unitssold", "soldcount", "salescount", "outqty"]

        for orig, norm in norm_headers:
            if not mapping["item_name"] and any(kw == norm or norm.startswith(kw) or norm.endswith(kw) for kw in name_kws):
                mapping["item_name"] = orig
            elif not mapping["stock"] and any(kw == norm or norm.startswith(kw) or norm.endswith(kw) for kw in stock_kws):
                mapping["stock"] = orig
            elif not mapping["sold"] and any(kw == norm or norm.startswith(kw) or norm.endswith(kw) for kw in sold_kws):
                mapping["sold"] = orig

    return mapping


def parse_biz_amount_value(val: Any) -> Optional[float]:
    """
    Parses numeric strings or floats into strictly positive float amounts.
    Negative amounts and zero return None.
    """
    if val is None:
        return None
    if isinstance(val, (int, float)):
        if val <= 0:
            return None
        return float(val)

    clean = str(val).replace("₹", "").replace(",", "").replace(" ", "").strip()
    if clean.startswith(("-", "(")):
        return None
    try:
        f = float(clean)
        return f if f > 0 else None
    except Exception:
        return None


def parse_integer_value(val: Any, default_val: int = 0) -> Tuple[Optional[int], Optional[str]]:
    """
    Parses a non-negative integer for inventory counts.
    Returns: (parsed_int, error_message)
    """
    if val is None or str(val).strip() == "":
        return default_val, None
    try:
        clean = str(val).replace(",", "").replace(" ", "").strip()
        f_val = float(clean)
        if f_val < 0:
            return None, f"Value '{val}' cannot be negative"
        return int(round(f_val)), None
    except Exception:
        return None, f"Invalid numeric count '{val}'"


def normalize_payment_method(val: Any) -> str:
    """Normalizes payment method into standard UPI, Cash, Card, Net Banking."""
    if not val:
        return "UPI"
    s = str(val).strip().lower()
    if "cash" in s:
        return "Cash"
    if "card" in s or "debit" in s or "credit" in s:
        return "Card"
    if "net" in s or "bank" in s or "neft" in s or "rtgs" in s or "imps" in s:
        return "Net Banking"
    return "UPI"


def normalize_biz_category(val: Any) -> str:
    """Matches freeform category string against standard Business Expense categories."""
    if not val:
        return "Miscellaneous"
    s = str(val).strip().lower()
    for cat in BIZ_EXPENSE_CATEGORIES:
        if cat.lower() in s or s in cat.lower():
            return cat
    return str(val).strip().title()


def validate_and_prepare_biz_rows(
    headers: List[str],
    rows: List[List[Any]],
    entity_type: str,
    mapping: Dict[str, Optional[str]],
    existing_records: Optional[List[Any]] = None
) -> Dict[str, Any]:
    """
    Validates spreadsheet rows against the chosen Business ERP entity (Sales, Expenses, Inventory).
    Reports row-level errors and duplicate warnings.
    """
    header_indices = {h: idx for idx, h in enumerate(headers)}

    valid_rows: List[Dict[str, Any]] = []
    error_rows: List[Dict[str, Any]] = []
    duplicate_rows: List[Dict[str, Any]] = []

    # Build DB duplicate signatures based on entity
    db_signatures = set()
    db_secondary = set()

    if existing_records:
        for r in existing_records:
            if entity_type == ENTITY_SALES:
                # r: id, date, customer_name, invoice_no, amount, payment_method, ...
                dt = str(r[1] if len(r) > 1 else "")
                cust = str(r[2] if len(r) > 2 else "").strip().lower()
                inv = str(r[3] if len(r) > 3 and r[3] else "").strip().lower()
                amt = round(float(r[4] if len(r) > 4 else 0), 2)
                if inv:
                    db_signatures.add(inv)  # Invoice match
                db_secondary.add((dt, cust, amt))

            elif entity_type == ENTITY_EXPENSES:
                # r: id, date, category, amount, notes, ...
                dt = str(r[1] if len(r) > 1 else "")
                cat = str(r[2] if len(r) > 2 else "").strip().lower()
                amt = round(float(r[3] if len(r) > 3 else 0), 2)
                note = str(r[4] if len(r) > 4 and r[4] else "").strip().lower()
                db_signatures.add((dt, cat, amt, note))
                db_secondary.add((dt, cat, amt))

            elif entity_type == ENTITY_INVENTORY:
                # r: id, item_name, stock, sold, ...
                name = str(r[1] if len(r) > 1 else "").strip().lower()
                db_signatures.add(name)

    # In-file duplicate signatures
    file_signatures = set()
    file_secondary = set()

    for row_idx, r in enumerate(rows):
        excel_row_num = row_idx + 2  # Accounting for header row
        row_errors = []

        # =====================================================================
        # 1. SALES VALIDATION
        # =====================================================================
        if entity_type == ENTITY_SALES:
            col_date = mapping.get("date")
            col_cust = mapping.get("customer")
            col_amt = mapping.get("amount")
            col_inv = mapping.get("invoice")
            col_pay = mapping.get("payment_method")

            raw_date = r[header_indices[col_date]] if col_date in header_indices and header_indices[col_date] < len(r) else None
            raw_cust = r[header_indices[col_cust]] if col_cust in header_indices and header_indices[col_cust] < len(r) else None
            raw_amt = r[header_indices[col_amt]] if col_amt in header_indices and header_indices[col_amt] < len(r) else None
            raw_inv = r[header_indices[col_inv]] if col_inv in header_indices and header_indices[col_inv] < len(r) else None
            raw_pay = r[header_indices[col_pay]] if col_pay in header_indices and header_indices[col_pay] < len(r) else None

            # Validate Date
            p_date = parse_date_value(raw_date)
            if not p_date:
                row_errors.append(f"Invalid date '{raw_date}'")

            # Validate Customer
            p_cust = str(raw_cust).strip() if raw_cust is not None else ""
            if not p_cust:
                row_errors.append("Customer Name is required")

            # Validate Amount
            p_amt = parse_biz_amount_value(raw_amt)
            if p_amt is None:
                row_errors.append(f"Invalid amount '{raw_amt}'. Must be numeric and > 0")

            if row_errors:
                error_rows.append({
                    "row_number": excel_row_num,
                    "row_num": excel_row_num,
                    "error": " | ".join(row_errors),
                    "errors": row_errors,
                    "date": str(raw_date or "-"),
                    "primary_name": p_cust or str(raw_cust or "-"),
                    "amount": str(raw_amt or "-"),
                    "details": str(raw_inv or "-")
                })
                continue

            # Optional / defaults
            p_inv = str(raw_inv).strip() if raw_inv is not None and str(raw_inv).strip() else f"INV-AUTO-{excel_row_num}"
            p_pay = normalize_payment_method(raw_pay)

            # Duplicate check
            inv_lower = p_inv.lower()
            core_sig = (p_date, p_cust.lower(), round(p_amt, 2))
            is_dup = False
            dup_reason = ""

            if p_inv and not p_inv.startswith("INV-AUTO") and inv_lower in db_signatures:
                is_dup = True
                dup_reason = f"Invoice '{p_inv}' already exists in Business Sales ledger"
            elif core_sig in db_secondary:
                is_dup = True
                dup_reason = f"Suspected duplicate: existing sale for '{p_cust}' on {p_date} for ₹{p_amt:,.2f}"
            elif (p_inv and not p_inv.startswith("INV-AUTO") and inv_lower in file_signatures) or core_sig in file_secondary:
                is_dup = True
                dup_reason = f"Duplicate sale record within this spreadsheet file ({p_cust}, ₹{p_amt:,.2f})"

            if p_inv and not p_inv.startswith("INV-AUTO"):
                file_signatures.add(inv_lower)
            file_secondary.add(core_sig)

            record = {
                "row_number": excel_row_num,
                "row_num": excel_row_num,
                "date": p_date,
                "customer_name": p_cust,
                "primary_name": p_cust,
                "invoice_no": p_inv,
                "details": p_inv,
                "amount": p_amt,
                "payment_method": p_pay,
                "is_duplicate": is_dup,
                "duplicate_reason": dup_reason,
                "warning": dup_reason
            }

            if is_dup:
                duplicate_rows.append(record)
            else:
                valid_rows.append(record)

        # =====================================================================
        # 2. EXPENSES VALIDATION
        # =====================================================================
        elif entity_type == ENTITY_EXPENSES:
            col_date = mapping.get("date")
            col_cat = mapping.get("category")
            col_amt = mapping.get("amount")
            col_notes = mapping.get("notes")

            raw_date = r[header_indices[col_date]] if col_date in header_indices and header_indices[col_date] < len(r) else None
            raw_cat = r[header_indices[col_cat]] if col_cat in header_indices and header_indices[col_cat] < len(r) else None
            raw_amt = r[header_indices[col_amt]] if col_amt in header_indices and header_indices[col_amt] < len(r) else None
            raw_notes = r[header_indices[col_notes]] if col_notes in header_indices and header_indices[col_notes] < len(r) else None

            p_date = parse_date_value(raw_date)
            if not p_date:
                row_errors.append(f"Invalid date '{raw_date}'")

            p_amt = parse_biz_amount_value(raw_amt)
            if p_amt is None:
                row_errors.append(f"Invalid amount '{raw_amt}'. Must be numeric and > 0")

            if row_errors:
                error_rows.append({
                    "row_number": excel_row_num,
                    "row_num": excel_row_num,
                    "error": " | ".join(row_errors),
                    "errors": row_errors,
                    "date": str(raw_date or "-"),
                    "primary_name": str(raw_cat or "Miscellaneous"),
                    "amount": str(raw_amt or "-"),
                    "details": str(raw_notes or "-")
                })
                continue

            p_cat = normalize_biz_category(raw_cat)
            p_notes = str(raw_notes).strip() if raw_notes is not None else ""

            sig_exact = (p_date, p_cat.lower(), round(p_amt, 2), p_notes.lower())
            sig_core = (p_date, p_cat.lower(), round(p_amt, 2))
            is_dup = False
            dup_reason = ""

            if sig_exact in db_signatures:
                is_dup = True
                dup_reason = f"Matches existing business expense record ({p_cat}, ₹{p_amt:,.2f})"
            elif sig_core in db_secondary:
                is_dup = True
                dup_reason = f"Suspected duplicate: existing expense on {p_date} for ₹{p_amt:,.2f} ({p_cat})"
            elif sig_exact in file_signatures or sig_core in file_secondary:
                is_dup = True
                dup_reason = "Duplicate expense within this spreadsheet file"

            file_signatures.add(sig_exact)
            file_secondary.add(sig_core)

            record = {
                "row_number": excel_row_num,
                "row_num": excel_row_num,
                "date": p_date,
                "category": p_cat,
                "primary_name": p_cat,
                "amount": p_amt,
                "notes": p_notes,
                "details": p_notes or "-",
                "is_duplicate": is_dup,
                "duplicate_reason": dup_reason,
                "warning": dup_reason
            }

            if is_dup:
                duplicate_rows.append(record)
            else:
                valid_rows.append(record)

        # =====================================================================
        # 3. INVENTORY VALIDATION
        # =====================================================================
        elif entity_type == ENTITY_INVENTORY:
            col_name = mapping.get("item_name")
            col_stock = mapping.get("stock")
            col_sold = mapping.get("sold")

            raw_name = r[header_indices[col_name]] if col_name in header_indices and header_indices[col_name] < len(r) else None
            raw_stock = r[header_indices[col_stock]] if col_stock in header_indices and header_indices[col_stock] < len(r) else None
            raw_sold = r[header_indices[col_sold]] if col_sold in header_indices and header_indices[col_sold] < len(r) else None

            p_name = str(raw_name).strip() if raw_name is not None else ""
            if not p_name:
                row_errors.append("Product / Item Name is required")

            p_stock, err_stock = parse_integer_value(raw_stock, default_val=0)
            if err_stock:
                row_errors.append(err_stock)

            p_sold, err_sold = parse_integer_value(raw_sold, default_val=0)
            if err_sold:
                row_errors.append(err_sold)

            if p_stock is not None and p_sold is not None and p_sold > p_stock:
                row_errors.append(f"Sold units ({p_sold}) cannot exceed stock units ({p_stock})")

            if row_errors:
                error_rows.append({
                    "row_number": excel_row_num,
                    "row_num": excel_row_num,
                    "error": " | ".join(row_errors),
                    "errors": row_errors,
                    "date": "-",
                    "primary_name": p_name or str(raw_name or "-"),
                    "amount": str(raw_stock or "0"),
                    "details": f"Sold: {raw_sold or '0'}"
                })
                continue

            name_lower = p_name.lower()
            is_dup = False
            dup_reason = ""

            if name_lower in db_signatures:
                is_dup = True
                dup_reason = f"Product '{p_name}' already exists in inventory (Importing will increment stock by {p_stock})"
            elif name_lower in file_signatures:
                is_dup = True
                dup_reason = f"Duplicate product name '{p_name}' appears multiple times in spreadsheet"

            file_signatures.add(name_lower)

            record = {
                "row_number": excel_row_num,
                "row_num": excel_row_num,
                "item_name": p_name,
                "primary_name": p_name,
                "stock": p_stock,
                "sold": p_sold,
                "amount": p_stock,
                "date": "-",
                "details": f"Stock: {p_stock} | Sold: {p_sold}",
                "is_duplicate": is_dup,
                "duplicate_reason": dup_reason,
                "warning": dup_reason
            }

            if is_dup:
                duplicate_rows.append(record)
            else:
                valid_rows.append(record)

    # Aggregates
    total_vol = 0.0
    dates = []
    if entity_type in [ENTITY_SALES, ENTITY_EXPENSES]:
        total_vol = sum(r["amount"] for r in valid_rows)
        dates = [r["date"] for r in valid_rows if r.get("date") and r.get("date") != "-"]
    elif entity_type == ENTITY_INVENTORY:
        total_vol = sum(r["stock"] for r in valid_rows)

    date_min = min(dates) if dates else None
    date_max = max(dates) if dates else None

    return {
        "entity_type": entity_type,
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
        "total_volume": total_vol
    }


def execute_biz_import_batch(
    db_manager=None,
    rows_to_import: Optional[List[Dict[str, Any]]] = None,
    entity_type: str = ENTITY_SALES,
    filename: str = "",
    skipped_count: int = 0,
    db=None,
    valid_rows=None
) -> Dict[str, Any]:
    """
    Executes atomic database import of validated Business ERP records
    into business_sales, business_expenses, or inventory,
    marking source and recording the batch.
    """
    active_db = db or db_manager
    if not active_db:
        raise ExcelImportError("Database manager is required.")

    records = valid_rows if valid_rows is not None else rows_to_import
    if not records:
        raise ExcelImportError(f"No valid {entity_type} records to import.")

    timestamp_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    batch_suffix = uuid.uuid4().hex[:6]
    batch_id = f"BIZ_{entity_type.upper()[:3]}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{batch_suffix}"
    source_label = f"Imported: {os.path.basename(filename)}"

    dates = []
    total_val = 0.0

    if entity_type == ENTITY_SALES:
        tuples = []
        for r in records:
            d = r["date"]
            cust = r["customer_name"]
            inv = r["invoice_no"]
            amt = float(r["amount"])
            meth = r["payment_method"]
            dates.append(d)
            total_val += amt
            tuples.append((d, cust, inv, amt, meth, batch_id, source_label))

        active_db.add_business_sales_batch(tuples)

    elif entity_type == ENTITY_EXPENSES:
        tuples = []
        for r in records:
            d = r["date"]
            cat = r["category"]
            amt = float(r["amount"])
            notes = r.get("notes") or ""
            dates.append(d)
            total_val += amt
            tuples.append((d, cat, amt, notes, batch_id, source_label))

        active_db.add_business_expenses_batch(tuples)

    elif entity_type == ENTITY_INVENTORY:
        tuples = []
        for r in records:
            name = r["item_name"]
            stock = int(r["stock"])
            sold = int(r.get("sold", 0))
            total_val += stock
            tuples.append((name, stock, sold, batch_id, source_label))

        active_db.add_inventory_batch(tuples)

    d_min = min(dates) if dates else timestamp_str[:10]
    d_max = max(dates) if dates else timestamp_str[:10]

    # Record in excel_imports
    active_db.record_excel_import(
        batch_id=batch_id,
        filename=os.path.basename(filename),
        import_date=timestamp_str,
        total_rows=len(records) + skipped_count,
        valid_rows=len(records),
        skipped_rows=skipped_count,
        date_min=d_min,
        date_max=d_max,
        total_amount=total_val,
        profile="Business",
        notes=f"Business ERP {entity_type} import from {os.path.basename(filename)}.",
        entity_type=entity_type
    )

    return {
        "batch_id": batch_id,
        "entity_type": entity_type,
        "imported_count": len(records),
        "total_volume": total_val,
        "date_min": d_min,
        "date_max": d_max,
        "import_date": timestamp_str
    }


def remove_biz_import_batch(db_manager=None, batch_id: str = "", db=None) -> int:
    """
    Safely rolls back and removes a Business ERP import batch,
    deleting its sales, expenses, or inventory records while preserving manual data.
    """
    active_db = db or db_manager
    if not active_db:
        raise ExcelImportError("Database manager is required.")
    return active_db.remove_business_import(batch_id)


def get_business_insights_scope(db_manager=None, db=None) -> Dict[str, Any]:
    """
    Computes comprehensive dataset scope for Business ERP:
    - Sales: total count, manual count, imported count, total revenue, imported revenue, date range
    - Expenses: total count, manual count, imported count, total expenses, imported expenses, date range
    - Inventory: total products, manual products, imported products, total stock, imported stock
    - Active import batches count
    """
    active_db = db or db_manager
    if not active_db:
        raise ExcelImportError("Database manager is required.")

    conn = active_db.get_connection()
    cursor = conn.cursor()

    try:
        # 1. Sales metrics
        cursor.execute("""
            SELECT 
                COUNT(*) as total_sales,
                SUM(CASE WHEN import_batch_id IS NOT NULL THEN 1 ELSE 0 END) as imported_sales,
                SUM(CASE WHEN import_batch_id IS NULL THEN 1 ELSE 0 END) as manual_sales,
                SUM(amount) as total_rev,
                SUM(CASE WHEN import_batch_id IS NOT NULL THEN amount ELSE 0 END) as imported_rev,
                MIN(date) as earliest_sale,
                MAX(date) as latest_sale
            FROM business_sales
        """)
        s_row = cursor.fetchone()
        sales_metrics = {
            "total_count": s_row[0] or 0,
            "imported_count": s_row[1] or 0,
            "manual_count": s_row[2] or 0,
            "total_amount": s_row[3] or 0.0,
            "imported_amount": s_row[4] or 0.0,
            "earliest_date": s_row[5],
            "latest_date": s_row[6]
        }

        # 2. Expenses metrics
        cursor.execute("""
            SELECT 
                COUNT(*) as total_exp,
                SUM(CASE WHEN import_batch_id IS NOT NULL THEN 1 ELSE 0 END) as imported_exp,
                SUM(CASE WHEN import_batch_id IS NULL THEN 1 ELSE 0 END) as manual_exp,
                SUM(amount) as total_exp_amt,
                SUM(CASE WHEN import_batch_id IS NOT NULL THEN amount ELSE 0 END) as imported_exp_amt,
                MIN(date) as earliest_exp,
                MAX(date) as latest_exp
            FROM business_expenses
        """)
        e_row = cursor.fetchone()
        expense_metrics = {
            "total_count": e_row[0] or 0,
            "imported_count": e_row[1] or 0,
            "manual_count": e_row[2] or 0,
            "total_amount": e_row[3] or 0.0,
            "imported_amount": e_row[4] or 0.0,
            "earliest_date": e_row[5],
            "latest_date": e_row[6]
        }

        # 3. Inventory metrics
        cursor.execute("""
            SELECT 
                COUNT(*) as total_items,
                SUM(CASE WHEN import_batch_id IS NOT NULL THEN 1 ELSE 0 END) as imported_items,
                SUM(CASE WHEN import_batch_id IS NULL THEN 1 ELSE 0 END) as manual_items,
                SUM(stock) as total_stock,
                SUM(CASE WHEN import_batch_id IS NOT NULL THEN stock ELSE 0 END) as imported_stock,
                SUM(sold) as total_sold
            FROM inventory
        """)
        i_row = cursor.fetchone()
        inventory_metrics = {
            "total_items": i_row[0] or 0,
            "imported_items": i_row[1] or 0,
            "manual_items": i_row[2] or 0,
            "total_stock": i_row[3] or 0,
            "imported_stock": i_row[4] or 0,
            "total_sold": i_row[5] or 0
        }

        # 4. Import batches count
        cursor.execute("""
            SELECT COUNT(*) FROM excel_imports WHERE profile = 'Business'
        """)
        batches_count = cursor.fetchone()[0] or 0

        # 5. Top imported categories
        cursor.execute("""
            SELECT category, SUM(amount)
            FROM business_expenses
            WHERE import_batch_id IS NOT NULL
            GROUP BY category
            ORDER BY SUM(amount) DESC
            LIMIT 3
        """)
        top_imported_exp_cats = cursor.fetchall()

        # 6. Top imported customers
        cursor.execute("""
            SELECT customer_name, SUM(amount)
            FROM business_sales
            WHERE import_batch_id IS NOT NULL
            GROUP BY customer_name
            ORDER BY SUM(amount) DESC
            LIMIT 3
        """)
        top_imported_customers = cursor.fetchall()

        return {
            "sales": sales_metrics,
            "expenses": expense_metrics,
            "inventory": inventory_metrics,
            "active_batches": batches_count,
            "top_imported_exp_cats": top_imported_exp_cats,
            "top_imported_customers": top_imported_customers
        }
    finally:
        conn.close()
