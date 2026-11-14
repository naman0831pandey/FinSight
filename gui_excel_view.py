"""
FINSIGHT - Personal Finance Excel Import & Insights GUI View (gui_excel_view.py)
--------------------------------------------------------------------------------
Provides a comprehensive, 100% local Excel (.xlsx) import workflow scoped
strictly to the Personal Finance portal:
- File selection & multi-sheet discovery
- Live sheet preview (first rows)
- Intelligent column mapping (Date, Type, Amount, Category, Description)
- Pre-import validation engine with row-level error reporting & duplicate warnings
- Confirmation checkpoint before committing valid rows
- Distinguishable tagging (import_batch_id, source='Imported')
- Import history management with clean rollback/removal of import batches
- Scope summary banner showing dataset coverage & date range
"""

import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import gui_theme as theme
import excel_import_service

def create_excel_import_page(app):
    """Creates the Excel Import & Insights page for Personal Finance."""
    from gui_app import ScrollableFrame
    page = ScrollableFrame(app.content_container)
    frame = page.scrollable_frame

    # State variables for import workflow
    app.excel_filepath = None
    app.excel_sheet_data = None
    app.excel_validation_results = None

    # =========================================================================
    # 1. Page Header & Privacy Guarantee
    # =========================================================================
    top_bar = tk.Frame(frame, bg=theme.BG_MAIN)
    top_bar.pack(fill="x", padx=25, pady=(20, 10))

    title_box = tk.Frame(top_bar, bg=theme.BG_MAIN)
    title_box.pack(side="left", fill="x", expand=True)

    tk.Label(
        title_box,
        text="📥 Excel Import & Insights (Personal Finance)",
        font=theme.FONT_TITLE,
        fg=theme.TEXT_PRIMARY,
        bg=theme.BG_MAIN
    ).pack(anchor="w")

    tk.Label(
        title_box,
        text="Import spreadsheet statements into your Personal Finance ledger with column mapping, validation, and insights.",
        font=theme.FONT_BODY,
        fg=theme.TEXT_MUTED,
        bg=theme.BG_MAIN
    ).pack(anchor="w", pady=(3, 0))

    # Privacy badge
    privacy_badge = tk.Label(
        top_bar,
        text="🔒 100% Local Processing\nNo data leaves your device",
        font=theme.FONT_SMALL,
        fg="#10b981",
        bg=theme.BG_CARD,
        padx=12,
        pady=6,
        relief="groove",
        bd=1,
        justify="center"
    )
    privacy_badge.pack(side="right")

    # =========================================================================
    # 2. Dataset Scope Banner
    # =========================================================================
    scope_card = theme.Card(frame, width=900, height=75)
    scope_card.pack(fill="x", padx=25, pady=(5, 15))
    scope_card.pack_propagate(False)

    app.excel_scope_lbl = tk.Label(
        scope_card,
        text="📊 Personal Finance Scope: Loading dataset details...",
        font=theme.FONT_BODY_BOLD,
        fg=theme.TEXT_PRIMARY,
        bg=theme.BG_CARD,
        justify="left"
    )
    app.excel_scope_lbl.pack(side="left", padx=20, pady=10)

    btn_view_engine = tk.Button(
        scope_card,
        text="⚡ Open Insight Engine",
        font=theme.FONT_SMALL,
        fg=theme.TEXT_DARK,
        bg=theme.ACCENT,
        activebackground=theme.ACCENT,
        bd=0,
        padx=12,
        pady=6,
        cursor="hand2",
        command=lambda: app.switch_page("ai_insights")
    )
    btn_view_engine.pack(side="right", padx=20)

    # =========================================================================
    # 3. Step 1: File Selection & Sheet Chooser
    # =========================================================================
    step1_card = theme.Card(frame, width=900, height=130)
    step1_card.pack(fill="x", padx=25, pady=5)
    step1_card.pack_propagate(False)

    step1_header = tk.Label(
        step1_card,
        text="Step 1: Select Local Excel (.xlsx) File",
        font=theme.FONT_SUBTITLE,
        fg=theme.ACCENT,
        bg=theme.BG_CARD
    )
    step1_header.pack(anchor="w", padx=20, pady=(12, 6))

    file_row = tk.Frame(step1_card, bg=theme.BG_CARD)
    file_row.pack(fill="x", padx=20, pady=5)

    btn_browse = tk.Button(
        file_row,
        text="📂 Browse File (.xlsx)",
        font=theme.FONT_BODY_BOLD,
        bg=theme.ACCENT,
        fg=theme.TEXT_DARK,
        activebackground=theme.ACCENT,
        bd=0,
        padx=14,
        pady=6,
        cursor="hand2",
        command=lambda: _choose_excel_file(app)
    )
    btn_browse.pack(side="left", padx=(0, 10))

    app.lbl_excel_path = tk.Label(
        file_row,
        text="No file selected. Please choose an .xlsx workbook.",
        font=theme.FONT_BODY,
        fg=theme.TEXT_MUTED,
        bg=theme.BG_CARD,
        anchor="w"
    )
    app.lbl_excel_path.pack(side="left", fill="x", expand=True)

    # Sheet chooser row
    sheet_row = tk.Frame(step1_card, bg=theme.BG_CARD)
    sheet_row.pack(fill="x", padx=20, pady=(2, 10))

    tk.Label(
        sheet_row,
        text="Active Sheet:",
        font=theme.FONT_BODY_BOLD,
        fg=theme.TEXT_PRIMARY,
        bg=theme.BG_CARD
    ).pack(side="left", padx=(0, 8))

    app.excel_sheet_var = tk.StringVar(value="")
    app.excel_sheet_cmb = ttk.Combobox(
        sheet_row,
        textvariable=app.excel_sheet_var,
        state="readonly",
        width=25
    )
    app.excel_sheet_cmb.pack(side="left", padx=(0, 15))
    app.excel_sheet_cmb.bind("<<ComboboxSelected>>", lambda e: _load_selected_sheet(app))

    app.lbl_sheet_info = tk.Label(
        sheet_row,
        text="",
        font=theme.FONT_SMALL,
        fg=theme.TEXT_MUTED,
        bg=theme.BG_CARD
    )
    app.lbl_sheet_info.pack(side="left")

    # =========================================================================
    # 4. Step 2: Spreadsheet Preview & Column Mapping
    # =========================================================================
    step2_card = theme.Card(frame, width=900, height=360)
    step2_card.pack(fill="x", padx=25, pady=10)
    step2_card.pack_propagate(False)

    step2_header = tk.Label(
        step2_card,
        text="Step 2: Preview Spreadsheet & Map Columns to Transaction Fields",
        font=theme.FONT_SUBTITLE,
        fg=theme.ACCENT,
        bg=theme.BG_CARD
    )
    step2_header.pack(anchor="w", padx=20, pady=(12, 6))

    # Preview Table
    preview_box = tk.Frame(step2_card, bg=theme.BG_CARD)
    preview_box.pack(fill="x", padx=20, pady=5)

    app.excel_preview_tree = ttk.Treeview(preview_box, show="headings", height=5)
    app.excel_preview_tree.pack(side="left", fill="x", expand=True)

    preview_scroll = ttk.Scrollbar(preview_box, orient="vertical", command=app.excel_preview_tree.yview)
    app.excel_preview_tree.configure(yscrollcommand=preview_scroll.set)
    preview_scroll.pack(side="right", fill="y")

    # Mapping controls container
    mapping_box = tk.Frame(step2_card, bg=theme.BG_CARD)
    mapping_box.pack(fill="x", padx=20, pady=(10, 5))

    tk.Label(
        mapping_box,
        text="Map Spreadsheet Columns to FINSIGHT Fields:",
        font=theme.FONT_BODY_BOLD,
        fg=theme.TEXT_PRIMARY,
        bg=theme.BG_CARD
    ).grid(row=0, column=0, columnspan=6, sticky="w", pady=(0, 8))

    fields = [
        ("Date *", "date", "Transaction date (e.g. YYYY-MM-DD or DD/MM/YYYY)"),
        ("Type", "type", "Income or Expense (auto-inferred if omitted)"),
        ("Amount *", "amount", "Transaction amount (numeric, > 0)"),
        ("Category", "category", "Category (Food, Rent, Shopping, etc.)"),
        ("Description", "description", "Notes, merchant, or narration")
    ]

    app.excel_map_vars = {}
    app.excel_map_cmbs = {}

    for idx, (label, field_key, tip) in enumerate(fields):
        col_frame = tk.Frame(mapping_box, bg=theme.BG_CARD)
        col_frame.grid(row=1, column=idx, padx=6, sticky="ew")
        mapping_box.columnconfigure(idx, weight=1)

        tk.Label(
            col_frame,
            text=label,
            font=theme.FONT_BODY_BOLD,
            fg=theme.ACCENT if "*" in label else theme.TEXT_PRIMARY,
            bg=theme.BG_CARD
        ).pack(anchor="w")

        var = tk.StringVar(value="(None)")
        cmb = ttk.Combobox(col_frame, textvariable=var, state="readonly", width=18)
        cmb.pack(fill="x", pady=2)

        app.excel_map_vars[field_key] = var
        app.excel_map_cmbs[field_key] = cmb

    # Mapping actions row
    map_actions = tk.Frame(step2_card, bg=theme.BG_CARD)
    map_actions.pack(fill="x", padx=20, pady=(8, 10))

    btn_auto_detect = tk.Button(
        map_actions,
        text="⚡ Auto-Detect Columns",
        font=theme.FONT_SMALL,
        fg=theme.TEXT_PRIMARY,
        bg="#374151",
        bd=0,
        padx=10,
        pady=5,
        cursor="hand2",
        command=lambda: _auto_detect_mapping(app)
    )
    btn_auto_detect.pack(side="left")

    tk.Label(
        map_actions,
        text="* Required fields: Date and Amount. Other fields will use defaults if unmapped.",
        font=theme.FONT_SMALL,
        fg=theme.TEXT_MUTED,
        bg=theme.BG_CARD
    ).pack(side="left", padx=15)

    # =========================================================================
    # 5. Step 3: Validation, Duplicate Detection & Pre-Import Review
    # =========================================================================
    step3_card = theme.Card(frame, width=900, height=330)
    step3_card.pack(fill="x", padx=25, pady=10)
    step3_card.pack_propagate(False)

    step3_top = tk.Frame(step3_card, bg=theme.BG_CARD)
    step3_top.pack(fill="x", padx=20, pady=(12, 6))

    tk.Label(
        step3_top,
        text="Step 3: Validate Records & Review Warnings",
        font=theme.FONT_SUBTITLE,
        fg=theme.ACCENT,
        bg=theme.BG_CARD
    ).pack(side="left")

    btn_validate = tk.Button(
        step3_top,
        text="🔍 Validate & Inspect Rows",
        font=theme.FONT_BODY_BOLD,
        bg=theme.ACCENT,
        fg=theme.TEXT_DARK,
        activebackground=theme.ACCENT,
        bd=0,
        padx=14,
        pady=5,
        cursor="hand2",
        command=lambda: _run_validation(app)
    )
    btn_validate.pack(side="right")

    # Validation KPI Badges Row
    app.val_kpi_frame = tk.Frame(step3_card, bg=theme.BG_CARD)
    app.val_kpi_frame.pack(fill="x", padx=20, pady=5)

    app.lbl_val_total = tk.Label(app.val_kpi_frame, text="Total Rows: -", font=theme.FONT_BODY_BOLD, fg=theme.TEXT_PRIMARY, bg=theme.BG_CARD)
    app.lbl_val_total.pack(side="left", padx=(0, 15))

    app.lbl_val_valid = tk.Label(app.val_kpi_frame, text="✅ Valid: -", font=theme.FONT_BODY_BOLD, fg="#10b981", bg=theme.BG_CARD)
    app.lbl_val_valid.pack(side="left", padx=15)

    app.lbl_val_dups = tk.Label(app.val_kpi_frame, text="⚠️ Duplicates: -", font=theme.FONT_BODY_BOLD, fg="#f59e0b", bg=theme.BG_CARD)
    app.lbl_val_dups.pack(side="left", padx=15)

    app.lbl_val_errors = tk.Label(app.val_kpi_frame, text="❌ Errors: -", font=theme.FONT_BODY_BOLD, fg=theme.EXPENSE, bg=theme.BG_CARD)
    app.lbl_val_errors.pack(side="left", padx=15)

    # Options row
    opt_row = tk.Frame(step3_card, bg=theme.BG_CARD)
    opt_row.pack(fill="x", padx=20, pady=(2, 6))

    app.skip_dups_var = tk.BooleanVar(value=True)
    app.chk_skip_dups = tk.Checkbutton(
        opt_row,
        text="Skip suspected duplicate rows during import (Recommended)",
        variable=app.skip_dups_var,
        font=theme.FONT_SMALL,
        fg=theme.TEXT_PRIMARY,
        bg=theme.BG_CARD,
        selectcolor=theme.BG_MAIN,
        activebackground=theme.BG_CARD,
        activeforeground=theme.TEXT_PRIMARY
    )
    app.chk_skip_dups.pack(side="left")

    # Validation details treeview
    val_tree_box = tk.Frame(step3_card, bg=theme.BG_CARD)
    val_tree_box.pack(fill="both", expand=True, padx=20, pady=(2, 12))

    val_cols = ("row", "status", "date", "type", "amount", "category", "notes", "issue")
    app.val_tree = ttk.Treeview(val_tree_box, columns=val_cols, show="headings", height=5)
    app.val_tree.heading("row", text="Row #")
    app.val_tree.heading("status", text="Status")
    app.val_tree.heading("date", text="Date")
    app.val_tree.heading("type", text="Type")
    app.val_tree.heading("amount", text="Amount")
    app.val_tree.heading("category", text="Category")
    app.val_tree.heading("notes", text="Description")
    app.val_tree.heading("issue", text="Validation Message / Reason")

    app.val_tree.column("row", width=50, anchor="center")
    app.val_tree.column("status", width=80, anchor="center")
    app.val_tree.column("date", width=85, anchor="center")
    app.val_tree.column("type", width=65, anchor="center")
    app.val_tree.column("amount", width=80, anchor="e")
    app.val_tree.column("category", width=90, anchor="w")
    app.val_tree.column("notes", width=120, anchor="w")
    app.val_tree.column("issue", width=220, anchor="w")

    app.val_tree.pack(side="left", fill="both", expand=True)

    val_scroll = ttk.Scrollbar(val_tree_box, orient="vertical", command=app.val_tree.yview)
    app.val_tree.configure(yscrollcommand=val_scroll.set)
    val_scroll.pack(side="right", fill="y")

    # =========================================================================
    # 6. Step 4: Confirmation & Import Commit Action
    # =========================================================================
    step4_card = theme.Card(frame, width=900, height=80)
    step4_card.pack(fill="x", padx=25, pady=5)
    step4_card.pack_propagate(False)

    step4_box = tk.Frame(step4_card, bg=theme.BG_CARD)
    step4_box.pack(fill="both", expand=True, padx=20, pady=12)

    app.lbl_commit_info = tk.Label(
        step4_box,
        text="Validate rows above to enable final import confirmation.",
        font=theme.FONT_BODY,
        fg=theme.TEXT_MUTED,
        bg=theme.BG_CARD
    )
    app.lbl_commit_info.pack(side="left")

    app.btn_commit_import = tk.Button(
        step4_box,
        text="📥 Confirm & Import Valid Rows",
        font=theme.FONT_BODY_BOLD,
        bg="#10b981",
        fg=theme.TEXT_DARK,
        activebackground="#059669",
        bd=0,
        padx=18,
        pady=8,
        cursor="hand2",
        state="disabled",
        command=lambda: _confirm_and_import(app)
    )
    app.btn_commit_import.pack(side="right")

    # =========================================================================
    # 7. Past Excel Imports History & Removal / Rollback
    # =========================================================================
    history_card = theme.Card(frame, width=900, height=270)
    history_card.pack(fill="x", padx=25, pady=15)
    history_card.pack_propagate(False)

    hist_header_box = tk.Frame(history_card, bg=theme.BG_CARD)
    hist_header_box.pack(fill="x", padx=20, pady=(12, 6))

    tk.Label(
        hist_header_box,
        text="📑 Personal Finance Import History & Batch Management",
        font=theme.FONT_SUBTITLE,
        fg=theme.ACCENT,
        bg=theme.BG_CARD
    ).pack(side="left")

    app.btn_delete_batch = tk.Button(
        hist_header_box,
        text="🗑️ Remove Selected Import",
        font=theme.FONT_SMALL,
        fg=theme.TEXT_PRIMARY,
        bg=theme.EXPENSE,
        activebackground=theme.EXPENSE,
        bd=0,
        padx=12,
        pady=5,
        cursor="hand2",
        state="disabled",
        command=lambda: _remove_selected_batch(app)
    )
    app.btn_delete_batch.pack(side="right")

    hist_tree_box = tk.Frame(history_card, bg=theme.BG_CARD)
    hist_tree_box.pack(fill="both", expand=True, padx=20, pady=(2, 12))

    hist_cols = ("batch_id", "filename", "import_date", "valid_rows", "skipped_rows", "date_range", "total_amt")
    app.hist_tree = ttk.Treeview(hist_tree_box, columns=hist_cols, show="headings", height=5)
    app.hist_tree.heading("batch_id", text="Batch ID")
    app.hist_tree.heading("filename", text="Spreadsheet File")
    app.hist_tree.heading("import_date", text="Imported On")
    app.hist_tree.heading("valid_rows", text="Records Added")
    app.hist_tree.heading("skipped_rows", text="Skipped")
    app.hist_tree.heading("date_range", text="Transaction Date Range")
    app.hist_tree.heading("total_amt", text="Total Volume")

    app.hist_tree.column("batch_id", width=120, anchor="center")
    app.hist_tree.column("filename", width=160, anchor="w")
    app.hist_tree.column("import_date", width=110, anchor="center")
    app.hist_tree.column("valid_rows", width=85, anchor="center")
    app.hist_tree.column("skipped_rows", width=65, anchor="center")
    app.hist_tree.column("date_range", width=160, anchor="center")
    app.hist_tree.column("total_amt", width=110, anchor="e")

    app.hist_tree.pack(side="left", fill="both", expand=True)

    hist_scroll = ttk.Scrollbar(hist_tree_box, orient="vertical", command=app.hist_tree.yview)
    app.hist_tree.configure(yscrollcommand=hist_scroll.set)
    hist_scroll.pack(side="right", fill="y")

    app.hist_tree.bind("<<TreeviewSelect>>", lambda e: _on_history_select(app))

    return page


# =============================================================================
# Helper Event Handlers
# =============================================================================

def refresh_excel_import_page(app):
    """Refreshes the dynamic data on the Excel Import page."""
    _refresh_scope_banner(app)
    _refresh_import_history(app)

def _refresh_scope_banner(app):
    """Updates the dataset scope banner showing total records, manual vs imported, and date range."""
    try:
        scope = excel_import_service.get_personal_insights_scope(app.tracker.db)
        d_min = scope["earliest_date"] or "N/A"
        d_max = scope["latest_date"] or "N/A"
        date_range_str = f"{d_min} to {d_max}" if scope["total_transactions"] > 0 else "No records recorded"

        app.excel_scope_lbl.config(
            text=f"📊 Personal Finance Scope: {scope['total_transactions']} Total Transactions "
                 f"({scope['manual_count']} Manual ✍️ | {scope['imported_count']} Imported 📥 across {scope['active_batches']} active batches) "
                 f"• Date Range: {date_range_str}"
        )
    except Exception as e:
        app.excel_scope_lbl.config(text=f"📊 Personal Finance Scope: Error loading scope ({str(e)})")

def _refresh_import_history(app):
    """Refreshes the past imports table."""
    app.hist_tree.delete(*app.hist_tree.get_children())
    app.btn_delete_batch.config(state="disabled")
    try:
        imports = app.tracker.get_excel_imports(profile="Personal")
        for imp in imports:
            # batch_id, filename, import_date, total_rows, valid_rows, skipped_rows, date_min, date_max, total_amount, profile, notes
            bid = imp[0]
            fname = os.path.basename(imp[1]) if imp[1] else "-"
            idate = (imp[2] or "-")[:16]
            vrows = imp[4]
            srows = imp[5]
            d_range = f"{imp[6] or '-'} to {imp[7] or '-'}"
            tot_amt = f"₹{imp[8]:,.2f}" if imp[8] is not None else "-"
            app.hist_tree.insert("", "end", values=(bid, fname, idate, vrows, srows, d_range, tot_amt))
    except Exception as e:
        print(f"Error loading import history: {e}")

def _on_history_select(app):
    sel = app.hist_tree.selection()
    if sel:
        app.btn_delete_batch.config(state="normal")
    else:
        app.btn_delete_batch.config(state="disabled")

def _choose_excel_file(app):
    """Prompts user to select a local .xlsx file."""
    filepath = filedialog.askopenfilename(
        title="Select Excel File to Import",
        filetypes=[("Excel Workbooks (*.xlsx)", "*.xlsx")]
    )
    if not filepath:
        return

    # Check extension
    if not filepath.lower().endswith(".xlsx"):
        messagebox.showerror(
            "Unsupported Format",
            "FINSIGHT only supports standard Excel .xlsx workbooks processed 100% locally.\n"
            "Legacy .xls (Excel 97-2003), encrypted files, and macro/binary formats (.xlsb) are unsupported."
        )
        return

    app.excel_filepath = filepath
    filename = os.path.basename(filepath)
    app.lbl_excel_path.config(text=f"Selected: {filename} ({filepath})", fg=theme.TEXT_PRIMARY)

    try:
        sheets = excel_import_service.get_sheet_names(filepath)
        app.excel_sheet_cmb.config(values=sheets)
        if sheets:
            app.excel_sheet_var.set(sheets[0])
            _load_selected_sheet(app)
    except Exception as e:
        messagebox.showerror("Error Reading Workbook", f"Could not open Excel file:\n{str(e)}")

def _load_selected_sheet(app):
    """Loads sheet data, headers, and previews the first rows."""
    if not app.excel_filepath or not app.excel_sheet_var.get():
        return

    sheet_name = app.excel_sheet_var.get()
    try:
        headers, rows = excel_import_service.read_excel_sheet(app.excel_filepath, sheet_name)
        app.excel_sheet_data = (headers, rows)
        app.lbl_sheet_info.config(text=f"({len(headers)} columns detected, {len(rows)} data rows)")

        # Clear and configure preview treeview
        app.excel_preview_tree.delete(*app.excel_preview_tree.get_children())
        app.excel_preview_tree["columns"] = headers
        for h in headers:
            app.excel_preview_tree.heading(h, text=h)
            app.excel_preview_tree.column(h, width=110, anchor="w")

        # Preview first 8 rows
        for row in rows[:8]:
            row_vals = [str(cell) if cell is not None else "" for cell in row]
            app.excel_preview_tree.insert("", "end", values=row_vals)

        # Update column mapping combobox options
        col_options = ["(None)"] + headers
        for cmb in app.excel_map_cmbs.values():
            cmb.config(values=col_options)

        # Auto-detect mappings
        _auto_detect_mapping(app)

        # Reset validation state
        app.excel_validation_results = None
        app.val_tree.delete(*app.val_tree.get_children())
        app.lbl_val_total.config(text=f"Total Rows: {len(rows)}")
        app.lbl_val_valid.config(text="✅ Valid: -")
        app.lbl_val_dups.config(text="⚠️ Duplicates: -")
        app.lbl_val_errors.config(text="❌ Errors: -")
        app.btn_commit_import.config(state="disabled")
        app.lbl_commit_info.config(text="Validate rows above to review results before importing.")

    except Exception as e:
        messagebox.showerror("Error Reading Sheet", f"Failed to read sheet '{sheet_name}':\n{str(e)}")

def _auto_detect_mapping(app):
    """Runs column detection heuristics on the loaded headers."""
    if not app.excel_sheet_data:
        return
    headers, _ = app.excel_sheet_data
    detected = excel_import_service.detect_column_mapping(headers)

    for field_key, var in app.excel_map_vars.items():
        mapped_col = detected.get(field_key)
        if mapped_col and mapped_col in headers:
            var.set(mapped_col)
        else:
            var.set("(None)")

def _run_validation(app):
    """Validates the rows against the selected mapping and checks for duplicates."""
    if not app.excel_sheet_data:
        messagebox.showwarning("No Data", "Please select and load an Excel sheet first.")
        return

    headers, rows = app.excel_sheet_data
    mapping = {k: v.get() for k, v in app.excel_map_vars.items()}

    # Check required mapping
    if mapping.get("date") == "(None)" or mapping.get("amount") == "(None)":
        messagebox.showwarning(
            "Missing Required Mapping",
            "Both 'Date' and 'Amount' columns must be mapped to valid spreadsheet headers."
        )
        return

    try:
        # Fetch existing Personal transactions for duplicate checking
        existing_txs = app.tracker.get_transactions()
        results = excel_import_service.validate_and_prepare_rows(headers, rows, mapping, existing_txs)
        app.excel_validation_results = results

        # Update KPI labels
        app.lbl_val_total.config(text=f"Total Rows: {results['total_rows']}")
        app.lbl_val_valid.config(text=f"✅ Valid: {results['valid_count']}")
        app.lbl_val_dups.config(text=f"⚠️ Duplicates: {results['duplicate_count']}")
        app.lbl_val_errors.config(text=f"❌ Errors: {results['error_count']}")

        # Populate validation report treeview
        app.val_tree.delete(*app.val_tree.get_children())

        # Show invalid rows first
        error_list = results.get("invalid_rows") or results.get("error_rows") or []
        for err in error_list:
            row_idx = err.get("row_num") or err.get("row_number") or "-"
            d_val = err.get("date") or "-"
            t_val = err.get("type") or "-"
            a_val = err.get("amount") or "-"
            c_val = err.get("category") or "-"
            desc_val = err.get("description") or err.get("notes") or "-"
            err_msg = "; ".join(err.get("errors", [])) if err.get("errors") else err.get("error", "Validation error")
            app.val_tree.insert(
                "",
                "end",
                values=(
                    row_idx,
                    "❌ Error",
                    d_val,
                    t_val,
                    a_val,
                    c_val,
                    desc_val,
                    err_msg
                )
            )

        # Show duplicate warning rows
        dup_list = results.get("duplicate_rows") or []
        for dup in dup_list:
            row_idx = dup.get("row_num") or dup.get("row_number") or "-"
            d_val = dup.get("date") or "-"
            t_val = dup.get("type") or "-"
            a_val = f"₹{dup.get('amount', 0):,.2f}"
            c_val = dup.get("category") or "-"
            desc_val = dup.get("description") or dup.get("notes") or "-"
            warn_msg = dup.get("warning") or dup.get("duplicate_reason") or "Suspected duplicate record"
            app.val_tree.insert(
                "",
                "end",
                values=(
                    row_idx,
                    "⚠️ Duplicate",
                    d_val,
                    t_val,
                    a_val,
                    c_val,
                    desc_val,
                    warn_msg
                )
            )

        # Show valid clean rows (sample first 15)
        valid_list = results.get("valid_rows") or []
        for val in valid_list[:15]:
            row_idx = val.get("row_num") or val.get("row_number") or "-"
            d_val = val.get("date") or "-"
            t_val = val.get("type") or "-"
            a_val = f"₹{val.get('amount', 0):,.2f}"
            c_val = val.get("category") or "-"
            desc_val = val.get("description") or val.get("notes") or "-"
            app.val_tree.insert(
                "",
                "end",
                values=(
                    row_idx,
                    "✅ Valid",
                    d_val,
                    t_val,
                    a_val,
                    c_val,
                    desc_val,
                    "Ready to import"
                )
            )

        total_importable = results["valid_count"]
        if not app.skip_dups_var.get():
            total_importable += results["duplicate_count"]

        if total_importable > 0:
            app.btn_commit_import.config(state="normal")
            app.lbl_commit_info.config(
                text=f"Ready to import {total_importable} transactions. Click 'Confirm & Import' to proceed.",
                fg="#10b981"
            )
        else:
            app.btn_commit_import.config(state="disabled")
            app.lbl_commit_info.config(
                text="No valid records available to import. Please check column mappings and spreadsheet errors.",
                fg=theme.EXPENSE
            )

    except Exception as e:
        messagebox.showerror("Validation Error", f"Failed during data validation:\n{str(e)}")

def _confirm_and_import(app):
    """Asks for confirmation and executes the batch import into Personal Finance."""
    if not app.excel_validation_results:
        return

    results = app.excel_validation_results
    skip_dups = app.skip_dups_var.get()
    
    rows_to_import = list(results["valid_rows"])
    if not skip_dups:
        rows_to_import.extend(results["duplicate_rows"])

    if not rows_to_import:
        messagebox.showinfo("Nothing to Import", "There are no valid transactions to import.")
        return

    # Confirmation Prompt
    confirm_msg = (
        f"Ready to import {len(rows_to_import)} transactions into your Personal Finance ledger.\n\n"
        f"• Spreadsheet: {os.path.basename(app.excel_filepath)}\n"
        f"• Valid Records: {len(results['valid_rows'])}\n"
        f"• Suspected Duplicates: {results['duplicate_count']} ({'Skipped' if skip_dups else 'Included'})\n"
        f"• Skipped Errors: {results['error_count']}\n\n"
        f"These transactions will be tagged as 'Imported' with a unique batch ID.\n"
        f"You can review or remove this entire batch at any time.\n\n"
        f"Proceed with import?"
    )

    if not messagebox.askyesno("Confirm Personal Finance Import", confirm_msg):
        return

    try:
        batch_summary = excel_import_service.execute_import_batch(
            db=app.tracker.db,
            valid_rows=rows_to_import,
            filename=os.path.basename(app.excel_filepath),
            profile="Personal",
            skipped_count=results["error_count"] + (results["duplicate_count"] if skip_dups else 0)
        )

        messagebox.showinfo(
            "Import Successful",
            f"Successfully imported {batch_summary['imported_count']} transactions!\n\n"
            f"• Batch ID: {batch_summary['batch_id']}\n"
            f"• Total Volume: ₹{batch_summary['total_amount']:,.2f}\n"
            f"• Date Range: {batch_summary['date_min']} to {batch_summary['date_max']}\n\n"
            f"Your Personal Finance Dashboard and Insight Engine have been updated."
        )

        # Reset preview/validation UI
        app.excel_validation_results = None
        app.btn_commit_import.config(state="disabled")
        app.lbl_commit_info.config(text="Import complete. Select another sheet or review past imports below.")

        # Refresh all views
        refresh_excel_import_page(app)
        if hasattr(app, "refresh_page_data"):
            app.refresh_page_data(app.current_page)

    except Exception as e:
        messagebox.showerror("Import Execution Error", f"Failed to commit transactions:\n{str(e)}")

def _remove_selected_batch(app):
    """Removes the selected import batch, rolling back its transactions."""
    sel = app.hist_tree.selection()
    if not sel:
        return
    item = app.hist_tree.item(sel[0])
    values = item.get("values", [])
    if not values:
        return

    batch_id = str(values[0])
    filename = str(values[1])
    count = values[3]

    confirm_msg = (
        f"Are you sure you want to remove import batch '{batch_id}'?\n\n"
        f"• File: {filename}\n"
        f"• Transactions to Delete: {count}\n\n"
        f"This will permanently delete ONLY the transactions imported in this batch.\n"
        f"All manual records and transactions from other imports will be strictly preserved.\n\n"
        f"Proceed with rollback?"
    )

    if not messagebox.askyesno("Confirm Import Removal", confirm_msg):
        return

    try:
        deleted_count = excel_import_service.remove_import_batch(app.tracker.db, batch_id)
        messagebox.showinfo(
            "Import Removed",
            f"Successfully rolled back batch '{batch_id}'.\nRemoved {deleted_count} imported transactions."
        )
        refresh_excel_import_page(app)
        if hasattr(app, "refresh_page_data"):
            app.refresh_page_data(app.current_page)
    except Exception as e:
        messagebox.showerror("Removal Error", f"Failed to remove import batch:\n{str(e)}")
