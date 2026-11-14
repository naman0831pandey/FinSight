"""
FINSIGHT - Business ERP Excel Import & Insights GUI View (gui_biz_excel_view.py)
--------------------------------------------------------------------------------
Provides a comprehensive, 100% local Excel (.xlsx) import workflow scoped
strictly to the Business ERP portal:
- File selection & multi-sheet discovery
- Target entity selection: Sales, Expenses, or Inventory
- Live sheet preview (first rows)
- Intelligent column mapping dynamically tailored to the selected entity
- Pre-import validation engine with row-level error reporting & duplicate warnings
- Confirmation checkpoint before committing valid rows
- Distinguishable tagging (import_batch_id, source='Imported: <filename>')
- Import history management with clean rollback/removal of import batches
- Business scope banner and imported insights metrics
"""

import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import gui_theme as theme
import excel_business_import_service as biz_service
from excel_business_import_service import (
    ENTITY_SALES,
    ENTITY_EXPENSES,
    ENTITY_INVENTORY,
    VALID_ENTITIES
)

def create_biz_excel_import_page(app):
    """Creates the Excel Import & Insights page for Business ERP."""
    from gui_app import ScrollableFrame
    page = ScrollableFrame(app.content_container)
    frame = page.scrollable_frame

    # State variables
    app.biz_excel_filepath = None
    app.biz_excel_sheet_data = None
    app.biz_excel_validation_results = None
    app.biz_entity_var = tk.StringVar(value=ENTITY_SALES)

    # =========================================================================
    # 1. Page Header & Privacy Guarantee
    # =========================================================================
    top_bar = tk.Frame(frame, bg=theme.BG_MAIN)
    top_bar.pack(fill="x", padx=25, pady=(20, 10))

    title_box = tk.Frame(top_bar, bg=theme.BG_MAIN)
    title_box.pack(side="left", fill="x", expand=True)

    tk.Label(
        title_box,
        text="📥 Business ERP Excel Import & Insights",
        font=theme.FONT_TITLE,
        fg=theme.TEXT_PRIMARY,
        bg=theme.BG_MAIN
    ).pack(anchor="w")

    tk.Label(
        title_box,
        text="Import spreadsheet statements into your Business ERP (Sales, Expenses, or Inventory) with column mapping, validation, and insights.",
        font=theme.FONT_BODY,
        fg=theme.TEXT_MUTED,
        bg=theme.BG_MAIN
    ).pack(anchor="w", pady=(3, 0))

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
    # 2. Business Dataset Scope Banner
    # =========================================================================
    scope_card = theme.Card(frame, width=900, height=105)
    scope_card.pack(fill="x", padx=25, pady=(5, 12))
    scope_card.pack_propagate(False)

    app.biz_scope_lbl = tk.Label(
        scope_card,
        text="📊 Business ERP Scope: Loading dataset details...",
        font=theme.FONT_SMALL,
        fg=theme.TEXT_PRIMARY,
        bg=theme.BG_CARD,
        justify="left"
    )
    app.biz_scope_lbl.pack(side="left", padx=20, pady=10)

    btn_links_box = tk.Frame(scope_card, bg=theme.BG_CARD)
    btn_links_box.pack(side="right", padx=15, pady=10)

    tk.Button(
        btn_links_box,
        text="💼 ERP Dashboard",
        font=theme.FONT_SMALL,
        fg=theme.TEXT_DARK,
        bg=theme.INCOME,
        bd=0,
        padx=10,
        pady=5,
        cursor="hand2",
        command=lambda: app.switch_page("biz_dashboard")
    ).pack(side="left", padx=4)

    tk.Button(
        btn_links_box,
        text="📊 P&L Calculator",
        font=theme.FONT_SMALL,
        fg=theme.TEXT_PRIMARY,
        bg="#374151",
        bd=0,
        padx=10,
        pady=5,
        cursor="hand2",
        command=lambda: app.switch_page("biz_pl")
    ).pack(side="left", padx=4)

    # =========================================================================
    # 3. Step 1: File Selection & Choose Entity Type
    # =========================================================================
    step1_card = theme.Card(frame, width=900, height=170)
    step1_card.pack(fill="x", padx=25, pady=5)
    step1_card.pack_propagate(False)

    step1_top = tk.Frame(step1_card, bg=theme.BG_CARD)
    step1_top.pack(fill="x", padx=20, pady=(10, 4))

    tk.Label(
        step1_top,
        text="Step 1: Select File (.xlsx) & Target Business Module",
        font=theme.FONT_SUBTITLE,
        fg=theme.INCOME,
        bg=theme.BG_CARD
    ).pack(side="left")

    # File browsing row
    file_row = tk.Frame(step1_card, bg=theme.BG_CARD)
    file_row.pack(fill="x", padx=20, pady=3)

    btn_browse = tk.Button(
        file_row,
        text="📂 Browse File (.xlsx)",
        font=theme.FONT_BODY_BOLD,
        bg=theme.INCOME,
        fg=theme.TEXT_DARK,
        activebackground=theme.INCOME,
        bd=0,
        padx=14,
        pady=5,
        cursor="hand2",
        command=lambda: _choose_biz_excel_file(app)
    )
    btn_browse.pack(side="left", padx=(0, 10))

    app.lbl_biz_excel_path = tk.Label(
        file_row,
        text="No file selected. Please select an .xlsx spreadsheet workbook.",
        font=theme.FONT_BODY,
        fg=theme.TEXT_MUTED,
        bg=theme.BG_CARD,
        anchor="w"
    )
    app.lbl_biz_excel_path.pack(side="left", fill="x", expand=True)

    # Sheet chooser and entity radio row
    select_row = tk.Frame(step1_card, bg=theme.BG_CARD)
    select_row.pack(fill="x", padx=20, pady=5)

    tk.Label(
        select_row,
        text="Active Sheet:",
        font=theme.FONT_BODY_BOLD,
        fg=theme.TEXT_PRIMARY,
        bg=theme.BG_CARD
    ).pack(side="left", padx=(0, 6))

    app.biz_excel_sheet_var = tk.StringVar(value="")
    app.biz_excel_sheet_cmb = ttk.Combobox(
        select_row,
        textvariable=app.biz_excel_sheet_var,
        state="readonly",
        width=20
    )
    app.biz_excel_sheet_cmb.pack(side="left", padx=(0, 15))
    app.biz_excel_sheet_cmb.bind("<<ComboboxSelected>>", lambda e: _load_selected_biz_sheet(app))

    app.lbl_biz_sheet_info = tk.Label(
        select_row,
        text="",
        font=theme.FONT_SMALL,
        fg=theme.TEXT_MUTED,
        bg=theme.BG_CARD
    )
    app.lbl_biz_sheet_info.pack(side="left", padx=(0, 20))

    # Entity selection radio buttons
    tk.Label(
        select_row,
        text="What does this sheet contain? 👉",
        font=theme.FONT_BODY_BOLD,
        fg=theme.TEXT_PRIMARY,
        bg=theme.BG_CARD
    ).pack(side="left", padx=(10, 8))

    entity_options = [
        ("💼 Daily Sales", ENTITY_SALES),
        ("💸 Expenses", ENTITY_EXPENSES),
        ("📦 Inventory / Stock", ENTITY_INVENTORY)
    ]

    for label, val in entity_options:
        rb = tk.Radiobutton(
            select_row,
            text=label,
            value=val,
            variable=app.biz_entity_var,
            font=theme.FONT_BODY_BOLD,
            fg=theme.TEXT_PRIMARY,
            bg=theme.BG_CARD,
            selectcolor=theme.BG_MAIN,
            activebackground=theme.BG_CARD,
            activeforeground=theme.TEXT_PRIMARY,
            command=lambda: _on_entity_type_change(app)
        )
        rb.pack(side="left", padx=8)

    # =========================================================================
    # 4. Step 2: Spreadsheet Preview & Column Mapping
    # =========================================================================
    step2_card = theme.Card(frame, width=900, height=360)
    step2_card.pack(fill="x", padx=25, pady=8)
    step2_card.pack_propagate(False)

    step2_header = tk.Label(
        step2_card,
        text="Step 2: Preview Spreadsheet & Map Columns to Business Fields",
        font=theme.FONT_SUBTITLE,
        fg=theme.INCOME,
        bg=theme.BG_CARD
    )
    step2_header.pack(anchor="w", padx=20, pady=(10, 4))

    # Preview Table
    preview_box = tk.Frame(step2_card, bg=theme.BG_CARD)
    preview_box.pack(fill="x", padx=20, pady=3)

    app.biz_preview_tree = ttk.Treeview(preview_box, show="headings", height=5)
    app.biz_preview_tree.pack(side="left", fill="x", expand=True)

    preview_scroll = ttk.Scrollbar(preview_box, orient="vertical", command=app.biz_preview_tree.yview)
    app.biz_preview_tree.configure(yscrollcommand=preview_scroll.set)
    preview_scroll.pack(side="right", fill="y")

    # Mapping controls container (dynamically rebuilt on entity change)
    app.biz_mapping_container = tk.Frame(step2_card, bg=theme.BG_CARD)
    app.biz_mapping_container.pack(fill="x", padx=20, pady=(8, 4))

    app.biz_map_vars = {}
    app.biz_map_cmbs = {}

    # Initial mapping build
    _rebuild_mapping_controls(app)

    # Mapping actions row
    map_actions = tk.Frame(step2_card, bg=theme.BG_CARD)
    map_actions.pack(fill="x", padx=20, pady=(6, 8))

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
        command=lambda: _auto_detect_biz_mapping(app)
    )
    btn_auto_detect.pack(side="left")

    app.lbl_mapping_legend = tk.Label(
        map_actions,
        text="* Required fields marked with asterisk. Optional fields will use default values if unmapped.",
        font=theme.FONT_SMALL,
        fg=theme.TEXT_MUTED,
        bg=theme.BG_CARD
    )
    app.lbl_mapping_legend.pack(side="left", padx=15)

    # =========================================================================
    # 5. Step 3: Validation, Duplicate Detection & Pre-Import Review
    # =========================================================================
    step3_card = theme.Card(frame, width=900, height=330)
    step3_card.pack(fill="x", padx=25, pady=8)
    step3_card.pack_propagate(False)

    step3_top = tk.Frame(step3_card, bg=theme.BG_CARD)
    step3_top.pack(fill="x", padx=20, pady=(10, 4))

    tk.Label(
        step3_top,
        text="Step 3: Validate Records & Review Warnings",
        font=theme.FONT_SUBTITLE,
        fg=theme.INCOME,
        bg=theme.BG_CARD
    ).pack(side="left")

    btn_validate = tk.Button(
        step3_top,
        text="🔍 Validate & Inspect Rows",
        font=theme.FONT_BODY_BOLD,
        bg=theme.INCOME,
        fg=theme.TEXT_DARK,
        activebackground=theme.INCOME,
        bd=0,
        padx=14,
        pady=5,
        cursor="hand2",
        command=lambda: _run_biz_validation(app)
    )
    btn_validate.pack(side="right")

    # Validation KPI Badges Row
    app.biz_val_kpi_frame = tk.Frame(step3_card, bg=theme.BG_CARD)
    app.biz_val_kpi_frame.pack(fill="x", padx=20, pady=4)

    app.lbl_biz_val_total = tk.Label(app.biz_val_kpi_frame, text="Total Rows: -", font=theme.FONT_BODY_BOLD, fg=theme.TEXT_PRIMARY, bg=theme.BG_CARD)
    app.lbl_biz_val_total.pack(side="left", padx=(0, 15))

    app.lbl_biz_val_valid = tk.Label(app.biz_val_kpi_frame, text="✅ Valid: -", font=theme.FONT_BODY_BOLD, fg="#10b981", bg=theme.BG_CARD)
    app.lbl_biz_val_valid.pack(side="left", padx=15)

    app.lbl_biz_val_dups = tk.Label(app.biz_val_kpi_frame, text="⚠️ Duplicates: -", font=theme.FONT_BODY_BOLD, fg="#f59e0b", bg=theme.BG_CARD)
    app.lbl_biz_val_dups.pack(side="left", padx=15)

    app.lbl_biz_val_errors = tk.Label(app.biz_val_kpi_frame, text="❌ Errors: -", font=theme.FONT_BODY_BOLD, fg=theme.EXPENSE, bg=theme.BG_CARD)
    app.lbl_biz_val_errors.pack(side="left", padx=15)

    # Options row
    opt_row = tk.Frame(step3_card, bg=theme.BG_CARD)
    opt_row.pack(fill="x", padx=20, pady=(2, 5))

    app.biz_skip_dups_var = tk.BooleanVar(value=True)
    app.chk_biz_skip_dups = tk.Checkbutton(
        opt_row,
        text="Skip suspected duplicate rows during import (Recommended)",
        variable=app.biz_skip_dups_var,
        font=theme.FONT_SMALL,
        fg=theme.TEXT_PRIMARY,
        bg=theme.BG_CARD,
        selectcolor=theme.BG_MAIN,
        activebackground=theme.BG_CARD,
        activeforeground=theme.TEXT_PRIMARY
    )
    app.chk_biz_skip_dups.pack(side="left")

    # Validation details treeview
    val_tree_box = tk.Frame(step3_card, bg=theme.BG_CARD)
    val_tree_box.pack(fill="both", expand=True, padx=20, pady=(2, 10))

    val_cols = ("row", "status", "date", "primary", "value", "details", "issue")
    app.biz_val_tree = ttk.Treeview(val_tree_box, columns=val_cols, show="headings", height=5)
    app.biz_val_tree.heading("row", text="Row #")
    app.biz_val_tree.heading("status", text="Status")
    app.biz_val_tree.heading("date", text="Date")
    app.biz_val_tree.heading("primary", text="Customer / Category / Item")
    app.biz_val_tree.heading("value", text="Amount / Stock")
    app.biz_val_tree.heading("details", text="Details")
    app.biz_val_tree.heading("issue", text="Validation Message / Reason")

    app.biz_val_tree.column("row", width=50, anchor="center")
    app.biz_val_tree.column("status", width=80, anchor="center")
    app.biz_val_tree.column("date", width=85, anchor="center")
    app.biz_val_tree.column("primary", width=140, anchor="w")
    app.biz_val_tree.column("value", width=95, anchor="e")
    app.biz_val_tree.column("details", width=120, anchor="w")
    app.biz_val_tree.column("issue", width=240, anchor="w")

    app.biz_val_tree.pack(side="left", fill="both", expand=True)

    val_scroll = ttk.Scrollbar(val_tree_box, orient="vertical", command=app.biz_val_tree.yview)
    app.biz_val_tree.configure(yscrollcommand=val_scroll.set)
    val_scroll.pack(side="right", fill="y")

    # =========================================================================
    # 6. Step 4: Confirmation & Import Commit Action
    # =========================================================================
    step4_card = theme.Card(frame, width=900, height=75)
    step4_card.pack(fill="x", padx=25, pady=5)
    step4_card.pack_propagate(False)

    step4_box = tk.Frame(step4_card, bg=theme.BG_CARD)
    step4_box.pack(fill="both", expand=True, padx=20, pady=10)

    app.lbl_biz_commit_info = tk.Label(
        step4_box,
        text="Validate rows above to review results and enable final import confirmation.",
        font=theme.FONT_BODY,
        fg=theme.TEXT_MUTED,
        bg=theme.BG_CARD
    )
    app.lbl_biz_commit_info.pack(side="left")

    app.btn_biz_commit_import = tk.Button(
        step4_box,
        text="📥 Confirm & Import Valid Rows",
        font=theme.FONT_BODY_BOLD,
        bg="#10b981",
        fg=theme.TEXT_DARK,
        activebackground="#059669",
        bd=0,
        padx=18,
        pady=7,
        cursor="hand2",
        state="disabled",
        command=lambda: _confirm_and_import_biz(app)
    )
    app.btn_biz_commit_import.pack(side="right")

    # =========================================================================
    # 7. Past Business Imports History & Removal / Rollback
    # =========================================================================
    history_card = theme.Card(frame, width=900, height=270)
    history_card.pack(fill="x", padx=25, pady=12)
    history_card.pack_propagate(False)

    hist_header_box = tk.Frame(history_card, bg=theme.BG_CARD)
    hist_header_box.pack(fill="x", padx=20, pady=(10, 4))

    tk.Label(
        hist_header_box,
        text="📑 Business ERP Import History & Batch Rollback",
        font=theme.FONT_SUBTITLE,
        fg=theme.INCOME,
        bg=theme.BG_CARD
    ).pack(side="left")

    app.btn_biz_delete_batch = tk.Button(
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
        command=lambda: _remove_selected_biz_batch(app)
    )
    app.btn_biz_delete_batch.pack(side="right")

    hist_tree_box = tk.Frame(history_card, bg=theme.BG_CARD)
    hist_tree_box.pack(fill="both", expand=True, padx=20, pady=(2, 10))

    hist_cols = ("batch_id", "entity", "filename", "import_date", "valid_rows", "skipped_rows", "range", "volume")
    app.biz_hist_tree = ttk.Treeview(hist_tree_box, columns=hist_cols, show="headings", height=5)
    app.biz_hist_tree.heading("batch_id", text="Batch ID")
    app.biz_hist_tree.heading("entity", text="Entity")
    app.biz_hist_tree.heading("filename", text="Spreadsheet File")
    app.biz_hist_tree.heading("import_date", text="Imported On")
    app.biz_hist_tree.heading("valid_rows", text="Records")
    app.biz_hist_tree.heading("skipped_rows", text="Skipped")
    app.biz_hist_tree.heading("range", text="Date / Details Range")
    app.biz_hist_tree.heading("volume", text="Volume / Total")

    app.biz_hist_tree.column("batch_id", width=120, anchor="center")
    app.biz_hist_tree.column("entity", width=85, anchor="center")
    app.biz_hist_tree.column("filename", width=160, anchor="w")
    app.biz_hist_tree.column("import_date", width=110, anchor="center")
    app.biz_hist_tree.column("valid_rows", width=70, anchor="center")
    app.biz_hist_tree.column("skipped_rows", width=65, anchor="center")
    app.biz_hist_tree.column("range", width=160, anchor="center")
    app.biz_hist_tree.column("volume", width=110, anchor="e")

    app.biz_hist_tree.pack(side="left", fill="both", expand=True)

    hist_scroll = ttk.Scrollbar(hist_tree_box, orient="vertical", command=app.biz_hist_tree.yview)
    app.biz_hist_tree.configure(yscrollcommand=hist_scroll.set)
    hist_scroll.pack(side="right", fill="y")

    app.biz_hist_tree.bind("<<TreeviewSelect>>", lambda e: _on_biz_history_select(app))

    return page


# =============================================================================
# Helper Event Handlers & Dynamic UI Builders
# =============================================================================

def refresh_biz_excel_import_page(app):
    """Refreshes dynamic data on the Business Excel Import page."""
    _refresh_biz_scope_banner(app)
    _refresh_biz_import_history(app)


def _refresh_biz_scope_banner(app):
    """Updates the dataset scope banner showing Sales, Expenses, and Inventory breakdown."""
    try:
        scope = biz_service.get_business_insights_scope(app.tracker.db)
        s = scope["sales"]
        e = scope["expenses"]
        inv = scope["inventory"]

        s_range = f"{s['earliest_date']} to {s['latest_date']}" if s["total_count"] > 0 and s["earliest_date"] else "No sales recorded"
        e_range = f"{e['earliest_date']} to {e['latest_date']}" if e["total_count"] > 0 and e["earliest_date"] else "No expenses recorded"

        text = (
            f"💼 Sales: {s['total_count']} records ({s['manual_count']} Manual ✍️ | {s['imported_count']} Imported 📥) • Revenue: ₹{s['total_amount']:,.2f} (₹{s['imported_amount']:,.2f} Imported) • Dates: {s_range}\n"
            f"💸 Expenses: {e['total_count']} records ({e['manual_count']} Manual ✍️ | {e['imported_count']} Imported 📥) • Total: ₹{e['total_amount']:,.2f} (₹{e['imported_amount']:,.2f} Imported) • Dates: {e_range}\n"
            f"📦 Inventory: {inv['total_items']} items ({inv['manual_items']} Manual ✍️ | {inv['imported_items']} Imported 📥) • Stock: {inv['total_stock']} units ({inv['imported_stock']} Imported) • Active Batches: {scope['active_batches']}"
        )
        app.biz_scope_lbl.config(text=text)
    except Exception as ex:
        app.biz_scope_lbl.config(text=f"📊 Business ERP Scope: Error loading scope ({str(ex)})")


def _refresh_biz_import_history(app):
    """Refreshes the past imports table for Business ERP."""
    app.biz_hist_tree.delete(*app.biz_hist_tree.get_children())
    app.btn_biz_delete_batch.config(state="disabled")
    try:
        imports = app.tracker.get_business_excel_imports()
        for imp in imports:
            # batch_id, filename, import_date, total_rows, valid_rows, skipped_rows, date_min, date_max, total_amount, profile, notes, entity_type
            bid = imp[0]
            fname = os.path.basename(imp[1]) if imp[1] else "-"
            idate = (imp[2] or "-")[:16]
            vrows = imp[4]
            srows = imp[5]
            d_range = f"{imp[6] or '-'} to {imp[7] or '-'}"
            ent = imp[11] if len(imp) > 11 and imp[11] else "Sales"
            tot = f"₹{imp[8]:,.2f}" if ent != ENTITY_INVENTORY else f"{int(imp[8])} units"
            app.biz_hist_tree.insert("", "end", values=(bid, ent, fname, idate, vrows, srows, d_range, tot))
    except Exception as ex:
        print(f"Error loading business import history: {ex}")


def _on_biz_history_select(app):
    sel = app.biz_hist_tree.selection()
    if sel:
        app.btn_biz_delete_batch.config(state="normal")
    else:
        app.btn_biz_delete_batch.config(state="disabled")


def _choose_biz_excel_file(app):
    """Prompts user to select a local .xlsx file."""
    filepath = filedialog.askopenfilename(
        title="Select Excel File to Import into Business ERP",
        filetypes=[("Excel Workbooks (*.xlsx)", "*.xlsx")]
    )
    if not filepath:
        return

    if not filepath.lower().endswith(".xlsx"):
        messagebox.showerror(
            "Unsupported Format",
            "FINSIGHT only supports standard Excel .xlsx workbooks processed 100% locally.\n"
            "Legacy .xls (Excel 97-2003), encrypted workbooks, and binary formats (.xlsb) are unsupported."
        )
        return

    app.biz_excel_filepath = filepath
    filename = os.path.basename(filepath)
    app.lbl_biz_excel_path.config(text=f"Selected: {filename} ({filepath})", fg=theme.TEXT_PRIMARY)

    try:
        sheets = biz_service.get_sheet_names(filepath)
        app.biz_excel_sheet_cmb.config(values=sheets)
        if sheets:
            app.biz_excel_sheet_var.set(sheets[0])
            _load_selected_biz_sheet(app)
    except Exception as e:
        messagebox.showerror("Error Reading Workbook", f"Could not open Excel file:\n{str(e)}")


def _load_selected_biz_sheet(app):
    """Loads sheet data and previews first rows."""
    if not app.biz_excel_filepath or not app.biz_excel_sheet_var.get():
        return

    sheet_name = app.biz_excel_sheet_var.get()
    try:
        headers, rows = biz_service.read_excel_sheet(app.biz_excel_filepath, sheet_name)
        app.biz_excel_sheet_data = (headers, rows)
        app.lbl_biz_sheet_info.config(text=f"({len(headers)} columns detected, {len(rows)} data rows)")

        # Clear preview
        app.biz_preview_tree.delete(*app.biz_preview_tree.get_children())
        app.biz_preview_tree["columns"] = headers
        for h in headers:
            app.biz_preview_tree.heading(h, text=h)
            app.biz_preview_tree.column(h, width=110, anchor="w")

        for row in rows[:8]:
            row_vals = [str(cell) if cell is not None else "" for cell in row]
            app.biz_preview_tree.insert("", "end", values=row_vals)

        # Rebuild and populate mapping controls
        _rebuild_mapping_controls(app)
        _auto_detect_biz_mapping(app)

        # Reset validation state
        app.biz_excel_validation_results = None
        app.biz_val_tree.delete(*app.biz_val_tree.get_children())
        app.lbl_biz_val_total.config(text=f"Total Rows: {len(rows)}")
        app.lbl_biz_val_valid.config(text="✅ Valid: -")
        app.lbl_biz_val_dups.config(text="⚠️ Duplicates: -")
        app.lbl_biz_val_errors.config(text="❌ Errors: -")
        app.btn_biz_commit_import.config(state="disabled")
        app.lbl_biz_commit_info.config(text="Validate rows above to review results before importing.")

    except Exception as e:
        messagebox.showerror("Error Reading Sheet", f"Failed to read sheet '{sheet_name}':\n{str(e)}")


def _on_entity_type_change(app):
    """Called when user switches between Sales, Expenses, and Inventory."""
    _rebuild_mapping_controls(app)
    _auto_detect_biz_mapping(app)
    # Reset validation state
    app.biz_excel_validation_results = None
    app.biz_val_tree.delete(*app.biz_val_tree.get_children())
    app.lbl_biz_val_valid.config(text="✅ Valid: -")
    app.lbl_biz_val_dups.config(text="⚠️ Duplicates: -")
    app.lbl_biz_val_errors.config(text="❌ Errors: -")
    app.btn_biz_commit_import.config(state="disabled")


def _rebuild_mapping_controls(app):
    """Rebuilds mapping comboboxes for the active entity type."""
    for child in app.biz_mapping_container.winfo_children():
        child.destroy()

    app.biz_map_vars = {}
    app.biz_map_cmbs = {}

    entity = app.biz_entity_var.get()
    fields = biz_service.get_entity_fields(entity)

    headers = []
    if app.biz_excel_sheet_data:
        headers, _ = app.biz_excel_sheet_data
    col_options = ["(None)"] + headers

    grid_box = tk.Frame(app.biz_mapping_container, bg=theme.BG_CARD)
    grid_box.pack(fill="x")

    for idx, (label, field_key, is_req, desc) in enumerate(fields):
        col_frame = tk.Frame(grid_box, bg=theme.BG_CARD)
        col_frame.grid(row=0, column=idx, padx=6, sticky="ew")
        grid_box.columnconfigure(idx, weight=1)

        tk.Label(
            col_frame,
            text=label,
            font=theme.FONT_BODY_BOLD,
            fg=theme.INCOME if is_req else theme.TEXT_PRIMARY,
            bg=theme.BG_CARD
        ).pack(anchor="w")

        var = tk.StringVar(value="(None)")
        cmb = ttk.Combobox(col_frame, textvariable=var, values=col_options, state="readonly", width=18)
        cmb.pack(fill="x", pady=2)

        app.biz_map_vars[field_key] = var
        app.biz_map_cmbs[field_key] = cmb


def _auto_detect_biz_mapping(app):
    """Auto-detects mapping for the selected entity."""
    if not app.biz_excel_sheet_data:
        return
    headers, _ = app.biz_excel_sheet_data
    entity = app.biz_entity_var.get()
    detected = biz_service.detect_biz_column_mapping(headers, entity)

    for field_key, var in app.biz_map_vars.items():
        mapped_col = detected.get(field_key)
        if mapped_col and mapped_col in headers:
            var.set(mapped_col)
        else:
            var.set("(None)")


def _run_biz_validation(app):
    """Validates rows against the selected entity and mapping."""
    if not app.biz_excel_sheet_data:
        messagebox.showwarning("No Data", "Please select and load an Excel sheet first.")
        return

    headers, rows = app.biz_excel_sheet_data
    entity = app.biz_entity_var.get()
    mapping = {k: v.get() for k, v in app.biz_map_vars.items()}

    # Check required mappings
    req_fields = [f_key for _, f_key, is_req, _ in biz_service.get_entity_fields(entity) if is_req]
    for rf in req_fields:
        if mapping.get(rf) in [None, "", "(None)"]:
            messagebox.showwarning(
                "Missing Required Mapping",
                f"The '{rf}' column is required for {entity} import. Please select a spreadsheet header."
            )
            return

    try:
        # Fetch existing records for duplicate checking
        existing = []
        if entity == ENTITY_SALES:
            existing = app.tracker.db.get_all_business_sales()
        elif entity == ENTITY_EXPENSES:
            existing = app.tracker.db.get_all_business_expenses()
        elif entity == ENTITY_INVENTORY:
            existing = app.tracker.db.get_all_inventory()

        results = biz_service.validate_and_prepare_biz_rows(headers, rows, entity, mapping, existing)
        app.biz_excel_validation_results = results

        # Update KPI labels
        app.lbl_biz_val_total.config(text=f"Total Rows: {results['total_rows']}")
        app.lbl_biz_val_valid.config(text=f"✅ Valid: {results['valid_count']}")
        app.lbl_biz_val_dups.config(text=f"⚠️ Duplicates: {results['duplicate_count']}")
        app.lbl_biz_val_errors.config(text=f"❌ Errors: {results['error_count']}")

        # Populate treeview
        app.biz_val_tree.delete(*app.biz_val_tree.get_children())

        # Show errors first
        error_list = results.get("invalid_rows") or results.get("error_rows") or []
        for err in error_list:
            row_idx = err.get("row_num") or err.get("row_number") or "-"
            d_val = err.get("date") or "-"
            p_val = err.get("primary_name") or "-"
            v_val = str(err.get("amount") or "-")
            det_val = err.get("details") or "-"
            err_msg = "; ".join(err.get("errors", [])) if err.get("errors") else err.get("error", "Validation error")
            app.biz_val_tree.insert("", "end", values=(row_idx, "❌ Error", d_val, p_val, v_val, det_val, err_msg))

        # Show duplicate warning rows
        dup_list = results.get("duplicate_rows") or []
        for dup in dup_list:
            row_idx = dup.get("row_num") or dup.get("row_number") or "-"
            d_val = dup.get("date") or "-"
            p_val = dup.get("primary_name") or "-"
            if entity == ENTITY_INVENTORY:
                v_val = f"{dup.get('stock', 0)} units"
            else:
                v_val = f"₹{dup.get('amount', 0):,.2f}"
            det_val = dup.get("details") or "-"
            warn_msg = dup.get("warning") or dup.get("duplicate_reason") or "Suspected duplicate record"
            app.biz_val_tree.insert("", "end", values=(row_idx, "⚠️ Duplicate", d_val, p_val, v_val, det_val, warn_msg))

        # Show sample valid rows
        valid_list = results.get("valid_rows") or []
        for val in valid_list[:15]:
            row_idx = val.get("row_num") or val.get("row_number") or "-"
            d_val = val.get("date") or "-"
            p_val = val.get("primary_name") or "-"
            if entity == ENTITY_INVENTORY:
                v_val = f"{val.get('stock', 0)} units"
            else:
                v_val = f"₹{val.get('amount', 0):,.2f}"
            det_val = val.get("details") or "-"
            app.biz_val_tree.insert("", "end", values=(row_idx, "✅ Valid", d_val, p_val, v_val, det_val, "Ready to import"))

        total_importable = results["valid_count"]
        if not app.biz_skip_dups_var.get():
            total_importable += results["duplicate_count"]

        if total_importable > 0:
            app.btn_biz_commit_import.config(state="normal")
            app.lbl_biz_commit_info.config(
                text=f"Ready to import {total_importable} {entity} records. Click 'Confirm & Import' to proceed.",
                fg="#10b981"
            )
        else:
            app.btn_biz_commit_import.config(state="disabled")
            app.lbl_biz_commit_info.config(
                text="No valid records available to import. Please check column mappings and spreadsheet errors.",
                fg=theme.EXPENSE
            )

    except Exception as e:
        messagebox.showerror("Validation Error", f"Failed during data validation:\n{str(e)}")


def _confirm_and_import_biz(app):
    """Executes the batch import into Business ERP."""
    if not app.biz_excel_validation_results:
        return

    results = app.biz_excel_validation_results
    entity = results["entity_type"]
    skip_dups = app.biz_skip_dups_var.get()

    rows_to_import = list(results["valid_rows"])
    if not skip_dups:
        rows_to_import.extend(results["duplicate_rows"])

    if not rows_to_import:
        messagebox.showinfo("Nothing to Import", f"There are no valid {entity} records to import.")
        return

    vol_label = f"₹{results['total_volume']:,.2f}" if entity != ENTITY_INVENTORY else f"{int(results['total_volume'])} units"

    confirm_msg = (
        f"Ready to import {len(rows_to_import)} {entity} records into Business ERP.\n\n"
        f"• File: {os.path.basename(app.biz_excel_filepath)}\n"
        f"• Target Module: {entity}\n"
        f"• Total Volume: {vol_label}\n"
        f"• Suspected Duplicates: {results['duplicate_count']} ({'Skipped' if skip_dups else 'Included'})\n"
        f"• Skipped Errors: {results['error_count']}\n\n"
        f"These records will be tagged as 'Imported' with a unique batch ID.\n"
        f"You can review or remove this import batch at any time.\n\n"
        f"Proceed with import?"
    )

    if not messagebox.askyesno(f"Confirm {entity} Import", confirm_msg):
        return

    try:
        batch_summary = biz_service.execute_biz_import_batch(
            db=app.tracker.db,
            valid_rows=rows_to_import,
            entity_type=entity,
            filename=os.path.basename(app.biz_excel_filepath),
            skipped_count=results["error_count"] + (results["duplicate_count"] if skip_dups else 0)
        )

        messagebox.showinfo(
            "Import Successful",
            f"Successfully imported {batch_summary['imported_count']} {entity} records into Business ERP!\n\n"
            f"• Batch ID: {batch_summary['batch_id']}\n"
            f"• Module: {entity}\n"
            f"• Total Volume: {vol_label}\n\n"
            f"Your Business ERP Dashboard, P&L, and Ledgers have been updated."
        )

        app.biz_excel_validation_results = None
        app.btn_biz_commit_import.config(state="disabled")
        app.lbl_biz_commit_info.config(text="Import complete. Select another sheet or review past imports below.")

        refresh_biz_excel_import_page(app)
        if hasattr(app, "refresh_page_data"):
            app.refresh_page_data(app.current_page)

    except Exception as e:
        messagebox.showerror("Import Execution Error", f"Failed to commit {entity} records:\n{str(e)}")


def _remove_selected_biz_batch(app):
    """Removes the selected Business ERP import batch, rolling back its records."""
    sel = app.biz_hist_tree.selection()
    if not sel:
        return
    item = app.biz_hist_tree.item(sel[0])
    values = item.get("values", [])
    if not values:
        return

    batch_id = str(values[0])
    entity = str(values[1])
    filename = str(values[2])
    count = values[4]

    confirm_msg = (
        f"Are you sure you want to remove Business import batch '{batch_id}'?\n\n"
        f"• Module: {entity}\n"
        f"• File: {filename}\n"
        f"• Records to Delete: {count}\n\n"
        f"This will permanently delete ONLY the records created in this import batch.\n"
        f"All manual records and transactions from other imports will be strictly preserved.\n\n"
        f"Proceed with rollback?"
    )

    if not messagebox.askyesno("Confirm Import Rollback", confirm_msg):
        return

    try:
        deleted = biz_service.remove_biz_import_batch(app.tracker.db, batch_id)
        messagebox.showinfo(
            "Import Removed",
            f"Successfully rolled back batch '{batch_id}'.\nRemoved {deleted} imported {entity} records."
        )
        refresh_biz_excel_import_page(app)
        if hasattr(app, "refresh_page_data"):
            app.refresh_page_data(app.current_page)
    except Exception as e:
        messagebox.showerror("Removal Error", f"Failed to remove business import batch:\n{str(e)}")
