# ⚡ FINSIGHT - Smart Financial Intelligence & Business ERP System

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)
[![GUI](https://img.shields.io/badge/GUI-Tkinter%20%2F%20Custom%20Theme-cyan.svg)]()
[![Database](https://img.shields.io/badge/Database-SQLite3-lightgrey.svg)]()
[![Intelligence](https://img.shields.io/badge/Engine-Scikit--Learn%20%7C%20Pandas%20%7C%20NumPy-orange.svg)]()
[![Spreadsheets](https://img.shields.io/badge/Spreadsheets-OpenPyXL-brightgreen.svg)]()
[![Reports](https://img.shields.io/badge/Exports-PDF%20%7C%20Excel%20%7C%20CSV-green.svg)]()
[![Security](https://img.shields.io/badge/Security-Hardened%20%26%20Zero--Loss-purple.svg)]()

**FINSIGHT** is a comprehensive, production-ready desktop application engineered to bridge modern personal wealth management with enterprise-grade small business operations. Featuring a custom dark-mode interface, an applied machine learning **Insight Engine**, offline **Excel Ingestion & Analytics**, and real-time **Smart Shopping Comparisons**, FINSIGHT empowers individuals and entrepreneurs to track cash flow, forecast spending velocity, detect anomalies, and streamline daily ERP workflows with complete data privacy.

---

## 🌟 Key Features

### 👤 1. Personal Finance Portal

- **📊 Comprehensive Financial Dashboard**:
  - Live income, total expenses, and net balance summary cards.
  - Interactive Financial Health Gauge with real-time score indicators.
  - **Dynamic Budget Tracker**: Responsive progress bar with auto-flowing "Set Limit" controls that prevent label overlapping or truncation.
  - **Intelligent Expense Distribution Chart**: Clean, collision-free pie chart with dynamic category consolidation—automatically bundles long-tail minor categories ($< 5\%$ or beyond the top 5) into an **"Others"** slice, guaranteeing distinct, readable labels and uncrowded percentages.
  - **Income vs. Expense Trend**: 3-month trailing comparative bar visualizer.
- **⚡ Insight Engine (Applied AI & Analytics)**:
  - **📈 Monthly Spending Forecast**: Projects end-of-month expenditure, daily burn pace, and recommended daily spending allowances.
  - **⚠️ Unusual Spending Alerts**: Statistically detects category spending anomalies and spikes, presenting spacious, non-congested tabular alerts with plain-language explanations.
  - **🧬 Financial Profile Archetype**: Classifies user spending behavior (*e.g., Disciplined Wealth Builder*, *Balanced Lifestyle Spender*), scores financial health out of 100, and splits essential vs. discretionary expenses.
  - **🎯 Savings Goal Planner**: Tracks savings milestones, target completion dates, and dynamic daily/monthly savings targets.
  - **💡 Smart Category Assistant**: Suggests appropriate categories from transaction descriptions.
- **🛒 Smart Shopping Suggestions (Product Price Comparison)**:
  - Compares real-time product prices across **Amazon India** and **Flipkart**.
  - Dynamic budget evaluation: Identifies verified offers below budget as well as cheapest available options exceeding budget without bias.
  - One-click **`🛒 Visit Cheapest Deal`** button to open the lowest verified offer directly in your default browser.
  - Direct store search links and optional SerpApi Google Shopping engine integration via `SERPAPI_API_KEY`.
- **💸 Ledger & Transaction Management**:
  - Full CRUD operations with instant search, category filtering, and source tagging (`Manual` vs. `Imported`).
- **⏰ Digital Subscription Manager**:
  - Tracks recurring digital subscriptions, billing cycles, and upcoming renewal deadlines.
- **📥 Personal Excel Import & Insights (Exclusive)**:
  - **Spreadsheet Ingestion**: Select local `.xlsx` workbooks with multi-sheet discovery and tabular row previews.
  - **Column Auto-Detection & Mapping**: Intelligent matching for `Date`, `Type`, `Amount`, `Category`, and `Description`.
  - **Row-Level Validation & Duplicate Detection**: Checks for malformed dates and invalid amounts; flags intra-file duplicates as well as duplicates matching existing database records.
  - **Safe Confirmation & Rollback**: Displays pre-import validation summaries before commit. Easily review import history and roll back (delete) any imported batch with one click while preserving manual records.
  - **Scope Visibility**: Highlights active dataset coverage (manual vs. imported transaction counts and date ranges) across the portal and Insight Engine.
  - *100% Local*: All parsing and analysis occur strictly on your machine.
- **📄 Reports & Document Generation**:
  - One-click exportable **PDF Financial Health Statements** with embedded charts and summaries.
  - Data backups, one-click restores, and CSV ledger exports.

---

### 🏢 2. Business ERP Portal

- **💼 Executive Suite Dashboard**:
  - Gross Revenue, Operational Expenses (OpEx), Net Profit & Operating Margin overview.
  - Trend visualizations for sales, category expense distributions, and net profit trajectories.
- **📈 Daily Sales Ledger**:
  - Multi-payment tracking (Cash, UPI, Card, Net Banking) with instant record entry.
- **📊 Profit & Loss (P&L) Statements**:
  - Comprehensive P&L tracking, Cost of Goods Sold (COGS), and operating expense breakdowns.
- **📦 Inventory Tracker**:
  - Real-time stock counts, units sold, valuation, and automatic low-stock restock alerts ($< 5$ units).
- **👥 Customer & Vendor Management**:
  - Customer directory with receivables and outstanding balance tracking.
  - Vendor procurement ledger and purchase history tracking.
- **🧮 GST & Cash Flow Analyzers**:
  - Tax calculation supporting Indian GST slabs (5%, 12%, 18%, 28%) with CGST/SGST breakdown.
  - Inflow vs. outflow cash projection modeling.
- **👥 Employee Payroll**:
  - Salary records, department assignments, and pending vs. paid compensation tracking.
- **🧾 Automated PDF Invoice Generator**:
  - Compiles professional GST tax invoices with client details, itemized totals, and authorized signatures.
- **📥 Business ERP Excel Import & Insights (Exclusive)**:
  - **Multi-Entity Ingestion**: Import `.xlsx` spreadsheets directly into **Daily Sales**, **Business Expenses**, or **Inventory / Stock**.
  - **Entity-Specific Auto-Mapping**: Auto-detects headers tailored to the selected entity (e.g., customer, invoice, payment method, supplier, stock count).
  - **Strict Business Validation**: Verifies positive numeric values, dates, required fields, and warns about existing duplicate invoices or inventory items.
  - **Batch Traceability & Rollback**: Commits imports with unique batch IDs and source tagging (`source='Imported: <filename>'`). Allows reviewing and safely removing past batches at any time.
  - **Scope & ERP Insights**: Aggregates manual vs. imported sales, expenses, and inventory metrics across all ERP views.

---

### 🛡️ 3. Security, Data Protection & Integrity

FINSIGHT is built with security and data loss prevention at its core:

- **🔒 Forensic SQLite Disk Security**: Enforces `PRAGMA secure_delete = ON;` so deleted records and rolled-back imports are actively zeroed out on disk rather than lingering in unallocated freelist blocks. Enforces relational integrity via `PRAGMA foreign_keys = ON;`.
- **🛡️ Accidental Data Loss Prevention**:
  - **Pre-Restore Safety Backups**: Restoring a backup automatically creates a timestamped safety backup (`backups/auto_backup_before_restore_*.db`) of the active database before overwriting.
  - **Pre-Import Safety Backups**: Importing a database file captures an automated safety backup (`backups/auto_backup_before_import_*.db`) before replacement.
  - **Pre-Reset Safety Backups**: Clearing data creates `backups/auto_backup_before_reset_*.db`.
- **🔍 SQLite File Integrity Verification**: Before any database file can be restored or imported, `DatabaseManager.validate_sqlite_db_file()` verifies file size, tests the 16-byte magic header (`SQLite format 3\x00`), and executes `PRAGMA quick_check;` in read-only URI mode to reject corrupted or non-SQLite files without endangering existing data.
- **🛡️ CSV Formula Injection (DDE) Mitigation**: When exporting data to CSV, text cells beginning with formula trigger characters (`=`, `+`, `-`, `@`, `\t`, `\r`) are sanitized by prepending a single quote (`'`), neutralizing spreadsheet formula attacks in Microsoft Excel and LibreOffice Calc.
- **📄 ReportLab XML Sanitization**: Dynamic user input (transaction notes, customer names, categories, product names) is escaped with `xml.sax.saxutils.escape()` before rendering in ReportLab `Paragraph` flowables, preventing XML parsing crashes and markup injection.
- **🌐 Input Sanitization on External Services**: Search queries sent to external shopping engines and SerpApi are bounded to 100 characters and stripped of control characters/newlines to prevent parameter injection. Sensitive financial records are never transmitted externally.
- **📦 Protected Runtime Files**: Updated `.gitignore` rules prevent SQLite database files, WAL logs, backups, and user reports from accidental repository exposure.

---

## 🛠 Tech Stack

| Domain | Technologies |
|---|---|
| **Core Language** | Python 3.10+ |
| **Graphical Interface** | Tkinter, ttk, Custom Vector Widgets & Palette (`gui_theme.py`) |
| **Data & Visualization** | Matplotlib, Pandas, NumPy |
| **Spreadsheet Engine** | OpenPyXL (100% offline `.xlsx` parsing & validation) |
| **Database** | SQLite3 (Parameterized SQL, Secure Delete, Foreign Keys, Quick Checks) |
| **Financial Engine** | Scikit-learn (Linear Regression, Isolation Forest, K-Means) |
| **Web Scraping** | Requests, BeautifulSoup4, Urllib |
| **Document Generation** | ReportLab (Vector PDF Statements & Invoices), CSV |
| **Architecture** | Model-View-Controller (MVC) + Object-Oriented Domain Layer |

---

## 📂 Project Architecture

```text
├── main.py                           # Application entry point & initial seed manager
├── gui_app.py                        # Primary Tkinter GUI layout, router & controller
├── gui_theme.py                      # Modern dark-mode palette, typography & vector widgets
├── gui_ai_view.py                    # Dedicated Insight Engine dashboard view
├── gui_excel_view.py                 # Personal Finance Excel import & insights interface
├── gui_biz_excel_view.py             # Business ERP Excel import & insights interface
├── excel_import_service.py           # Personal Finance Excel parsing, validation & rollback engine
├── excel_business_import_service.py  # Business ERP multi-entity Excel import engine
├── ml_engine.py                      # Time-series regression, clustering & anomaly detection
├── price_comparison_service.py       # Independent price comparison engine (Amazon & Flipkart)
├── tracker_oop.py                    # Domain OOP classes (Transaction, Goal, Subscription, etc.)
├── db_manager.py                     # SQLite connection manager, integrity checker & migrations
├── analytics.py                      # Pandas financial metrics, summarization & trends
├── pdf_generator.py                  # Automated ReportLab PDF statement & invoice generator
├── test_price_comparison.py          # Automated tests for price comparison logic
├── test_excel_import.py              # Automated tests for personal finance Excel import
├── test_biz_excel_import.py          # Automated tests for business ERP Excel import
├── requirements.txt                  # Python package dependencies
├── run_app.bat                       # Windows launch script
└── settings.json                     # Application runtime configuration
```

---

## 🚀 Quick Start & Installation

### 1. Clone the Repository
```bash
git clone https://github.com/bhardwajkeshav208-ai/Smart--finance--management--system.git
cd Smart--finance--management--system
```

### 2. Set Up a Virtual Environment & Install Dependencies
```bash
python -m venv .venv
# On Windows:
.venv\Scripts\activate
# On macOS/Linux:
source .venv/bin/activate

pip install -r requirements.txt
```

### 3. Launch FINSIGHT
```bash
python main.py
```
*On Windows, you can also launch the app by double-clicking `run_app.bat`.*

---

## 🧪 Running Automated Tests

Run the test suites to verify functionality:

```bash
# Test Price Comparison Engine (Amazon & Flipkart logic)
python test_price_comparison.py

# Test Personal Finance Excel Import Service
python test_excel_import.py

# Test Business ERP Excel Import Service
python test_biz_excel_import.py
```

---

## 👨‍💻 Developed By

**Naman Pandey**  
<<<<<<< HEAD
- GitHub: [@naman0831pandey](https://github.com/naman0831pandey)
=======


>>>>>>> e8866e372e242f0f2fb9fca56296fcd9057c8d62
