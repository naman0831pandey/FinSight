import os
import sys
import sqlite3

# Ensure working directory is always the project directory
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Automatically detect and inject local virtual environment packages if running outside venv
_venv_site = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".venv", "Lib", "site-packages")
if os.path.exists(_venv_site) and _venv_site not in sys.path:
    sys.path.insert(0, _venv_site)

from datetime import datetime, timedelta, date
from db_manager import DatabaseManager
from gui_app import SmartBudgetTrackerApp

def seed_sample_data(db, force=False):
    """
    Seeds database with realistic sample transactions, saving goals, and
    subscriptions for the current and recent months so the UI displays graphs,
    metrics, and month-over-month trend insights out-of-the-box.
    """
    import json
    import os
    
    seeded_flag = False
    try:
        if os.path.exists("settings.json"):
            with open("settings.json", "r") as f:
                settings = json.load(f)
                seeded_flag = settings.get("seeded", False)
    except Exception:
        pass
        
    if seeded_flag and not force:
        return

    conn = db.get_connection()
    cursor = conn.cursor()

    if force:
        # Clear existing tables to ensure clean, coherent multi-month demo dataset
        cursor.execute("DELETE FROM transactions")
        cursor.execute("DELETE FROM saving_goals")
        cursor.execute("DELETE FROM subscriptions")
        conn.commit()

    # Check if personal transactions are empty
    cursor.execute("SELECT COUNT(*) FROM transactions WHERE profile = 'Personal'")
    personal_count = cursor.fetchone()[0]

    # Check if business transactions are empty
    cursor.execute("SELECT COUNT(*) FROM transactions WHERE profile = 'Business'")
    business_count = cursor.fetchone()[0]

    today = date.today()
    curr_yr = today.year
    curr_mo = today.month

    # 1 month ago
    prev_yr, prev_mo = (curr_yr - 1, 12) if curr_mo == 1 else (curr_yr, curr_mo - 1)
    # 2 months ago
    prev2_yr, prev2_mo = (prev_yr - 1, 12) if prev_mo == 1 else (prev_yr, prev_mo - 1)

    to_insert = []

    if personal_count == 0 or force:
        print("Seeding rich multi-month sample personal transactions...")
        # 1. Current Month Transactions
        to_insert.extend([
            ("Income", 85000.0, "Salary", f"{curr_yr}-{curr_mo:02d}-01", "Monthly Corporate Salary", "Personal"),
            ("Income", 18000.0, "Freelance", f"{curr_yr}-{curr_mo:02d}-08", "Full-Stack Web Development Gig", "Personal"),
            
            ("Expense", 16000.0, "Rent", f"{curr_yr}-{curr_mo:02d}-02", "Apartment monthly rent payment", "Personal"),
            ("Expense", 4800.0, "Food", f"{curr_yr}-{curr_mo:02d}-03", "Weekly organic grocery shopping", "Personal"),
            ("Expense", 3400.0, "Bills & Utilities", f"{curr_yr}-{curr_mo:02d}-05", "Electricity and optical fiber bill", "Personal"),
            ("Expense", 649.0, "Bills & Utilities", f"{curr_yr}-{curr_mo:02d}-07", "Netflix Premium monthly plan", "Personal"),
            ("Expense", 2500.0, "Entertainment", f"{curr_yr}-{curr_mo:02d}-08", "Concert tickets and beverages", "Personal"),
            ("Expense", 5600.0, "Shopping", f"{curr_yr}-{curr_mo:02d}-09", "Zara summer jackets and clothes", "Personal"),
            ("Expense", 3200.0, "Travel", f"{curr_yr}-{curr_mo:02d}-09", "Petrol and highway tolls", "Personal"),
            ("Expense", 1500.0, "Health & Fitness", f"{curr_yr}-{curr_mo:02d}-10", "Gym monthly access pass", "Personal"),
            ("Expense", 2400.0, "Food", f"{curr_yr}-{curr_mo:02d}-10", "Dinner with team colleagues", "Personal"),
            ("Expense", 900.0, "Others", f"{curr_yr}-{curr_mo:02d}-10", "Dry cleaning & laundry", "Personal"),
        ])
        
        # 2. Previous Month Transactions (for MoM trend comparisons)
        to_insert.extend([
            ("Income", 85000.0, "Salary", f"{prev_yr}-{prev_mo:02d}-01", "Monthly Corporate Salary", "Personal"),
            ("Income", 12000.0, "Freelance", f"{prev_yr}-{prev_mo:02d}-15", "Consulting Gig", "Personal"),
            
            ("Expense", 16000.0, "Rent", f"{prev_yr}-{prev_mo:02d}-02", "Apartment rent", "Personal"),
            ("Expense", 5200.0, "Food", f"{prev_yr}-{prev_mo:02d}-05", "Groceries and supermarkets", "Personal"),
            ("Expense", 1800.0, "Food", f"{prev_yr}-{prev_mo:02d}-20", "Weekend family dining", "Personal"),
            ("Expense", 6900.0, "Shopping", f"{prev_yr}-{prev_mo:02d}-15", "Apparel & electronics", "Personal"),
            ("Expense", 3200.0, "Bills & Utilities", f"{prev_yr}-{prev_mo:02d}-07", "Electricity & broadband", "Personal"),
            ("Expense", 649.0, "Bills & Utilities", f"{prev_yr}-{prev_mo:02d}-10", "Netflix", "Personal"),
            ("Expense", 4100.0, "Travel", f"{prev_yr}-{prev_mo:02d}-18", "Weekend roadtrip fuel & tolls", "Personal"),
            ("Expense", 3000.0, "Entertainment", f"{prev_yr}-{prev_mo:02d}-12", "Movie & event passes", "Personal"),
            ("Expense", 1500.0, "Health & Fitness", f"{prev_yr}-{prev_mo:02d}-21", "Gym pass", "Personal"),
            ("Expense", 1200.0, "Others", f"{prev_yr}-{prev_mo:02d}-25", "Home maintenance", "Personal"),
        ])

        # 3. Two Months Ago Transactions (for extended trend trajectory)
        to_insert.extend([
            ("Income", 85000.0, "Salary", f"{prev2_yr}-{prev2_mo:02d}-01", "Monthly Corporate Salary", "Personal"),
            ("Expense", 16000.0, "Rent", f"{prev2_yr}-{prev2_mo:02d}-02", "Apartment rent", "Personal"),
            ("Expense", 6400.0, "Food", f"{prev2_yr}-{prev2_mo:02d}-06", "Monthly pantry stock", "Personal"),
            ("Expense", 8200.0, "Shopping", f"{prev2_yr}-{prev2_mo:02d}-14", "Home decor & appliances", "Personal"),
            ("Expense", 3600.0, "Bills & Utilities", f"{prev2_yr}-{prev2_mo:02d}-08", "Utilities & phone", "Personal"),
            ("Expense", 2800.0, "Travel", f"{prev2_yr}-{prev2_mo:02d}-19", "City commute fuel", "Personal"),
            ("Expense", 2200.0, "Entertainment", f"{prev2_yr}-{prev2_mo:02d}-22", "Gaming & subscriptions", "Personal"),
            ("Expense", 1500.0, "Health & Fitness", f"{prev2_yr}-{prev2_mo:02d}-21", "Gym pass", "Personal"),
        ])

    if business_count == 0 or force:
        print("Seeding sample business transactions...")
        to_insert.extend([
            # Business Revenues (Incomes)
            ("Income", 150000.0, "Sales", f"{curr_yr}-{curr_mo:02d}-02", "Product Sales Revenue", "Business"),
            ("Income", 45000.0, "Services", f"{curr_yr}-{curr_mo:02d}-08", "Consulting Fee Client A", "Business"),
            
            # Business Expenses
            ("Expense", 35000.0, "Inventory", f"{curr_yr}-{curr_mo:02d}-03", "Bulk warehouse packaging purchase", "Business"),
            ("Expense", 8500.0, "Shipping & Logistics", f"{curr_yr}-{curr_mo:02d}-06", "FedEx delivery services", "Business"),
            ("Expense", 12000.0, "Marketing & Ads", f"{curr_yr}-{curr_mo:02d}-05", "Facebook & Instagram Ads campaign", "Business"),
            ("Expense", 18000.0, "Office Rent", f"{curr_yr}-{curr_mo:02d}-01", "Co-working space desk rental", "Business"),
            ("Expense", 25000.0, "Salaries & Payroll", f"{curr_yr}-{curr_mo:02d}-09", "Freelance assistant monthly salary", "Business"),
            ("Expense", 4500.0, "Software & Tools", f"{curr_yr}-{curr_mo:02d}-08", "Google Workspace & Slack tools subscription", "Business")
        ])

    if to_insert:
        cursor.executemany("""
            INSERT INTO transactions (type, amount, category, date, notes, profile)
            VALUES (?, ?, ?, ?, ?, ?)
        """, to_insert)
        conn.commit()

    # Check if goals table is empty
    cursor.execute("SELECT COUNT(*) FROM saving_goals")
    goal_count = cursor.fetchone()[0]

    if goal_count == 0 or force:
        print("Seeding sample savings goals...")
        today = date.today()
        date_6mo = (today + timedelta(days=180)).strftime("%Y-%m-%d")
        date_12mo = (today + timedelta(days=365)).strftime("%Y-%m-%d")
        date_3mo = (today + timedelta(days=90)).strftime("%Y-%m-%d")

        sample_goals = [
            ("MacBook Pro 16", 140000.0, 65000.0, date_6mo),
            ("Europe Summer Tour", 150000.0, 55000.0, date_12mo),
            ("Emergency Safety Reserve", 60000.0, 40000.0, date_3mo)
        ]
        cursor.executemany("""
            INSERT INTO saving_goals (name, target_amount, current_savings, target_date)
            VALUES (?, ?, ?, ?)
        """, sample_goals)
        conn.commit()

    # Check if subscriptions table is empty
    cursor.execute("SELECT COUNT(*) FROM subscriptions")
    sub_count = cursor.fetchone()[0]

    if sub_count == 0 or force:
        print("Seeding sample subscriptions...")
        today = date.today()
        date_netflix = (today + timedelta(days=4)).strftime("%Y-%m-%d")
        date_spotify = (today + timedelta(days=12)).strftime("%Y-%m-%d")
        date_prime = (today + timedelta(days=45)).strftime("%Y-%m-%d")

        sample_subs = [
            ("Netflix Premium", 649.0, "Monthly", date_netflix),
            ("Spotify Premium Duo", 179.0, "Monthly", date_spotify),
            ("Amazon Prime Membership", 1499.0, "Yearly", date_prime)
        ]
        cursor.executemany("""
            INSERT INTO subscriptions (name, amount, billing_cycle, renewal_date)
            VALUES (?, ?, ?, ?)
        """, sample_subs)
        conn.commit()

    conn.close()

    # Save initialized flag to settings
    try:
        settings = {}
        if os.path.exists("settings.json"):
            with open("settings.json", "r") as f:
                settings = json.load(f)
        settings["seeded"] = True
        with open("settings.json", "w") as f:
            json.dump(settings, f)
    except Exception:
        pass


def main():
    print("=========================================================")
    print("Smart Budget Tracker & Personal Finance Assistant Booter")
    print("=========================================================")
    
    # 1. Initialize DB Manager
    try:
        db = DatabaseManager()
        # Seed Mock Data to display beautiful graphs on startup
        seed_sample_data(db)
    except Exception as e:
        print(f"CRITICAL ERROR: Failed to instantiate database: {e}")
        sys.exit(1)

    # 2. Boot Tkinter Dashboard App
    print("Loading UI Interface Frame...")
    try:
        app = SmartBudgetTrackerApp()
        print("Application window loaded successfully. Main loop running.")
        app.mainloop()
    except Exception as e:
        print(f"CRITICAL ERROR: Failed to load Tkinter UI engine: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()