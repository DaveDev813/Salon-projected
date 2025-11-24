import os
import pandas as pd
import numpy as np
from datetime import datetime

# --------- Define services and prices (approx from menu) ---------
hair_services = [
    ("Blowdry", 99),
    ("Haircut", 149),
    ("Iron", 199),
    ("Hair Spa", 349),
    ("Hair Botox Treatment", 499),
    ("Keratin Mask", 499),
    ("Color", 699),
    ("Rebond", 1499),
    ("Rebond Hair Botox", 1999),
    ("Rebond / Color", 2199),
    ("Rebond / Color / Brazilian", 3199),
    ("Kerabond", 1999),
    ("Brazilian", 1499),
    ("Brazilian / Botox", 1499),
    ("Brazilian / Keratin", 1499),
    ("LUXLISS Cysteine Treatment", 1999),
    ("LUXLISS Smooth Wonder Treatment", 1999),
    ("Highlights (per foil)", 100),
    ("Balayage / Color", 2499),
]

nail_services = [
    ("Manicure", 119),
    ("Pedicure", 149),
    ("Manicure / Pedicure", 269),
    ("Manicure / Pedicure / Foot Spa", 499),
    ("Foot Spa", 299),
    ("Foot Spa / Pedicure", 399),
    ("Manicure / Gel Polish", 349),
    ("Pedicure / Gel Polish", 399),
    ("Nails Extension Ordinary Polish", 399),
    ("Nails Extension Gel Polish", 599),
]

other_services = [
    ("Eyelash Extension", 399),
    ("Hair & Make up (Wedding/Debut)", 1499),
]

# --------- Settings for simulation ---------
# Example range: 15–30 services per day (you can change these)
MIN_SERVICES_PER_DAY = 7
MAX_SERVICES_PER_DAY = 15

# Combine pricelist for separate sheet
price_rows = []
for name, price in hair_services:
    price_rows.append(("Hair", name, price))
for name, price in nail_services:
    price_rows.append(("Nail", name, price))
for name, price in other_services:
    price_rows.append(("Other", name, price))

price_df = pd.DataFrame(price_rows, columns=["Category", "Service", "Price"])

# --------- Simulate a month's worth of transactions ---------
np.random.seed(42)

# Weighted probabilities so cheaper services occur more often than expensive ones
hair_prices = np.array([price for name, price in hair_services])
hair_weights = 1 / hair_prices
hair_weights = hair_weights / hair_weights.sum()

nail_prices = np.array([price for name, price in nail_services])
nail_weights = 1 / nail_prices
nail_weights = nail_weights / nail_weights.sum()

# January 1–30 as sample month
dates = pd.date_range("2025-01-01", "2025-01-30", freq="D")
# Open every day from Tuesday to Monday (7 days a week), so we include all dates
working_dates = list(dates)

transactions = []

for d in working_dates:
    # Number of services for this day, within the configured range
    services_count = np.random.randint(MIN_SERVICES_PER_DAY, MAX_SERVICES_PER_DAY + 1)

    for _ in range(services_count):
        # choose if hair or nail service for this slot
        visit_type = np.random.choice(["Hair", "Nail"], p=[0.7, 0.3])

        if visit_type == "Hair":
            # pick one hair service, biased toward more common/cheaper services
            idx = np.random.choice(len(hair_services), p=hair_weights)
            service_name, price = hair_services[idx]
            category = "Hair"
            staff = np.random.choice(["Senior Stylist", "Junior Stylist"], p=[0.6, 0.4])
        else:
            # pick one nail service, biased toward more common/cheaper services
            idx = np.random.choice(len(nail_services), p=nail_weights)
            service_name, price = nail_services[idx]
            category = "Nail"
            staff = "Nail Tech"

        transactions.append({
            "Date": d.date(),
            "Day": d.strftime("%a"),
            "Service": service_name,
            "Category": category,
            "Staff": staff,
            "Price": float(price),
        })

transactions_df = pd.DataFrame(transactions)

# Sort by date
transactions_df.sort_values("Date", inplace=True)
transactions_df.reset_index(drop=True, inplace=True)

# --------- Create Excel with multiple sheets and formulas ---------
base_dir = os.path.dirname(__file__)
file_path = os.path.join(base_dir, "salon_month_simulation_v2.xlsx")

with pd.ExcelWriter(file_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as writer:
    # Price list sheet
    price_df.to_excel(writer, sheet_name="PriceList", index=False)

    workbook = writer.book

    # Settings sheet
    settings_sheet = workbook.add_worksheet("Settings")
    settings_sheet.write("A1", "Product Cost % of Sales")
    settings_sheet.write("B1", 0.20)

    settings_sheet.write("A2", "Fixed Rent")
    settings_sheet.write("B2", 15000)

    settings_sheet.write("A3", "Electricity")
    settings_sheet.write("B3", 5000)

    settings_sheet.write("A4", "Water")
    settings_sheet.write("B4", 500)

    # Base monthly salaries (you can adjust these in Excel)
    settings_sheet.write("A6", "Senior Base Salary")
    settings_sheet.write("B6", 12000)

    settings_sheet.write("A7", "Junior Base Salary")
    settings_sheet.write("B7", 9000)

    settings_sheet.write("A8", "Nail Tech Base Salary")
    settings_sheet.write("B8", 8000)

    # Transactions sheet (no 40/60 split here – just raw sales)
    transactions_df.to_excel(writer, sheet_name="Transactions", index=False, startrow=0)
    trans_sheet = writer.sheets["Transactions"]

    # Add per-transaction commission (10%) column
    trans_sheet.write(0, 6, "Commission (10%)")  # column G header

    start_row = 1
    end_row = start_row + len(transactions_df)
    for row in range(start_row, end_row):
        excel_row = row + 1  # Excel is 1-based
        # Commission is 10% of the service price in column F
        trans_sheet.write_formula(row, 6, f"=F{excel_row}*0.10")

    # Summary sheet
    summary = workbook.add_worksheet("Summary")
    summary.write("A1", "Metric")
    summary.write("B1", "Amount (PHP)")

    # Total sales
    summary.write("A2", "Total Sales")
    summary.write_formula("B2", "=SUM(Transactions!F:F)")

    # Product cost (percentage of sales)
    summary.write("A3", "Product Cost")
    summary.write_formula("B3", "=B2*Settings!$B$1")

    # Fixed expenses
    summary.write("A4", "Fixed Salon Expenses (Rent+Utilities)")
    summary.write_formula("B4", "=Settings!$B$2+Settings!$B$3+Settings!$B$4")

    # Per-staff sales
    summary.write("A5", "Senior Sales")
    summary.write_formula("B5", '=SUMIF(Transactions!E:E,"Senior Stylist",Transactions!F:F)')

    summary.write("A6", "Junior Sales")
    summary.write_formula("B6", '=SUMIF(Transactions!E:E,"Junior Stylist",Transactions!F:F)')

    summary.write("A7", "Nail Tech Sales")
    summary.write_formula("B7", '=SUMIF(Transactions!E:E,"Nail Tech",Transactions!F:F)')

    # Base salaries (from Settings)
    summary.write("A8", "Senior Base Salary")
    summary.write_formula("B8", "=Settings!$B$6")

    summary.write("A9", "Junior Base Salary")
    summary.write_formula("B9", "=Settings!$B$7")

    summary.write("A10", "Nail Tech Base Salary")
    summary.write_formula("B10", "=Settings!$B$8")

    # Incentives
    # Senior Stylist: if sales >= 100,000 → 2,000 + 500 per extra 5,000
    summary.write("A11", "Senior Incentive")
    summary.write_formula("B11", "=IF(B5<100000,0,2000+500*INT((B5-100000)/5000))")

    # Junior Stylist: same rules as Senior
    summary.write("A12", "Junior Incentive")
    summary.write_formula("B12", "=IF(B6<100000,0,2000+500*INT((B6-100000)/5000))")

    # Nail Tech: if sales >= 45,000 → 2,000 at 45k, plus 500 per extra 5,000 above 45k
    summary.write("A13", "Nail Tech Incentive")
    summary.write_formula("B13", "=IF(B7<45000,0,1500+500*(1+INT((B7-45000)/5000)))")

    # 10% commissions on individual sales (sum of per-transaction commissions)
    summary.write("A14", "Senior Commission (10%)")
    summary.write_formula("B14", '=SUMIF(Transactions!E:E,"Senior Stylist",Transactions!G:G)')

    summary.write("A15", "Junior Commission (10%)")
    summary.write_formula("B15", '=SUMIF(Transactions!E:E,"Junior Stylist",Transactions!G:G)')

    summary.write("A16", "Nail Tech Commission (10%)")
    summary.write_formula("B16", '=SUMIF(Transactions!E:E,"Nail Tech",Transactions!G:G)')

    # Total staff pay (base + incentives + commissions)
    summary.write("A17", "Total Staff Pay (Base + Incentives + Commission)")
    summary.write_formula("B17", "=B8+B9+B10+B11+B12+B13+B14+B15+B16")

    # Total expenses
    summary.write("A18", "Total Expenses")
    summary.write_formula("B18", "=B3+B4+B17")

    # Owner net income
    summary.write("A19", "Owner Net Income")
    summary.write_formula("B19", "=B2-B18")

    # Optional: compact per-staff table at bottom
    summary.write("A21", "Staff")
    summary.write("B21", "Total Sales")
    summary.write("C21", "Base Salary")
    summary.write("D21", "Incentive")
    summary.write("E21", "Commission")
    summary.write("F21", "Total Pay")

    # Senior row
    summary.write("A22", "Senior Stylist")
    summary.write_formula("B22", "=B5")
    summary.write_formula("C22", "=B8")
    summary.write_formula("D22", "=B11")
    summary.write_formula("E22", "=B14")
    summary.write_formula("F22", "=C22+D22+E22")

    # Junior row
    summary.write("A23", "Junior Stylist")
    summary.write_formula("B23", "=B6")
    summary.write_formula("C23", "=B9")
    summary.write_formula("D23", "=B12")
    summary.write_formula("E23", "=B15")
    summary.write_formula("F23", "=C23+D23+E23")

    # Nail Tech row
    summary.write("A24", "Nail Tech")
    summary.write_formula("B24", "=B7")
    summary.write_formula("C24", "=B10")
    summary.write_formula("D24", "=B13")
    summary.write_formula("E24", "=B16")
    summary.write_formula("F24", "=C24+D24+E24")

print(f"Excel file generated at: {file_path}")