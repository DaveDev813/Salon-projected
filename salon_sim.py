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
# Example range: 5–7 services per day (you can change these)
MIN_SERVICES_PER_DAY = 5
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

dates = pd.date_range("2025-01-01", "2025-01-30", freq="D")
# Open every day from Tuesday to Monday (7 days a week), so we include all dates
working_dates = list(dates)

staff_members = ["Senior Stylist", "Junior Stylist", "Nail Tech"]

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
file_path = os.path.join(base_dir, "salon_month_simulation.xlsx")

with pd.ExcelWriter(file_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as writer:
    # Price list sheet
    price_df.to_excel(writer, sheet_name="PriceList", index=False)
    
    # Settings sheet
    workbook  = writer.book
    settings_sheet = workbook.add_worksheet("Settings")
    settings_sheet.write("A1", "Staff Share %")
    settings_sheet.write("B1", 0.40)
    settings_sheet.write("A2", "Owner Share %")
    settings_sheet.write("B2", 0.60)
    settings_sheet.write("A4", "Fixed Rent")
    settings_sheet.write("B4", 15000)
    settings_sheet.write("A5", "Electricity")
    settings_sheet.write("B5", 5000)
    settings_sheet.write("A6", "Water")
    settings_sheet.write("B6", 500)
    settings_sheet.write("A8", "Product Cost % of Sales")
    settings_sheet.write("B8", 0.20)
    settings_sheet.write("A11", "Nail Tech Min Monthly Wage")
    settings_sheet.write("B11", 7000)
    
    # Transactions sheet
    transactions_df.to_excel(writer, sheet_name="Transactions", index=False, startrow=0)
    trans_sheet = writer.sheets["Transactions"]
    
    # Add calculated columns for Staff Share and Owner Share
    # Assuming header row is 0, data starts row 1 in pandas but row 2 in Excel
    start_row = 1
    end_row = start_row + len(transactions_df)
    
    # Write headers for new columns
    trans_sheet.write(0, 6, "Staff Share (40%)")
    trans_sheet.write(0, 7, "Owner Share (60%)")
    
    for row in range(start_row, end_row):
        excel_row = row + 1  # Excel rows are 1-based
        # Price is column F (index 5); Staff is column E
        # If Staff is "Nail Tech", staff share is 0 and owner keeps 100% of the service price
        trans_sheet.write_formula(row, 6, f'=IF(E{excel_row}="Nail Tech",0,F{excel_row}*Settings!$B$1)')
        trans_sheet.write_formula(row, 7, f'=IF(E{excel_row}="Nail Tech",F{excel_row},F{excel_row}*Settings!$B$2)')
    
    # Summary sheet
    summary = workbook.add_worksheet("Summary")
    summary.write("A1", "Metric")
    summary.write("B1", "Amount (PHP)")
    
    # Formulas for summary
    summary.write("A2", "Total Sales")
    summary.write_formula("B2", "=SUM(Transactions!F:F)")
    
    summary.write("A3", "Total Staff Salary (Commissions)")
    summary.write_formula("B3", "=SUM(Transactions!G:G)")
    
    summary.write("A4", "Owner Gross Share")
    summary.write_formula("B4", "=SUM(Transactions!H:H)")
    
    summary.write("A5", "Product Cost")
    summary.write_formula("B5", "=B2*Settings!$B$8")
    
    summary.write("A6", "Fixed Salon Expenses (Rent+Utilities)")
    summary.write_formula("B6", "=Settings!$B$4+Settings!$B$5+Settings!$B$6")
    
    # We'll use an adjusted staff salary that respects the Nail Tech minimum wage
    summary.write("A7", "Total Expenses")
    summary.write_formula("B7", "=B14+B5+B6")
    
    summary.write("A8", "Owner Net Income")
    summary.write_formula("B8", "=B2-B7")
    
    # Nail Tech minimum wage logic
    summary.write("A11", "Nail Tech Min Monthly Wage (Setting)")
    summary.write_formula("B11", "=Settings!$B$11")
    
    summary.write("A12", "Nail Tech Commission (0%)")
    summary.write_formula("B12", "0")
    
    summary.write("A13", "Nail Tech Actual Pay")
    summary.write_formula("B13", "=MAX(B11,B12)")
    
    summary.write("A14", "Adjusted Total Staff Salary")
    summary.write_formula("B14", "=B3-B12+B13")
    
    # Per-staff breakdown (informational)
    summary.write("A17", "Staff")
    summary.write("B17", "Total Sales")
    summary.write("C17", "Commission (40%)")
    
    # Senior
    summary.write("A18", "Senior Stylist")
    summary.write_formula("B18", '=SUMIF(Transactions!E:E,"Senior Stylist",Transactions!F:F)')
    summary.write_formula("C18", "=B18*Settings!$B$1")
    
    # Junior
    summary.write("A19", "Junior Stylist")
    summary.write_formula("B19", '=SUMIF(Transactions!E:E,"Junior Stylist",Transactions!F:F)')
    summary.write_formula("C19", "=B19*Settings!$B$1")
    
    # Nail Tech
    summary.write("A20", "Nail Tech")
    summary.write_formula("B20", '=SUMIF(Transactions!E:E,"Nail Tech",Transactions!F:F)')
    # No percentage commission for Nail Tech in the breakdown (fixed wage only)
    summary.write_formula("C20", "0")

print(f"Excel file generated at: {file_path}")
