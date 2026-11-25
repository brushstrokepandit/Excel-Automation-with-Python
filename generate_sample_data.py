"""Generate six messy monthly branch-sales files (SYNTHETIC) in raw_data/.

The files deliberately contain the kinds of problems seen in real month-end data:
different header names, title rows above the header, blank rows, trailing spaces,
inconsistent case, text amounts (with currency symbols and commas), mixed date formats,
invalid dates, duplicates, negative quantities, unknown regions and amount mismatches.
"""
import random
import datetime as dt
from pathlib import Path
import pandas as pd

random.seed(2026)
RAW = Path(__file__).parent / "raw_data"
RAW.mkdir(exist_ok=True)

PRICES = {"Laptop Stand": 1800, "Wireless Mouse": 650, "USB-C Hub": 2200, "Webcam": 3100, "Keyboard": 1450}
REGIONS = ["North", "South", "East", "West"]
HEADERS = {
    1: ["Invoice No", "Date", "Region", "Product", "Qty", "Unit Price", "Amount"],
    2: ["Inv No.", "Invoice Date", "Region ", "Product Name", "Quantity", "Rate", "Total"],
    3: ["invoice_number", "date", "region", "product", "qty", "unit_price", "amount"],
    4: ["Invoice No", "Date", "Region", "Product", "Qty", "Unit Price", "Amount"],
    5: ["Invoice No", "Date", "Region", "Product", "Qty", "Unit Price", "Amount"],
    6: ["Inv No.", "Invoice Date", "Region", "Product Name", "Quantity", "Rate", "Total"],
}


def base_rows(month, n=110):
    rows = []
    for k in range(1, n + 1):
        day = random.randint(1, 28)
        prod = random.choice(list(PRICES))
        price = PRICES[prod] + random.choice([0, 0, 0, 50, -50])
        qty = random.randint(1, 20)
        rows.append([f"INV-{month:02d}{k:04d}", dt.date(2026, month, day), random.choice(REGIONS), prod, qty, price, qty * price])
    return rows


for idx, month in enumerate(range(5, 11), start=1):
    rows = base_rows(month)
    if idx == 1:  # unknown region typo
        rows[10][2] = "Nrth"; rows[40][2] = "Nrth"
    if idx == 2:  # trailing spaces + lowercase + amount mismatch
        for r in rows[:30]:
            r[2] = r[2].lower() + " "; r[3] = r[3].upper() + " "
        rows[15][6] = rows[15][6] + 5000
    if idx == 3:  # negative quantity
        rows[22][4] = -3; rows[22][6] = -3 * rows[22][5]
    if idx == 4:  # amounts as text, one missing
        for r in rows:
            r[6] = f"\u20b9{r[6]:,.2f}"
        rows[33][6] = None
    if idx == 5:  # mixed date formats and invalid dates
        for i, r in enumerate(rows):
            d = r[1]
            r[1] = d.strftime("%d/%m/%Y") if i % 2 == 0 else d.strftime("%Y-%m-%d")
        rows[12][1] = "31/02/2026"; rows[60][1] = "not a date"
    if idx == 6:  # duplicates
        rows += [rows[5][:], rows[18][:], rows[44][:], rows[70][:]]
        random.shuffle(rows)
    for pos in sorted(random.sample(range(len(rows)), 3), reverse=True):  # blank rows
        rows.insert(pos, [None] * 7)
    df = pd.DataFrame(rows, columns=HEADERS[idx])
    path = RAW / f"branch_sales_2026-{month:02d}.xlsx"
    if idx == 3:  # title rows above the header
        with pd.ExcelWriter(path, engine="openpyxl") as xw:
            df.to_excel(xw, index=False, startrow=3)
            ws = xw.sheets["Sheet1"]
            ws["A1"] = "Branch Sales Report"; ws["A2"] = f"Month: 2026-{month:02d}"; ws["A3"] = "Prepared by: Branch Accounts (synthetic)"
    else:
        df.to_excel(path, index=False)
    print("wrote", path.name, len(df), "rows")
