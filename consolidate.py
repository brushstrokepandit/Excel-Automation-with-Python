"""Consolidate and clean monthly branch sales files into one validated Excel report.

Usage:  python consolidate.py            (reads raw_data/*.xlsx, writes output/)

Steps per file: detect header row -> standardise column names -> drop blank rows ->
clean text/date/amount fields -> validate -> separate rejects -> de-duplicate across files.
A control-total check proves that rows read = blank + rejected + duplicate + clean.
"""
from pathlib import Path
import sys
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

BASE = Path(__file__).parent
RAW, OUT = BASE / "raw_data", BASE / "output"
OUT.mkdir(exist_ok=True)

ALIASES = {
    "invoice_no": {"invoice no", "inv no.", "inv no", "invoice_number", "invoice number"},
    "date": {"date", "invoice date", "inv date"},
    "region": {"region", "branch region"},
    "product": {"product", "product name", "item"},
    "quantity": {"qty", "quantity", "units"},
    "unit_price": {"unit price", "unit_price", "rate", "price"},
    "amount": {"amount", "total", "value", "net amount"},
}
ALIAS_LOOKUP = {a: std for std, names in ALIASES.items() for a in names}
REGIONS = {"north": "North", "south": "South", "east": "East", "west": "West"}
PRODUCTS = {p.lower(): p for p in ["Laptop Stand", "Wireless Mouse", "USB-C Hub", "Webcam", "Keyboard"]}
DATE_FORMATS = ["%d/%m/%Y", "%Y-%m-%d", "%d-%b-%Y"]
TOLERANCE = 0.01  # 1% allowed difference between amount and qty x price


def find_header(raw: pd.DataFrame) -> int:
    for i in range(min(15, len(raw))):
        cells = {str(v).strip().lower() for v in raw.iloc[i] if pd.notna(v)}
        if len(cells & set(ALIAS_LOOKUP)) >= 4:
            return i
    raise ValueError("header row not found")


def parse_date(v):
    if pd.isna(v):
        return pd.NaT
    if isinstance(v, (pd.Timestamp,)) or hasattr(v, "year"):
        return pd.Timestamp(v)
    for fmt in DATE_FORMATS:
        try:
            return pd.to_datetime(str(v).strip(), format=fmt)
        except ValueError:
            continue
    return pd.NaT


def parse_amount(v):
    if pd.isna(v):
        return float("nan")
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace("\u20b9", "").replace("INR", "").replace(",", "").strip()
    return pd.to_numeric(s, errors="coerce")


def load_file(path: Path):
    raw = pd.read_excel(path, header=None)
    h = find_header(raw)
    df = raw.iloc[h + 1:].copy()
    df.columns = [ALIAS_LOOKUP.get(str(c).strip().lower(), str(c).strip().lower()) for c in raw.iloc[h]]
    df = df[[c for c in ALIASES if c in df.columns]]
    rows_read = len(df)
    blank = df.isna().all(axis=1)
    df = df[~blank].copy()
    df.insert(0, "source_file", path.name)
    return df, rows_read, int(blank.sum())


def clean_and_validate(df: pd.DataFrame):
    df = df.copy()
    df["invoice_no"] = df["invoice_no"].astype(str).str.strip()
    df["date"] = df["date"].map(parse_date)
    df["region"] = df["region"].astype(str).str.strip().str.lower().map(REGIONS)
    df["product"] = df["product"].astype(str).str.strip().str.lower().map(PRODUCTS)
    df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce")
    df["unit_price"] = pd.to_numeric(df["unit_price"], errors="coerce")
    df["amount"] = df["amount"].map(parse_amount)
    reason = pd.Series("", index=df.index)
    checks = [
        (df["invoice_no"].isin(["", "nan", "None"]), "Missing invoice number"),
        (df["date"].isna(), "Invalid or missing date"),
        (df["region"].isna(), "Unknown region"),
        (df["product"].isna(), "Unknown product"),
        (df["quantity"].isna() | (df["quantity"] <= 0), "Quantity missing or not positive"),
        (df["amount"].isna(), "Missing amount"),
    ]
    for mask, text in checks:  # first failing rule is the reason recorded
        reason = reason.where(reason != "", other=pd.Series(text, index=df.index).where(mask, ""))
    df["reject_reason"] = reason
    return df


def main():
    files = sorted(RAW.glob("*.xlsx"))
    if not files:
        sys.exit("No input files in raw_data/. Run generate_sample_data.py first.")
    frames, log = [], []
    for f in files:
        df, n_read, n_blank = load_file(f)
        df = clean_and_validate(df)
        frames.append(df)
        log.append({"Source file": f.name, "Rows read": n_read, "Blank rows skipped": n_blank})
    allrows = pd.concat(frames, ignore_index=True)
    rejected = allrows[allrows["reject_reason"] != ""].copy()
    good = allrows[allrows["reject_reason"] == ""].copy()
    dup_mask = good.duplicated(subset="invoice_no", keep="first")
    dups = good[dup_mask].copy(); dups["reject_reason"] = "Duplicate invoice number"
    clean = good[~dup_mask].copy()
    expected = (clean["quantity"] * clean["unit_price"]).round(2)
    clean["amount_check"] = ((clean["amount"] - expected).abs() <= TOLERANCE * expected.abs()).map({True: "OK", False: "CHECK - amount differs from qty x price"})
    clean["month"] = clean["date"].dt.strftime("%Y-%m")
    clean = clean.sort_values(["date", "invoice_no"]).reset_index(drop=True)

    # ---- per-file control totals ----
    for row in log:
        s = row["Source file"]
        row["Rejected (validation)"] = int((rejected["source_file"] == s).sum())
        row["Duplicates removed"] = int((dups["source_file"] == s).sum())
        row["Clean rows loaded"] = int((clean["source_file"] == s).sum())
        row["Reconciles?"] = "YES" if row["Rows read"] == row["Blank rows skipped"] + row["Rejected (validation)"] + row["Duplicates removed"] + row["Clean rows loaded"] else "NO"
    log_df = pd.DataFrame(log)
    assert (log_df["Reconciles?"] == "YES").all(), "Control totals do not reconcile"
    reasons = pd.concat([rejected, dups]).groupby(["source_file", "reject_reason"]).size().reset_index(name="Rows")
    reasons.columns = ["Source file", "Issue", "Rows"]
    mism = clean[clean["amount_check"] != "OK"].groupby("source_file").size().reset_index(name="Rows")
    if len(mism):
        mism.columns = ["Source file", "Rows"]; mism.insert(1, "Issue", "Amount differs from qty x price (kept, flagged)")
        reasons = pd.concat([reasons, mism], ignore_index=True)

    summary_month = clean.pivot_table(index="month", columns="region", values="amount", aggfunc="sum", margins=True, margins_name="Total").round(2)
    summary_prod = clean.groupby("product").agg(Invoices=("invoice_no", "count"), Units=("quantity", "sum"), Revenue=("amount", "sum")).round(2).sort_values("Revenue", ascending=False)

    out_cols = ["source_file", "invoice_no", "date", "month", "region", "product", "quantity", "unit_price", "amount", "amount_check"]
    xlsx = OUT / "consolidated_sales_report.xlsx"
    with pd.ExcelWriter(xlsx, engine="openpyxl") as xw:
        clean[out_cols].to_excel(xw, sheet_name="Clean_Data", index=False)
        summary_month.to_excel(xw, sheet_name="Summary_Month_Region")
        summary_prod.to_excel(xw, sheet_name="Summary_Product")
        log_df.to_excel(xw, sheet_name="File_Control_Totals", index=False)
        reasons.to_excel(xw, sheet_name="Data_Quality_Log", index=False)
        pd.concat([rejected, dups])[["source_file", "invoice_no", "date", "region", "product", "quantity", "amount", "reject_reason"]].to_excel(xw, sheet_name="Rejected_Rows", index=False)
    clean[out_cols].to_csv(OUT / "clean_sales.csv", index=False)
    style(xlsx, n_clean=len(clean))
    print(log_df.to_string(index=False))
    print(f"\nClean rows: {len(clean)} | Rejected: {len(rejected)} | Duplicates: {len(dups)} | Revenue: {clean['amount'].sum():,.2f}")
    print("Control totals reconcile for every file.")


def style(path, n_clean):
    wb = load_workbook(path)
    head, fill = Font(name="Arial", bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F3864")
    for ws in wb.worksheets:
        for c in ws[1]:
            c.font, c.fill, c.alignment = head, fill, Alignment(horizontal="center", wrap_text=True)
        for col in ws.columns:
            width = max(len(str(c.value)) if c.value is not None else 0 for c in col[:60])
            ws.column_dimensions[col[0].column_letter].width = min(max(12, width + 2), 48)
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.font = Font(name="Arial", size=10, bold=(c.row == ws.max_row and ws.title.startswith("Summary_Month")))
        ws.freeze_panes = "A2"
    cd = wb["Clean_Data"]
    for r in range(2, n_clean + 2):
        cd[f"C{r}"].number_format = "dd-mmm-yyyy"
        for col in "HI":
            cd[f"{col}{r}"].number_format = "#,##0.00"
    for ws in (wb["Summary_Month_Region"], wb["Summary_Product"]):
        for row in ws.iter_rows(min_row=2, min_col=2):
            for c in row:
                if isinstance(c.value, float):
                    c.number_format = "#,##0.00"
    sm = wb["Summary_Month_Region"]
    last = sm.max_row
    sm[f"A{last+2}"] = "Check: grand total here less SUM of Clean_Data amount (must be nil)"
    grand_col = sm.max_column
    from openpyxl.utils import get_column_letter as L
    sm[f"{L(grand_col)}{last+2}"] = f"=ROUND({L(grand_col)}{last}-SUM(Clean_Data!I2:I{n_clean+1}),2)"
    sm[f"A{last+2}"].font = Font(name="Arial", italic=True)
    wb.save(path)


if __name__ == "__main__":
    main()
