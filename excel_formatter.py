from __future__ import annotations

import io
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from formulas import YEARS


HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
SECTION_FILL = PatternFill("solid", fgColor="D9E1F2")
TOTAL_FILL = PatternFill("solid", fgColor="FCE4D6")
THIN = Border(
    left=Side(style="thin", color="808080"),
    right=Side(style="thin", color="808080"),
    top=Side(style="thin", color="808080"),
    bottom=Side(style="thin", color="808080"),
)


def _write_table(ws, title: str, rows: list[tuple[str, list[float] | None]], percent: bool = False) -> None:
    ws.append([title] + YEARS)
    for c in range(1, 7):
        cell = ws.cell(row=ws.max_row, column=c)
        cell.fill = HEADER_FILL
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center")
        cell.border = THIN

    for label, vals in rows:
        ws.append([label] + (vals if vals is not None else [""] * 5))
        r = ws.max_row
        fill = SECTION_FILL if vals is None else None
        if label.lower().startswith(("total", "sub-total", "net ", "mpbf", "balance status")):
            fill = TOTAL_FILL
        for c in range(1, 7):
            cell = ws.cell(row=r, column=c)
            cell.border = THIN
            if c == 1:
                cell.alignment = Alignment(horizontal="left")
                if vals is None:
                    cell.font = Font(bold=True)
            else:
                cell.number_format = "0.00%" if percent else "#,##0.00"
                cell.alignment = Alignment(horizontal="right")
            if fill:
                cell.fill = fill
                if c == 1:
                    cell.font = Font(bold=True)

    ws.column_dimensions["A"].width = 52
    for c in "BCDEFG":
        ws.column_dimensions[c].width = 16


def build_excel_report(result: dict[str, Any]) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    pl = result["pl"]
    bs = result["bs"]
    wc = result["wc"]
    ratios = result["ratios"]

    ws_pl = wb.create_sheet("PROJECTED PL")
    pl_rows = [
        ("1. Gross Income", None),
        ("(i) Sales (Net of Returns)", None),
        ("(a) Domestic Sales", pl["domestic_sales"]),
        ("(b) Export Sales", pl["export_sales"]),
        ("(c) Sub-Total", pl["sales"]),
        ("(d) Percentage Rise/Fall in Sales", pl["sales_growth_pct"]),
        ("(ii) Other Income", None),
        ("(a) Duty Drawback", [0.0] * 5),
        ("(b) Cash Assistance", [0.0] * 5),
        ("(c) Commission & Brokerage", [0.0] * 5),
        ("(d) Sub-Total", pl["other_income"]),
        ("(iii) Total Gross Income", [s + oi for s, oi in zip(pl["sales"], pl["other_income"])]),
        ("2. Cost of Sales", None),
        ("(i) Purchases", pl["purchases"]),
        ("(ii) Carriage Inward", pl["carriage_inward"]),
        ("(iii) Sub-Total", [p + c for p, c in zip(pl["purchases"], pl["carriage_inward"])]),
        ("(iv) Add Opening Stock", pl["opening_stock"]),
        ("(vi) Less Closing Stock", pl["closing_stock"]),
        ("(vii) Total Cost of Sales", pl["cost_of_sales"]),
        ("3. Operating Expenses", None),
        ("Salary", pl["salary"]),
        ("Rent", pl["rent"]),
        ("Power & Fuel", pl["power_fuel"]),
        ("Travelling & Conveyance", pl["travelling"]),
        ("Telephone & Internet", pl["telephone"]),
        ("Office Expenses", pl["office"]),
        ("Printing & Stationery", pl["printing"]),
        ("Repairs & Maintenance", pl["repairs"]),
        ("Other Operating Expenses", pl["other_operating"]),
        ("4. Operating Profit (Before Interest & Depreciation)", pl["op_profit_before_int_dep"]),
        ("5. Interest on Cash Credit / OD", pl["interest_cc"]),
        ("6. Depreciation", pl["depreciation"]),
        ("7. Operating Profit (After Interest & Depreciation)", pl["op_profit_after_int_dep"]),
        ("9. Profit Before Tax", pl["pbt"]),
        ("10. Provision for Tax", pl["tax"]),
        ("11. Net Profit", pl["net_profit"]),
        ("12. Dividend", pl["dividend"]),
        ("13. Retained Profit", pl["retained_profit"]),
    ]
    _write_table(ws_pl, "PROJECTED PL", pl_rows)

    ws_bs = wb.create_sheet("PROJECTED BS")
    bs_rows = [
        ("CURRENT LIABILITIES", None),
        ("Short Term Borrowings from Applicant Bank", bs["cc_from_applicant_bank"]),
        ("Sundry Creditors (Trade)", bs["sundry_creditors"]),
        ("Other Current Liabilities", bs["other_current_liabilities"]),
        ("Total Current Liabilities (B)", bs["total_current_liabilities"]),
        ("TERM LIABILITIES", None),
        ("Term Loans", bs["term_loans"]),
        ("Total Term Liabilities (C)", bs["term_loans"]),
        ("TOTAL OUTSIDE LIABILITIES (D)", bs["total_outside_liabilities"]),
        ("SHAREHOLDERS FUNDS", None),
        ("Share Capital", bs["share_capital"]),
        ("Surplus / Deficit in P&L", bs["surplus_pl"]),
        ("Net Worth (E)", bs["net_worth"]),
        ("TOTAL LIABILITIES (F)", bs["total_liabilities"]),
        ("CURRENT ASSETS", None),
        ("Cash & Bank Balances", bs["cash_bank"]),
        ("Receivables", bs["receivables"]),
        ("Stocks in Trade", bs["stocks"]),
        ("Other Current Assets", bs["other_current_assets"]),
        ("Total Current Assets (G)", bs["total_current_assets"]),
        ("FIXED ASSETS", None),
        ("Gross Block", bs["gross_block"]),
        ("Depreciation to Date", bs["dep_to_date"]),
        ("Net Block (H)", bs["net_block"]),
        ("TOTAL ASSETS (J)", bs["total_assets"]),
        ("NET WORKING CAPITAL", bs["net_working_capital"]),
    ]
    _write_table(ws_bs, "PROJECTED BS", bs_rows)
    ws_bs.append(["Balance Status"] + bs["balance_status"])

    ws_wc = wb.create_sheet("WORKING CAPITAL ANALYSIS")
    wc_rows = [
        ("A. CURRENT ASSETS", None),
        ("Raw Material", wc["raw_material"]),
        ("Work in Progress", wc["wip"]),
        ("Finished Goods", wc["finished_goods"]),
        ("Receivables / Sundry Debtors", wc["receivables"]),
        ("Cash & Bank", wc["cash_bank"]),
        ("Other Current Assets", wc["other_current_assets"]),
        ("Total Current Assets (A)", wc["total_current_assets"]),
        ("B. CURRENT LIABILITIES (OTHER THAN BANK)", None),
        ("Sundry Creditors", wc["sundry_creditors"]),
        ("Outstanding Expenses", wc["outstanding_expenses"]),
        ("Total Current Liabilities (B)", wc["total_current_liabilities"]),
        ("C. WORKING CAPITAL GAP (A - B)", wc["wc_gap"]),
        ("D. BORROWER CONTRIBUTION (25% of CA)", wc["borrower_contribution"]),
        ("E. MAXIMUM PERMISSIBLE BANK FINANCE (MPBF)", wc["mpbf"]),
        ("F. PROPOSED CC LIMIT", wc["proposed_cc_limit"]),
        ("Alternative - Projected Annual Turnover", wc["projected_annual_turnover"]),
        ("Working Capital Requirement @25%", wc["nayak_wc_requirement"]),
        ("Borrower Contribution @5%", wc["nayak_borrower_contribution"]),
        ("Eligible Bank Finance @20%", wc["nayak_eligible_bank_finance"]),
    ]
    _write_table(ws_wc, "WORKING CAPITAL ANALYSIS", wc_rows)

    ws_ratios = wb.create_sheet("FINANCIAL RATIOS ANALYSIS")
    ws_ratios.append(["S.No", "Particulars", "Numerator", "Denominator", *YEARS, "Bank Acceptable Benchmark", "Status"])
    for cell in ws_ratios[1]:
        cell.fill = HEADER_FILL
        cell.font = Font(bold=True, color="FFFFFF")
        cell.border = THIN
        cell.alignment = Alignment(horizontal="center")

    for row in ratios:
        ws_ratios.append([
            row["S.No"], row["Particulars"], row["Numerator"], row["Denominator"],
            row["FY1"], row["FY2"], row["FY3"], row["FY4"], row["FY5"],
            row["Bank Acceptable Benchmark"], row["Status"],
        ])
    for r in ws_ratios.iter_rows(min_row=2):
        for c in r:
            c.border = THIN
    ws_ratios.column_dimensions["B"].width = 30
    ws_ratios.column_dimensions["C"].width = 22
    ws_ratios.column_dimensions["D"].width = 22
    for c in "EFGHI":
        ws_ratios.column_dimensions[c].width = 12

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()
