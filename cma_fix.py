"""Create a fixed CMA workbook with only two user inputs (CC Amount and ROI).

This script does not overwrite source workbook. It always writes a new file.
"""

from __future__ import annotations

import argparse
import csv
from pathlib import Path

from openpyxl import Workbook


def safe_div(a: float, b: float) -> float:
    return 0.0 if b == 0 else a / b


def build_projection(cc_amount: float, roi: float, years: int = 5):
    assumptions = {
        "sales_to_cc": 4.0,
        "sales_growth": 0.10,
        "material_pct": 0.72,
        "opex_pct": 0.12,
        "dep_pct": 0.02,
        "tax_rate": 0.25,
        "debtor_days": 60,
        "creditor_days": 45,
        "inventory_days": 50,
        "term_loan_pct": 0.40,
        "repayment_pct": 0.10,
    }

    rows = []
    for i in range(years):
        sales = cc_amount * assumptions["sales_to_cc"] * ((1 + assumptions["sales_growth"]) ** i)
        material = sales * assumptions["material_pct"]
        opex = sales * assumptions["opex_pct"]
        dep = sales * assumptions["dep_pct"]
        ebit = sales - material - opex

        creditors = material * assumptions["creditor_days"] / 365
        stock = material * assumptions["inventory_days"] / 365
        debtors = sales * assumptions["debtor_days"] / 365

        cc = cc_amount if i == 0 else max(0.0, (stock + debtors - creditors) * 0.75)
        interest_cc = cc * (roi / 100)
        term_loan = cc * assumptions["term_loan_pct"]
        interest_tl = term_loan * (roi / 100)
        interest = interest_cc + interest_tl

        ebt = ebit - interest
        tax = max(0.0, ebt * assumptions["tax_rate"])
        pat = ebt - tax
        repayment = term_loan * assumptions["repayment_pct"]

        opening_stock = stock * 0.9 if i == 0 else rows[-1]["closing_stock"]
        closing_stock = stock

        current_assets = stock + debtors
        current_liabilities = cc + creditors
        wc = current_assets - current_liabilities

        rows.append(
            {
                "year": f"Year {i+1}",
                "sales": sales,
                "opening_stock": opening_stock,
                "closing_stock": closing_stock,
                "debtors": debtors,
                "creditors": creditors,
                "interest": interest,
                "ebit": ebit,
                "ebt": ebt,
                "pat": pat,
                "wc": wc,
                "current_ratio": safe_div(current_assets, current_liabilities),
                "dscr": safe_div(pat + dep + interest_tl, interest_tl + repayment),
            }
        )
    return rows


def write_workbook(output: Path, cc_amount: float, roi: float, years: int = 5) -> None:
    data = build_projection(cc_amount, roi, years)
    wb = Workbook()

    ws = wb.active
    ws.title = "Dashboard"
    ws.append(["Input", "Value"])
    ws.append(["CC Amount", cc_amount])
    ws.append(["ROI", roi])

    def add_sheet(name: str, headers: list[str], values: list[list[float]]):
        s = wb.create_sheet(name)
        s.append(["Particulars", *[r["year"] for r in data]])
        for h, vals in zip(headers, values):
            s.append([h, *vals])

    add_sheet(
        "Projected P&L",
        ["Sales", "Opening Stock", "Closing Stock", "Interest", "Projected PAT"],
        [
            [r["sales"] for r in data],
            [r["opening_stock"] for r in data],
            [r["closing_stock"] for r in data],
            [r["interest"] for r in data],
            [r["pat"] for r in data],
        ],
    )
    add_sheet(
        "Projected Balance Sheet",
        ["Stock", "Debtors", "Creditors", "Working Capital"],
        [[r["closing_stock"] for r in data], [r["debtors"] for r in data], [r["creditors"] for r in data], [r["wc"] for r in data]],
    )
    add_sheet(
        "Financial Ratios",
        ["Current Ratio", "DSCR"],
        [[r["current_ratio"] for r in data], [r["dscr"] for r in data]],
    )
    add_sheet(
        "Working Capital Analysis",
        ["Stock", "Debtors", "Creditors", "Working Capital Gap"],
        [
            [r["closing_stock"] for r in data],
            [r["debtors"] for r in data],
            [r["creditors"] for r in data],
            [r["wc"] for r in data],
        ],
    )

    wb.save(output)


def write_csv_reports(output_dir: Path, cc_amount: float, roi: float, years: int = 5) -> None:
    """Write text-based CSV outputs for environments where binary files are not supported."""
    output_dir.mkdir(parents=True, exist_ok=True)
    data = build_projection(cc_amount, roi, years)

    with (output_dir / "dashboard.csv").open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Input", "Value"])
        writer.writerow(["CC Amount", cc_amount])
        writer.writerow(["ROI", roi])

    with (output_dir / "projected_pl.csv").open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Particulars", *[r["year"] for r in data]])
        writer.writerow(["Sales", *[r["sales"] for r in data]])
        writer.writerow(["Opening Stock", *[r["opening_stock"] for r in data]])
        writer.writerow(["Closing Stock", *[r["closing_stock"] for r in data]])
        writer.writerow(["Interest", *[r["interest"] for r in data]])
        writer.writerow(["Projected PAT", *[r["pat"] for r in data]])

    with (output_dir / "projected_balance_sheet.csv").open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Particulars", *[r["year"] for r in data]])
        writer.writerow(["Stock", *[r["closing_stock"] for r in data]])
        writer.writerow(["Debtors", *[r["debtors"] for r in data]])
        writer.writerow(["Creditors", *[r["creditors"] for r in data]])
        writer.writerow(["Working Capital", *[r["wc"] for r in data]])

    with (output_dir / "financial_ratios.csv").open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Particulars", *[r["year"] for r in data]])
        writer.writerow(["Current Ratio", *[r["current_ratio"] for r in data]])
        writer.writerow(["DSCR", *[r["dscr"] for r in data]])

    with (output_dir / "working_capital_analysis.csv").open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Particulars", *[r["year"] for r in data]])
        writer.writerow(["Stock", *[r["closing_stock"] for r in data]])
        writer.writerow(["Debtors", *[r["debtors"] for r in data]])
        writer.writerow(["Creditors", *[r["creditors"] for r in data]])
        writer.writerow(["Working Capital Gap", *[r["wc"] for r in data]])


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--cc-amount", type=float, required=True)
    parser.add_argument("--roi", type=float, required=True)
    parser.add_argument("--years", type=int, default=5)
    parser.add_argument("--source", default="FINAL_CMA_REPORT_FINAL.xlsx")
    parser.add_argument("--output", default="FINAL_CMA_REPORT_FINAL_FIXED.xlsx")
    parser.add_argument(
        "--format",
        choices=["xlsx", "csv"],
        default="xlsx",
        help="Output format. Use csv for text-only outputs when binary files are not supported.",
    )
    args = parser.parse_args()

    source = Path(args.source)
    output = Path(args.output)
    if source.exists() and source.resolve() == output.resolve():
        raise ValueError("Output file must be different from source file.")

    if args.format == "csv":
        write_csv_reports(output.with_suffix(""), args.cc_amount, args.roi, args.years)
        print(f"CSV reports created in: {output.with_suffix('')}")
    else:
        write_workbook(output, args.cc_amount, args.roi, args.years)
        print(f"Fixed file created: {output}")


if __name__ == "__main__":
    main()
