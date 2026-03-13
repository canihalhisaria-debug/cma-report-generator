from __future__ import annotations

import html

import pandas as pd
import streamlit as st

from cma_engine import generate_projection
from excel_formatter import build_excel_report
from formulas import YEARS


def fmt_currency(value: float | str, percent: bool = False) -> str:
    if isinstance(value, str):
        return value
    if percent:
        return f"{value:,.2f}%"
    return f"{value:,.2f}"


def inject_table_css() -> None:
    st.markdown(
        """
        <style>
        .bank-table-wrap{
            overflow-x:auto;
            margin-top:8px;
            margin-bottom:16px;
        }
        table.bank-table{
            border-collapse:collapse;
            width:100%;
            min-width:980px;
            font-size:15px;
        }
        table.bank-table th, table.bank-table td{
            border:1px solid #8b8b8b;
            padding:8px 10px;
        }
        table.bank-table thead th{
            background:#103C72;
            color:#ffffff;
            text-align:center;
            font-weight:700;
        }
        table.bank-table td:first-child{
            text-align:left;
            min-width:360px;
        }
        table.bank-table td:not(:first-child){
            text-align:right;
        }
        table.bank-table tr.section td{
            background:#CFDCC8;
            font-weight:700;
        }
        table.bank-table tr.subsection td{
            background:#C7D1DE;
            font-weight:700;
        }
        table.bank-table tr.total td{
            background:#EAD8C5;
            font-weight:700;
        }
        table.bank-table tr.status-ok td:last-child{
            color:#166534;
            font-weight:700;
        }
        table.bank-table tr.status-alert td:last-child{
            color:#b91c1c;
            font-weight:700;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_statement_table(
    title: str,
    rows: list[tuple[str, list[float] | list[str] | None, str | None, bool]],
) -> None:
    html_parts = [
        f"<h3>{html.escape(title)}</h3>",
        '<div class="bank-table-wrap">',
        '<table class="bank-table">',
        "<thead><tr><th>Particulars</th>",
    ]
    for fy in YEARS:
        html_parts.append(f"<th>{html.escape(fy.replace('FY', 'FY '))}</th>")
    html_parts.append("</tr></thead><tbody>")

    for label, values, row_style, is_percent in rows:
        row_class = row_style or ""
        html_parts.append(f'<tr class="{row_class}">')
        html_parts.append(f"<td>{html.escape(label)}</td>")
        if values is None:
            for _ in YEARS:
                html_parts.append("<td></td>")
        else:
            for v in values:
                if isinstance(v, str):
                    html_parts.append(f"<td>{html.escape(v)}</td>")
                else:
                    html_parts.append(f"<td>{fmt_currency(v, percent=is_percent)}</td>")
        html_parts.append("</tr>")

    html_parts.append("</tbody></table></div>")
    st.markdown("".join(html_parts), unsafe_allow_html=True)


def render_ratios_table(ratios_df: pd.DataFrame) -> None:
    html_parts = [
        "<h3>Financial Ratios Analysis</h3>",
        '<div class="bank-table-wrap">',
        '<table class="bank-table">',
        "<thead><tr>",
        "<th>S.No</th>",
        "<th>Particulars</th>",
        "<th>Numerator</th>",
        "<th>Denominator</th>",
    ]
    for fy in YEARS:
        html_parts.append(f"<th>{html.escape(fy.replace('FY', 'FY '))}</th>")
    html_parts.append("<th>Bank Acceptable Benchmark</th><th>Status</th></tr></thead><tbody>")

    for _, row in ratios_df.iterrows():
        status = str(row["Status"])
        row_class = "status-ok" if status == "OK" else "status-alert"
        html_parts.append(f'<tr class="{row_class}">')
        html_parts.append(f"<td>{row['S.No']}</td>")
        html_parts.append(f"<td>{html.escape(str(row['Particulars']))}</td>")
        html_parts.append(f"<td>{html.escape(str(row['Numerator']))}</td>")
        html_parts.append(f"<td>{html.escape(str(row['Denominator']))}</td>")
        for fy in YEARS:
            val = row[fy]
            html_parts.append(f"<td>{val:,.2f}</td>")
        html_parts.append(f"<td>{html.escape(str(row['Bank Acceptable Benchmark']))}</td>")
        html_parts.append(f"<td>{html.escape(status)}</td>")
        html_parts.append("</tr>")

    html_parts.append("</tbody></table></div>")
    st.markdown("".join(html_parts), unsafe_allow_html=True)


def build_pl_rows(pl: dict) -> list[tuple[str, list[float] | None, str | None, bool]]:
    return [
        ("1. Gross Income", None, "section", False),
        ("(i) Sales (Net of Returns)", None, "subsection", False),
        ("(a) Domestic Sales", pl["domestic_sales"], None, False),
        ("(b) Export Sales", pl["export_sales"], None, False),
        ("(c) Sub-Total", pl["sales"], "total", False),
        ("(d) Percentage Rise/Fall in Sales", pl["sales_growth_pct"], None, True),
        ("(ii) Other Income", None, "subsection", False),
        ("(a) Duty Drawback", [0.0] * 5, None, False),
        ("(b) Cash Assistance", [0.0] * 5, None, False),
        ("(c) Commission & Brokerage", [0.0] * 5, None, False),
        ("(d) Sub-Total", pl["other_income"], "total", False),
        (
            "(iii) Total Gross Income",
            [s + oi for s, oi in zip(pl["sales"], pl["other_income"])],
            "section",
            False,
        ),
        ("2. Cost of Sales", None, "section", False),
        ("(i) Purchases", pl["purchases"], None, False),
        ("(ii) Carriage Inward", pl["carriage_inward"], None, False),
        (
            "(iii) Sub-Total",
            [p + c for p, c in zip(pl["purchases"], pl["carriage_inward"])],
            "total",
            False,
        ),
        ("(iv) Add Opening Stock", pl["opening_stock"], None, False),
        ("(vi) Less Closing Stock", pl["closing_stock"], None, False),
        ("(vii) Total Cost of Sales", pl["cost_of_sales"], "section", False),
        ("3. Operating Expenses", None, "section", False),
        ("Salary", pl["salary"], None, False),
        ("Rent", pl["rent"], None, False),
        ("Power & Fuel", pl["power_fuel"], None, False),
        ("Travelling & Conveyance", pl["travelling"], None, False),
        ("Telephone & Internet", pl["telephone"], None, False),
        ("Office Expenses", pl["office"], None, False),
        ("Printing & Stationery", pl["printing"], None, False),
        ("Repairs & Maintenance", pl["repairs"], None, False),
        ("Other Operating Expenses", pl["other_operating"], None, False),
        (
            "4. Operating Profit (Before Interest & Depreciation)",
            pl["op_profit_before_int_dep"],
            "section",
            False,
        ),
        ("5. Interest on Cash Credit / OD", pl["interest_cc"], "section", False),
        ("6. Depreciation", pl["depreciation"], "section", False),
        (
            "7. Operating Profit (After Interest & Depreciation)",
            pl["op_profit_after_int_dep"],
            "section",
            False,
        ),
        ("9. Profit Before Tax", pl["pbt"], "section", False),
        ("10. Provision for Tax", pl["tax"], "subsection", False),
        ("11. Net Profit", pl["net_profit"], "section", False),
        ("12. Dividend", pl["dividend"], "subsection", False),
        ("13. Retained Profit", pl["retained_profit"], "total", False),
    ]


def build_bs_rows(bs: dict) -> list[tuple[str, list[float] | list[str] | None, str | None, bool]]:
    return [
        ("CURRENT LIABILITIES", None, "section", False),
        ("(i) From Applicant Bank", bs["cc_from_applicant_bank"], None, False),
        ("(ii) From Other Banks", bs["cc_from_other_banks"], None, False),
        ("(iii) Of Which Bills Purchased & Discounted", bs["bills_purchased"], None, False),
        (
            "Short Term Borrowings from Others",
            bs["short_term_borrowings_others"],
            None,
            False,
        ),
        ("Sundry Creditors (Trade)", bs["sundry_creditors"], None, False),
        ("Advances from Customers / Deposits", bs["advances_customers"], None, False),
        ("Provision for Taxation", bs["provision_tax"], None, False),
        ("Dividend Payable", bs["dividend_payable"], None, False),
        ("Other Statutory Liabilities", bs["other_statutory"], None, False),
        ("Other Current Liabilities & Provisions", bs["other_current_liabilities"], None, False),
        ("Total Current Liabilities (B)", bs["total_current_liabilities"], "total", False),
        ("TERM LIABILITIES", None, "section", False),
        ("Term Loans", bs["term_loans"], None, False),
        ("Other Term Liabilities", bs["other_term_liabilities"], None, False),
        (
            "Total Term Liabilities (C)",
            [tl + ot for tl, ot in zip(bs["term_loans"], bs["other_term_liabilities"])],
            "total",
            False,
        ),
        ("Total Outside Liabilities (D)", bs["total_outside_liabilities"], "section", False),
        ("NET WORTH", None, "section", False),
        ("Share Capital", bs["share_capital"], None, False),
        ("General Reserve", bs["general_reserve"], None, False),
        ("Revaluation Reserve", bs["revaluation_reserve"], None, False),
        ("Other Reserves", bs["other_reserves"], None, False),
        ("Surplus / (Deficit) in P&L A/c", bs["surplus_pl"], None, False),
        ("Net Worth (E)", bs["net_worth"], "total", False),
        ("Total Liabilities (F = D + E)", bs["total_liabilities"], "section", False),
        ("CURRENT ASSETS", None, "section", False),
        ("Cash & Bank Balances", bs["cash_bank"], None, False),
        ("Government & Trustee Securities", bs["govt_securities"], None, False),
        ("Fixed Deposits with Banks", bs["fixed_deposits"], None, False),
        ("Receivables", bs["receivables"], None, False),
        ("Export Receivables", bs["export_receivables"], None, False),
        ("Deferred Receivables", bs["deferred_receivables"], None, False),
        ("Stocks-in-Trade", bs["stocks"], None, False),
        ("Advances to Suppliers of Merchandise", bs["advances_suppliers"], None, False),
        ("Advance Payment of Taxes", bs["advance_tax"], None, False),
        ("Other Current Assets", bs["other_current_assets"], None, False),
        ("Total Current Assets (G)", bs["total_current_assets"], "total", False),
        ("FIXED ASSETS", None, "section", False),
        ("Gross Block", bs["gross_block"], None, False),
        ("Depreciation to Date", bs["dep_to_date"], None, False),
        ("Net Block (H)", bs["net_block"], "total", False),
        ("OTHER NON-CURRENT ASSETS", None, "section", False),
        ("Other Investments", bs["other_investments"], None, False),
        ("Security Deposits / Tender Deposits", bs["security_deposits"], None, False),
        ("Other Non-Current Assets", bs["other_non_current_assets"], None, False),
        ("Total Other Non-Current Assets (I)", bs["total_other_non_current"], "total", False),
        ("INTANGIBLE ASSETS", None, "section", False),
        ("Intangible Assets", bs["intangible_assets"], None, False),
        ("Total Assets (J)", bs["total_assets"], "section", False),
        ("Net Working Capital (L)", bs["net_working_capital"], "total", False),
        ("Diff Check Rounded (M)", bs["balance_diff"], None, False),
        ("Balance Status (N)", bs["balance_status"], "subsection", False),
    ]


def build_wc_rows(wc: dict) -> list[tuple[str, list[float] | None, str | None, bool]]:
    return [
        ("A. CURRENT ASSETS", None, "section", False),
        ("Raw Material", wc["raw_material"], None, False),
        ("Work in Progress", wc["wip"], None, False),
        ("Finished Goods", wc["finished_goods"], None, False),
        ("Receivables / Sundry Debtors", wc["receivables"], None, False),
        ("Cash & Bank", wc["cash_bank"], None, False),
        ("Other Current Assets", wc["other_current_assets"], None, False),
        ("Total Current Assets (A)", wc["total_current_assets"], "total", False),
        ("B. CURRENT LIABILITIES (OTHER THAN BANK)", None, "section", False),
        ("Sundry Creditors", wc["sundry_creditors"], None, False),
        ("Outstanding Expenses", wc["outstanding_expenses"], None, False),
        ("Statutory Liabilities", wc["statutory_liabilities"], None, False),
        ("Total Current Liabilities (B)", wc["total_current_liabilities"], "total", False),
        ("C. WORKING CAPITAL GAP (A - B)", wc["wc_gap"], "section", False),
        ("D. BORROWER CONTRIBUTION (25% of CA)", wc["borrower_contribution"], "subsection", False),
        ("E. MAXIMUM PERMISSIBLE BANK FINANCE (MPBF)", wc["mpbf"], "total", False),
        ("F. PROPOSED CC LIMIT", wc["proposed_cc_limit"], "section", False),
        ("Alternative - Projected Annual Turnover", wc["projected_annual_turnover"], "subsection", False),
        ("Working Capital Requirement @25%", wc["nayak_wc_requirement"], None, False),
        ("Borrower Contribution @5%", wc["nayak_borrower_contribution"], None, False),
        ("Eligible Bank Finance @20%", wc["nayak_eligible_bank_finance"], "total", False),
    ]


st.set_page_config(page_title="CMA Projection Software | Phase-1", layout="wide")
inject_table_css()

st.title("CMA Projection Software (Phase-1)")
st.caption("Phase-1: Proposal profile + New Business projection engine + Excel output")

with st.form("cma_form"):
    st.subheader("SECTION A – PROPOSAL PROFILE")
    c1, c2, c3 = st.columns(3)
    with c1:
        client_name = st.text_input("Client Name", value="ABC Traders")
        firm_name = st.text_input("Firm Name", value="ABC Traders")
        constitution = st.selectbox("Constitution", ["Proprietorship", "Partnership", "LLP", "Private Limited", "Public Limited"])
        business_type = st.selectbox("Business Type", ["New Business", "Existing Business"], index=0)
    with c2:
        loan_type = st.selectbox("Loan Type", ["CC", "Term Loan", "CC + Term Loan"], index=0)
        requirement = st.selectbox("Requirement", ["New CC", "Enhancement", "Renewal", "New Term Loan"], index=0)
        industry = st.selectbox("Business / Industry", ["Trading", "Manufacturing", "Services", "Retail", "Agri", "Other"])
        bank_name = st.text_input("Bank Name", value="State Bank")
    with c3:
        branch_name = st.text_input("Branch", value="Main Branch")
        existing_limit = st.number_input("Existing Limit", min_value=0.0, value=0.0, step=100000.0)
        proposed_limit = st.number_input("Proposed Limit", min_value=0.0, value=1000000.0, step=100000.0)
        starting_fy = st.text_input("Projection Starting FY", value="2025-26")

    st.subheader("SECTION B – FINANCIAL DRIVERS")
    f1, f2 = st.columns(2)
    with f1:
        cc_amount = st.number_input("CC Amount", min_value=100000.0, value=max(proposed_limit, 1000000.0), step=100000.0)
    with f2:
        roi = st.number_input("ROI (%)", min_value=1.0, max_value=30.0, value=10.5, step=0.1)

    with st.expander("SECTION C – OPTIONAL ADVANCED INPUTS (Collapsed)", expanded=False):
        a1, a2, a3 = st.columns(3)
        with a1:
            sales_growth_pct = st.number_input("Sales Growth %", min_value=0.0, max_value=40.0, value=0.0, step=0.5)
            gross_profit_pct = st.number_input("Gross Profit %", min_value=0.0, max_value=60.0, value=0.0, step=0.5)
        with a2:
            debtors_days = st.number_input("Debtors Days", min_value=0.0, max_value=180.0, value=0.0, step=1.0)
            creditors_days = st.number_input("Creditors Days", min_value=0.0, max_value=180.0, value=0.0, step=1.0)
        with a3:
            inventory_days = st.number_input("Inventory Days", min_value=0.0, max_value=365.0, value=0.0, step=1.0)
            expense_loading_factor = st.number_input("Expense Loading Factor", min_value=0.0, max_value=2.0, value=0.0, step=0.05)

    submitted = st.form_submit_button("Generate Phase-1 Projections", type="primary")

if submitted:
    payload = {
        "client_name": client_name,
        "firm_name": firm_name,
        "constitution": constitution,
        "business_type": business_type,
        "loan_type": loan_type,
        "requirement": requirement,
        "industry": industry,
        "bank_name": bank_name,
        "branch_name": branch_name,
        "cc_amount": cc_amount,
        "roi": roi,
        "starting_fy": starting_fy,
        "sales_growth_pct": sales_growth_pct if sales_growth_pct > 0 else None,
        "gross_profit_pct": gross_profit_pct if gross_profit_pct > 0 else None,
        "debtors_days": debtors_days if debtors_days > 0 else None,
        "creditors_days": creditors_days if creditors_days > 0 else None,
        "inventory_days": inventory_days if inventory_days > 0 else None,
        "expense_loading_factor": expense_loading_factor if expense_loading_factor > 0 else None,
    }

    result = generate_projection(payload)
    st.success("Projections generated. Banker-style formatting restored.")

    t1, t2, t3, t4 = st.tabs(["Projected PL", "Projected BS", "Working Capital", "Ratios"])

    with t1:
        render_statement_table("PROJECTED PL", build_pl_rows(result["pl"]))

    with t2:
        render_statement_table("PROJECTED BS", build_bs_rows(result["bs"]))

    with t3:
        render_statement_table("WORKING CAPITAL ANALYSIS", build_wc_rows(result["wc"]))

    with t4:
        ratios_df = pd.DataFrame(result["ratios"])
        render_ratios_table(ratios_df)

    excel_bytes = build_excel_report(result)
    st.download_button(
        "Download Excel (PROJECTED PL / BS / RATIOS / WC / DEPRECIATION)",
        data=excel_bytes,
        file_name=f"CMA_Phase1_{client_name.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Fill proposal details and click Generate Phase-1 Projections.")
