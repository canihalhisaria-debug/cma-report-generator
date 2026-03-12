from __future__ import annotations

import pandas as pd
import streamlit as st

from cma_engine import generate_projection
from excel_formatter import build_excel_report
from formulas import YEARS


def fmt_currency(value: float | str) -> str:
    if isinstance(value, str):
        return value
    return f"{value:,.2f}"


def projected_bs_view(bs: dict[str, list[float]]) -> pd.DataFrame:
    fy_cols = [f"FY {idx}" for idx in range(1, 6)]

    def section(title: str) -> dict[str, str]:
        return {"Particulars": title, **{col: "" for col in fy_cols}}

    def row(title: str, key: str) -> dict[str, str]:
        return {"Particulars": title, **{col: fmt_currency(bs[key][idx]) for idx, col in enumerate(fy_cols)}}

    def row_values(title: str, values: list[float]) -> dict[str, str]:
        return {"Particulars": title, **{col: fmt_currency(values[idx]) for idx, col in enumerate(fy_cols)}}

    rows = [
        section("CURRENT LIABILITIES"),
        row("(i) From Applicant Bank", "cc_from_applicant_bank"),
        row("(ii) From Other Banks", "cc_from_other_banks"),
        row("(iii) Of Which Bills Purchased & Discounted", "bills_purchased"),
        row("Short Term Borrowings from Others", "short_term_borrowings_others"),
        row("Sundry Creditors (Trade)", "sundry_creditors"),
        row("Advances from Customers / Deposits", "advances_customers"),
        row("Provision for Taxation", "provision_tax"),
        row("Dividend Payable", "dividend_payable"),
        row("Other Statutory Liabilities", "other_statutory"),
        row("Other Current Liabilities & Provisions", "other_current_liabilities"),
        row("Total Current Liabilities (B)", "total_current_liabilities"),
        section("TERM LIABILITIES"),
        row("Term Loans", "term_loans"),
        row("Other Term Liabilities", "other_term_liabilities"),
        row_values("Total Term Liabilities (C)", [tl + ot for tl, ot in zip(bs["term_loans"], bs["other_term_liabilities"])]),
        row("Total Outside Liabilities (D)", "total_outside_liabilities"),
        section("NET WORTH"),
        row("Share Capital", "share_capital"),
        row("General Reserve", "general_reserve"),
        row("Revaluation Reserve", "revaluation_reserve"),
        row("Other Reserves", "other_reserves"),
        row("Surplus / (Deficit) in P&L A/c", "surplus_pl"),
        row("Net Worth (E)", "net_worth"),
        row("Total Liabilities (F = D + E)", "total_liabilities"),
        section("CURRENT ASSETS"),
        row("Cash & Bank Balances", "cash_bank"),
        row("Government & Trustee Securities", "govt_securities"),
        row("Fixed Deposits with Banks", "fixed_deposits"),
        row("Receivables", "receivables"),
        row("Export Receivables", "export_receivables"),
        row("Deferred Receivables", "deferred_receivables"),
        row("Stocks-in-Trade", "stocks"),
        row("Advances to Suppliers of Merchandise", "advances_suppliers"),
        row("Advance Payment of Taxes", "advance_tax"),
        row("Other Current Assets", "other_current_assets"),
        row("Total Current Assets (G)", "total_current_assets"),
        section("FIXED ASSETS"),
        row("Gross Block", "gross_block"),
        row("Depreciation to Date", "dep_to_date"),
        row("Net Block (H)", "net_block"),
        section("OTHER NON-CURRENT ASSETS"),
        row("Other Investments", "other_investments"),
        row("Security Deposits / Tender Deposits", "security_deposits"),
        row("Other Non-Current Assets", "other_non_current_assets"),
        row("Total Other Non-Current Assets (I)", "total_other_non_current"),
        section("INTANGIBLE ASSETS"),
        row("Intangible Assets", "intangible_assets"),
        row("Total Assets (J)", "total_assets"),
        section("WORKING CAPITAL CHECK"),
        row("Net Working Capital (L)", "net_working_capital"),
        row("Diff Check Rounded (M)", "balance_diff"),
        row("Balance Status (N)", "balance_status"),
    ]
    return pd.DataFrame(rows)

st.set_page_config(page_title="CMA Projection Software | Phase-1", layout="wide")
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

    st.success("Projections generated. Engine used only CC Amount + ROI (with optional profile-tuned assumptions).")

    pl_df = pd.DataFrame(result["pl"], index=YEARS).T[[*YEARS]]
    bs_display_df = projected_bs_view(result["bs"])
    wc_df = pd.DataFrame(result["wc"], index=YEARS).T[[*YEARS]]
    ratios_df = pd.DataFrame(result["ratios"])

    t1, t2, t3, t4 = st.tabs(["Projected PL", "Projected BS", "Working Capital", "Ratios"])
    with t1:
        st.dataframe(pl_df, use_container_width=True)
    with t2:
        st.markdown("### PROJECTED BS")

        section_rows = bs_display_df["FY 1"].eq("")
        total_rows = bs_display_df["Particulars"].str.startswith("Total") | bs_display_df["Particulars"].str.contains("Net Worth \(E\)|Net Block \(H\)|Net Working Capital \(L\)")

        def highlight_rows(row: pd.Series) -> list[str]:
            if row.name in bs_display_df.index[section_rows]:
                return ["background-color: #d9e8d3; font-weight: 700;"] * len(row)
            if row.name in bs_display_df.index[total_rows]:
                return ["background-color: #f6e3d3; font-weight: 700;"] * len(row)
            return [""] * len(row)

        styled_bs = (
            bs_display_df.style
            .hide(axis="index")
            .set_properties(**{"text-align": "right"}, subset=["FY 1", "FY 2", "FY 3", "FY 4", "FY 5"])
            .set_properties(**{"text-align": "left"}, subset=["Particulars"])
            .apply(highlight_rows, axis=1)
        )
        st.dataframe(styled_bs, use_container_width=True, height=1000)
    with t3:
        st.dataframe(wc_df, use_container_width=True)
    with t4:
        st.dataframe(ratios_df, use_container_width=True)

    excel_bytes = build_excel_report(result)
    st.download_button(
        "Download Excel (PROJECTED PL / BS / RATIOS / WC)",
        data=excel_bytes,
        file_name=f"CMA_Phase1_{client_name.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Fill proposal details and click **Generate Phase-1 Projections**.")
