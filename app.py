from __future__ import annotations

import pandas as pd
import streamlit as st

from cma_engine import generate_projection
from excel_formatter import build_excel_report
from formulas import YEARS

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
            gross_profit_pct = st.number_input("Gross Profit %", min_value=5.0, max_value=60.0, value=0.0, step=0.5)
        with a2:
            debtors_days = st.number_input("Debtors Days", min_value=0.0, max_value=180.0, value=0.0, step=1.0)
            creditors_days = st.number_input("Creditors Days", min_value=0.0, max_value=180.0, value=0.0, step=1.0)
        with a3:
            inventory_days = st.number_input("Inventory Days", min_value=0.0, max_value=365.0, value=0.0, step=1.0)
            expense_loading_factor = st.number_input("Expense Loading Factor", min_value=0.3, max_value=2.0, value=0.0, step=0.05)

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
    bs_df = pd.DataFrame(result["bs"], index=YEARS).T[[*YEARS]]
    wc_df = pd.DataFrame(result["wc"], index=YEARS).T[[*YEARS]]
    ratios_df = pd.DataFrame(result["ratios"])

    t1, t2, t3, t4 = st.tabs(["Projected PL", "Projected BS", "Working Capital", "Ratios"])
    with t1:
        st.dataframe(pl_df, use_container_width=True)
    with t2:
        st.dataframe(bs_df, use_container_width=True)
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
