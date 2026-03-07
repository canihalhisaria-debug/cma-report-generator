import io
from dataclasses import dataclass

import pandas as pd
import streamlit as st

st.set_page_config(page_title="CMA Dashboard", layout="wide")

# --- Premium CSS ---
st.markdown(
    """
    <style>
        :root {
            --bg: #f3f6fb;
            --card: #ffffff;
            --text: #1b2a41;
            --muted: #64748b;
            --primary: #0f4c81;
            --accent: #2e7ad7;
            --success: #0f766e;
            --border: #e2e8f0;
        }

        .stApp {
            background: radial-gradient(circle at top left, #eef4ff 0%, var(--bg) 45%, #f8fafc 100%);
            color: var(--text);
            font-family: "Inter", "Segoe UI", sans-serif;
        }

        .top-shell {
            background: linear-gradient(135deg, #0f172a 0%, #102a43 40%, #123b74 100%);
            border-radius: 18px;
            padding: 24px 28px;
            box-shadow: 0 14px 28px rgba(15, 23, 42, 0.18);
            margin-bottom: 20px;
        }

        .business-name {
            font-size: 2.2rem;
            font-weight: 800;
            line-height: 1.1;
            background: linear-gradient(90deg, #93c5fd, #e2e8f0, #a7f3d0);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            letter-spacing: 0.4px;
            margin-bottom: 6px;
        }

        .business-address {
            color: #dbeafe;
            font-size: 0.98rem;
            letter-spacing: 0.2px;
        }

        .main-heading {
            font-size: 1.28rem;
            font-weight: 700;
            color: #0f172a;
            margin: 6px 0 18px 0;
            padding-left: 6px;
            border-left: 4px solid #2563eb;
        }

        .section-title {
            font-size: 1.08rem;
            font-weight: 700;
            color: #1e3a8a;
            margin: 12px 0 8px 2px;
            letter-spacing: 0.2px;
        }

        .card {
            background: var(--card);
            border: 1px solid var(--border);
            border-radius: 16px;
            padding: 14px 16px 8px;
            margin-bottom: 14px;
            box-shadow: 0 8px 22px rgba(15, 23, 42, 0.08);
        }

        .card-title {
            font-size: 1rem;
            font-weight: 700;
            color: #0f4c81;
            margin-bottom: 10px;
        }

        .summary-box {
            background: linear-gradient(160deg, #ffffff 0%, #f8fbff 100%);
            border: 1px solid #dbe7f7;
            border-radius: 18px;
            box-shadow: 0 10px 24px rgba(15, 23, 42, 0.08);
            padding: 16px;
            position: sticky;
            top: 1rem;
        }

        .summary-title {
            font-size: 1.06rem;
            font-weight: 800;
            color: #0f4c81;
            margin-bottom: 10px;
        }

        .metric-row {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 1px dashed #e2e8f0;
            padding: 7px 0;
            font-size: 0.95rem;
        }

        .metric-row:last-child {
            border-bottom: none;
        }

        .metric-label {
            color: #334155;
            font-weight: 500;
        }

        .metric-val {
            color: #0f172a;
            font-weight: 700;
        }

        div[data-testid="stNumberInput"] input,
        div[data-testid="stTextInput"] input,
        div[data-testid="stSelectbox"] div[data-baseweb="select"] {
            border-radius: 10px !important;
        }

        .button-shell {
            background: #ffffff;
            border-radius: 16px;
            border: 1px solid #e2e8f0;
            box-shadow: 0 8px 20px rgba(15, 23, 42, 0.06);
            padding: 16px;
            margin-top: 10px;
        }

        .stButton > button, .stDownloadButton > button {
            width: 100%;
            border-radius: 10px;
            border: 0;
            font-weight: 600;
            padding: 0.55rem 0.8rem;
        }

        .stButton > button[kind="primary"] {
            background: linear-gradient(90deg, #2563eb, #1d4ed8);
            color: white;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


@dataclass
class Summary:
    total_income: float
    material_cost: float
    ebit: float
    total_interest: float
    ebt: float
    pat: float
    cash_accrual: float
    working_capital: float


def num(label: str, key: str) -> float:
    return st.number_input(label, min_value=0.0, value=0.0, step=1000.0, key=key, format="%.2f")


# --- Header ---
header_left, header_right = st.columns([2.4, 1.4], vertical_alignment="center")
with header_left:
    st.markdown(
        """
        <div class='top-shell'>
            <div class='business-name'>Apex Capital Advisory LLP</div>
            <div class='business-address'>904, Trade Center, BKC Annex, Mumbai • Maharashtra • India</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
with header_right:
    st.markdown("### Dashboard Controls")
    party_name = st.text_input("Party Name", value="")
    constitution = st.selectbox("Constitution", ["Proprietorship", "Partnership", "Pvt Ltd"], index=0)

st.markdown("<div class='main-heading'>Actual Year Input (Locked for Projection)</div>", unsafe_allow_html=True)

main_col, summary_col = st.columns([3.5, 1.2], gap="large")

with main_col:
    # A. Profit & Loss Inputs
    st.markdown("<div class='section-title'>A. Profit &amp; Loss Inputs</div>", unsafe_allow_html=True)
    a1, a2 = st.columns(2, gap="medium")

    with a1:
        st.markdown("<div class='card'><div class='card-title'>Income</div>", unsafe_allow_html=True)
        domestic_sale = num("Domestic Sale", "domestic_sale")
        export_sale = num("Export Sale", "export_sale")
        other_income = num("Other Income", "other_income")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='card'><div class='card-title'>Operating Expenses</div>", unsafe_allow_html=True)
        salary_wages = num("Salary & Wages", "salary_wages")
        power_fuel = num("Power & Fuel", "power_fuel")
        rent_exp = num("Rent Expenses", "rent_exp")
        printing_stationery = num("Printing & Stationery", "printing_stationery")
        depreciation = num("Depreciation", "depreciation")
        other_expenditure = num("Other Expenditure", "other_expenditure")
        st.markdown("</div>", unsafe_allow_html=True)

    with a2:
        st.markdown("<div class='card'><div class='card-title'>Direct Cost</div>", unsafe_allow_html=True)
        opening_stock = num("Opening Stock", "opening_stock")
        purchases = num("Purchases", "purchases")
        carriage_outward = num("Carriage Outward", "carriage_outward")
        unloading_expenses = num("Unloading Expenses", "unloading_expenses")
        direct_expenses = num("Direct Expenses", "direct_expenses")
        closing_stock = num("Closing Stock", "closing_stock")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='card'><div class='card-title'>Finance & Tax</div>", unsafe_allow_html=True)
        interest_cc = num("Interest on CC", "interest_cc")
        interest_tl = num("Interest on TL", "interest_tl")
        tax_expense = num("Tax Expense", "tax_expense")
        repayment_tl = num("Repayment of Term Loan", "repayment_tl")
        st.markdown("</div>", unsafe_allow_html=True)

    # B. Balance Sheet Inputs
    st.markdown("<div class='section-title'>B. Balance Sheet Inputs</div>", unsafe_allow_html=True)
    b1, b2 = st.columns(2, gap="medium")
    with b1:
        st.markdown("<div class='card'><div class='card-title'>Assets</div>", unsafe_allow_html=True)
        gross_fa = num("Gross Fixed Assets", "gross_fa")
        accum_dep = num("Accumulated Depreciation", "accum_dep")
        inventory = num("Inventory", "inventory")
        debtors = num("Debtors", "debtors")
        cash_bank = num("Cash & Bank", "cash_bank")
        other_ca = num("Other Current Assets", "other_ca")
        loans_adv = num("Loans & Advances", "loans_adv")
        st.markdown("</div>", unsafe_allow_html=True)

    with b2:
        st.markdown("<div class='card'><div class='card-title'>Liabilities</div>", unsafe_allow_html=True)
        capital = num("Capital", "capital")
        reserves = num("Reserves & Surplus", "reserves")
        term_loan = num("Term Loan", "term_loan")
        bank_cc = num("Bank CC / OD", "bank_cc")
        sundry_creditors = num("Sundry Creditors", "sundry_creditors")
        other_cl = num("Other Current Liabilities", "other_cl")
        statutory_liab = num("Statutory Liabilities", "statutory_liab")
        st.markdown("</div>", unsafe_allow_html=True)

    # C. Projection Controls
    st.markdown("<div class='section-title'>C. Projection Controls</div>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3, gap="medium")
    with c1:
        st.markdown("<div class='card'><div class='card-title'>Growth</div>", unsafe_allow_html=True)
        domestic_growth = num("Domestic Sale Growth %", "domestic_growth")
        export_growth = num("Export Sale Growth %", "export_growth")
        other_income_growth = num("Other Income Growth %", "other_income_growth")
        material_growth = num("Material Cost Growth %", "material_growth")
        operating_growth = num("Operating Expense Growth %", "operating_growth")
        st.markdown("</div>", unsafe_allow_html=True)

    with c2:
        st.markdown("<div class='card'><div class='card-title'>Working Capital Days</div>", unsafe_allow_html=True)
        debtor_days = num("Debtor Days", "debtor_days")
        creditor_days = num("Creditor Days", "creditor_days")
        inventory_days = num("Inventory Days", "inventory_days")
        st.markdown("</div>", unsafe_allow_html=True)

    with c3:
        st.markdown("<div class='card'><div class='card-title'>Rates</div>", unsafe_allow_html=True)
        cc_rate = num("Interest Rate on CC", "cc_rate")
        tl_rate = num("Interest Rate on TL", "tl_rate")
        tax_rate = num("Tax Rate", "tax_rate")
        st.markdown("</div>", unsafe_allow_html=True)

    # Buttons
    st.markdown("<div class='button-shell'>", unsafe_allow_html=True)
    btn1, btn2, btn3 = st.columns(3)
    with btn1:
        if st.button("Generate CMA", type="primary"):
            st.success(f"CMA prepared for {party_name or 'selected party'} ({constitution}).")
    with btn2:
        if st.button("Reset Inputs"):
            for k in list(st.session_state.keys()):
                if k not in ["party_name", "constitution"]:
                    st.session_state[k] = 0.0
            st.experimental_rerun()

    # Build export data
    export_data = {
        "Party Name": party_name,
        "Constitution": constitution,
        "Total Income": domestic_sale + export_sale + other_income,
        "Material Cost": opening_stock + purchases + carriage_outward + unloading_expenses + direct_expenses - closing_stock,
        "EBIT": 0.0,
        "Total Interest": interest_cc + interest_tl,
        "EBT": 0.0,
        "PAT": 0.0,
    }
    export_df = pd.DataFrame([export_data])
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        export_df.to_excel(writer, index=False, sheet_name="CMA Summary")
    with btn3:
        st.download_button(
            label="Export Excel",
            data=output.getvalue(),
            file_name="cma_dashboard_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    st.markdown("</div>", unsafe_allow_html=True)

# Summary calculations
material_cost = opening_stock + purchases + carriage_outward + unloading_expenses + direct_expenses - closing_stock
operating_exp = salary_wages + power_fuel + rent_exp + printing_stationery + depreciation + other_expenditure
total_income = domestic_sale + export_sale + other_income
ebit = total_income - material_cost - operating_exp
total_interest = interest_cc + interest_tl
ebt = ebit - total_interest
tax_calc = max(0.0, ebt) * (tax_rate / 100.0) if tax_rate else tax_expense
pat = ebt - tax_calc
cash_accrual = pat + depreciation
working_capital = inventory + debtors + other_ca + loans_adv - sundry_creditors - other_cl - statutory_liab

summary = Summary(
    total_income=total_income,
    material_cost=material_cost,
    ebit=ebit,
    total_interest=total_interest,
    ebt=ebt,
    pat=pat,
    cash_accrual=cash_accrual,
    working_capital=working_capital,
)

with summary_col:
    st.markdown("<div class='summary-box'><div class='summary-title'>Calculated Summary</div>", unsafe_allow_html=True)

    def metric(label: str, value: float):
        st.markdown(
            f"<div class='metric-row'><span class='metric-label'>{label}</span><span class='metric-val'>{value:,.2f}</span></div>",
            unsafe_allow_html=True,
        )

    metric("Total Income", summary.total_income)
    metric("Material Cost", summary.material_cost)
    metric("EBIT", summary.ebit)
    metric("Total Interest", summary.total_interest)
    metric("EBT", summary.ebt)
    metric("PAT", summary.pat)
    metric("Cash Accrual", summary.cash_accrual)
    metric("Working Capital", summary.working_capital)

    st.markdown("</div>", unsafe_allow_html=True)
