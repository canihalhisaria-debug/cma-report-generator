from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from formulas import YEARS, benchmark_status, bounded_variation, growth_series, pct, profile_seed, rolling_opening_closing, safe_div


@dataclass
class ProposalInput:
    client_name: str
    firm_name: str
    constitution: str
    business_type: str
    loan_type: str
    requirement: str
    industry: str
    bank_name: str
    branch_name: str
    cc_amount: float
    roi: float
    starting_fy: str
    sales_growth_pct: float | None = None
    gross_profit_pct: float | None = None
    debtors_days: float | None = None
    creditors_days: float | None = None
    inventory_days: float | None = None
    expense_loading_factor: float | None = None


class CMAEngine:
    def __init__(self, proposal: ProposalInput) -> None:
        self.p = proposal

    def run(self) -> dict[str, Any]:
        assumptions = self._build_assumptions()
        sales = self._project_sales(assumptions)
        pl = self._build_pl(sales, assumptions)
        bs = self._build_bs(pl, assumptions)
        wc = self._build_wc(bs, sales)
        ratios = self._build_ratios(pl, bs, wc)
        return {
            "meta": self.p.__dict__,
            "assumptions": assumptions,
            "pl": pl,
            "bs": bs,
            "wc": wc,
            "ratios": ratios,
        }

    def _build_assumptions(self) -> dict[str, float]:
        seed = profile_seed(
            self.p.client_name,
            self.p.constitution,
            self.p.industry,
            self.p.requirement,
            self.p.business_type,
            self.p.loan_type,
        )
        industry_growth = {
            "Trading": 0.11,
            "Manufacturing": 0.13,
            "Services": 0.10,
            "Retail": 0.12,
            "Agri": 0.09,
            "Other": 0.10,
        }.get(self.p.industry, 0.10)
        constitution_mod = {
            "Proprietorship": -0.005,
            "Partnership": 0.0,
            "LLP": 0.005,
            "Private Limited": 0.01,
            "Public Limited": 0.012,
        }.get(self.p.constitution, 0.0)
        requirement_mod = {
            "New CC": 0.0,
            "Enhancement": 0.01,
            "Renewal": -0.005,
            "New Term Loan": 0.006,
        }.get(self.p.requirement, 0.0)

        growth = self.p.sales_growth_pct / 100 if self.p.sales_growth_pct is not None else industry_growth + constitution_mod + requirement_mod
        gp_ratio = self.p.gross_profit_pct / 100 if self.p.gross_profit_pct is not None else bounded_variation(seed, 0.23, 0.33)
        debtors_days = self.p.debtors_days if self.p.debtors_days is not None else bounded_variation(seed // 3, 35, 65)
        creditors_days = self.p.creditors_days if self.p.creditors_days is not None else bounded_variation(seed // 5, 25, 50)
        inventory_days = self.p.inventory_days if self.p.inventory_days is not None else bounded_variation(seed // 7, 40, 90)
        expense_load = self.p.expense_loading_factor if self.p.expense_loading_factor is not None else bounded_variation(seed // 11, 0.75, 1.1)
        return {
            "seed": seed,
            "growth": growth,
            "gp_ratio": gp_ratio,
            "debtors_days": debtors_days,
            "creditors_days": creditors_days,
            "inventory_days": inventory_days,
            "expense_load": expense_load,
        }

    def _project_sales(self, a: dict[str, float]) -> list[float]:
        wc_cycle_days = max(a["inventory_days"] + a["debtors_days"] - a["creditors_days"], 40)
        base_utilization = bounded_variation(int(a["seed"]), 0.52, 0.72)
        base_sales = self.p.cc_amount * (360 / wc_cycle_days) * base_utilization
        return growth_series(base_sales, a["growth"], years=5)

    def _build_pl(self, sales: list[float], a: dict[str, float]) -> dict[str, Any]:
        export_share = bounded_variation(a["seed"] // 13, 0.0, 0.18)
        other_income_pct = bounded_variation(a["seed"] // 17, 0.0, 0.015)
        tax_rate = 0.25

        closing_stock = [s * (a["inventory_days"] / 360) for s in sales]
        opening_stock, closing_stock = rolling_opening_closing(closing_stock)
        debtors = [s * (a["debtors_days"] / 360) for s in sales]
        creditors = [s * (a["creditors_days"] / 360) for s in sales]

        purchases = [max(s * (1 - a["gp_ratio"]) + cs - os, 0) for s, cs, os in zip(sales, closing_stock, opening_stock)]
        carriage = [p * 0.008 for p in purchases]
        commission_purchase = [p * 0.0 for p in purchases]
        brokerage_purchase = [p * 0.0 for p in purchases]
        total_cogs = [
            p + c + cp + bp + os - cs
            for p, c, cp, bp, os, cs in zip(
                purchases,
                carriage,
                commission_purchase,
                brokerage_purchase,
                opening_stock,
                closing_stock,
            )
        ]

        op_base = [s * 0.18 * a["expense_load"] for s in sales]
        min_op_margin = max(0.08, pct(self.p.roi) + 0.03)
        op_exp_total = [min(ob, max(s * (1 - a["gp_ratio"] - min_op_margin), s * 0.08)) for ob, s in zip(op_base, sales)]
        salary = [v * 0.28 for v in op_base]
        rent = [v * 0.15 for v in op_base]
        power = [v * 0.03 for v in op_base]
        travel = [v * 0.02 for v in op_base]
        tel = [v * 0.015 for v in op_base]
        office = [v * 0.05 for v in op_base]
        print_stationery = [v * 0.008 for v in op_base]
        repairs = [v * 0.015 for v in op_base]
        other_op = [max(v - (sa + re + po + tr + te + of + ps + rp), 0.0) for v, sa, re, po, tr, te, of, ps, rp in zip(op_exp_total, salary, rent, power, travel, tel, office, print_stationery, repairs)]

        pbdt = [s - c - o for s, c, o in zip(sales, total_cogs, op_exp_total)]
        interest_cc = [self.p.cc_amount * pct(self.p.roi) for _ in YEARS]
        dep_schedule = self._build_depreciation_schedule(sales)
        depreciation = dep_schedule["depreciation"]
        pat_before_other = [p - i - d for p, i, d in zip(pbdt, interest_cc, depreciation)]
        other_income = [s * other_income_pct for s in sales]
        pbt = [x + oi for x, oi in zip(pat_before_other, other_income)]
        tax = [max(v * tax_rate, 0) for v in pbt]
        net_profit = [v - t for v, t in zip(pbt, tax)]
        dividend = [0.0 for _ in YEARS]
        retained = [np - dv for np, dv in zip(net_profit, dividend)]

        return {
            "sales": sales,
            "domestic_sales": [s * (1 - export_share) for s in sales],
            "export_sales": [s * export_share for s in sales],
            "sales_growth_pct": [0.0] + [safe_div(sales[i] - sales[i - 1], sales[i - 1]) * 100 for i in range(1, 5)],
            "other_income": other_income,
            "purchases": purchases,
            "carriage_inward": carriage,
            "commission_on_purchases": commission_purchase,
            "brokerage_on_purchases": brokerage_purchase,
            "opening_stock": opening_stock,
            "closing_stock": closing_stock,
            "cost_of_sales": total_cogs,
            "salary": salary,
            "rent": rent,
            "power_fuel": power,
            "travelling": travel,
            "telephone": tel,
            "office": office,
            "printing": print_stationery,
            "repairs": repairs,
            "other_operating": other_op,
            "operating_expenses": op_exp_total,
            "op_profit_before_int_dep": pbdt,
            "interest_cc": interest_cc,
            "interest_tl": [0.0] * 5,
            "interest_unsecured": [0.0] * 5,
            "depreciation": depreciation,
            "dep_opening_gross_block": dep_schedule["opening_gross_block"],
            "dep_additions": dep_schedule["additions"],
            "dep_gross_block": dep_schedule["gross_block"],
            "dep_opening_accumulated": dep_schedule["opening_accumulated_dep"],
            "dep_accumulated": dep_schedule["accumulated_dep"],
            "dep_net_block": dep_schedule["net_block"],
            "op_profit_after_int_dep": pat_before_other,
            "pbt": pbt,
            "tax": tax,
            "net_profit": net_profit,
            "dividend": dividend,
            "retained_profit": retained,
            "receivables": debtors,
            "creditors": creditors,
        }

    def _build_depreciation_schedule(self, sales: list[float]) -> dict[str, list[float]]:
        dep_rate = 0.10
        opening_gross = self.p.cc_amount * 0.22

        opening_gross_block: list[float] = []
        additions: list[float] = []
        gross_block: list[float] = []
        opening_acc_dep: list[float] = []
        dep_charge: list[float] = []
        accumulated_dep: list[float] = []
        net_block: list[float] = []

        prev_gross = opening_gross
        prev_acc_dep = opening_gross * 0.08
        for i in range(5):
            addition = (opening_gross * 0.08) if i == 0 else max((sales[i] - sales[i - 1]) * 0.06, 0.0)
            current_gross = prev_gross + addition
            current_dep = (prev_gross * dep_rate) + (addition * dep_rate * 0.5)
            current_acc_dep = prev_acc_dep + current_dep
            current_net = max(current_gross - current_acc_dep, 0.0)

            opening_gross_block.append(prev_gross)
            additions.append(addition)
            gross_block.append(current_gross)
            opening_acc_dep.append(prev_acc_dep)
            dep_charge.append(current_dep)
            accumulated_dep.append(current_acc_dep)
            net_block.append(current_net)

            prev_gross = current_gross
            prev_acc_dep = current_acc_dep

        return {
            "opening_gross_block": opening_gross_block,
            "additions": additions,
            "gross_block": gross_block,
            "opening_accumulated_dep": opening_acc_dep,
            "depreciation": dep_charge,
            "accumulated_dep": accumulated_dep,
            "net_block": net_block,
        }

    def _build_bs(self, pl: dict[str, Any], a: dict[str, float]) -> dict[str, Any]:
        share_capital = [self.p.cc_amount * 0.3] * 5
        surplus = []
        run = 0.0
        for v in pl["retained_profit"]:
            run += v
            surplus.append(run)
        net_worth = [sc + sp for sc, sp in zip(share_capital, surplus)]

        cc_bank = [self.p.cc_amount] * 5
        sundry_creditors = pl["creditors"]
        other_cl = [v * 0.05 for v in sundry_creditors]
        provision_tax = pl["tax"]
        dividend_payable = pl["dividend"]

        receivables = pl["receivables"]
        inventory = pl["closing_stock"]
        gross_block = pl["dep_gross_block"]
        dep_to_date = pl["dep_accumulated"]
        net_block = pl["dep_net_block"]

        cash_bank: list[float] = []
        short_term_borrowings_others: list[float] = []
        other_term_liabilities: list[float] = []
        other_ca: list[float] = []
        total_cl: list[float] = []
        total_liabilities: list[float] = []
        total_ca: list[float] = []
        total_assets: list[float] = []

        opening_cash = self.p.cc_amount * 0.02
        prev_wc_investment = 0.0
        for i in range(5):
            wc_investment = receivables[i] + inventory[i] - sundry_creditors[i]
            delta_wc = wc_investment - prev_wc_investment
            operating_cash = pl["retained_profit"][i] + (pl["depreciation"][i] * 0.6)
            raw_cash = opening_cash + operating_cash - max(delta_wc, 0.0)

            funding_gap = max(-raw_cash, 0.0)
            cash_value = max(raw_cash, 0.0)

            cl_other = sundry_creditors[i] + other_cl[i] + provision_tax[i] + dividend_payable[i]
            min_ca_for_mpbf = (cc_bank[i] + cl_other) / 0.75
            ca_without_other = receivables[i] + inventory[i] + cash_value
            mpbf_shortfall_assets = max(min_ca_for_mpbf - ca_without_other, 0.0)

            st_borr = funding_gap
            liabilities_base = cc_bank[i] + sundry_creditors[i] + other_cl[i] + provision_tax[i] + dividend_payable[i] + st_borr + net_worth[i]
            balancing_other_ca = max(liabilities_base - (ca_without_other + net_block[i]), 0.0)

            term_support = 0.0
            if mpbf_shortfall_assets > balancing_other_ca:
                term_support = mpbf_shortfall_assets - balancing_other_ca
                liabilities_base += term_support
                balancing_other_ca = mpbf_shortfall_assets

            residual_gap = liabilities_base - (ca_without_other + balancing_other_ca + net_block[i])
            if residual_gap < 0:
                extra_term_support = abs(residual_gap)
                term_support += extra_term_support
                liabilities_base += extra_term_support
            balancing_other_ca = max(liabilities_base - (ca_without_other + net_block[i]), 0.0)

            current_liabilities_total = cc_bank[i] + sundry_creditors[i] + other_cl[i] + provision_tax[i] + dividend_payable[i] + st_borr
            current_assets_total = ca_without_other + balancing_other_ca
            liabilities_total = liabilities_base
            assets_total = current_assets_total + net_block[i]

            cash_bank.append(cash_value)
            short_term_borrowings_others.append(st_borr)
            other_term_liabilities.append(term_support)
            other_ca.append(balancing_other_ca)
            total_cl.append(current_liabilities_total)
            total_liabilities.append(liabilities_total)
            total_ca.append(current_assets_total)
            total_assets.append(assets_total)

            opening_cash = cash_value
            prev_wc_investment = wc_investment

        return {
            "cc_from_applicant_bank": cc_bank,
            "cc_from_other_banks": [0.0] * 5,
            "bills_purchased": [0.0] * 5,
            "short_term_borrowings_others": short_term_borrowings_others,
            "sundry_creditors": sundry_creditors,
            "advances_customers": [0.0] * 5,
            "provision_tax": provision_tax,
            "dividend_payable": dividend_payable,
            "other_statutory": [0.0] * 5,
            "other_current_liabilities": other_cl,
            "total_current_liabilities": total_cl,
            "term_loans": [0.0] * 5,
            "other_term_liabilities": other_term_liabilities,
            "total_outside_liabilities": [cl + ot for cl, ot in zip(total_cl, other_term_liabilities)],
            "share_capital": share_capital,
            "general_reserve": [0.0] * 5,
            "revaluation_reserve": [0.0] * 5,
            "other_reserves": [0.0] * 5,
            "surplus_pl": surplus,
            "net_worth": net_worth,
            "total_liabilities": total_liabilities,
            "cash_bank": cash_bank,
            "govt_securities": [0.0] * 5,
            "fixed_deposits": [0.0] * 5,
            "receivables": receivables,
            "export_receivables": [0.0] * 5,
            "deferred_receivables": [0.0] * 5,
            "stocks": inventory,
            "advances_suppliers": [0.0] * 5,
            "advance_tax": [0.0] * 5,
            "other_current_assets": other_ca,
            "total_current_assets": total_ca,
            "gross_block": gross_block,
            "dep_to_date": dep_to_date,
            "net_block": net_block,
            "other_investments": [0.0] * 5,
            "security_deposits": [0.0] * 5,
            "other_non_current_assets": [0.0] * 5,
            "total_other_non_current": [0.0] * 5,
            "intangible_assets": [0.0] * 5,
            "total_assets": total_assets,
            "net_working_capital": [ca - cl for ca, cl in zip(total_ca, total_cl)],
            "balance_diff": [round(ta - tl, 2) for ta, tl in zip(total_assets, total_liabilities)],
            "balance_status": ["OK" if round(ta - tl, 2) == 0 else "CHECK" for ta, tl in zip(total_assets, total_liabilities)],
        }

    def _build_wc(self, bs: dict[str, Any], sales: list[float]) -> dict[str, Any]:
        raw_material = [0.0] * 5
        wip = [0.0] * 5
        fg = bs["stocks"]
        debtors = bs["receivables"]
        cash = bs["cash_bank"]
        oca = bs["other_current_assets"]
        total_ca = bs["total_current_assets"]

        creditors = bs["sundry_creditors"]
        outstanding = bs["other_current_liabilities"]
        statutory = [p + d for p, d in zip(bs["provision_tax"], bs["dividend_payable"])]
        other_cl = [0.0] * 5
        total_cl_other = [c + o + s + oc for c, o, s, oc in zip(creditors, outstanding, statutory, other_cl)]

        wc_gap = [a - b for a, b in zip(total_ca, total_cl_other)]
        contribution = [a * 0.25 for a in total_ca]
        mpbf = [max(a - d - b, 0.0) for a, d, b in zip(total_ca, contribution, total_cl_other)]
        proposed_cc = [self.p.cc_amount] * 5

        nayak_wc_req = [s * 0.25 for s in sales]
        nayak_contrib = [s * 0.05 for s in sales]
        nayak_eligible = [s * 0.20 for s in sales]

        return {
            "raw_material": raw_material,
            "wip": wip,
            "finished_goods": fg,
            "receivables": debtors,
            "cash_bank": cash,
            "other_current_assets": oca,
            "total_current_assets": total_ca,
            "sundry_creditors": creditors,
            "outstanding_expenses": outstanding,
            "statutory_liabilities": statutory,
            "other_current_liabilities": other_cl,
            "total_current_liabilities": total_cl_other,
            "wc_gap": wc_gap,
            "borrower_contribution": contribution,
            "mpbf": mpbf,
            "proposed_cc_limit": proposed_cc,
            "projected_annual_turnover": sales,
            "nayak_wc_requirement": nayak_wc_req,
            "nayak_borrower_contribution": nayak_contrib,
            "nayak_eligible_bank_finance": nayak_eligible,
        }

    def _build_ratios(self, pl: dict[str, Any], bs: dict[str, Any], wc: dict[str, Any]) -> list[dict[str, Any]]:
        rows = []

        def add(sno: int, name: str, num_label: str, den_label: str, values: list[float]) -> None:
            bm, st = benchmark_status(name, values[0])
            rows.append(
                {
                    "S.No": sno,
                    "Particulars": name,
                    "Numerator": num_label,
                    "Denominator": den_label,
                    **{fy: values[i] for i, fy in enumerate(YEARS)},
                    "Bank Acceptable Benchmark": bm,
                    "Status": st,
                }
            )

        add(1, "Current Ratio", "Current Assets", "Current Liabilities", [safe_div(a, b) for a, b in zip(bs["total_current_assets"], bs["total_current_liabilities"])])
        add(2, "Quick Ratio", "Current Assets - Inventory", "Current Liabilities", [safe_div(a - inv, b) for a, inv, b in zip(bs["total_current_assets"], bs["stocks"], bs["total_current_liabilities"])])
        add(3, "Debt Equity Ratio", "Total Debt", "Net Worth", [safe_div(d, nw) for d, nw in zip(bs["total_outside_liabilities"], bs["net_worth"])])
        add(4, "Total Indebtedness Ratio", "Total Outside Liabilities", "Net Worth", [safe_div(d, nw) for d, nw in zip(bs["total_outside_liabilities"], bs["net_worth"])])
        add(5, "Debt Asset Ratio", "Total Debt", "Total Assets", [safe_div(d, a) for d, a in zip(bs["total_outside_liabilities"], bs["total_assets"])])
        add(6, "Gross Profit Ratio", "Gross Profit", "Net Sales", [safe_div(s - c, s) for s, c in zip(pl["sales"], pl["cost_of_sales"])])
        add(7, "Operating Profit Ratio", "Operating Profit", "Net Sales", [safe_div(p, s) for p, s in zip(pl["op_profit_before_int_dep"], pl["sales"])])
        add(8, "Net Profit Ratio", "Net Profit", "Net Sales", [safe_div(n, s) for n, s in zip(pl["net_profit"], pl["sales"])])
        add(9, "Interest Coverage Ratio", "EBIT", "Interest", [safe_div(p + i, i) for p, i in zip(pl["op_profit_after_int_dep"], pl["interest_cc"])])
        add(10, "DSCR", "Cash Profit", "Debt Service", [safe_div(np + d, i) for np, d, i in zip(pl["net_profit"], pl["depreciation"], pl["interest_cc"])])
        add(11, "Inventory Turnover", "COGS", "Average Inventory", [safe_div(c, inv) for c, inv in zip(pl["cost_of_sales"], bs["stocks"])])
        add(12, "Debtors Turnover", "Sales", "Debtors", [safe_div(s, d) for s, d in zip(pl["sales"], bs["receivables"])])
        add(13, "Creditors Turnover", "Purchases", "Creditors", [safe_div(p, c) for p, c in zip(pl["purchases"], bs["sundry_creditors"])])
        add(14, "Working Capital Turnover", "Sales", "Working Capital", [safe_div(s, w) for s, w in zip(pl["sales"], bs["net_working_capital"])])
        return rows


def generate_projection(data: dict[str, Any]) -> dict[str, Any]:
    return CMAEngine(ProposalInput(**data)).run()
