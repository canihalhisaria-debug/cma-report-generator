const num = (id) => Number(document.getElementById(id).value || 0);
const txt = (id) => document.getElementById(id).value.trim();

const inputIds = [
  "in-domestic-sale", "in-export-sale", "in-other-operating-income", "in-non-operating-income",
  "in-opening-stock", "in-purchases", "in-carriage-outward", "in-unloading-expenses", "in-direct-expenses", "in-closing-stock",
  "in-salary-wages", "in-power-fuel", "in-rent-expenses", "in-printing-stationery", "in-depreciation", "in-other-expenditure",
  "in-tl-repayment", "in-income-tax", "bs-gross-fixed-assets", "bs-accumulated-depreciation", "bs-sundry-debtors", "bs-cash-bank",
  "bs-other-current-assets", "bs-loans-advances", "bs-capital", "bs-reserves-surplus", "bs-term-loan", "bs-bank-cc-od",
  "bs-sundry-creditors", "bs-other-current-liabilities", "bs-statutory-liabilities",
  "g-domestic", "g-export", "g-other-income", "g-material", "g-opex", "debtor-days", "creditor-days", "inventory-days",
  "rate-cc", "rate-tl", "tax-rate", "projection-years",
];

const sheetDefs = {
  "Profit & Loss": [
    "A. INCOME", "Domestic Sale", "Export Sale", "Other Operating Income", "Non Operating Income", "Total Income",
    "B. DIRECT / MATERIAL COST", "Opening Stock", "Purchases", "Carriage Outward", "Unloading Expenses", "Direct Expenses", "Less: Closing Stock", "Material Cost",
    "C. INDIRECT / OPERATING EXPENSES", "Salary & Wages", "Power & Fuel", "Rent Expenses", "Printing & Stationery", "Depreciation", "Other Expenditure", "Total Operating Expenses",
    "D. EBIT", "E. INTEREST EXPENSE", "Interest on CC", "Interest on TL", "Total Interest", "F. EBT", "G. TAX EXPENSE", "Income Tax", "H. PROFIT AFTER TAX", "I. DEPRECIATION ADDED BACK", "J. CASH ACCRUAL", "K. REPAYMENT OF TERM LOAN", "L. NET CASH AVAILABLE"
  ],
  "Balance Sheet": [
    "ASSETS", "Gross Fixed Assets", "Less Accumulated Depreciation", "Net Fixed Assets",
    "Inventory / Closing Stock", "Sundry Debtors", "Cash & Bank", "Other Current Assets", "Loans & Advances", "Current Assets",
    "LIABILITIES", "Capital", "Reserves & Surplus", "Net Worth", "Term Loan", "Bank CC / OD", "Sundry Creditors", "Other Current Liabilities", "Statutory Liabilities", "Current Liabilities", "Working Capital"
  ],
  "Ratios": ["Current Ratio", "Quick Ratio", "GP", "GP Ratio", "EBIT Ratio", "Net Profit Ratio", "Interest Coverage Ratio", "DSCR"],
};

function safeDiv(a, b) { return b === 0 ? 0 : a / b; }
function money(v) { return Number(v || 0).toLocaleString("en-IN", { maximumFractionDigits: 0 }); }
function pct(v) { return `${Number(v || 0).toFixed(2)}%`; }

function calcYear(prev, growth, index, actualLabel) {
  if (index === 0) {
    const dom = num("in-domestic-sale");
    const exp = num("in-export-sale");
    const ooi = num("in-other-operating-income");
    const noi = num("in-non-operating-income");
    const os = num("in-opening-stock");
    const pur = num("in-purchases");
    const co = num("in-carriage-outward");
    const ue = num("in-unloading-expenses");
    const de = num("in-direct-expenses");
    const cs = num("in-closing-stock");
    const sw = num("in-salary-wages");
    const pf = num("in-power-fuel");
    const rent = num("in-rent-expenses");
    const ps = num("in-printing-stationery");
    const dep = num("in-depreciation");
    const oe = num("in-other-expenditure");
    const repay = num("in-tl-repayment");

    const totalIncome = dom + exp + ooi + noi;
    const materialCost = os + pur + co + ue + de - cs;
    const totalOpex = sw + pf + rent + ps + dep + oe;
    const ebit = totalIncome - materialCost - totalOpex;
    const ccInt = num("bs-bank-cc-od") * (num("rate-cc") / 100);
    const tlInt = num("bs-term-loan") * (num("rate-tl") / 100);
    const totalInterest = ccInt + tlInt;
    const ebt = ebit - totalInterest;
    const tax = num("in-income-tax") > 0 ? num("in-income-tax") : Math.max(0, ebt * num("tax-rate") / 100);
    const pat = ebt - tax;
    const cashAccrual = pat + dep;
    const netCash = cashAccrual - repay;

    const gfa = num("bs-gross-fixed-assets");
    const accDep = num("bs-accumulated-depreciation");
    const nfa = gfa - accDep;
    const inv = num("in-closing-stock");
    const debt = num("bs-sundry-debtors");
    const cash = num("bs-cash-bank");
    const oca = num("bs-other-current-assets");
    const la = num("bs-loans-advances");
    const ca = inv + debt + cash + oca + la;
    const cap = num("bs-capital");
    const res = num("bs-reserves-surplus");
    const nw = cap + res;
    const tl = num("bs-term-loan");
    const cc = num("bs-bank-cc-od");
    const cred = num("bs-sundry-creditors");
    const ocl = num("bs-other-current-liabilities");
    const sl = num("bs-statutory-liabilities");
    const cl = cc + cred + ocl + sl;
    const wc = ca - cl;

    return { year: actualLabel, locked: true, dom, exp, ooi, noi, os, pur, co, ue, de, cs, sw, pf, rent, ps, dep, oe, totalIncome, materialCost, totalOpex, ebit, ccInt, tlInt, totalInterest, ebt, tax, pat, cashAccrual, repay, netCash, gfa, accDep, nfa, inv, debt, cash, oca, la, ca, cap, res, nw, tl, cc, cred, ocl, sl, cl, wc };
  }

  const grow = (v, pct) => v * (1 + pct / 100);
  const dom = grow(prev.dom, growth.gDom);
  const exp = grow(prev.exp, growth.gExp);
  const ooi = grow(prev.ooi, growth.gOther);
  const noi = grow(prev.noi, growth.gOther);
  const os = grow(prev.os, growth.gMaterial);
  const pur = grow(prev.pur, growth.gMaterial);
  const co = grow(prev.co, growth.gMaterial);
  const ue = grow(prev.ue, growth.gMaterial);
  const de = grow(prev.de, growth.gMaterial);
  const sw = grow(prev.sw, growth.gOpex);
  const pf = grow(prev.pf, growth.gOpex);
  const rent = grow(prev.rent, growth.gOpex);
  const ps = grow(prev.ps, growth.gOpex);
  const dep = grow(prev.dep, growth.gOpex);
  const oe = grow(prev.oe, growth.gOpex);
  const repay = prev.repay;

  const tradingSales = dom + exp;
  const cs = tradingSales * growth.inventoryDays / 365;
  const totalIncome = dom + exp + ooi + noi;
  const materialCost = os + pur + co + ue + de - cs;
  const totalOpex = sw + pf + rent + ps + dep + oe;

  const gfa = prev.gfa;
  const accDep = prev.accDep + dep;
  const nfa = gfa - accDep;
  const inv = cs;
  const debt = tradingSales * growth.debtorDays / 365;
  const oca = grow(prev.oca, growth.gOpex);
  const la = grow(prev.la, growth.gOpex);

  const cred = materialCost * growth.creditorDays / 365;
  const ocl = grow(prev.ocl, growth.gOpex);
  const sl = grow(prev.sl, growth.gOpex);
  const cc = grow(prev.cc, growth.gMaterial);
  const ccInt = ((prev.cc + cc) / 2) * growth.rateCc / 100;

  const tl = Math.max(prev.tl - repay, 0);
  const tlInt = ((prev.tl + tl) / 2) * growth.rateTl / 100;
  const totalInterest = ccInt + tlInt;

  const ebit = totalIncome - materialCost - totalOpex;
  const ebt = ebit - totalInterest;
  const tax = Math.max(0, ebt * growth.taxRate / 100);
  const pat = ebt - tax;
  const cashAccrual = pat + dep;
  const netCash = cashAccrual - repay;

  const cash = Math.max(0, prev.cash + netCash * 0.15);
  const ca = inv + debt + cash + oca + la;
  const cap = prev.cap;
  const res = prev.res + pat;
  const nw = cap + res;
  const cl = cc + cred + ocl + sl;
  const wc = ca - cl;

  return { year: `FY+${index}`, locked: false, dom, exp, ooi, noi, os, pur, co, ue, de, cs, sw, pf, rent, ps, dep, oe, totalIncome, materialCost, totalOpex, ebit, ccInt, tlInt, totalInterest, ebt, tax, pat, cashAccrual, repay, netCash, gfa, accDep, nfa, inv, debt, cash, oca, la, ca, cap, res, nw, tl, cc, cred, ocl, sl, cl, wc };
}

function buildModel() {
  const years = Number(num("projection-years"));
  const actualLabel = txt("actual-year") || "Actual Year";
  const growth = {
    gDom: num("g-domestic"), gExp: num("g-export"), gOther: num("g-other-income"), gMaterial: num("g-material"), gOpex: num("g-opex"),
    debtorDays: num("debtor-days"), creditorDays: num("creditor-days"), inventoryDays: num("inventory-days"),
    rateCc: num("rate-cc"), rateTl: num("rate-tl"), taxRate: num("tax-rate"),
  };

  const arr = [];
  for (let i = 0; i <= years; i += 1) arr.push(calcYear(arr[i - 1], growth, i, actualLabel));
  return arr.map((y) => {
    const gp = (y.dom + y.exp) - y.materialCost;
    return {
      ...y,
      currentRatio: safeDiv(y.ca, y.cl),
      quickRatio: safeDiv(y.ca - y.inv, y.cl),
      gp,
      gpRatio: safeDiv(gp, (y.dom + y.exp)) * 100,
      ebitRatio: safeDiv(y.ebit, y.totalIncome) * 100,
      netProfitRatio: safeDiv(y.pat, y.totalIncome) * 100,
      interestCoverageRatio: safeDiv(y.ebit, y.totalInterest),
      dscr: safeDiv(y.pat + y.dep + y.tlInt, y.tlInt + y.repay),
    };
  });
}

function row(label, arr, kind = "money") {
  const format = (v) => (kind === "pct" ? pct(v) : kind === "num" ? Number(v || 0).toFixed(2) : money(v));
  return `<tr><td>${label}</td>${arr.map((x) => `<td>${format(x)}</td>`).join("")}</tr>`;
}

function render(data) {
  const periods = data.map((d) => d.year);
  const table = (title, rows) => `<h3>${title}</h3><table><thead><tr><th>Head</th>${periods.map((p) => `<th>${p}</th>`).join("")}</tr></thead><tbody>${rows}</tbody></table>`;

  const plRows = [
    `<tr class="group"><td>A. INCOME</td>${periods.map(() => "<td></td>").join("")}</tr>`,
    row("Domestic Sale", data.map((d) => d.dom)),
    row("Export Sale", data.map((d) => d.exp)),
    row("Other Operating Income", data.map((d) => d.ooi)),
    row("Non Operating Income", data.map((d) => d.noi)),
    `<tr class="formula"><td>Total Income = Domestic Sale + Export Sale + Other Income</td>${data.map((d) => `<td>${money(d.totalIncome)}</td>`).join("")}</tr>`,
    `<tr class="group"><td>B. DIRECT / MATERIAL COST</td>${periods.map(() => "<td></td>").join("")}</tr>`,
    row("Opening Stock", data.map((d) => d.os)), row("Purchases", data.map((d) => d.pur)), row("Carriage Outward", data.map((d) => d.co)), row("Unloading Expenses", data.map((d) => d.ue)), row("Direct Expenses", data.map((d) => d.de)), row("Less: Closing Stock", data.map((d) => d.cs)),
    `<tr class="formula"><td>Material Cost = Opening Stock + Purchases + Carriage Outward + Unloading Expenses + Direct Expenses - Closing Stock</td>${data.map((d) => `<td>${money(d.materialCost)}</td>`).join("")}</tr>`,
    `<tr class="group"><td>C. INDIRECT / OPERATING EXPENSES</td>${periods.map(() => "<td></td>").join("")}</tr>`,
    row("Salary & Wages", data.map((d) => d.sw)), row("Power & Fuel", data.map((d) => d.pf)), row("Rent Expenses", data.map((d) => d.rent)), row("Printing & Stationery", data.map((d) => d.ps)), row("Depreciation", data.map((d) => d.dep)), row("Other Expenditure", data.map((d) => d.oe)),
    `<tr class="formula"><td>Total Operating Expenses = Salary & Wages + Power & Fuel + Rent Expenses + Printing & Stationery + Depreciation + Other Expenditure</td>${data.map((d) => `<td>${money(d.totalOpex)}</td>`).join("")}</tr>`,
    `<tr class="formula"><td>D. EBIT = Total Income - Material Cost - Total Operating Expenses</td>${data.map((d) => `<td>${money(d.ebit)}</td>`).join("")}</tr>`,
    row("Interest on CC", data.map((d) => d.ccInt)), row("Interest on TL", data.map((d) => d.tlInt)),
    `<tr class="formula"><td>E. Total Interest = Interest on CC + Interest on TL</td>${data.map((d) => `<td>${money(d.totalInterest)}</td>`).join("")}</tr>`,
    `<tr class="formula"><td>F. EBT = EBIT - Total Interest</td>${data.map((d) => `<td>${money(d.ebt)}</td>`).join("")}</tr>`,
    row("Income Tax", data.map((d) => d.tax)),
    `<tr class="formula"><td>H. PAT = EBT - Tax Expense</td>${data.map((d) => `<td>${money(d.pat)}</td>`).join("")}</tr>`,
    `<tr class="formula"><td>I. Depreciation Added Back = Depreciation</td>${data.map((d) => `<td>${money(d.dep)}</td>`).join("")}</tr>`,
    `<tr class="formula"><td>J. Cash Accrual = PAT + Depreciation Added Back</td>${data.map((d) => `<td>${money(d.cashAccrual)}</td>`).join("")}</tr>`,
    row("K. Repayment of Term Loan", data.map((d) => d.repay)),
    `<tr class="formula"><td>L. Net Cash Available = Cash Accrual - Repayment of Term Loan</td>${data.map((d) => `<td>${money(d.netCash)}</td>`).join("")}</tr>`,
  ].join("");

  const bsRows = [
    row("Gross Fixed Assets", data.map((d) => d.gfa)), row("Less Accumulated Depreciation", data.map((d) => d.accDep)), row("Net Fixed Assets", data.map((d) => d.nfa)),
    row("Inventory / Closing Stock", data.map((d) => d.inv)), row("Sundry Debtors", data.map((d) => d.debt)), row("Cash & Bank", data.map((d) => d.cash)), row("Other Current Assets", data.map((d) => d.oca)), row("Loans & Advances", data.map((d) => d.la)),
    `<tr class="formula"><td>Current Assets = Inventory + Debtors + Cash & Bank + Other Current Assets + Loans & Advances</td>${data.map((d) => `<td>${money(d.ca)}</td>`).join("")}</tr>`,
    row("Capital", data.map((d) => d.cap)), row("Reserves & Surplus", data.map((d) => d.res)), `<tr class="formula"><td>Net Worth = Capital + Reserves & Surplus</td>${data.map((d) => `<td>${money(d.nw)}</td>`).join("")}</tr>`,
    row("Term Loan", data.map((d) => d.tl)), row("Bank CC / OD", data.map((d) => d.cc)), row("Sundry Creditors", data.map((d) => d.cred)), row("Other Current Liabilities", data.map((d) => d.ocl)), row("Statutory Liabilities", data.map((d) => d.sl)),
    `<tr class="formula"><td>Current Liabilities = Bank CC / OD + Sundry Creditors + Other Current Liabilities + Statutory Liabilities</td>${data.map((d) => `<td>${money(d.cl)}</td>`).join("")}</tr>`,
    `<tr class="formula"><td>Working Capital = Current Assets - Current Liabilities</td>${data.map((d) => `<td>${money(d.wc)}</td>`).join("")}</tr>`,
  ].join("");

  const ratioRows = [
    row("Current Ratio", data.map((d) => d.currentRatio), "num"), row("Quick Ratio", data.map((d) => d.quickRatio), "num"), row("GP", data.map((d) => d.gp)),
    row("GP Ratio", data.map((d) => d.gpRatio), "pct"), row("EBIT Ratio", data.map((d) => d.ebitRatio), "pct"), row("Net Profit Ratio", data.map((d) => d.netProfitRatio), "pct"),
    row("Interest Coverage Ratio", data.map((d) => d.interestCoverageRatio), "num"), row("DSCR", data.map((d) => d.dscr), "num"),
  ].join("");

  const validation = data.map((d) => {
    const checks = {
      "Total Income": d.totalIncome === d.dom + d.exp + d.ooi + d.noi,
      "Material Cost": d.materialCost === d.os + d.pur + d.co + d.ue + d.de - d.cs,
      EBIT: d.ebit === d.totalIncome - d.materialCost - d.totalOpex,
      Interest: d.totalInterest === d.ccInt + d.tlInt,
      EBT: d.ebt === d.ebit - d.totalInterest,
      PAT: d.pat === d.ebt - d.tax,
      "Current Assets": d.ca === d.inv + d.debt + d.cash + d.oca + d.la,
      "Current Liabilities": d.cl === d.cc + d.cred + d.ocl + d.sl,
      "Working Capital": d.wc === d.ca - d.cl,
    };
    return Object.entries(checks).map(([k, v]) => `<tr><td>${d.year} - ${k}</td><td class="${v ? "ok" : "bad"}">${v ? "OK" : "Mismatch"}</td></tr>`).join("");
  }).join("");

  document.getElementById("output").innerHTML =
    table("Profit & Loss", plRows) +
    table("Balance Sheet", bsRows) +
    table("Working Capital", `<tr><td>Working Capital = Current Assets - Current Liabilities</td>${data.map((d) => `<td>${money(d.wc)}</td>`).join("")}</tr>`) +
    table("Ratios", ratioRows) +
    `<h3>DSCR</h3><table><thead><tr><th>Year</th><th>DSCR</th><th>Status</th></tr></thead><tbody>${data.map((d) => `<tr><td>${d.year}</td><td>${d.dscr.toFixed(2)}</td><td class="${d.dscr >= 1.25 ? "ok" : "bad"}">${d.dscr >= 1.25 ? "Acceptable" : "Watch"}</td></tr>`).join("")}</tbody></table>` +
    `<h3>Validation</h3><table><thead><tr><th>Check</th><th>Status</th></tr></thead><tbody>${validation}</tbody></table>`;
}

let model = [];

function buildWorkbook(data) {
  const wb = XLSX.utils.book_new();
  const periods = data.map((d) => d.year);
  const addSheet = (name, rows) => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([[name], ["Head", ...periods], ...rows]), name);

  addSheet("Dashboard", [["Actual year is locked input", "Yes", ...periods.slice(1).map(() => "Projected")]]);

  addSheet("Profit & Loss", sheetDefs["Profit & Loss"].map((h) => {
    const m = {
      "Domestic Sale": "dom", "Export Sale": "exp", "Other Operating Income": "ooi", "Non Operating Income": "noi", "Total Income": "totalIncome",
      "Opening Stock": "os", "Purchases": "pur", "Carriage Outward": "co", "Unloading Expenses": "ue", "Direct Expenses": "de", "Less: Closing Stock": "cs", "Material Cost": "materialCost",
      "Salary & Wages": "sw", "Power & Fuel": "pf", "Rent Expenses": "rent", "Printing & Stationery": "ps", "Depreciation": "dep", "Other Expenditure": "oe", "Total Operating Expenses": "totalOpex",
      "D. EBIT": "ebit", "Interest on CC": "ccInt", "Interest on TL": "tlInt", "Total Interest": "totalInterest", "F. EBT": "ebt", "Income Tax": "tax",
      "H. PROFIT AFTER TAX": "pat", "J. CASH ACCRUAL": "cashAccrual", "K. REPAYMENT OF TERM LOAN": "repay", "L. NET CASH AVAILABLE": "netCash",
    };
    return [h, ...data.map((d) => (m[h] ? d[m[h]] : ""))];
  }));

  addSheet("Balance Sheet", sheetDefs["Balance Sheet"].map((h) => {
    const m = { "Gross Fixed Assets": "gfa", "Less Accumulated Depreciation": "accDep", "Net Fixed Assets": "nfa", "Inventory / Closing Stock": "inv", "Sundry Debtors": "debt", "Cash & Bank": "cash", "Other Current Assets": "oca", "Loans & Advances": "la", "Current Assets": "ca", "Capital": "cap", "Reserves & Surplus": "res", "Net Worth": "nw", "Term Loan": "tl", "Bank CC / OD": "cc", "Sundry Creditors": "cred", "Other Current Liabilities": "ocl", "Statutory Liabilities": "sl", "Current Liabilities": "cl", "Working Capital": "wc" };
    return [h, ...data.map((d) => (m[h] ? d[m[h]] : ""))];
  }));

  addSheet("Working Capital", [["Current Assets", ...data.map((d) => d.ca)], ["Current Liabilities", ...data.map((d) => d.cl)], ["Working Capital", ...data.map((d) => d.wc)]]);
  addSheet("Ratios", sheetDefs.Ratios.map((h) => {
    const m = { "Current Ratio": "currentRatio", "Quick Ratio": "quickRatio", GP: "gp", "GP Ratio": "gpRatio", "EBIT Ratio": "ebitRatio", "Net Profit Ratio": "netProfitRatio", "Interest Coverage Ratio": "interestCoverageRatio", DSCR: "dscr" };
    return [h, ...data.map((d) => d[m[h]])];
  }));
  addSheet("DSCR", [["Formula", "(PAT + Depreciation + Interest on TL) / (Interest on TL + Repayment of Term Loan)"], ["Year", "DSCR", "Status"], ...data.map((d) => [d.year, d.dscr, d.dscr >= 1.25 ? "Acceptable" : "Watch"])]);
  addSheet("Validation", [["Check", ...periods], ["Total Income", ...data.map((d) => d.totalIncome - (d.dom + d.exp + d.ooi + d.noi))], ["Material Cost", ...data.map((d) => d.materialCost - (d.os + d.pur + d.co + d.ue + d.de - d.cs))], ["EBIT", ...data.map((d) => d.ebit - (d.totalIncome - d.materialCost - d.totalOpex))], ["Interest", ...data.map((d) => d.totalInterest - (d.ccInt + d.tlInt))], ["EBT", ...data.map((d) => d.ebt - (d.ebit - d.totalInterest))], ["PAT", ...data.map((d) => d.pat - (d.ebt - d.tax))], ["Current Assets", ...data.map((d) => d.ca - (d.inv + d.debt + d.cash + d.oca + d.la))], ["Current Liabilities", ...data.map((d) => d.cl - (d.cc + d.cred + d.ocl + d.sl))], ["Working Capital", ...data.map((d) => d.wc - (d.ca - d.cl))]]);

  return wb;
}

function run() {
  model = buildModel();
  render(model);
}

document.getElementById("run").addEventListener("click", run);
document.getElementById("download").addEventListener("click", () => {
  if (!model.length) run();
  XLSX.writeFile(buildWorkbook(model), "banker-grade-cma.xlsx");
});

inputIds.forEach((id) => {
  const el = document.getElementById(id);
  if (el) el.addEventListener("change", run);
});
run();
