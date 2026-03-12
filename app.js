const num = (id) => Number(document.getElementById(id).value || 0);

const INPUT_IDS = ["cc-amount", "roi", "projection-years"];

function safeDiv(a, b) {
  return b === 0 ? 0 : a / b;
}

function money(v) {
  return Number(v || 0).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function pct(v) {
  return `${Number(v || 0).toFixed(2)}%`;
}

function buildModel() {
  const ccBase = num("cc-amount");
  const roi = num("roi");
  const projectionYears = Math.max(2, Number(num("projection-years") || 5));

  const assumptions = {
    salesToCc: 4,
    salesGrowth: 10,
    materialCostPct: 72,
    opexPct: 12,
    depreciationPct: 2,
    taxRate: 25,
    payoutRatio: 20,
    debtorDays: 60,
    creditorDays: 45,
    inventoryDays: 50,
    cashSalesPct: 5,
    otherCurrentAssetsPct: 1.5,
    loansAdvPct: 1,
    termLoanPct: 40,
    repaymentPct: 10,
    equityPct: 30,
    reserveStartPct: 8,
    statutoryLiabilityPct: 1.5,
  };

  const periods = [];
  const actualYear = document.getElementById("actual-year").value || "FY 2024-25";

  for (let i = 0; i < projectionYears; i += 1) {
    const sales = ccBase * assumptions.salesToCc * Math.pow(1 + assumptions.salesGrowth / 100, i);
    const totalIncome = sales;

    const materialCost = totalIncome * (assumptions.materialCostPct / 100);
    const totalOpex = totalIncome * (assumptions.opexPct / 100);
    const depreciation = totalIncome * (assumptions.depreciationPct / 100);
    const ebit = totalIncome - materialCost - totalOpex;

    const creditors = materialCost * (assumptions.creditorDays / 365);
    const inventory = materialCost * (assumptions.inventoryDays / 365);
    const debtors = totalIncome * (assumptions.debtorDays / 365) * (1 - assumptions.cashSalesPct / 100);
    const otherCurrentAssets = totalIncome * (assumptions.otherCurrentAssetsPct / 100);
    const loansAdvances = totalIncome * (assumptions.loansAdvPct / 100);

    const currentAssetsBeforeCash = inventory + debtors + otherCurrentAssets + loansAdvances;
    const statutoryLiabilities = totalIncome * (assumptions.statutoryLiabilityPct / 100);
    const currentLiabilitiesNonCc = creditors + statutoryLiabilities;
    const workingCapitalGap = Math.max(0, currentAssetsBeforeCash - currentLiabilitiesNonCc);
    const cc = i === 0 ? ccBase : workingCapitalGap * 0.75;

    const interestCc = cc * (roi / 100);
    const termLoan = cc * (assumptions.termLoanPct / 100);
    const interestTl = termLoan * (roi / 100);
    const totalInterest = interestCc + interestTl;
    const ebt = ebit - totalInterest;
    const tax = Math.max(0, ebt * (assumptions.taxRate / 100));
    const pat = ebt - tax;
    const cashAccrual = pat + depreciation;
    const repaymentTl = termLoan * (assumptions.repaymentPct / 100);
    const netCash = cashAccrual - repaymentTl;

    const openingStock = i === 0 ? inventory * 0.9 : periods[i - 1].cs;
    const closingStock = inventory;
    const purchases = materialCost - openingStock + closingStock;

    const grossFixedAssets = totalIncome * 0.22;
    const accumulatedDepreciation = totalIncome * 0.06;
    const netFixedAssets = grossFixedAssets - accumulatedDepreciation;
    const cashBank = Math.max(0, currentLiabilitiesNonCc + cc - currentAssetsBeforeCash);
    const currentAssets = currentAssetsBeforeCash + cashBank;
    const capital = cc * (assumptions.equityPct / 100);
    const reserves = cc * (assumptions.reserveStartPct / 100) + pat * (1 - assumptions.payoutRatio / 100);
    const netWorth = capital + reserves;
    const currentLiabilities = cc + currentLiabilitiesNonCc;
    const workingCapital = currentAssets - currentLiabilities;

    const currentRatio = safeDiv(currentAssets, currentLiabilities);
    const quickRatio = safeDiv(currentAssets - inventory, currentLiabilities);
    const gp = totalIncome - materialCost;
    const gpRatio = safeDiv(gp * 100, totalIncome);
    const ebitRatio = safeDiv(ebit * 100, totalIncome);
    const netProfitRatio = safeDiv(pat * 100, totalIncome);
    const interestCoverageRatio = safeDiv(ebit, totalInterest);
    const dscr = safeDiv(pat + depreciation + interestTl, interestTl + repaymentTl);

    periods.push({
      year: i === 0 ? actualYear : `Year ${i + 1}`,
      dom: totalIncome,
      exp: 0,
      ooi: 0,
      noi: 0,
      totalIncome,
      os: openingStock,
      pur: purchases,
      co: 0,
      ue: 0,
      de: 0,
      cs: closingStock,
      materialCost,
      sw: totalOpex * 0.45,
      pf: totalOpex * 0.2,
      rent: totalOpex * 0.15,
      ps: totalOpex * 0.05,
      dep: depreciation,
      oe: totalOpex * 0.15,
      totalOpex,
      ebit,
      ccInt: interestCc,
      tlInt: interestTl,
      totalInterest,
      ebt,
      tax,
      pat,
      cashAccrual,
      repay: repaymentTl,
      netCash,
      gfa: grossFixedAssets,
      accDep: accumulatedDepreciation,
      nfa: netFixedAssets,
      inv: inventory,
      debt: debtors,
      cash: cashBank,
      oca: otherCurrentAssets,
      la: loansAdvances,
      ca: currentAssets,
      cap: capital,
      res: reserves,
      nw: netWorth,
      tl: termLoan,
      cc,
      cred: creditors,
      ocl: 0,
      sl: statutoryLiabilities,
      cl: currentLiabilities,
      wc: workingCapital,
      currentRatio,
      quickRatio,
      gp,
      gpRatio,
      ebitRatio,
      netProfitRatio,
      interestCoverageRatio,
      dscr,
    });
  }

  return periods;
}

function render(data) {
  const periods = data.map((d) => d.year);
  const formatValue = (value, kind = "money") => {
    if (kind === "pct") return pct(value);
    if (kind === "num") return Number(value || 0).toFixed(2);
    if (kind === "blank") return "";
    return money(value);
  };

  const row = ({ label, vals, kind = "money", className = "", labelClass = "" }) => `
    <tr class="${className}">
      <td class="${labelClass}">${label}</td>
      ${vals.map((v) => `<td>${formatValue(v, kind)}</td>`).join("")}
    </tr>
  `;

  const table = (title, body, firstColumnLabel = "Head", tableClass = "", headers = periods) => `
    <h3>${title}</h3>
    <table class="${tableClass}">
      <thead>
        <tr>
          <th>${firstColumnLabel}</th>
          ${headers.map((p) => `<th>${p}</th>`).join("")}
        </tr>
      </thead>
      <tbody>${body}</tbody>
    </table>
  `;

  const blankVals = () => periods.map(() => 0);
  const projectedPlRows = [
    { label: "1. Gross Income", vals: data.map((d) => d.totalIncome), className: "pl-section" },
    { label: "(i) Sales (Net of Returns)", vals: data.map((d) => d.totalIncome), className: "pl-subsection" },
    { label: "(a) Domestic Sales", vals: data.map((d) => d.dom), labelClass: "pl-indent-1" },
    { label: "(b) Export Sales", vals: data.map((d) => d.exp), labelClass: "pl-indent-1" },
    { label: "(c) Sub-Total (a + b)", vals: data.map((d) => d.dom + d.exp), className: "pl-total", labelClass: "pl-indent-1" },
    { label: "(d) Percentage Rise (+) or Fall (−) in Sales Turnover Compared to Previous Year", vals: data.map((_, i) => (i === 0 ? 0 : 12)), kind: "num", labelClass: "pl-indent-1" },
    { label: "(ii) Other Income", vals: data.map((d) => d.ooi), className: "pl-subsection" },
    { label: "(a) Duty Drawback", vals: blankVals(), labelClass: "pl-indent-1" },
    { label: "(b) Cash Assistance", vals: blankVals(), labelClass: "pl-indent-1" },
    { label: "(c) Commission & Brokerage Received", vals: blankVals(), labelClass: "pl-indent-1" },
    { label: "(d) Sub-Total (a + b + c)", vals: blankVals(), className: "pl-total", labelClass: "pl-indent-1" },
    { label: "(iii) Total Gross Income (i + ii)", vals: data.map((d) => d.totalIncome), className: "pl-total" },

    { label: "2. Cost of Sales", vals: data.map((d) => d.materialCost), className: "pl-section" },
    { label: "(i) Purchases", vals: data.map((d) => d.pur), className: "pl-subsection" },
    { label: "(ii) Other Trading Expenses", vals: blankVals(), className: "pl-subsection" },
    { label: "• Carriage Inward", vals: data.map((d) => d.ue), labelClass: "pl-indent-1" },
    { label: "• Commission on Purchases", vals: blankVals(), labelClass: "pl-indent-1" },
    { label: "• Brokerage on Purchases", vals: blankVals(), labelClass: "pl-indent-1" },
    { label: "(iii) Sub-Total (i + ii)", vals: data.map((d) => d.pur + d.ue), className: "pl-total" },
    { label: "(iv) Add : Opening Stock", vals: data.map((d) => d.os), className: "pl-subsection" },
    { label: "(v) Sub-Total (iii + iv)", vals: data.map((d) => d.pur + d.ue + d.os), className: "pl-total" },
    { label: "(vi) Less : Closing Stock", vals: data.map((d) => d.cs), className: "pl-subsection" },
    { label: "(vii) Total Cost of Sales (v − vi)", vals: data.map((d) => d.materialCost), className: "pl-total" },

    { label: "3. Operating Expenses", vals: data.map((d) => d.totalOpex), className: "pl-section" },
    { label: "• Salary", vals: data.map((d) => d.sw), labelClass: "pl-indent-1" },
    { label: "• Rent", vals: data.map((d) => d.rent), labelClass: "pl-indent-1" },
    { label: "• Power & Fuel", vals: data.map((d) => d.pf), labelClass: "pl-indent-1" },
    { label: "• Travelling & Conveyance", vals: data.map((d) => d.ps), labelClass: "pl-indent-1" },
    { label: "• Telephone & Internet", vals: data.map((d) => d.oe * 0.2), labelClass: "pl-indent-1" },
    { label: "• Office Expenses", vals: data.map((d) => d.oe * 0.25), labelClass: "pl-indent-1" },
    { label: "• Printing & Stationery", vals: data.map((d) => d.oe * 0.15), labelClass: "pl-indent-1" },
    { label: "• Repairs & Maintenance", vals: data.map((d) => d.oe * 0.2), labelClass: "pl-indent-1" },
    { label: "• Other Operating Expenses", vals: data.map((d) => d.oe * 0.2), labelClass: "pl-indent-1" },

    { label: "4. Operating Profit (Before Interest & Depreciation)", vals: data.map((d) => d.ebit), className: "pl-section" },
    { label: "5. Interest", vals: data.map((d) => d.totalInterest), className: "pl-section" },
    { label: "• Interest on Cash Credit / OD", vals: data.map((d) => d.ccInt), labelClass: "pl-indent-1" },
    { label: "• Interest on Term Loans", vals: data.map((d) => d.tlInt), labelClass: "pl-indent-1" },
    { label: "• Interest on Unsecured Loans", vals: blankVals(), labelClass: "pl-indent-1" },
    { label: "6. Depreciation", vals: data.map((d) => d.dep), className: "pl-section" },
    { label: "7. Operating Profit (After Interest & Depreciation)", vals: data.map((d) => d.ebt), className: "pl-section" },

    { label: "8. Other Non-Operating Income / Expenses", vals: blankVals(), className: "pl-section" },
    { label: "(i) Add : Other Non-Operating Income", vals: blankVals(), className: "pl-subsection" },
    { label: "(a) Interest Income", vals: blankVals(), labelClass: "pl-indent-1" },
    { label: "(b) Incentive / Subsidy", vals: blankVals(), labelClass: "pl-indent-1" },
    { label: "(c) Other Income", vals: blankVals(), labelClass: "pl-indent-1" },
    { label: "(d) Sub-Total (Incomes)", vals: blankVals(), className: "pl-total" },
    { label: "(ii) Less : Other Non-Operating Expenses", vals: blankVals(), className: "pl-subsection" },
    { label: "(a) Loss on Sale of Assets", vals: blankVals(), labelClass: "pl-indent-1" },
    { label: "(b) Donation / CSR / Charity", vals: blankVals(), labelClass: "pl-indent-1" },
    { label: "(c) Penalty / Late Fees", vals: blankVals(), labelClass: "pl-indent-1" },
    { label: "(d) Misc. Non-Operating Expenses", vals: blankVals(), labelClass: "pl-indent-1" },
    { label: "(e) Sub-Total (Expenses)", vals: blankVals(), className: "pl-total" },
    { label: "(iii) Net Other Non-Operating Income / Expenses", vals: blankVals(), className: "pl-subsection" },

    { label: "9. Profit Before Tax (PBT)", vals: data.map((d) => d.ebt), className: "pl-section" },
    { label: "10. Provision for Taxes", vals: data.map((d) => d.tax), className: "pl-subsection" },
    { label: "11. Net Profit / Loss", vals: data.map((d) => d.pat), className: "pl-section" },
    { label: "12. Dividend", vals: blankVals(), className: "pl-subsection" },
    { label: "(a) Equity Dividend Paid", vals: blankVals(), labelClass: "pl-indent-1" },
    { label: "(b) Dividend Rate", vals: blankVals(), labelClass: "pl-indent-1" },
    { label: "13. Retained Profit", vals: data.map((d) => d.pat), className: "pl-total" },
    { label: "14. Retained Profit / Net Profit (%)", vals: data.map(() => 100), kind: "num", className: "pl-subsection" },
  ];

  const projectedPlBody = projectedPlRows.map((r) => row(r)).join("");

  const bsRows = [
    row({ label: "Inventory", vals: data.map((d) => d.inv) }),
    row({ label: "Debtors", vals: data.map((d) => d.debt) }),
    row({ label: "Creditors", vals: data.map((d) => d.cred) }),
    row({ label: "Current Assets", vals: data.map((d) => d.ca) }),
    row({ label: "Current Liabilities", vals: data.map((d) => d.cl) }),
    row({ label: "Working Capital", vals: data.map((d) => d.wc) }),
  ].join("");

  const ratioRows = [
    row({ label: "Current Ratio", vals: data.map((d) => d.currentRatio), kind: "num" }),
    row({ label: "Quick Ratio", vals: data.map((d) => d.quickRatio), kind: "num" }),
    row({ label: "EBIT Ratio", vals: data.map((d) => d.ebitRatio), kind: "pct" }),
    row({ label: "Net Profit Ratio", vals: data.map((d) => d.netProfitRatio), kind: "pct" }),
    row({ label: "DSCR", vals: data.map((d) => d.dscr), kind: "num" }),
  ].join("");

  const fyHeaders = periods.map((_, i) => `FY ${i + 1}`);

  document.getElementById("output").innerHTML =
    table("Projected PL", projectedPlBody, "Particulars", "projected-pl", fyHeaders) +
    table("Projected Balance Sheet", bsRows) +
    table("Financial Ratios", ratioRows);
}

function buildWorkbook(data) {
  const wb = XLSX.utils.book_new();
  const years = data.map((_, i) => `FY ${i + 1}`);
  const blank = () => data.map(() => 0);

  const plRows = [
    ["1. Gross Income", ...data.map((d) => d.totalIncome)],
    ["(i) Sales (Net of Returns)", ...data.map((d) => d.totalIncome)],
    ["(a) Domestic Sales", ...data.map((d) => d.dom)],
    ["(b) Export Sales", ...data.map((d) => d.exp)],
    ["(c) Sub-Total (a + b)", ...data.map((d) => d.dom + d.exp)],
    ["(d) Percentage Rise (+) or Fall (−) in Sales Turnover Compared to Previous Year", ...data.map((_, i) => (i === 0 ? 0 : 12))],
    ["(ii) Other Income", ...blank()],
    ["(a) Duty Drawback", ...blank()],
    ["(b) Cash Assistance", ...blank()],
    ["(c) Commission & Brokerage Received", ...blank()],
    ["(d) Sub-Total (a + b + c)", ...blank()],
    ["(iii) Total Gross Income (i + ii)", ...data.map((d) => d.totalIncome)],
    ["2. Cost of Sales", ...data.map((d) => d.materialCost)],
    ["(i) Purchases", ...data.map((d) => d.pur)],
    ["(ii) Other Trading Expenses", ...blank()],
    ["• Carriage Inward", ...blank()],
    ["• Commission on Purchases", ...blank()],
    ["• Brokerage on Purchases", ...blank()],
    ["(iii) Sub-Total (i + ii)", ...data.map((d) => d.pur)],
    ["(iv) Add : Opening Stock", ...data.map((d) => d.os)],
    ["(v) Sub-Total (iii + iv)", ...data.map((d) => d.pur + d.os)],
    ["(vi) Less : Closing Stock", ...data.map((d) => d.cs)],
    ["(vii) Total Cost of Sales (v − vi)", ...data.map((d) => d.materialCost)],
    ["3. Operating Expenses", ...data.map((d) => d.totalOpex)],
    ["• Salary", ...data.map((d) => d.sw)],
    ["• Rent", ...data.map((d) => d.rent)],
    ["• Power & Fuel", ...data.map((d) => d.pf)],
    ["• Travelling & Conveyance", ...data.map((d) => d.ps)],
    ["• Telephone & Internet", ...data.map((d) => d.oe * 0.2)],
    ["• Office Expenses", ...data.map((d) => d.oe * 0.25)],
    ["• Printing & Stationery", ...data.map((d) => d.oe * 0.15)],
    ["• Repairs & Maintenance", ...data.map((d) => d.oe * 0.2)],
    ["• Other Operating Expenses", ...data.map((d) => d.oe * 0.2)],
    ["4. Operating Profit (Before Interest & Depreciation)", ...data.map((d) => d.ebit)],
    ["5. Interest", ...data.map((d) => d.totalInterest)],
    ["• Interest on Cash Credit / OD", ...data.map((d) => d.ccInt)],
    ["• Interest on Term Loans", ...data.map((d) => d.tlInt)],
    ["• Interest on Unsecured Loans", ...blank()],
    ["6. Depreciation", ...data.map((d) => d.dep)],
    ["7. Operating Profit (After Interest & Depreciation)", ...data.map((d) => d.ebt)],
    ["8. Other Non-Operating Income / Expenses", ...blank()],
    ["(i) Add : Other Non-Operating Income", ...blank()],
    ["(a) Interest Income", ...blank()],
    ["(b) Incentive / Subsidy", ...blank()],
    ["(c) Other Income", ...blank()],
    ["(d) Sub-Total (Incomes)", ...blank()],
    ["(ii) Less : Other Non-Operating Expenses", ...blank()],
    ["(a) Loss on Sale of Assets", ...blank()],
    ["(b) Donation / CSR / Charity", ...blank()],
    ["(c) Penalty / Late Fees", ...blank()],
    ["(d) Misc. Non-Operating Expenses", ...blank()],
    ["(e) Sub-Total (Expenses)", ...blank()],
    ["(iii) Net Other Non-Operating Income / Expenses", ...blank()],
    ["9. Profit Before Tax (PBT)", ...data.map((d) => d.ebt)],
    ["10. Provision for Taxes", ...data.map((d) => d.tax)],
    ["11. Net Profit / Loss", ...data.map((d) => d.pat)],
    ["12. Dividend", ...blank()],
    ["(a) Equity Dividend Paid", ...blank()],
    ["(b) Dividend Rate", ...blank()],
    ["13. Retained Profit", ...data.map((d) => d.pat)],
    ["14. Retained Profit / Net Profit (%)", ...data.map(() => 100)],
  ];

  const bsRows = [
    ["CURRENT LIABILITIES", ...blank()],
    ["Short Term Borrowings from Banks", ...blank()],
    ["(i) From Applicant Bank", ...data.map((d) => d.cc)],
    ["(ii) From Other Banks", ...blank()],
    ["(iii) Of Which Bills Purchased & Bills Discounted", ...blank()],
    ["Sub-Total (A)", ...data.map((d) => d.cc)],
    ["Short Term Borrowings from Others", ...blank()],
    ["Sundry Creditors (Trade)", ...data.map((d) => d.cred)],
    ["Advances from Customers / Deposits from Dealers", ...blank()],
    ["Provision for Taxation", ...data.map((d) => d.tax)],
    ["Dividend Payable", ...blank()],
    ["Other Statutory Liabilities (Due within One Year)", ...data.map((d) => d.sl)],
    ["Deposits / Debentures / Installments under Term Loans / DPGs etc. (Due within One Year)", ...blank()],
    ["Other Current Liabilities & Provisions (Due within One Year)", ...blank()],
    ["Total Current Liabilities (B)", ...data.map((d) => d.cl)],
    ["TERM LIABILITIES", ...blank()],
    ["Term Loans (Excluding Installments Payable within One Year)", ...data.map((d) => d.tl)],
    ["Other Term Liabilities", ...blank()],
    ["Total Term Liabilities (C)", ...data.map((d) => d.tl)],
    ["TOTAL OUTSIDE LIABILITIES", ...blank()],
    ["Total Outside Liabilities (D)", ...data.map((d) => d.cl + d.tl)],
    ["", ...blank()],
    ["Share Capital", ...data.map((d) => d.cap)],
    ["General Reserve", ...data.map((d) => d.res)],
    ["Revaluation Reserve", ...blank()],
    ["Other Reserves (Excluding Provisions)", ...blank()],
    ["Surplus / (Deficit) in P&L A/c (adjusted after drawings)", ...data.map((d) => d.pat)],
    ["", ...blank()],
    ["Net Worth", ...data.map((d) => d.nw)],
    ["TOTAL LIABILITIES", ...blank()],
    ["Total Liabilities (F)", ...data.map((d) => d.cl + d.tl + d.nw)],
    ["CURRENT ASSETS", ...blank()],
    ["Cash & Bank Balances", ...data.map((d) => d.cash)],
    ["Government & Trustee Securities", ...blank()],
    ["Fixed Deposits with Banks", ...blank()],
    ["Receivables other than Deferred & Export Receivables", ...data.map((d) => d.debt)],
    ["Export Receivables", ...blank()],
    ["Installments of Deferred Receivables (Due within One Year)", ...blank()],
    ["Stocks-in-Trade", ...data.map((d) => d.inv)],
    ["Advances to Suppliers of Merchandise", ...blank()],
    ["Advance Payment of Taxes", ...blank()],
    ["Other Current Assets", ...data.map((d) => d.oca + d.la)],
    ["Total Current Assets (G)", ...data.map((d) => d.ca)],
    ["FIXED ASSETS", ...blank()],
    ["Gross Block", ...data.map((d) => d.gfa)],
    ["Depreciation to Date", ...data.map((d) => d.accDep)],
    ["Net Block (H)", ...data.map((d) => d.nfa)],
    ["OTHER NON-CURRENT ASSETS", ...blank()],
    ["Other Investments", ...blank()],
    ["Security Deposits / Tender Deposits", ...blank()],
    ["Other Non-Current Assets", ...blank()],
    ["Total Other Non-Current Assets (I)", ...blank()],
    ["INTANGIBLE ASSETS", ...blank()],
    ["Intangible Assets", ...blank()],
    ["TOTAL ASSETS", ...blank()],
    ["Total Assets (J)", ...data.map((d) => d.ca + d.nfa)],
    ["TANGIBLE NET WORTH", ...blank()],
    ["Tangible Net Worth (K)", ...data.map((d) => d.nw)],
    ["NET WORKING CAPITAL", ...blank()],
    ["Net Working Capital (L)", ...data.map((d) => d.wc)],
    ["Diff Check Rounded (M)", ...data.map((d) => Math.round((d.cl + d.tl + d.nw) - (d.ca + d.nfa)))],
    ["Balance Status (N)", ...data.map((d) => (Math.round((d.cl + d.tl + d.nw) - (d.ca + d.nfa)) === 0 ? "OK" : "CHECK"))],
  ];

  const plSheet = XLSX.utils.aoa_to_sheet([["PROJECTED PL"], ["Particulars", ...years], ...plRows]);
  const bsSheet = XLSX.utils.aoa_to_sheet([["PROJECTED BS"], ["Particulars", ...years], ...bsRows]);

  plSheet["!cols"] = [{ wch: 72 }, ...years.map(() => ({ wch: 14 }))];
  bsSheet["!cols"] = [{ wch: 72 }, ...years.map(() => ({ wch: 14 }))];

  XLSX.utils.book_append_sheet(wb, plSheet, "Projected PL");
  XLSX.utils.book_append_sheet(wb, bsSheet, "Projected BS");

  return wb;
}

let model = [];
function run() {
  model = buildModel();
  render(model);
}

document.getElementById("run").addEventListener("click", run);
document.getElementById("download").addEventListener("click", () => {
  if (!model.length) run();
  XLSX.writeFile(buildWorkbook(model), "FINAL_CMA_REPORT_FINAL_FIXED.xlsx");
});

INPUT_IDS.forEach((id) => {
  const el = document.getElementById(id);
  if (el) el.addEventListener("change", run);
});

run();
