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
  const periods = data.map((d) => d.year);
  const addSheet = (name, rows) => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([[name], ["Head", ...periods], ...rows]), name);

  addSheet("Dashboard", [["User Input: CC Amount", num("cc-amount")], ["User Input: ROI %", num("roi")]]);
  addSheet("Profit & Loss", [["Sales", ...data.map((d) => d.totalIncome)], ["Interest", ...data.map((d) => d.totalInterest)], ["PAT", ...data.map((d) => d.pat)]]);
  addSheet("Balance Sheet", [["Stock", ...data.map((d) => d.inv)], ["Debtors", ...data.map((d) => d.debt)], ["Creditors", ...data.map((d) => d.cred)], ["Working Capital", ...data.map((d) => d.wc)]]);
  addSheet("Working Capital Analysis", [["Current Assets", ...data.map((d) => d.ca)], ["Current Liabilities", ...data.map((d) => d.cl)], ["Gap", ...data.map((d) => d.wc)]]);
  addSheet("Financial Ratios", [["Current Ratio", ...data.map((d) => d.currentRatio)], ["Quick Ratio", ...data.map((d) => d.quickRatio)], ["DSCR", ...data.map((d) => d.dscr)]]);

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
