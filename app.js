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

function buildDepreciationSchedule(projectionYears) {
  const categories = [
    { key: "land", label: "Land & Site Development", initialAddition: 0 },
    { key: "building", label: "Building", initialAddition: 0 },
    { key: "plant", label: "Plant & Machinery", initialAddition: 0 },
    { key: "electrical", label: "Electrical Installation", initialAddition: 24685 },
    { key: "furniture", label: "Furniture & Fixtures", initialAddition: 82175 },
    { key: "office", label: "Office Equipment", initialAddition: 21240 },
    { key: "computer", label: "Computer / Laptop", initialAddition: 57920 },
    { key: "printer", label: "Printer / Scanner", initialAddition: 16430 },
    { key: "vehicle", label: "Vehicle", initialAddition: 0 },
    { key: "generator", label: "Generator / UPS", initialAddition: 0 },
    { key: "other", label: "Other Fixed Assets", initialAddition: 0 },
  ];
  const depreciationRate = 0.15;
  const years = [];
  const openingByCategory = Object.fromEntries(categories.map((c) => [c.key, 0]));

  for (let i = 0; i < projectionYears; i += 1) {
    const additionsByCategory = Object.fromEntries(categories.map((c) => [c.key, i === 0 ? c.initialAddition : 0]));
    const totalByCategory = Object.fromEntries(categories.map((c) => [c.key, openingByCategory[c.key] + additionsByCategory[c.key]]));
    const depreciationByCategory = Object.fromEntries(categories.map((c) => [c.key, totalByCategory[c.key] * depreciationRate]));
    const closingByCategory = Object.fromEntries(categories.map((c) => [c.key, totalByCategory[c.key] - depreciationByCategory[c.key]]));
    const calcTotal = (source) => categories.reduce((sum, c) => sum + source[c.key], 0);

    years.push({
      label: `FY-${i + 1}`,
      opening: { ...openingByCategory, total: calcTotal(openingByCategory) },
      additions: { ...additionsByCategory, total: calcTotal(additionsByCategory) },
      total: { ...totalByCategory, total: calcTotal(totalByCategory) },
      depreciation: { ...depreciationByCategory, total: calcTotal(depreciationByCategory) },
      closing: { ...closingByCategory, total: calcTotal(closingByCategory) },
    });

    categories.forEach((c) => {
      openingByCategory[c.key] = closingByCategory[c.key];
    });
  }

  return { categories, years };
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
    const currentLiabilities = cc + currentLiabilitiesNonCc;
    const capital = cc * (assumptions.equityPct / 100);
    const outsideLiabilities = currentLiabilities + termLoan;
    const totalAssets = currentAssets + netFixedAssets;
    const netWorth = totalAssets - outsideLiabilities;
    const reserves = netWorth - capital;
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
  const depreciationSchedule = buildDepreciationSchedule(periods.length);
  const formatValue = (value, kind = "money") => {
    if (value === "" || value === null || value === undefined) return "";
    if (kind === "pct") return pct(value);
    if (kind === "num") return Number(value || 0).toFixed(2);
    if (kind === "blank") return "";
    if (kind === "dash") return value ? money(value) : "-";
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

  const blankVals = () => periods.map(() => "");
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
    row({ label: "CURRENT LIABILITIES", vals: blankVals(), className: "pl-section" }),
    row({ label: "Short Term Borrowings from Banks", vals: blankVals() }),
    row({ label: "(i) From Applicant Bank", vals: data.map((d) => d.cc), className: "pl-subsection" }),
    row({ label: "(ii) From Other Banks", vals: blankVals() }),
    row({ label: "(iii) Of Which Bills Purchased & Bills Discounted", vals: blankVals() }),
    row({ label: "Sub-Total (A)", vals: data.map((d) => d.cc), className: "pl-total" }),
    row({ label: "Short Term Borrowings from Others", vals: blankVals() }),
    row({ label: "Sundry Creditors (Trade)", vals: data.map((d) => d.cred) }),
    row({ label: "Advances from Customers / Deposits from Dealers", vals: blankVals() }),
    row({ label: "Provision for Taxation", vals: data.map((d) => d.tax) }),
    row({ label: "Dividend Payable", vals: blankVals() }),
    row({ label: "Other Statutory Liabilities (Due within One Year)", vals: data.map((d) => d.sl) }),
    row({ label: "Other Current Liabilities & Provisions (Due within One Year)", vals: blankVals() }),
    row({ label: "Total Current Liabilities (B)", vals: data.map((d) => d.cl), className: "pl-total" }),
    row({ label: "TERM LIABILITIES", vals: blankVals(), className: "pl-section" }),
    row({ label: "Term Loans (Excluding Installments Payable within One Year)", vals: data.map((d) => d.tl) }),
    row({ label: "Other Term Liabilities", vals: blankVals() }),
    row({ label: "Total Term Liabilities (C)", vals: data.map((d) => d.tl), className: "pl-total" }),
    row({ label: "TOTAL OUTSIDE LIABILITIES", vals: blankVals(), className: "pl-section" }),
    row({ label: "Total Outside Liabilities (D)", vals: data.map((d) => d.cl + d.tl), className: "pl-total" }),
    row({ label: "Share Capital", vals: data.map((d) => d.cap) }),
    row({ label: "General Reserve", vals: data.map((d) => d.res) }),
    row({ label: "Surplus / (Deficit) in P&L A/c (adjusted after drawings)", vals: data.map((d) => d.pat), className: "pl-subsection" }),
    row({ label: "Net Worth", vals: data.map((d) => d.nw), className: "pl-total" }),
    row({ label: "TOTAL LIABILITIES", vals: blankVals(), className: "pl-section" }),
    row({ label: "Total Liabilities (F)", vals: data.map((d) => d.cl + d.tl + d.nw), className: "pl-total" }),
    row({ label: "CURRENT ASSETS", vals: blankVals(), className: "pl-section" }),
    row({ label: "Cash & Bank Balances", vals: data.map((d) => d.cash) }),
    row({ label: "Receivables other than Deferred & Export Receivables", vals: data.map((d) => d.debt) }),
    row({ label: "Stocks-in-Trade", vals: data.map((d) => d.inv) }),
    row({ label: "Other Current Assets", vals: data.map((d) => d.oca + d.la) }),
    row({ label: "Total Current Assets (G)", vals: data.map((d) => d.ca), className: "pl-total" }),
    row({ label: "FIXED ASSETS", vals: blankVals(), className: "pl-section" }),
    row({ label: "Gross Block", vals: data.map((d) => d.gfa) }),
    row({ label: "Depreciation to Date", vals: data.map((d) => d.accDep) }),
    row({ label: "Net Block (H)", vals: data.map((d) => d.nfa), className: "pl-total" }),
    row({ label: "TOTAL ASSETS", vals: blankVals(), className: "pl-section" }),
    row({ label: "Total Assets (J)", vals: data.map((d) => d.ca + d.nfa), className: "pl-total" }),
    row({ label: "NET WORKING CAPITAL", vals: blankVals(), className: "pl-section" }),
    row({ label: "Net Working Capital (L)", vals: data.map((d) => d.wc), className: "pl-total" }),
    row({ label: "Diff Check Rounded (M)", vals: data.map((d) => Math.round((d.cl + d.tl + d.nw) - (d.ca + d.nfa))), kind: "num" }),
    row({ label: "Balance Status (N)", vals: data.map((d) => (Math.round((d.cl + d.tl + d.nw) - (d.ca + d.nfa)) === 0 ? "OK" : "CHECK")), kind: "blank", className: "pl-subsection" }),
  ].join("");

  const ratioInput = [
    { name: "Current Ratio", num: "Current Assets", den: "Current Liabilities", vals: data.map((d) => d.currentRatio), acceptable: ">=1.33", isPct: false },
    { name: "Quick Ratio", num: "Quick Assets", den: "Current Liabilities", vals: data.map((d) => d.quickRatio), acceptable: ">=1.00", isPct: false },
    { name: "Debt Equity Ratio", num: "Total Debt", den: "Net Worth", vals: data.map((d) => safeDiv(d.cl + d.tl, d.nw)), acceptable: "<=2.00", isPct: false },
    { name: "Debt Asset Ratio", num: "Total Debt", den: "Total Assets", vals: data.map((d) => safeDiv(d.cl + d.tl, d.ca + d.nfa)), acceptable: "<=0.75", isPct: false },
    { name: "Gross Profit Ratio", num: "Gross Profit", den: "Sales", vals: data.map((d) => d.gpRatio), acceptable: ">=15%", isPct: true },
    { name: "Operating Profit Ratio", num: "EBIT", den: "Sales", vals: data.map((d) => d.ebitRatio), acceptable: ">=10%", isPct: true },
    { name: "Net Profit Ratio", num: "PAT", den: "Sales", vals: data.map((d) => d.netProfitRatio), acceptable: ">=5%", isPct: true },
    { name: "DSCR", num: "Cash Profit", den: "Debt Service", vals: data.map((d) => d.dscr), acceptable: "2.00+", isPct: false },
  ];

  const ratioRows = ratioInput
    .map((r, i) => {
      const fy1 = r.vals[0] || 0;
      const status = (r.isPct ? fy1 >= Number(r.acceptable.replace(/[^\d.]/g, "")) : fy1 >= 1) ? "OK" : "Weak";
      return `
      <tr>
        <td>${i + 1}</td>
        <td>${r.name}</td>
        <td>${r.num}</td>
        <td>${r.den}</td>
        ${r.vals.map((v) => `<td>${r.isPct ? pct(v) : Number(v || 0).toFixed(2)}</td>`).join("")}
        <td>${r.acceptable}</td>
        <td class="${status === "OK" ? "ok" : "bad"}">${status}</td>
      </tr>`;
    })
    .join("");

  const ratioTable = `
    <h3>Financial Ratios Analysis</h3>
    <table class="projected-pl">
      <thead>
        <tr>
          <th>S.No</th>
          <th>Particulars</th>
          <th>Numerator</th>
          <th>Denominator</th>
          ${periods.map((_, i) => `<th>FY-${i + 1}</th>`).join("")}
          <th>Bank Acceptable</th>
          <th>Status FY-1</th>
        </tr>
      </thead>
      <tbody>${ratioRows}</tbody>
    </table>
  `;

  const fyHeaders = periods.map((_, i) => `FY ${i + 1}`);

  const wcRows = [
    row({ label: "A. Current Assets", vals: blankVals(), kind: "blank", className: "pl-subsection" }),
    row({ label: "Raw Material", vals: blankVals(), kind: "dash" }),
    row({ label: "Work in Progress", vals: blankVals(), kind: "dash" }),
    row({ label: "Finished Goods", vals: data.map((d) => d.inv) }),
    row({ label: "Receivables / Sundry Debtors", vals: data.map((d) => d.debt) }),
    row({ label: "Cash & Bank", vals: data.map((d) => d.cash) }),
    row({ label: "Other Current Assets", vals: data.map((d) => d.oca + d.la) }),
    row({ label: "Total Current Assets (A)", vals: data.map((d) => d.ca), className: "pl-total" }),
    row({ label: "", vals: blankVals(), kind: "blank" }),
    row({ label: "B. Current Liabilities (Other than Bank)", vals: blankVals(), kind: "blank", className: "pl-subsection" }),
    row({ label: "Sundry Creditors", vals: data.map((d) => d.cred) }),
    row({ label: "Outstanding Expenses", vals: blankVals(), kind: "dash" }),
    row({ label: "Statutory Liabilities", vals: data.map((d) => d.sl), kind: "dash" }),
    row({ label: "Other Current Liabilities", vals: blankVals(), kind: "dash" }),
    row({ label: "Total Current Liabilities (B)", vals: data.map((d) => d.cl - d.cc), className: "pl-total" }),
    row({ label: "", vals: blankVals(), kind: "blank" }),
    row({ label: "C. Working Capital Gap (A − B)", vals: data.map((d) => d.ca - (d.cl - d.cc)), className: "pl-total" }),
    row({ label: "", vals: blankVals(), kind: "blank" }),
    row({ label: "D. Borrower Contribution (Margin)", vals: blankVals(), kind: "blank", className: "pl-subsection" }),
    row({ label: "25% of Current Assets (Tandon Method II)", vals: data.map((d) => d.ca * 0.25), className: "pl-total" }),
    row({ label: "", vals: blankVals(), kind: "blank" }),
    row({ label: "E. Maximum Permissible Bank Finance (MPBF)", vals: blankVals(), kind: "blank", className: "pl-subsection" }),
    row({ label: "MPBF = A − D − B", vals: data.map((d) => d.ca - d.ca * 0.25 - (d.cl - d.cc)), className: "pl-total" }),
    row({ label: "", vals: blankVals(), kind: "blank" }),
    row({ label: "F. Proposed CC Limit", vals: data.map((d) => d.cc), className: "pl-subsection" }),
    row({ label: "", vals: blankVals(), kind: "blank" }),
    row({ label: "Alternative (Nayak Committee Method – MSME)", vals: blankVals(), kind: "blank", className: "pl-subsection" }),
    row({ label: "Used when Turnover ≤ ₹5 Crore (commonly for MSME CC)", vals: blankVals(), kind: "blank" }),
    row({ label: "Projected Annual Turnover", vals: data.map((d) => d.totalIncome) }),
    row({ label: "Working Capital Requirement @25%", vals: data.map((d) => d.totalIncome * 0.25), className: "pl-total" }),
    row({ label: "Borrower Contribution @5%", vals: data.map((d) => d.totalIncome * 0.05), className: "pl-total" }),
    row({ label: "Eligible Bank Finance @20%", vals: data.map((d) => d.totalIncome * 0.2), className: "pl-total" }),
  ].join("");

  const workingCapitalTable = table(
    "Working Capital Analysis (CMA Format – Tandon / Nayak Based)",
    wcRows,
    "Particulars",
    "projected-pl",
    fyHeaders
  );

  const depreciationRows = depreciationSchedule.years
    .map((year) => {
      const rowValues = (record) => depreciationSchedule.categories.map((c) => `<td>${money(record[c.key])}</td>`).join("");
      return `
        <tr>
          <td rowspan="5" class="dep-year">${year.label}</td>
          <td>Opening WDV</td>
          ${rowValues(year.opening)}
          <td>${money(year.opening.total)}</td>
        </tr>
        <tr>
          <td>Additions</td>
          ${rowValues(year.additions)}
          <td>${money(year.additions.total)}</td>
        </tr>
        <tr class="dep-total-row">
          <td>Total</td>
          ${rowValues(year.total)}
          <td>${money(year.total.total)}</td>
        </tr>
        <tr>
          <td>Depreciation</td>
          ${rowValues(year.depreciation)}
          <td>${money(year.depreciation.total)}</td>
        </tr>
        <tr>
          <td>Closing WDV</td>
          ${rowValues(year.closing)}
          <td>${money(year.closing.total)}</td>
        </tr>
      `;
    })
    .join("");

  const depreciationTable = `
    <h3>Depreciation Schedule - Print Format</h3>
    <table class="projected-pl depreciation-schedule">
      <thead>
        <tr>
          <th>Year</th>
          <th>Particulars</th>
          ${depreciationSchedule.categories.map((c) => `<th>${c.label}</th>`).join("")}
          <th>Total</th>
        </tr>
      </thead>
      <tbody>${depreciationRows}</tbody>
    </table>
  `;

  document.getElementById("output").innerHTML =
    table("Projected PL", projectedPlBody, "Particulars", "projected-pl", fyHeaders) +
    table("Projected BS", bsRows, "Particulars", "projected-pl", fyHeaders) +
    workingCapitalTable +
    ratioTable +
    depreciationTable;
}

function buildWorkbook(data) {
  const wb = XLSX.utils.book_new();
  const years = data.map((_, i) => `FY ${i + 1}`);
  const depreciationSchedule = buildDepreciationSchedule(years.length);
  const blank = () => data.map(() => "");

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

  const wcWorkbookRows = [
    ["A. Current Assets", ...blank()],
    ["Raw Material", ...blank()],
    ["Work in Progress", ...blank()],
    ["Finished Goods", ...data.map((d) => d.inv)],
    ["Receivables / Sundry Debtors", ...data.map((d) => d.debt)],
    ["Cash & Bank", ...data.map((d) => d.cash)],
    ["Other Current Assets", ...data.map((d) => d.oca + d.la)],
    ["Total Current Assets (A)", ...data.map((d) => d.ca)],
    ["", ...blank()],
    ["B. Current Liabilities (Other than Bank)", ...blank()],
    ["Sundry Creditors", ...data.map((d) => d.cred)],
    ["Outstanding Expenses", ...blank()],
    ["Statutory Liabilities", ...data.map((d) => d.sl)],
    ["Other Current Liabilities", ...blank()],
    ["Total Current Liabilities (B)", ...data.map((d) => d.cl - d.cc)],
    ["", ...blank()],
    ["C. Working Capital Gap (A − B)", ...data.map((d) => d.ca - (d.cl - d.cc))],
    ["", ...blank()],
    ["D. Borrower Contribution (Margin)", ...blank()],
    ["25% of Current Assets (Tandon Method II)", ...data.map((d) => d.ca * 0.25)],
    ["", ...blank()],
    ["E. Maximum Permissible Bank Finance (MPBF)", ...blank()],
    ["MPBF = A − D − B", ...data.map((d) => d.ca - d.ca * 0.25 - (d.cl - d.cc))],
    ["", ...blank()],
    ["F. Proposed CC Limit", ...data.map((d) => d.cc)],
    ["", ...blank()],
    ["Alternative (Nayak Committee Method – MSME)", ...blank()],
    ["Used when Turnover ≤ ₹5 Crore (commonly for MSME CC)", ...blank()],
    ["Projected Annual Turnover", ...data.map((d) => d.totalIncome)],
    ["Working Capital Requirement @25%", ...data.map((d) => d.totalIncome * 0.25)],
    ["Borrower Contribution @5%", ...data.map((d) => d.totalIncome * 0.05)],
    ["Eligible Bank Finance @20%", ...data.map((d) => d.totalIncome * 0.2)],
  ];
  const wcSheet = XLSX.utils.aoa_to_sheet([["WORKING CAPITAL ANALYSIS"], ["Particulars", ...years], ...wcWorkbookRows]);

  const depreciationRows = [];

  depreciationSchedule.years.forEach((year) => {
    depreciationRows.push([year.label, "Opening WDV", ...depreciationSchedule.categories.map((c) => year.opening[c.key]), year.opening.total]);
    depreciationRows.push(["", "Additions", ...depreciationSchedule.categories.map((c) => year.additions[c.key]), year.additions.total]);
    depreciationRows.push(["", "Total", ...depreciationSchedule.categories.map((c) => year.total[c.key]), year.total.total]);
    depreciationRows.push(["", "Depreciation", ...depreciationSchedule.categories.map((c) => year.depreciation[c.key]), year.depreciation.total]);
    depreciationRows.push(["", "Closing WDV", ...depreciationSchedule.categories.map((c) => year.closing[c.key]), year.closing.total]);
  });

  const depreciationSheet = XLSX.utils.aoa_to_sheet([
    ["DEPRECIATION SCHEDULE - PRINT FORMAT"],
    ["Year", "Particulars", ...depreciationSchedule.categories.map((c) => c.label), "Total"],
    ...depreciationRows,
  ]);

  plSheet["!cols"] = [{ wch: 72 }, ...years.map(() => ({ wch: 14 }))];
  bsSheet["!cols"] = [{ wch: 72 }, ...years.map(() => ({ wch: 14 }))];
  wcSheet["!cols"] = [{ wch: 72 }, ...years.map(() => ({ wch: 14 }))];
  depreciationSheet["!cols"] = [{ wch: 12 }, { wch: 18 }, ...depreciationSchedule.categories.map(() => ({ wch: 18 })), { wch: 14 }];

  XLSX.utils.book_append_sheet(wb, plSheet, "Projected PL");
  XLSX.utils.book_append_sheet(wb, bsSheet, "Projected BS");
  XLSX.utils.book_append_sheet(wb, wcSheet, "Working Capital Analysis");
  XLSX.utils.book_append_sheet(wb, depreciationSheet, "Depreciation Schedule");

  return wb;
}

async function downloadDisplayPdf() {
  const output = document.getElementById("output");
  if (!output || !window.html2canvas || !window.jspdf) {
    alert("PDF library load nahi hui. Please page refresh karke dobara try karein.");
    return;
  }

  const originalOverflow = output.style.overflow;
  const originalMaxHeight = output.style.maxHeight;
  output.style.overflow = "visible";
  output.style.maxHeight = "none";

  try {
    const canvas = await window.html2canvas(output, {
      scale: 2,
      useCORS: true,
      backgroundColor: "#ffffff",
      windowWidth: output.scrollWidth,
      windowHeight: output.scrollHeight,
    });

    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF("p", "pt", "a4");

    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const margin = 24;
    const usableWidth = pageWidth - margin * 2;
    const usableHeight = pageHeight - margin * 2;

    const imgWidth = usableWidth;
    const imgHeight = (canvas.height * imgWidth) / canvas.width;

    let heightLeft = imgHeight;
    let positionY = margin;

    const imageData = canvas.toDataURL("image/png");
    pdf.addImage(imageData, "PNG", margin, positionY, imgWidth, imgHeight);
    heightLeft -= usableHeight;

    while (heightLeft > 0) {
      positionY = margin - (imgHeight - heightLeft);
      pdf.addPage();
      pdf.addImage(imageData, "PNG", margin, positionY, imgWidth, imgHeight);
      heightLeft -= usableHeight;
    }

    pdf.save("CMA_Display_Match_PL_BS.pdf");
  } catch (error) {
    console.error(error);
    alert("Download failed. Please try again.");
  } finally {
    output.style.overflow = originalOverflow;
    output.style.maxHeight = originalMaxHeight;
  }
}

function downloadExcelReport() {
  if (!window.XLSX) {
    alert("Excel export library not loaded yet. Please try again.");
    return;
  }
  if (!model.length) {
    run();
  }

  const wb = buildWorkbook(model);
  XLSX.writeFile(wb, "CMA_Display_Match_Report.xlsx");
}


const TOTAL_WIZARD_STEPS = 4;
let currentWizardStep = 1;

function getStepTarget(stepNumber) {
  const button = document.querySelector(`#wizard-progress .step[data-step="${stepNumber}"]`);
  return button ? button.dataset.target : null;
}

function setActiveStep(stepNumber) {
  const progressButtons = document.querySelectorAll("#wizard-progress .step");
  if (!progressButtons.length) return;

  progressButtons.forEach((btn) => {
    const isActive = Number(btn.dataset.step || 0) === stepNumber;
    btn.classList.toggle("active", isActive);
    btn.setAttribute("aria-current", isActive ? "step" : "false");
  });
}

function applyWizardStep(stepNumber) {
  currentWizardStep = Math.min(TOTAL_WIZARD_STEPS, Math.max(1, stepNumber));

  document.querySelectorAll("[data-form-step]").forEach((section) => {
    const sectionStep = Number(section.dataset.formStep);
    section.classList.toggle("wizard-hidden", sectionStep !== currentWizardStep);
  });

  const target = getStepTarget(currentWizardStep);
  if (target) setActiveStep(currentWizardStep);

  const backBtn = document.getElementById("wizard-back");
  const nextBtn = document.getElementById("wizard-next");
  if (backBtn) backBtn.disabled = currentWizardStep === 1;
  if (nextBtn) {
    nextBtn.textContent = currentWizardStep === TOTAL_WIZARD_STEPS ? "Finish" : "Next";
  }
}

function updateFormVisibility() {
  const businessTypeEl = document.getElementById("business-type");
  const loanTypeEl = document.getElementById("loan-type");
  const requirementEl = document.getElementById("requirement");

  const businessType = businessTypeEl ? businessTypeEl.value : "New Business";
  const loanType = loanTypeEl ? loanTypeEl.value : "CC";
  const requirement = requirementEl ? requirementEl.value : "New CC";

  const showExistingBusinessFields = businessType === "Existing Business";
  const showCCFields = loanType === "CC" || loanType === "CC + Term Loan";
  const showTermLoanFields = loanType === "Term Loan" || loanType === "CC + Term Loan";

  const toggle = (id, show) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.classList.toggle("hidden", !show);
  };

  toggle("cc-facility-section", showCCFields);
  toggle("term-loan-section", showTermLoanFields);
  toggle("machinery-section", showTermLoanFields);
  toggle("existing-business-section", showExistingBusinessFields);
  toggle("new-business-section", !showExistingBusinessFields);

  const visible = [];
  if (showCCFields) visible.push("CC");
  if (showTermLoanFields) visible.push("Term Loan", "Machinery");
  visible.push(showExistingBusinessFields ? "Existing Business" : "New Business");

  const setText = (id, value) => {
    const el = document.getElementById(id);
    if (el) el.textContent = value;
  };

  setText("preview-business-type", businessType);
  setText("preview-loan-type", loanType);
  setText("preview-requirement", requirement);
  setText("preview-visible-sections", visible.join(" + "));

  const step2 = document.querySelector('#wizard-progress .step[data-step="2"]');
  if (step2) {
    step2.dataset.target = showCCFields ? "cc-facility-section" : "term-loan-section";
  }

  const step3 = document.querySelector('#wizard-progress .step[data-step="3"]');
  if (step3) {
    if (showExistingBusinessFields) {
      step3.dataset.target = "existing-business-section";
    } else {
      step3.dataset.target = showTermLoanFields ? "term-loan-section" : "cc-facility-section";
    }
  }

  applyWizardStep(currentWizardStep);
}

let model = [];
function run() {
  model = buildModel();
  render(model);
}

document.getElementById("run").addEventListener("click", run);
document.getElementById("save-draft").addEventListener("click", downloadExcelReport);
["business-type", "loan-type", "requirement"].forEach((id) => {
  const el = document.getElementById(id);
  if (el) el.addEventListener("change", updateFormVisibility);
});
document.getElementById("download").addEventListener("click", async () => {
  if (!model.length) run();
  await downloadDisplayPdf();
});

document.querySelectorAll("#wizard-progress .step").forEach((btn) => {
  btn.addEventListener("click", () => {
    applyWizardStep(Number(btn.dataset.step || 1));
  });
});

document.getElementById("wizard-back").addEventListener("click", () => {
  applyWizardStep(currentWizardStep - 1);
});

document.getElementById("wizard-next").addEventListener("click", () => {
  if (currentWizardStep === TOTAL_WIZARD_STEPS) return;
  applyWizardStep(currentWizardStep + 1);
});

INPUT_IDS.forEach((id) => {
  const el = document.getElementById(id);
  if (el) el.addEventListener("change", run);
});

updateFormVisibility();
applyWizardStep(1);
run();

document.querySelectorAll(".nav-item").forEach((item) => {
  item.addEventListener("click", (event) => {
    event.preventDefault();
    document.querySelectorAll(".nav-item").forEach((nav) => nav.classList.remove("active"));
    item.classList.add("active");
  });
});
