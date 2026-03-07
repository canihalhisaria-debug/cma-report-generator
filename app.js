const workbookState = {
  workbook: null,
  parsed: null,
  mapping: {},
  generated: null,
  parseError: "",
};

const fileInput = document.getElementById("excel-file");
const sheetStatus = document.getElementById("sheet-status");
const generateBtn = document.getElementById("generate-btn");
const output = document.getElementById("output");
const downloadReportBtn = document.getElementById("download-report");
const downloadExcelBtn = document.getElementById("download-excel");
const downloadJsonBtn = document.getElementById("download-json");
const mappingReview = document.getElementById("mapping-review");
const mappingGrid = document.getElementById("mapping-grid");

const configIds = [
  "borrower-name", "constitution", "pan", "gstin", "bank-name", "branch-name", "existing-limit", "proposed-limit", "facility-type",
  "prepared-date", "projection-years-cc", "sales-growth", "other-income-ratio", "gross-margin", "employee-ratio", "admin-ratio",
  "selling-ratio", "depr-ratio", "tax-rate", "receivable-days", "inventory-days", "payable-days", "cc-limit", "proposed-cc-limit",
  "wc-method", "tl-amount", "tl-rate", "tl-years", "moratorium-years",
];
const configInputs = configIds.reduce((acc, id) => ({ ...acc, [id]: document.getElementById(id) }), {});

if (!configInputs["prepared-date"].value) configInputs["prepared-date"].value = new Date().toISOString().slice(0, 10);

const MAP_FIELDS = [
  { key: "sales", label: "Sales", aliases: ["salesaccounts", "sales", "turnover", "salesaccount"] },
  { key: "openingStock", label: "Opening Stock", aliases: ["openingstock", "opening inventory"] },
  { key: "purchases", label: "Purchases", aliases: ["purchaseaccounts", "purchases", "purchaseaccount"] },
  { key: "closingStock", label: "Closing Stock / Inventory", aliases: ["closingstock", "inventory", "stockinhand"] },
  { key: "directExpenses", label: "Direct Expenses", aliases: ["directexpenses", "carriageinward", "manufacturingexpenses"] },
  { key: "employeeCost", label: "Employee Cost", aliases: ["salary", "wages", "employeebenefits", "staffwelfare"] },
  { key: "adminExpenses", label: "Administrative Expenses", aliases: ["administrativeexpenses", "officeexpenses"] },
  { key: "sellingExpenses", label: "Selling Expenses", aliases: ["sellingexpenses", "marketing", "freightoutward"] },
  { key: "interest", label: "Interest", aliases: ["interest", "financecost", "bankinterest"] },
  { key: "depreciation", label: "Depreciation", aliases: ["depreciation"] },
  { key: "tax", label: "Tax", aliases: ["provisionfortax", "tax"] },
  { key: "tradeReceivables", label: "Trade Receivables", aliases: ["sundrydebtors", "tradereceivables", "debtors"] },
  { key: "tradeCreditors", label: "Trade Creditors", aliases: ["sundrycreditors", "tradecreditors", "creditors"] },
  { key: "workingCapitalBorrowings", label: "Working Capital Borrowings", aliases: ["bankod", "cashcredit", "workingcapitalborrowing"] },
  { key: "capital", label: "Capital", aliases: ["capitalaccount", "capital"] },
  { key: "reserves", label: "Reserves", aliases: ["reserve", "surplus", "retainedearnings"] },
  { key: "termLoan", label: "Term Loan", aliases: ["termloan", "vehicleloan", "machineryloan"] },
  { key: "unsecuredLoans", label: "Unsecured Loans", aliases: ["unsecuredloan", "loanfromdirectors"] },
  { key: "otherCurrentLiabilities", label: "Other Current Liabilities", aliases: ["othercurrentliabilities", "dutiesandtaxes"] },
  { key: "fixedAssets", label: "Fixed Assets", aliases: ["fixedassets", "plantandmachinery", "tangibleassets"] },
  { key: "investments", label: "Investments", aliases: ["investments"] },
  { key: "cashBank", label: "Cash & Bank", aliases: ["cashatbank", "cashinhand", "bankbalance", "cashbank"] },
  { key: "loansAdvances", label: "Loans & Advances", aliases: ["loansandadvances", "advancesrecoverable"] },
  { key: "otherCurrentAssets", label: "Other Current Assets", aliases: ["othercurrentassets", "prepaidexpenses"] },
];

fileInput.addEventListener("change", handleWorkbookUpload);
generateBtn.addEventListener("click", generateReport);
downloadReportBtn.addEventListener("click", downloadReport);
downloadExcelBtn.addEventListener("click", downloadExcel);
downloadJsonBtn.addEventListener("click", downloadJson);

function getNumeric(value, fallback = 0) {
  if (value === null || value === undefined || value === "") return fallback;
  if (typeof value === "number") return Number.isFinite(value) ? value : fallback;
  const cleaned = String(value).replace(/[₹,\s]/g, "").replace(/\(([^)]+)\)/, "-$1");
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : fallback;
}

function getConfigNumber(id, fallback = 0) { return getNumeric(configInputs[id]?.value, fallback); }
function getConfigText(id, fallback = "") { return (configInputs[id]?.value || fallback).toString().trim(); }
function safeDivide(a, b) { return b ? a / b : 0; }
function normalize(value) { return (value || "").toString().toLowerCase().replace(/[^a-z0-9]/g, ""); }
function fmtCurrency(v) { return v === null || v === undefined ? "" : new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", maximumFractionDigits: 0 }).format(v); }
function fmtPct(v) { return `${(v || 0).toFixed(2)}%`; }
function fmtNumber(v) { return (v || 0).toFixed(2); }

function detectSheetByContent(workbook, type) {
  const scoreSheet = (sheetName) => {
    const normName = normalize(sheetName);
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "" });
    const content = rows.flat().map(normalize).join("|");
    const plHints = ["profitandloss", "trading", "purchase", "sales", "grossprofit"];
    const bsHints = ["balancesheet", "capital", "sundrydebtors", "sundrycreditors", "liabilities", "assets"];
    const hints = type === "pl" ? plHints : bsHints;
    let score = 0;
    if (type === "pl" && (normName.includes("profit") || normName.includes("pl") || normName.includes("trading"))) score += 4;
    if (type === "bs" && (normName.includes("balance") || normName === "bs")) score += 4;
    hints.forEach((hint) => { if (content.includes(hint)) score += 2; });
    return score;
  };

  const ranked = workbook.SheetNames.map((name) => ({ name, score: scoreSheet(name) })).sort((a, b) => b.score - a.score);
  return ranked[0]?.score > 0 ? ranked[0].name : null;
}

function extractHeadsFromTwoSidedSheet(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const records = [];

  rows.forEach((row, rowIdx) => {
    const numericCols = row.map((val, colIdx) => ({ colIdx, value: getNumeric(val, NaN) })).filter((x) => Number.isFinite(x.value));
    row.forEach((val, colIdx) => {
      if (typeof val !== "string" || val.trim().length < 3) return;
      const head = val.trim();
      const norm = normalize(head);
      if (!norm || norm.includes("total") || norm === "particulars") return;

      const rightNumeric = numericCols.filter((n) => n.colIdx > colIdx).sort((a, b) => a.colIdx - b.colIdx)[0];
      if (!rightNumeric) return;
      if (Math.abs(rightNumeric.value) < 1) return;

      records.push({
        head,
        normalized: norm,
        amount: Math.abs(rightNumeric.value),
        rowIdx,
      });
    });
  });

  const latest = new Map();
  records.forEach((r) => latest.set(r.normalized, r));
  return Array.from(latest.values());
}

function bestMatchForField(field, heads) {
  let best = { head: "", score: 0 };
  heads.forEach((h) => {
    let score = 0;
    field.aliases.forEach((alias) => {
      const nAlias = normalize(alias);
      if (h.normalized === nAlias) score += 1;
      else if (h.normalized.includes(nAlias) || nAlias.includes(h.normalized)) score += 0.75;
    });
    if (score > best.score) best = { head: h.head, score };
  });
  return best;
}

function autoMapHeads(heads) {
  const mapping = {};
  const confidence = {};
  MAP_FIELDS.forEach((field) => {
    const best = bestMatchForField(field, heads);
    mapping[field.key] = best.head || "";
    confidence[field.key] = Math.min(1, best.score);
  });
  return { mapping, confidence };
}

function buildMappingReview(heads, mapping, confidence) {
  mappingGrid.innerHTML = "";
  let needsReview = false;
  MAP_FIELDS.forEach((field) => {
    const wrap = document.createElement("label");
    wrap.innerHTML = `<span>${field.label}</span>`;
    const select = document.createElement("select");
    select.id = `map-${field.key}`;
    select.innerHTML = `<option value="">-- Not mapped --</option>${heads.map((h) => `<option value="${h.head}">${h.head} (${fmtCurrency(h.amount)})</option>`).join("")}`;
    select.value = mapping[field.key] || "";
    select.addEventListener("change", () => {
      workbookState.mapping[field.key] = select.value;
    });
    wrap.appendChild(select);
    if ((confidence[field.key] || 0) < 0.75) {
      needsReview = true;
      const warn = document.createElement("small");
      warn.className = "bad";
      warn.textContent = "Low confidence mapping";
      wrap.appendChild(warn);
    }
    mappingGrid.appendChild(wrap);
    workbookState.mapping[field.key] = select.value;
  });

  if (needsReview) mappingReview.classList.remove("hidden");
  else mappingReview.classList.add("hidden");
  return needsReview;
}

function parseWorkbook(workbook) {
  const plSheetName = detectSheetByContent(workbook, "pl");
  const bsSheetName = detectSheetByContent(workbook, "bs");
  if (!plSheetName || !bsSheetName) throw new Error("Could not detect Profit & Loss and Balance Sheet sheets.");

  const plHeads = extractHeadsFromTwoSidedSheet(workbook.Sheets[plSheetName]);
  const bsHeads = extractHeadsFromTwoSidedSheet(workbook.Sheets[bsSheetName]);
  if (!plHeads.length || !bsHeads.length) throw new Error("Could not parse Tally left-right format statements.");

  const allHeads = [...plHeads, ...bsHeads];
  const { mapping, confidence } = autoMapHeads(allHeads);
  const needsReview = buildMappingReview(allHeads, mapping, confidence);

  return { plSheetName, bsSheetName, plHeads, bsHeads, allHeads, needsReview };
}

function handleWorkbookUpload(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      workbookState.workbook = XLSX.read(e.target.result, { type: "array" });
      workbookState.parsed = parseWorkbook(workbookState.workbook);
      workbookState.parseError = "";
      sheetStatus.innerHTML = `<span class="good">Detected P&L: <strong>${workbookState.parsed.plSheetName}</strong> | Balance Sheet: <strong>${workbookState.parsed.bsSheetName}</strong></span>`;
      output.innerHTML = `<p class="good">Workbook parsed. Review mapping if any low confidence fields are highlighted.</p>`;
      generateBtn.disabled = false;
    } catch (error) {
      workbookState.parsed = null;
      workbookState.parseError = error.message;
      generateBtn.disabled = true;
      mappingReview.classList.add("hidden");
      sheetStatus.innerHTML = `<span class="bad">Upload parsing failed: ${error.message}</span>`;
      output.innerHTML = `<p class="bad">Unable to generate CMA. Please upload valid Tally final accounts workbook.</p>`;
    }
  };
  reader.readAsArrayBuffer(file);
}

function getMappedAmount(selectedHead) {
  if (!selectedHead) return 0;
  return workbookState.parsed.allHeads.find((h) => h.head === selectedHead)?.amount || 0;
}

function mappedFinancialsFromParse() {
  if (!workbookState.parsed) throw new Error("Upload and parse workbook first.");
  const mapped = {};
  MAP_FIELDS.forEach((field) => { mapped[field.key] = getMappedAmount(workbookState.mapping[field.key]); });

  const mandatory = ["sales", "purchases", "closingStock", "tradeReceivables", "tradeCreditors", "capital"];
  const missing = mandatory.filter((k) => !mapped[k]);
  if (missing.length) throw new Error(`Missing mapped values for: ${missing.join(", ")}. Complete mapping review.`);

  return mapped;
}

function projectionPeriods() {
  const years = Math.max(1, Math.min(7, getConfigNumber("projection-years-cc", 5)));
  return Array.from({ length: years }, (_, idx) => `FY${idx + 1}`);
}

function buildTermLoanSchedule(years) {
  const amount = getConfigNumber("tl-amount", 0);
  if (amount <= 0) return { applicable: false, rows: [] };

  const rate = getConfigNumber("tl-rate", 11) / 100;
  const repaymentYears = Math.max(1, getConfigNumber("tl-years", years.length));
  const moratorium = Math.max(0, getConfigNumber("moratorium-years", 0));
  const annualInstallment = amount / repaymentYears;

  let opening = amount;
  const rows = years.map((year, idx) => {
    const installment = idx < moratorium ? 0 : Math.min(annualInstallment, opening);
    const interest = opening * rate;
    const closing = Math.max(opening - installment, 0);
    const row = { year, opening, installment, interest, closing };
    opening = closing;
    return row;
  });
  return { applicable: true, rows };
}

function buildCmaReport(mapped) {
  const years = projectionPeriods();
  const salesGrowth = getConfigNumber("sales-growth", 10) / 100;
  const grossMarginPct = getConfigNumber("gross-margin", 25) / 100;
  const employeePct = getConfigNumber("employee-ratio", 6) / 100;
  const adminPct = getConfigNumber("admin-ratio", 4) / 100;
  const sellingPct = getConfigNumber("selling-ratio", 3) / 100;
  const otherIncomePct = getConfigNumber("other-income-ratio", 1.5) / 100;
  const deprPct = getConfigNumber("depr-ratio", 10) / 100;
  const taxPct = getConfigNumber("tax-rate", 25) / 100;
  const debtorDays = getConfigNumber("receivable-days", 60);
  const inventoryDays = getConfigNumber("inventory-days", 45);
  const creditorDays = getConfigNumber("payable-days", 30);
  const existingCCLimit = getConfigNumber("cc-limit", mapped.workingCapitalBorrowings || 0);
  const proposedCCLimitInput = getConfigNumber("proposed-cc-limit", 0);
  const proposedCCLimit = proposedCCLimitInput > 0 ? proposedCCLimitInput : existingCCLimit;
  const wcMethod = getConfigText("wc-method", "max");

  const termLoan = buildTermLoanSchedule(years);
  const data = [];

  years.forEach((year, idx) => {
    const prev = data[idx - 1];
    const sales = idx === 0 ? mapped.sales : prev.pl.Sales * (1 + salesGrowth);
    const openingStock = idx === 0 ? (mapped.openingStock || mapped.closingStock * 0.9) : prev.pl["Closing Stock"];
    const purchases = idx === 0 ? mapped.purchases : prev.pl.Purchases * (1 + salesGrowth * 0.9);
    const directExpenses = idx === 0 ? (mapped.directExpenses || sales * 0.03) : prev.pl["Direct Expenses"] * (1 + salesGrowth * 0.8);
    const closingStock = idx === 0 ? mapped.closingStock : (sales * inventoryDays) / 365;
    const cogs = openingStock + purchases + directExpenses - closingStock;
    const grossProfit = sales - cogs;
    const targetGP = sales * grossMarginPct;
    const grossProfitAdjusted = Math.max(grossProfit, targetGP);
    const employeeCost = idx === 0 ? (mapped.employeeCost || sales * employeePct) : sales * employeePct;
    const adminExpenses = idx === 0 ? (mapped.adminExpenses || sales * adminPct) : sales * adminPct;
    const sellingExpenses = idx === 0 ? (mapped.sellingExpenses || sales * sellingPct) : sales * sellingPct;
    const ebitda = grossProfitAdjusted - employeeCost - adminExpenses - sellingExpenses;
    const interestWC = Math.max((existingCCLimit * 0.11), mapped.interest || 0);
    const termLoanInterest = termLoan.applicable ? termLoan.rows[idx].interest : 0;
    const interest = interestWC + termLoanInterest;
    const depreciation = idx === 0 ? (mapped.depreciation || Math.max(mapped.fixedAssets, sales * 0.15) * deprPct) : Math.max(data[idx - 1].bs["Fixed Assets"], sales * 0.15) * deprPct;
    const pbt = ebitda + (sales * otherIncomePct) - interest - depreciation;
    const tax = idx === 0 ? (mapped.tax || Math.max(pbt, 0) * taxPct) : Math.max(pbt, 0) * taxPct;
    const pat = pbt - tax;

    const capital = idx === 0 ? mapped.capital : data[idx - 1].bs.Capital;
    const reserves = idx === 0 ? (mapped.reserves || Math.max(capital * 0.2, 0)) : data[idx - 1].bs.Reserves + data[idx - 1].pl["Profit After Tax"] * 0.65;
    const netWorth = capital + reserves;
    const tlOutstanding = termLoan.applicable ? termLoan.rows[idx].opening : 0;
    const wcBorrowings = idx === 0 ? (mapped.workingCapitalBorrowings || existingCCLimit * 0.85) : Math.min(data[idx - 1].workingCapital.requiredFinance, proposedCCLimit);
    const unsecuredLoans = idx === 0 ? mapped.unsecuredLoans : data[idx - 1].bs["Unsecured Loans"];
    const tradeCreditors = idx === 0 ? mapped.tradeCreditors : (purchases * creditorDays) / 365;
    const ocl = idx === 0 ? (mapped.otherCurrentLiabilities || sales * 0.01) : sales * 0.01;

    const fixedAssets = idx === 0 ? (mapped.fixedAssets || sales * 0.2) : Math.max(data[idx - 1].bs["Fixed Assets"] * 0.95, sales * 0.18);
    const investments = idx === 0 ? (mapped.investments || sales * 0.01) : data[idx - 1].bs.Investments;
    const inventory = closingStock;
    const tradeReceivables = idx === 0 ? mapped.tradeReceivables : (sales * debtorDays) / 365;
    const cashBank = idx === 0 ? (mapped.cashBank || sales * 0.02) : Math.max(data[idx - 1].bs["Cash & Bank"] + pat * 0.1, 0);
    const loansAdvances = idx === 0 ? (mapped.loansAdvances || sales * 0.015) : data[idx - 1].bs["Loans & Advances"];
    const oca = idx === 0 ? (mapped.otherCurrentAssets || sales * 0.01) : sales * 0.01;

    const totalLiabilities = netWorth + tlOutstanding + wcBorrowings + unsecuredLoans + tradeCreditors + ocl;
    const totalAssets = fixedAssets + investments + inventory + tradeReceivables + cashBank + loansAdvances + oca;

    const nayakReq = sales * 0.2;
    const receivablesWc = (sales * debtorDays) / 365;
    const inventoryWc = (cogs * inventoryDays) / 365;
    const creditorsWc = (purchases * creditorDays) / 365;
    const cycleGap = Math.max(receivablesWc + inventoryWc - creditorsWc, 0);
    const requiredFinance = wcMethod === "nayak" ? nayakReq : wcMethod === "cycle" ? cycleGap : Math.max(nayakReq, cycleGap);

    const dscrInstallment = termLoan.applicable ? termLoan.rows[idx].installment : 0;
    const dscrInterest = termLoanInterest;
    const dscr = termLoan.applicable
      ? safeDivide((pat + depreciation + termLoanInterest), (dscrInstallment + dscrInterest))
      : null;

    data.push({
      year,
      pl: {
        Sales: sales,
        "Opening Stock": openingStock,
        Purchases: purchases,
        "Direct Expenses": directExpenses,
        "Closing Stock": closingStock,
        "Cost of Goods Sold": cogs,
        "Gross Profit": grossProfitAdjusted,
        "Employee Cost": employeeCost,
        "Administrative Expenses": adminExpenses,
        "Selling Expenses": sellingExpenses,
        "Operating Profit / EBITDA": ebitda,
        Interest: interest,
        Depreciation: depreciation,
        "Profit Before Tax": pbt,
        Tax: tax,
        "Profit After Tax": pat,
      },
      bs: {
        Capital: capital,
        Reserves: reserves,
        "Net Worth": netWorth,
        "Term Loan": tlOutstanding,
        "Working Capital Borrowings": wcBorrowings,
        "Unsecured Loans": unsecuredLoans,
        "Trade Creditors": tradeCreditors,
        "Other Current Liabilities": ocl,
        "Total Liabilities": totalLiabilities,
        "Fixed Assets": fixedAssets,
        Investments: investments,
        Inventory: inventory,
        "Trade Receivables": tradeReceivables,
        "Cash & Bank": cashBank,
        "Loans & Advances": loansAdvances,
        "Other Current Assets": oca,
        "Total Assets": totalAssets,
      },
      workingCapital: {
        projectedSales: sales,
        nayakRequired: nayakReq,
        cycleRequired: cycleGap,
        requiredFinance,
        existingCCLimit,
        shortfall: Math.max(requiredFinance - existingCCLimit, 0),
      },
      dscr,
    });
  });

  const avgDscr = termLoan.applicable ? safeDivide(data.reduce((a, r) => a + (r.dscr || 0), 0), data.length) : null;

  return {
    periods: years,
    mappedFinancials: mapped,
    years: data,
    termLoan,
    avgDscr,
    borrowerSummary: {
      "Borrower Name": getConfigText("borrower-name"),
      Constitution: getConfigText("constitution"),
      PAN: getConfigText("pan"),
      GSTIN: getConfigText("gstin"),
      "Bank / Branch": `${getConfigText("bank-name")} / ${getConfigText("branch-name")}`,
      "Existing CC Limit": existingCCLimit,
      "Proposed CC Limit": proposedCCLimit,
      "Facility Type": getConfigText("facility-type"),
      "Prepared Date": getConfigText("prepared-date"),
      "WC Method": wcMethod,
    },
  };
}

function buildStatementRows(years, key, rowLabels) {
  return rowLabels.map((label) => ({ label, values: years.map((y) => y[key][label]) }));
}

function buildSectionHtml(title, periods, rows, formatter = fmtCurrency) {
  const head = `<tr><th>Particulars</th>${periods.map((p) => `<th>${p}</th>`).join("")}</tr>`;
  const body = rows.map((row) => `<tr><td>${row.label}</td>${row.values.map((v) => `<td>${formatter(v, row.label)}</td>`).join("")}</tr>`).join("");
  return `<section class="report-section"><h3>${title}</h3><table><thead>${head}</thead><tbody>${body}</tbody></table></section>`;
}

function renderReport(report) {
  const periods = report.periods;
  const summaryRows = Object.entries(report.borrowerSummary)
    .map(([k, v]) => `<tr><td>${k}</td><td>${typeof v === "number" ? fmtCurrency(v) : v}</td></tr>`).join("");

  const plRows = buildStatementRows(report.years, "pl", [
    "Sales", "Opening Stock", "Purchases", "Direct Expenses", "Closing Stock", "Cost of Goods Sold", "Gross Profit",
    "Employee Cost", "Administrative Expenses", "Selling Expenses", "Operating Profit / EBITDA", "Interest", "Depreciation",
    "Profit Before Tax", "Tax", "Profit After Tax",
  ]);

  const bsRows = buildStatementRows(report.years, "bs", [
    "Capital", "Reserves", "Net Worth", "Term Loan", "Working Capital Borrowings", "Unsecured Loans",
    "Trade Creditors", "Other Current Liabilities", "Total Liabilities", "Fixed Assets", "Investments", "Inventory",
    "Trade Receivables", "Cash & Bank", "Loans & Advances", "Other Current Assets", "Total Assets",
  ]);

  const wcRows = [
    { label: "Projected Sales", values: report.years.map((y) => y.workingCapital.projectedSales) },
    { label: "Required WC Finance", values: report.years.map((y) => y.workingCapital.requiredFinance) },
    { label: "Existing CC Limit", values: report.years.map((y) => y.workingCapital.existingCCLimit) },
    { label: "Shortfall", values: report.years.map((y) => y.workingCapital.shortfall) },
  ];

  const ratioRows = report.years.map((y) => ({
    "Current Ratio": safeDivide(y.bs.Inventory + y.bs["Trade Receivables"] + y.bs["Cash & Bank"] + y.bs["Loans & Advances"] + y.bs["Other Current Assets"], y.bs["Trade Creditors"] + y.bs["Other Current Liabilities"]),
    "Debt Equity Ratio": safeDivide(y.bs["Term Loan"] + y.bs["Working Capital Borrowings"] + y.bs["Unsecured Loans"], y.bs["Net Worth"]),
    "TOL/TNW": safeDivide(y.bs["Term Loan"] + y.bs["Working Capital Borrowings"] + y.bs["Unsecured Loans"] + y.bs["Trade Creditors"] + y.bs["Other Current Liabilities"], y.bs["Net Worth"]),
    "GP Ratio": safeDivide(y.pl["Gross Profit"], y.pl.Sales) * 100,
    "NP Ratio": safeDivide(y.pl["Profit After Tax"], y.pl.Sales) * 100,
    "EBITDA Margin": safeDivide(y.pl["Operating Profit / EBITDA"], y.pl.Sales) * 100,
    "Interest Coverage Ratio": safeDivide(y.pl["Operating Profit / EBITDA"], y.pl.Interest),
    "Debtor Days": safeDivide(y.bs["Trade Receivables"], y.pl.Sales) * 365,
    "Creditor Days": safeDivide(y.bs["Trade Creditors"], y.pl.Purchases) * 365,
    "Inventory Days": safeDivide(y.bs.Inventory, y.pl["Cost of Goods Sold"]) * 365,
    DSCR: y.dscr,
  }));

  const ratioLabels = Object.keys(ratioRows[0]);
  const ratioTableRows = ratioLabels.map((label) => ({ label, values: ratioRows.map((r) => r[label]) }));

  const tlHtml = report.termLoan.applicable
    ? buildSectionHtml("Term Loan Schedule", ["Year", "Opening Balance", "Installment", "Interest", "Closing Balance"], report.termLoan.rows.map((r) => ({
      label: r.year,
      values: [r.opening, r.installment, r.interest, r.closing],
    })), (v, l) => l === "Year" ? v : fmtCurrency(v))
    : `<section class="report-section"><h3>Term Loan</h3><p class="hint">Term Loan not applicable.</p></section>`;

  output.innerHTML = `
    <article class="cma-report">
      <header class="report-header"><h2>Banker Grade CMA Report</h2></header>
      <section class="report-section"><h3>Summary</h3><table><tbody>${summaryRows}</tbody></table></section>
      ${buildSectionHtml("Profit & Loss", periods, plRows)}
      ${buildSectionHtml("Balance Sheet", periods, bsRows)}
      ${buildSectionHtml("Working Capital Limit Comparison", periods, wcRows)}
      ${tlHtml}
      ${buildSectionHtml("Ratio Analysis", periods, ratioTableRows, (v, label) => {
    if (label.includes("Ratio") || label === "DSCR" || label === "TOL/TNW") return fmtNumber(v);
    if (label.includes("Days")) return fmtNumber(v);
    if (label.includes("Margin") || label.includes("GP") || label.includes("NP")) return fmtPct(v);
    return fmtNumber(v);
  })}
      <section class="report-section"><h3>DSCR Status</h3>
        <p><strong>Average DSCR:</strong> ${report.avgDscr === null ? "Term Loan not applicable" : fmtNumber(report.avgDscr)}
        ${report.avgDscr === null ? "" : (report.avgDscr >= 1.25 ? "<span class='good'> (Acceptable)</span>" : "<span class='bad'> (Warning)</span>")}</p>
      </section>
    </article>`;
}

function generateReport() {
  try {
    if (!workbookState.parsed) throw new Error(workbookState.parseError || "Upload workbook first.");
    const mapped = mappedFinancialsFromParse();
    workbookState.generated = {
      meta: { sourcePL: workbookState.parsed.plSheetName, sourceBS: workbookState.parsed.bsSheetName, generatedAt: new Date().toISOString() },
      ...buildCmaReport(mapped),
    };
    renderReport(workbookState.generated);
    downloadReportBtn.disabled = false;
    downloadExcelBtn.disabled = false;
    downloadJsonBtn.disabled = false;
  } catch (error) {
    output.innerHTML = `<p class="bad">Cannot generate report: ${error.message}</p>`;
  }
}

function saveFile(content, filename, mime = "text/plain") {
  const blob = new Blob([content], { type: mime });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

function downloadReport() {
  if (!workbookState.generated) return;
  const html = `<!doctype html><html><head><meta charset="UTF-8"><title>CMA Report</title><style>${document.querySelector("style")?.innerHTML || ""}</style></head><body>${output.innerHTML}</body></html>`;
  saveFile(html, "cma-report.html", "text/html");
}

function toCell(row, col) { return XLSX.utils.encode_cell({ r: row - 1, c: col - 1 }); }
function makeSheet(title, periods, rowLabels, rowsByYear) {
  const ws = XLSX.utils.aoa_to_sheet([[title], ["Particulars", ...periods], ...rowLabels.map((l) => [l, ...rowsByYear.map((y) => y[l] ?? "")])]);
  ws["!cols"] = Array.from({ length: periods.length + 1 }, (_, idx) => ({ wch: idx === 0 ? 36 : 15 }));
  return ws;
}
function setFormula(ws, row, col, formula, format) {
  ws[toCell(row, col)] = { t: "n", f: formula, z: format || undefined };
}

function downloadExcel() {
  if (!workbookState.generated) return;
  const wb = XLSX.utils.book_new();
  const periods = workbookState.generated.periods;
  const years = workbookState.generated.years;

  const assumptions = XLSX.utils.aoa_to_sheet([
    ["Assumption", "Value"],
    ["Sales Growth (%)", getConfigNumber("sales-growth", 10)],
    ["Gross Margin (%)", getConfigNumber("gross-margin", 25)],
    ["Debtor Days", getConfigNumber("receivable-days", 60)],
    ["Inventory Days", getConfigNumber("inventory-days", 45)],
    ["Creditor Days", getConfigNumber("payable-days", 30)],
    ["Existing CC Limit", getConfigNumber("cc-limit", 0)],
    ["Proposed CC Limit", getConfigNumber("proposed-cc-limit", 0)],
  ]);

  const plLabels = ["Sales", "Opening Stock", "Purchases", "Direct Expenses", "Closing Stock", "Cost of Goods Sold", "Gross Profit", "Employee Cost", "Administrative Expenses", "Selling Expenses", "Operating Profit / EBITDA", "Interest", "Depreciation", "Profit Before Tax", "Tax", "Profit After Tax"];
  const bsLabels = ["Capital", "Reserves", "Net Worth", "Term Loan", "Working Capital Borrowings", "Unsecured Loans", "Trade Creditors", "Other Current Liabilities", "Total Liabilities", "Fixed Assets", "Investments", "Inventory", "Trade Receivables", "Cash & Bank", "Loans & Advances", "Other Current Assets", "Total Assets"];

  const plWs = makeSheet("Profit & Loss", periods, plLabels, years.map((y) => y.pl));
  const bsWs = makeSheet("Balance Sheet", periods, bsLabels, years.map((y) => y.bs));

  periods.forEach((_, idx) => {
    const c = idx + 2;
    setFormula(plWs, 7, c, `${toCell(3, c)}+${toCell(4, c)}+${toCell(5, c)}-${toCell(6, c)}`, "₹#,##0");
    setFormula(plWs, 8, c, `${toCell(2, c)}-${toCell(7, c)}`, "₹#,##0");
    setFormula(plWs, 12, c, `${toCell(8, c)}-${toCell(9, c)}-${toCell(10, c)}-${toCell(11, c)}`, "₹#,##0");
    setFormula(plWs, 15, c, `${toCell(12, c)}-${toCell(13, c)}-${toCell(14, c)}`, "₹#,##0");
    setFormula(plWs, 17, c, `${toCell(15, c)}-${toCell(16, c)}`, "₹#,##0");

    setFormula(bsWs, 5, c, `${toCell(3, c)}+${toCell(4, c)}`, "₹#,##0");
    setFormula(bsWs, 11, c, `${toCell(5, c)}+${toCell(6, c)}+${toCell(7, c)}+${toCell(8, c)}+${toCell(9, c)}+${toCell(10, c)}`, "₹#,##0");
    setFormula(bsWs, 19, c, `${toCell(12, c)}+${toCell(13, c)}+${toCell(14, c)}+${toCell(15, c)}+${toCell(16, c)}+${toCell(17, c)}+${toCell(18, c)}`, "₹#,##0");
  });

  const wcWs = XLSX.utils.aoa_to_sheet([
    ["Working Capital"],
    ["Year", "Projected Sales", "Nayak Required", "Cycle Required", "Required WC Finance", "Existing CC Limit", "Shortfall"],
    ...years.map((y) => [y.year, y.workingCapital.projectedSales, y.workingCapital.nayakRequired, y.workingCapital.cycleRequired, y.workingCapital.requiredFinance, y.workingCapital.existingCCLimit, y.workingCapital.shortfall]),
  ]);

  const tlWs = workbookState.generated.termLoan.applicable
    ? XLSX.utils.aoa_to_sheet([
      ["Term Loan"],
      ["Year", "Opening Balance", "Installment", "Interest", "Closing Balance"],
      ...workbookState.generated.termLoan.rows.map((r) => [r.year, r.opening, r.installment, r.interest, r.closing]),
    ])
    : XLSX.utils.aoa_to_sheet([["Term Loan"], ["Term Loan not applicable"]]);

  const dscrWs = XLSX.utils.aoa_to_sheet([
    ["DSCR"],
    ["Year", "PAT", "Depreciation", "Interest on TL", "Installment", "DSCR", "Status"],
    ...years.map((y, idx) => {
      const tlInterest = workbookState.generated.termLoan.applicable ? workbookState.generated.termLoan.rows[idx].interest : 0;
      const installment = workbookState.generated.termLoan.applicable ? workbookState.generated.termLoan.rows[idx].installment : 0;
      return [y.year, y.pl["Profit After Tax"], y.pl.Depreciation, tlInterest, installment, y.dscr || "", y.dscr === null ? "N/A" : (y.dscr >= 1.25 ? "Acceptable" : "Warning")];
    }),
    ["Average DSCR", "", "", "", "", workbookState.generated.avgDscr || "", workbookState.generated.avgDscr >= 1.25 ? "Acceptable" : "Warning"],
  ]);

  const ratioLabels = ["Current Ratio", "Debt Equity Ratio", "TOL/TNW", "GP Ratio", "NP Ratio", "EBITDA Margin", "Interest Coverage Ratio", "Debtor Days", "Creditor Days", "Inventory Days", "DSCR"];
  const ratioWs = XLSX.utils.aoa_to_sheet([["Ratio Analysis"], ["Particulars", ...periods], ...ratioLabels.map((r) => [r, ...periods.map(() => "")])]);

  periods.forEach((_, idx) => {
    const c = idx + 2;
    setFormula(ratioWs, 3, c, `('Balance Sheet'!${toCell(14, c)}+'Balance Sheet'!${toCell(15, c)}+'Balance Sheet'!${toCell(16, c)}+'Balance Sheet'!${toCell(17, c)}+'Balance Sheet'!${toCell(18, c)})/('Balance Sheet'!${toCell(8, c)}+'Balance Sheet'!${toCell(9, c)})`, "0.00");
    setFormula(ratioWs, 4, c, `('Balance Sheet'!${toCell(6, c)}+'Balance Sheet'!${toCell(7, c)}+'Balance Sheet'!${toCell(8, c)})/'Balance Sheet'!${toCell(5, c)}`, "0.00");
    setFormula(ratioWs, 5, c, `('Balance Sheet'!${toCell(6, c)}+'Balance Sheet'!${toCell(7, c)}+'Balance Sheet'!${toCell(8, c)}+'Balance Sheet'!${toCell(9, c)}+'Balance Sheet'!${toCell(10, c)})/'Balance Sheet'!${toCell(5, c)}`, "0.00");
    setFormula(ratioWs, 6, c, `'Profit & Loss'!${toCell(8, c)}/'Profit & Loss'!${toCell(2, c)}*100`, "0.00");
    setFormula(ratioWs, 7, c, `'Profit & Loss'!${toCell(17, c)}/'Profit & Loss'!${toCell(2, c)}*100`, "0.00");
    setFormula(ratioWs, 8, c, `'Profit & Loss'!${toCell(12, c)}/'Profit & Loss'!${toCell(2, c)}*100`, "0.00");
    setFormula(ratioWs, 9, c, `'Profit & Loss'!${toCell(12, c)}/'Profit & Loss'!${toCell(13, c)}`, "0.00");
    setFormula(ratioWs, 10, c, `'Balance Sheet'!${toCell(15, c)}/'Profit & Loss'!${toCell(2, c)}*365`, "0.00");
    setFormula(ratioWs, 11, c, `'Balance Sheet'!${toCell(8, c)}/'Profit & Loss'!${toCell(4, c)}*365`, "0.00");
    setFormula(ratioWs, 12, c, `'Balance Sheet'!${toCell(14, c)}/'Profit & Loss'!${toCell(7, c)}*365`, "0.00");
    setFormula(ratioWs, 13, c, `'DSCR'!${toCell(3 + idx, 6)}`, "0.00");
  });

  const summary = XLSX.utils.aoa_to_sheet([
    ["CMA Summary"],
    ...Object.entries(workbookState.generated.borrowerSummary).map(([k, v]) => [k, v]),
    ["Average DSCR", workbookState.generated.avgDscr || "Term Loan not applicable"],
  ]);

  XLSX.utils.book_append_sheet(wb, summary, "Summary");
  XLSX.utils.book_append_sheet(wb, assumptions, "Assumptions");
  XLSX.utils.book_append_sheet(wb, plWs, "Profit & Loss");
  XLSX.utils.book_append_sheet(wb, bsWs, "Balance Sheet");
  XLSX.utils.book_append_sheet(wb, wcWs, "Working Capital");
  XLSX.utils.book_append_sheet(wb, ratioWs, "Ratio Analysis");
  XLSX.utils.book_append_sheet(wb, tlWs, "Term Loan");
  XLSX.utils.book_append_sheet(wb, dscrWs, "DSCR");

  XLSX.writeFile(wb, "banker-grade-cma-report.xlsx");
}

function downloadJson() {
  if (!workbookState.generated) return;
  saveFile(JSON.stringify(workbookState.generated, null, 2), "cma-data.json", "application/json");
}
