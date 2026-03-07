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
  { key: "sales", label: "Sales", section: "pl", aliases: ["salesaccounts", "sales", "turnover", "salese", "revenuefromoperations"] },
  { key: "openingStock", label: "Opening Stock", section: "pl", aliases: ["openingstock", "opening inventory", "openingstockinhand"] },
  { key: "purchases", label: "Purchases", section: "pl", aliases: ["purchaseaccounts", "purchases", "purchaseaccount"] },
  { key: "closingStock", label: "Closing Stock / Inventory", section: "pl", aliases: ["closingstock", "inventory", "stockinhand"] },
  { key: "directExpenses", label: "Direct Expenses", section: "pl", aliases: ["directexpenses", "carriageinward", "powerandfuel", "manufacturingexpenses"] },
  { key: "employeeCost", label: "Employee Cost", section: "pl", aliases: ["salary", "wages", "employeebenefits", "staffwelfare"] },
  { key: "adminExpenses", label: "Administrative Expenses", section: "pl", aliases: ["administrativeexpenses", "officeexpenses", "generalandadmin"] },
  { key: "sellingExpenses", label: "Selling Expenses", section: "pl", aliases: ["sellingexpenses", "marketing", "freightoutward"] },
  { key: "interest", label: "Interest", section: "pl", aliases: ["interest", "financecost", "bankinterest"] },
  { key: "depreciation", label: "Depreciation", section: "pl", aliases: ["depreciation"] },
  { key: "tax", label: "Tax", section: "pl", aliases: ["provisionfortax", "incometax", "tax"] },
  { key: "tradeReceivables", label: "Trade Receivables", section: "bs", aliases: ["sundrydebtors", "tradereceivables", "debtors"] },
  { key: "tradeCreditors", label: "Trade Creditors", section: "bs", aliases: ["sundrycreditors", "tradecreditors", "creditors"] },
  { key: "workingCapitalBorrowings", label: "Working Capital Borrowings", section: "bs", aliases: ["bankod", "cashcredit", "workingcapitalborrowing"] },
  { key: "capital", label: "Capital", section: "bs", aliases: ["capitalaccount", "capital", "partnerscapital"] },
  { key: "reserves", label: "Reserves", section: "bs", aliases: ["reserve", "surplus", "retainedearnings", "placcount"] },
  { key: "termLoan", label: "Term Loan", section: "bs", aliases: ["termloan", "vehicleloan", "machineryloan", "securedloan"] },
  { key: "unsecuredLoans", label: "Unsecured Loans", section: "bs", aliases: ["unsecuredloan", "loanfromdirectors"] },
  { key: "otherCurrentLiabilities", label: "Other Current Liabilities", section: "bs", aliases: ["othercurrentliabilities", "dutiesandtaxes", "statutorydues"] },
  { key: "fixedAssets", label: "Fixed Assets", section: "bs", aliases: ["fixedassets", "plantandmachinery", "tangibleassets"] },
  { key: "investments", label: "Investments", section: "bs", aliases: ["investments"] },
  { key: "cashBank", label: "Cash & Bank", section: "bs", aliases: ["cashatbank", "cashinhand", "bankbalance", "cashbank"] },
  { key: "loansAdvances", label: "Loans & Advances", section: "bs", aliases: ["loansandadvances", "advancesrecoverable"] },
  { key: "otherCurrentAssets", label: "Other Current Assets", section: "bs", aliases: ["othercurrentassets", "prepaidexpenses"] },
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
function fmtNumber(v) { return v === null || v === undefined ? "" : (v || 0).toFixed(2); }

function detectSheetByContent(workbook, type) {
  const hintMap = {
    pl: ["profitandloss", "trading", "sales", "purchase", "grossprofit", "indirectexpenses"],
    bs: ["balancesheet", "assets", "liabilities", "capital", "sundrydebtors", "sundrycreditors"],
  };
  const ranked = workbook.SheetNames.map((sheetName) => {
    const normName = normalize(sheetName);
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "" });
    const content = rows.flat().map(normalize).join("|");
    let score = 0;
    if (type === "pl" && /pl|profit|trading/.test(normName)) score += 8;
    if (type === "bs" && /bs|balance|position/.test(normName)) score += 8;
    hintMap[type].forEach((hint) => { if (content.includes(hint)) score += 2; });
    return { sheetName, score };
  }).sort((a, b) => b.score - a.score);
  return ranked[0]?.score > 2 ? ranked[0].sheetName : null;
}

function extractHeadsFromTwoSidedSheet(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const records = [];

  rows.forEach((row, rowIndex) => {
    const numericCols = row.map((value, colIndex) => ({ colIndex, value: getNumeric(value, NaN) })).filter((x) => Number.isFinite(x.value));
    row.forEach((value, colIndex) => {
      if (typeof value !== "string" || value.trim().length < 3) return;
      const head = value.trim().replace(/\s+/g, " ");
      const normalized = normalize(head);
      if (!normalized || normalized === "particulars" || normalized.includes("openingbalance") || normalized.includes("closingbalance") || normalized.includes("total")) return;

      const amountCell = numericCols
        .filter((n) => Math.abs(n.colIndex - colIndex) <= 3 && n.colIndex > colIndex)
        .sort((a, b) => a.colIndex - b.colIndex)[0];
      if (!amountCell || Math.abs(amountCell.value) < 1) return;

      records.push({ head, normalized, amount: Math.abs(amountCell.value), rowIndex });
    });
  });

  const latestByHead = new Map();
  records.forEach((entry) => latestByHead.set(entry.normalized, entry));
  return Array.from(latestByHead.values());
}

function bestMatchForField(field, heads) {
  let best = { head: "", score: 0 };
  heads.forEach((headRow) => {
    let score = 0;
    field.aliases.forEach((alias) => {
      const aliasNorm = normalize(alias);
      if (headRow.normalized === aliasNorm) score += 1;
      else if (headRow.normalized.includes(aliasNorm) || aliasNorm.includes(headRow.normalized)) score += 0.7;
    });
    if (score > best.score) best = { head: headRow.head, score };
  });
  return best;
}

function autoMapHeads(parsed) {
  const mapping = {};
  const confidence = {};
  MAP_FIELDS.forEach((field) => {
    const sourceHeads = field.section === "pl" ? parsed.plHeads : parsed.bsHeads;
    const best = bestMatchForField(field, sourceHeads);
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

    const scopedHeads = heads.filter((h) => h.section === field.section);
    select.innerHTML = `<option value="">-- Not mapped --</option>${scopedHeads.map((h) => `<option value="${h.head}">${h.head} (${fmtCurrency(h.amount)})</option>`).join("")}`;
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

  const parsed = { plSheetName, bsSheetName, plHeads, bsHeads };
  const { mapping, confidence } = autoMapHeads(parsed);
  const allHeads = [
    ...plHeads.map((h) => ({ ...h, section: "pl" })),
    ...bsHeads.map((h) => ({ ...h, section: "bs" })),
  ];
  const needsReview = buildMappingReview(allHeads, mapping, confidence);

  return { ...parsed, allHeads, needsReview };
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
      output.innerHTML = `<p class="good">Workbook parsed successfully. Review the mapping section (if shown) before generating the CMA report.</p>`;
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

function getMappedAmount(selectedHead, section) {
  if (!selectedHead || !workbookState.parsed) return 0;
  return workbookState.parsed.allHeads.find((h) => h.head === selectedHead && h.section === section)?.amount || 0;
}

function mappedFinancialsFromParse() {
  if (!workbookState.parsed) throw new Error("Upload and parse workbook first.");

  const mapped = {};
  MAP_FIELDS.forEach((field) => {
    mapped[field.key] = getMappedAmount(workbookState.mapping[field.key], field.section);
  });

  const mandatory = ["sales", "purchases", "closingStock", "tradeReceivables", "tradeCreditors", "capital"];
  const missing = mandatory.filter((key) => !mapped[key]);
  if (missing.length) throw new Error(`Missing mapped values for: ${missing.join(", ")}. Complete mapping review.`);

  return mapped;
}

function projectionPeriods() {
  const years = Math.max(1, Math.min(7, getConfigNumber("projection-years-cc", 5)));
  return Array.from({ length: years }, (_, idx) => `FY${idx + 1}`);
}

function buildTermLoanSchedule(periods, fallbackAmount = 0) {
  const amount = getConfigNumber("tl-amount", 0) || fallbackAmount;
  if (amount <= 0) return { applicable: false, rows: [] };

  const rate = getConfigNumber("tl-rate", 11) / 100;
  const repaymentYears = Math.max(1, getConfigNumber("tl-years", periods.length));
  const moratoriumYears = Math.max(0, getConfigNumber("moratorium-years", 0));
  const annualInstallment = amount / repaymentYears;

  let opening = amount;
  const rows = periods.map((year, idx) => {
    const installment = idx < moratoriumYears ? 0 : Math.min(annualInstallment, opening);
    const interest = opening * rate;
    const closing = Math.max(opening - installment, 0);
    const row = { year, opening, installment, interest, closing };
    opening = closing;
    return row;
  });

  return { applicable: true, rows };
}

function buildCmaReport(mapped) {
  const periods = projectionPeriods();
  const salesGrowth = getConfigNumber("sales-growth", 10) / 100;
  const grossMarginPct = getConfigNumber("gross-margin", 28) / 100;
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
  const proposedCCInput = getConfigNumber("proposed-cc-limit", 0);
  const proposedCCLimit = proposedCCInput > 0 ? proposedCCInput : existingCCLimit;
  const wcMethod = getConfigText("wc-method", "max");

  const termLoan = buildTermLoanSchedule(periods, mapped.termLoan || 0);
  const years = [];

  periods.forEach((period, idx) => {
    const prev = years[idx - 1];
    const sales = idx === 0 ? mapped.sales : prev.pl.Sales * (1 + salesGrowth);
    const openingStock = idx === 0 ? (mapped.openingStock || Math.max(mapped.closingStock * 0.9, 0)) : prev.pl["Closing Stock"];
    const purchases = idx === 0 ? mapped.purchases : prev.pl.Purchases * (1 + salesGrowth * 0.95);
    const directExpenses = idx === 0 ? (mapped.directExpenses || sales * 0.03) : prev.pl["Direct Expenses"] * (1 + salesGrowth * 0.8);
    const closingStock = idx === 0 ? mapped.closingStock : (sales * inventoryDays) / 365;

    const cogs = openingStock + purchases + directExpenses - closingStock;
    const grossProfit = Math.max(sales - cogs, sales * grossMarginPct);
    const employeeCost = idx === 0 && mapped.employeeCost ? mapped.employeeCost : sales * employeePct;
    const adminExpenses = idx === 0 && mapped.adminExpenses ? mapped.adminExpenses : sales * adminPct;
    const sellingExpenses = idx === 0 && mapped.sellingExpenses ? mapped.sellingExpenses : sales * sellingPct;
    const ebitda = grossProfit - employeeCost - adminExpenses - sellingExpenses;

    const wcInterest = existingCCLimit * 0.11;
    const termLoanInterest = termLoan.applicable ? termLoan.rows[idx].interest : 0;
    const interest = Math.max(mapped.interest || 0, wcInterest) + termLoanInterest;
    const depreciation = idx === 0
      ? (mapped.depreciation || Math.max(mapped.fixedAssets, sales * 0.15) * deprPct)
      : Math.max(prev.bs["Fixed Assets"] * deprPct, sales * 0.012);

    const pbt = ebitda + (sales * otherIncomePct) - interest - depreciation;
    const tax = Math.max(pbt, 0) * taxPct;
    const pat = pbt - tax;

    const capital = idx === 0 ? mapped.capital : prev.bs.Capital;
    const reserves = idx === 0 ? (mapped.reserves || capital * 0.15) : prev.bs.Reserves + prev.pl["Profit After Tax"] * 0.65;
    const netWorth = capital + reserves;

    const termLoanOutstanding = termLoan.applicable ? termLoan.rows[idx].opening : 0;
    const tradeCreditors = idx === 0 ? mapped.tradeCreditors : (purchases * creditorDays) / 365;
    const otherCurrentLiabilities = idx === 0 ? (mapped.otherCurrentLiabilities || sales * 0.01) : sales * 0.01;

    const inventory = closingStock;
    const tradeReceivables = idx === 0 ? mapped.tradeReceivables : (sales * debtorDays) / 365;
    const cashBank = idx === 0 ? (mapped.cashBank || sales * 0.02) : Math.max(prev.bs["Cash & Bank"] + pat * 0.1, 0);
    const loansAdvances = idx === 0 ? (mapped.loansAdvances || sales * 0.01) : prev.bs["Loans & Advances"];
    const otherCurrentAssets = idx === 0 ? (mapped.otherCurrentAssets || sales * 0.01) : sales * 0.01;
    const fixedAssets = idx === 0 ? (mapped.fixedAssets || sales * 0.2) : Math.max(prev.bs["Fixed Assets"] * 0.95, sales * 0.18);
    const investments = idx === 0 ? (mapped.investments || 0) : prev.bs.Investments;

    const nayakRequired = sales * 0.2;
    const cycleReceivables = (sales * debtorDays) / 365;
    const cycleInventory = (cogs * inventoryDays) / 365;
    const cycleCreditors = (purchases * creditorDays) / 365;
    const cycleRequired = Math.max(cycleReceivables + cycleInventory - cycleCreditors, 0);
    const requiredFinance = wcMethod === "nayak"
      ? nayakRequired
      : wcMethod === "cycle"
        ? cycleRequired
        : Math.max(nayakRequired, cycleRequired);

    const sanctionedCcLimit = idx === 0
      ? Math.min(existingCCLimit, proposedCCLimit)
      : Math.max(existingCCLimit, proposedCCLimit > existingCCLimit ? proposedCCLimit : existingCCLimit);

    const workingCapitalBorrowings = Math.min(requiredFinance, sanctionedCcLimit);
    const unsecuredLoans = idx === 0 ? mapped.unsecuredLoans : prev.bs["Unsecured Loans"];

    const totalLiabilities = netWorth + termLoanOutstanding + workingCapitalBorrowings + unsecuredLoans + tradeCreditors + otherCurrentLiabilities;
    const totalAssets = fixedAssets + investments + inventory + tradeReceivables + cashBank + loansAdvances + otherCurrentAssets;

    const installment = termLoan.applicable ? termLoan.rows[idx].installment : 0;
    const dscr = termLoan.applicable ? safeDivide((pat + depreciation + termLoanInterest), (installment + termLoanInterest)) : null;

    years.push({
      year: period,
      pl: {
        Sales: sales,
        "Opening Stock": openingStock,
        Purchases: purchases,
        "Direct Expenses": directExpenses,
        "Closing Stock": closingStock,
        "Cost of Goods Sold": cogs,
        "Gross Profit": grossProfit,
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
        "Term Loan": termLoanOutstanding,
        "Working Capital Borrowings": workingCapitalBorrowings,
        "Unsecured Loans": unsecuredLoans,
        "Trade Creditors": tradeCreditors,
        "Other Current Liabilities": otherCurrentLiabilities,
        "Total Liabilities": totalLiabilities,
        "Fixed Assets": fixedAssets,
        Investments: investments,
        Inventory: inventory,
        "Trade Receivables": tradeReceivables,
        "Cash & Bank": cashBank,
        "Loans & Advances": loansAdvances,
        "Other Current Assets": otherCurrentAssets,
        "Total Assets": totalAssets,
      },
      workingCapital: {
        projectedSales: sales,
        nayakRequired,
        cycleRequired,
        requiredFinance,
        existingCCLimit,
        sanctionedCcLimit,
        shortfall: Math.max(requiredFinance - sanctionedCcLimit, 0),
      },
      dscr,
    });
  });

  const dscrValues = years.map((row) => row.dscr).filter((x) => x !== null);
  const avgDscr = termLoan.applicable ? safeDivide(dscrValues.reduce((sum, value) => sum + value, 0), dscrValues.length) : null;

  return {
    periods,
    mappedFinancials: mapped,
    years,
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
  const head = `<tr><th>Particulars</th>${periods.map((period) => `<th>${period}</th>`).join("")}</tr>`;
  const body = rows.map((row) => `<tr><td>${row.label}</td>${row.values.map((value) => `<td>${formatter(value, row.label)}</td>`).join("")}</tr>`).join("");
  return `<section class="report-section"><h3>${title}</h3><table><thead>${head}</thead><tbody>${body}</tbody></table></section>`;
}

function renderReport(report) {
  const periods = report.periods;
  const summaryRows = Object.entries(report.borrowerSummary)
    .map(([key, value]) => `<tr><td>${key}</td><td>${typeof value === "number" ? fmtCurrency(value) : value}</td></tr>`)
    .join("");

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

  const limitRows = [
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
  const ratioTableRows = ratioLabels.map((label) => ({ label, values: ratioRows.map((row) => row[label]) }));

  const termLoanHtml = report.termLoan.applicable
    ? buildSectionHtml(
      "Term Loan Schedule",
      ["Opening Balance", "Installment", "Interest", "Closing Balance"],
      report.termLoan.rows.map((row) => ({
        label: row.year,
        values: [row.opening, row.installment, row.interest, row.closing],
      })),
    )
    : `<section class="report-section"><h3>Term Loan</h3><p class="hint">Term Loan not applicable.</p></section>`;

  output.innerHTML = `
    <article class="cma-report">
      <header class="report-header"><h2>Banker Grade CMA Report</h2></header>
      <section class="report-section"><h3>Summary</h3><table><tbody>${summaryRows}</tbody></table></section>
      ${buildSectionHtml("Profit & Loss", periods, plRows)}
      ${buildSectionHtml("Balance Sheet", periods, bsRows)}
      ${buildSectionHtml("Working Capital Limit Comparison", periods, limitRows)}
      ${termLoanHtml}
      ${buildSectionHtml("Ratio Analysis", periods, ratioTableRows, (value, label) => {
    if (label.includes("Days")) return fmtNumber(value);
    if (label === "DSCR" || label.includes("Ratio") || label === "TOL/TNW") return fmtNumber(value);
    return fmtPct(value);
  })}
      <section class="report-section"><h3>DSCR Status</h3>
        <p><strong>Average DSCR:</strong> ${report.avgDscr === null ? "Term Loan not applicable" : fmtNumber(report.avgDscr)}
          ${report.avgDscr === null ? "" : report.avgDscr >= 1.25 ? "<span class='good'> (Acceptable)</span>" : "<span class='bad'> (Warning)</span>"}
        </p>
      </section>
    </article>`;
}

function generateReport() {
  try {
    if (!workbookState.parsed) throw new Error(workbookState.parseError || "Upload workbook first.");
    if (workbookState.parsed.needsReview) {
      const unmapped = MAP_FIELDS.filter((f) => !workbookState.mapping[f.key]).map((f) => f.label);
      if (unmapped.length) throw new Error(`Please complete mapping review. Unmapped heads: ${unmapped.join(", ")}.`);
    }

    const mapped = mappedFinancialsFromParse();
    workbookState.generated = {
      meta: {
        sourcePL: workbookState.parsed.plSheetName,
        sourceBS: workbookState.parsed.bsSheetName,
        generatedAt: new Date().toISOString(),
      },
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
  const html = `<!doctype html><html><head><meta charset="UTF-8"><title>CMA Report</title><link rel="stylesheet" href="styles.css"></head><body>${output.innerHTML}</body></html>`;
  saveFile(html, "cma-report.html", "text/html");
}

function toCell(row, col) { return XLSX.utils.encode_cell({ r: row - 1, c: col - 1 }); }

function makeSheet(title, periods, rowLabels, rowsByYear) {
  const ws = XLSX.utils.aoa_to_sheet([
    [title],
    ["Particulars", ...periods],
    ...rowLabels.map((label) => [label, ...rowsByYear.map((row) => row[label] ?? "")]),
  ]);
  ws["!cols"] = Array.from({ length: periods.length + 1 }, (_, idx) => ({ wch: idx === 0 ? 34 : 15 }));
  ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: periods.length } }];
  return ws;
}

function setFormula(ws, row, col, formula, format) {
  ws[toCell(row, col)] = { t: "n", f: formula, z: format || undefined };
}

function applyCurrencyFormat(ws) {
  const range = XLSX.utils.decode_range(ws["!ref"]);
  for (let r = range.s.r + 2; r <= range.e.r; r += 1) {
    for (let c = range.s.c + 1; c <= range.e.c; c += 1) {
      const address = XLSX.utils.encode_cell({ r, c });
      if (ws[address] && ws[address].t === "n") ws[address].z = "₹#,##0";
    }
  }
}

function downloadExcel() {
  if (!workbookState.generated) return;

  const wb = XLSX.utils.book_new();
  const periods = workbookState.generated.periods;
  const years = workbookState.generated.years;

  const assumptions = XLSX.utils.aoa_to_sheet([
    ["Assumptions"],
    ["Parameter", "Value"],
    ["Sales Growth (%)", getConfigNumber("sales-growth", 10)],
    ["Gross Margin (%)", getConfigNumber("gross-margin", 28)],
    ["Debtor Days", getConfigNumber("receivable-days", 60)],
    ["Inventory Days", getConfigNumber("inventory-days", 45)],
    ["Creditor Days", getConfigNumber("payable-days", 30)],
    ["Existing CC Limit", getConfigNumber("cc-limit", 0)],
    ["Proposed CC Limit", getConfigNumber("proposed-cc-limit", 0)],
  ]);

  const plLabels = [
    "Sales", "Opening Stock", "Purchases", "Direct Expenses", "Closing Stock",
    "Cost of Goods Sold", "Gross Profit", "Employee Cost", "Administrative Expenses",
    "Selling Expenses", "Operating Profit / EBITDA", "Interest", "Depreciation",
    "Profit Before Tax", "Tax", "Profit After Tax",
  ];

  const bsLabels = [
    "Capital", "Reserves", "Net Worth", "Term Loan", "Working Capital Borrowings", "Unsecured Loans",
    "Trade Creditors", "Other Current Liabilities", "Total Liabilities", "Fixed Assets", "Investments",
    "Inventory", "Trade Receivables", "Cash & Bank", "Loans & Advances", "Other Current Assets", "Total Assets",
  ];

  const plWs = makeSheet("Profit & Loss", periods, plLabels, years.map((y) => y.pl));
  const bsWs = makeSheet("Balance Sheet", periods, bsLabels, years.map((y) => y.bs));

  periods.forEach((_, idx) => {
    const col = idx + 2;
    setFormula(plWs, 8, col, `${toCell(3, col)}+${toCell(4, col)}+${toCell(5, col)}-${toCell(6, col)}`, "₹#,##0");
    setFormula(plWs, 9, col, `${toCell(2, col)}-${toCell(8, col)}`, "₹#,##0");
    setFormula(plWs, 13, col, `${toCell(9, col)}-${toCell(10, col)}-${toCell(11, col)}-${toCell(12, col)}`, "₹#,##0");
    setFormula(plWs, 16, col, `${toCell(13, col)}-${toCell(14, col)}-${toCell(15, col)}`, "₹#,##0");
    setFormula(plWs, 18, col, `${toCell(16, col)}-${toCell(17, col)}`, "₹#,##0");

    setFormula(bsWs, 5, col, `${toCell(3, col)}+${toCell(4, col)}`, "₹#,##0");
    setFormula(bsWs, 11, col, `${toCell(5, col)}+${toCell(6, col)}+${toCell(7, col)}+${toCell(8, col)}+${toCell(9, col)}+${toCell(10, col)}`, "₹#,##0");
    setFormula(bsWs, 19, col, `${toCell(12, col)}+${toCell(13, col)}+${toCell(14, col)}+${toCell(15, col)}+${toCell(16, col)}+${toCell(17, col)}+${toCell(18, col)}`, "₹#,##0");
  });

  applyCurrencyFormat(plWs);
  applyCurrencyFormat(bsWs);

  const wcWs = XLSX.utils.aoa_to_sheet([
    ["Working Capital"],
    ["Year", "Projected Sales", "Nayak Required", "WC Cycle Required", "Required WC Finance", "Existing CC Limit", "Shortfall"],
    ...years.map((y) => [
      y.year,
      y.workingCapital.projectedSales,
      y.workingCapital.nayakRequired,
      y.workingCapital.cycleRequired,
      y.workingCapital.requiredFinance,
      y.workingCapital.existingCCLimit,
      y.workingCapital.shortfall,
    ]),
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
    ["Year", "PAT", "Depreciation", "Interest on Term Loan", "Installment", "DSCR", "Status"],
    ...years.map((yearData, idx) => {
      const tlInterest = workbookState.generated.termLoan.applicable ? workbookState.generated.termLoan.rows[idx].interest : 0;
      const installment = workbookState.generated.termLoan.applicable ? workbookState.generated.termLoan.rows[idx].installment : 0;
      return [
        yearData.year,
        yearData.pl["Profit After Tax"],
        yearData.pl.Depreciation,
        tlInterest,
        installment,
        yearData.dscr || "",
        yearData.dscr === null ? "N/A" : (yearData.dscr >= 1.25 ? "Acceptable" : "Warning"),
      ];
    }),
    ["Average DSCR", "", "", "", "", workbookState.generated.avgDscr || "", workbookState.generated.avgDscr === null ? "N/A" : (workbookState.generated.avgDscr >= 1.25 ? "Acceptable" : "Warning")],
  ]);

  const ratioLabels = [
    "Current Ratio", "Debt Equity Ratio", "TOL/TNW", "GP Ratio", "NP Ratio", "EBITDA Margin",
    "Interest Coverage Ratio", "Debtor Days", "Creditor Days", "Inventory Days", "DSCR",
  ];
  const ratioWs = XLSX.utils.aoa_to_sheet([["Ratio Analysis"], ["Particulars", ...periods], ...ratioLabels.map((r) => [r, ...periods.map(() => "")])]);

  periods.forEach((_, idx) => {
    const c = idx + 2;
    setFormula(ratioWs, 3, c, `('Balance Sheet'!${toCell(14, c)}+'Balance Sheet'!${toCell(15, c)}+'Balance Sheet'!${toCell(16, c)}+'Balance Sheet'!${toCell(17, c)}+'Balance Sheet'!${toCell(18, c)})/('Balance Sheet'!${toCell(8, c)}+'Balance Sheet'!${toCell(9, c)})`, "0.00");
    setFormula(ratioWs, 4, c, `('Balance Sheet'!${toCell(6, c)}+'Balance Sheet'!${toCell(7, c)}+'Balance Sheet'!${toCell(8, c)})/'Balance Sheet'!${toCell(5, c)}`, "0.00");
    setFormula(ratioWs, 5, c, `('Balance Sheet'!${toCell(6, c)}+'Balance Sheet'!${toCell(7, c)}+'Balance Sheet'!${toCell(8, c)}+'Balance Sheet'!${toCell(9, c)}+'Balance Sheet'!${toCell(10, c)})/'Balance Sheet'!${toCell(5, c)}`, "0.00");
    setFormula(ratioWs, 6, c, `'Profit & Loss'!${toCell(9, c)}/'Profit & Loss'!${toCell(2, c)}*100`, "0.00");
    setFormula(ratioWs, 7, c, `'Profit & Loss'!${toCell(18, c)}/'Profit & Loss'!${toCell(2, c)}*100`, "0.00");
    setFormula(ratioWs, 8, c, `'Profit & Loss'!${toCell(13, c)}/'Profit & Loss'!${toCell(2, c)}*100`, "0.00");
    setFormula(ratioWs, 9, c, `'Profit & Loss'!${toCell(13, c)}/'Profit & Loss'!${toCell(14, c)}`, "0.00");
    setFormula(ratioWs, 10, c, `'Balance Sheet'!${toCell(15, c)}/'Profit & Loss'!${toCell(2, c)}*365`, "0.00");
    setFormula(ratioWs, 11, c, `'Balance Sheet'!${toCell(8, c)}/'Profit & Loss'!${toCell(4, c)}*365`, "0.00");
    setFormula(ratioWs, 12, c, `'Balance Sheet'!${toCell(14, c)}/'Profit & Loss'!${toCell(8, c)}*365`, "0.00");
    setFormula(ratioWs, 13, c, `'DSCR'!${toCell(3 + idx, 6)}`, "0.00");
  });

  const summaryWs = XLSX.utils.aoa_to_sheet([
    ["CMA Summary"],
    ...Object.entries(workbookState.generated.borrowerSummary).map(([k, v]) => [k, v]),
    ["Average DSCR", workbookState.generated.avgDscr || "Term Loan not applicable"],
  ]);

  XLSX.utils.book_append_sheet(wb, summaryWs, "Summary");
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
