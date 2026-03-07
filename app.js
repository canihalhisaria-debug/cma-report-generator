const workbookState = {
  workbook: null,
  selectedSheet: null,
  sourceRows: [],
  generated: null,
  detectedHeads: [],
  mapping: {
    tradeReceivables: "",
    inventory: "",
    tradePayables: "",
  },
};

const fileInput = document.getElementById("excel-file");
const sheetStatus = document.getElementById("sheet-status");
const sheetSelector = document.getElementById("sheet-selector");
const generateBtn = document.getElementById("generate-btn");
const output = document.getElementById("output");
const downloadReportBtn = document.getElementById("download-report");
const downloadExcelBtn = document.getElementById("download-excel");
const downloadJsonBtn = document.getElementById("download-json");

const mapReceivablesSelect = document.getElementById("map-trade-receivables");
const mapInventorySelect = document.getElementById("map-inventory");
const mapPayablesSelect = document.getElementById("map-trade-payables");

const configIds = [
  "borrower-name",
  "constitution",
  "pan",
  "gstin",
  "bank-name",
  "branch-name",
  "existing-limit",
  "proposed-limit",
  "facility-type",
  "prepared-date",
  "historical-count",
  "projection-years-cc",
  "projection-years-tl",
  "moratorium-years",
  "sales-growth",
  "sales-projection-mode",
  "other-income-ratio",
  "gross-margin",
  "employee-ratio",
  "admin-ratio",
  "selling-ratio",
  "depr-ratio",
  "tax-rate",
  "receivable-days",
  "inventory-days",
  "payable-days",
  "margin-ratio",
  "cc-limit",
  "cc-utilization",
  "tl-outstanding",
  "tl-rate",
];

const configInputs = configIds.reduce((acc, id) => {
  acc[id] = document.getElementById(id);
  return acc;
}, {});

if (!configInputs["prepared-date"].value) {
  configInputs["prepared-date"].value = new Date().toISOString().slice(0, 10);
}

fileInput.addEventListener("change", handleWorkbookUpload);
sheetSelector.addEventListener("change", handleSheetSelection);
generateBtn.addEventListener("click", generateReport);
downloadReportBtn.addEventListener("click", downloadReport);
downloadExcelBtn.addEventListener("click", downloadExcel);
downloadJsonBtn.addEventListener("click", downloadJson);
[mapReceivablesSelect, mapInventorySelect, mapPayablesSelect].forEach((el) => {
  el.addEventListener("change", () => {
    workbookState.mapping.tradeReceivables = mapReceivablesSelect.value;
    workbookState.mapping.inventory = mapInventorySelect.value;
    workbookState.mapping.tradePayables = mapPayablesSelect.value;
  });
});

function getNumeric(value, fallback = 0) {
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

function getConfigNumber(id, fallback = 0) {
  return getNumeric(configInputs[id]?.value, fallback);
}

function getConfigText(id, fallback = "") {
  return (configInputs[id]?.value || fallback).toString().trim();
}

function ratio(value, percent) {
  return value * (percent / 100);
}

function safeDivide(a, b) {
  if (!b) return 0;
  return a / b;
}

function formatFY(startYear, suffix = "") {
  return `FY ${startYear}-${String(startYear + 1).slice(-2)}${suffix}`;
}

function fmtCurrency(amount) {
  if (amount === null || amount === undefined) return "";
  return new Intl.NumberFormat("en-IN", {
    style: "currency",
    currency: "INR",
    maximumFractionDigits: 0,
  }).format(amount || 0);
}

function fmtPct(value) {
  if (value === null || value === undefined) return "";
  return `${(value || 0).toFixed(2)}%`;
}

function fmtNumber(value) {
  if (value === null || value === undefined) return "";
  return (value || 0).toFixed(2);
}

function normalizeHead(value) {
  return (value || "").toString().toLowerCase().replace(/[^a-z0-9]/g, "");
}

function guessHeadText(row) {
  for (const val of Object.values(row)) {
    if (typeof val === "string" && val.trim()) return val.trim();
  }
  return "";
}

function extractAmountFromRow(row) {
  const nums = Object.values(row).map((v) => getNumeric(v, NaN)).filter((v) => Number.isFinite(v));
  if (!nums.length) return 0;
  return nums.reduce((sum, n) => sum + n, 0) / nums.length;
}

function detectFinancialHeads(rows) {
  const heads = [];
  const seen = new Set();
  rows.forEach((row) => {
    const head = guessHeadText(row);
    if (!head) return;
    const norm = normalizeHead(head);
    if (!norm || seen.has(norm)) return;
    seen.add(norm);
    heads.push({ head, norm, amount: extractAmountFromRow(row) });
  });
  return heads;
}

function findByAliases(heads, aliases) {
  const normalizedAliases = aliases.map(normalizeHead);
  return heads.find((h) => normalizedAliases.some((a) => h.norm.includes(a) || a.includes(h.norm)));
}

function buildMappingOptions() {
  const heads = workbookState.detectedHeads;
  const autoReceivable = findByAliases(heads, ["trade receivables", "sundry debtors", "debtors"]);
  const autoInventory = findByAliases(heads, ["inventory", "closing stock", "stock"]);
  const autoPayables = findByAliases(heads, ["trade payables", "creditors", "sundry creditors"]);

  const options = ['<option value="">(Not mapped)</option>']
    .concat(heads.map((h) => `<option value="${h.head}">${h.head}</option>`))
    .join("");

  mapReceivablesSelect.innerHTML = options;
  mapInventorySelect.innerHTML = options;
  mapPayablesSelect.innerHTML = options;

  mapReceivablesSelect.value = autoReceivable?.head || "";
  mapInventorySelect.value = autoInventory?.head || "";
  mapPayablesSelect.value = autoPayables?.head || "";

  workbookState.mapping.tradeReceivables = mapReceivablesSelect.value;
  workbookState.mapping.inventory = mapInventorySelect.value;
  workbookState.mapping.tradePayables = mapPayablesSelect.value;
}

function lookupMappedAmount(heads, mappedHead, fallback) {
  if (!mappedHead) return fallback;
  const match = heads.find((h) => h.head === mappedHead);
  return match ? match.amount : fallback;
}

function handleWorkbookUpload(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      workbookState.workbook = XLSX.read(e.target.result, { type: "array" });
      sheetSelector.innerHTML = "";
      workbookState.workbook.SheetNames.forEach((name) => {
        const option = document.createElement("option");
        option.value = name;
        option.textContent = name;
        sheetSelector.appendChild(option);
      });
      sheetSelector.disabled = false;
      sheetStatus.textContent = `Loaded ${file.name} with ${workbookState.workbook.SheetNames.length} sheets.`;
      handleSheetSelection();
    } catch (error) {
      sheetStatus.textContent = `Failed to parse workbook: ${error.message}`;
    }
  };
  reader.readAsArrayBuffer(file);
}

function handleSheetSelection() {
  const selectedSheet = sheetSelector.value;
  if (!selectedSheet || !workbookState.workbook) return;
  workbookState.selectedSheet = selectedSheet;

  const ws = workbookState.workbook.Sheets[selectedSheet];
  workbookState.sourceRows = XLSX.utils.sheet_to_json(ws, { defval: null });
  workbookState.detectedHeads = detectFinancialHeads(workbookState.sourceRows);
  buildMappingOptions();

  generateBtn.disabled = false;
  output.innerHTML = `<p class="good">Sheet <strong>${selectedSheet}</strong> ready with ${workbookState.sourceRows.length} records and ${workbookState.detectedHeads.length} detected heads.</p>`;
}

function guessYearlySales(rows, fallback = 10000000) {
  if (!rows.length) return fallback;
  const candidateKeys = ["sales", "turnover", "revenue", "net_sales", "gross_sales"];
  const sample = rows[0];
  const key = Object.keys(sample).find((k) => candidateKeys.some((n) => k.toLowerCase().includes(n)));
  if (!key) return fallback;

  const values = rows.map((row) => getNumeric(row[key])).filter((n) => n > 0);
  if (!values.length) return fallback;
  return values.reduce((sum, n) => sum + n, 0) / values.length;
}

function buildPeriods() {
  const historicalCount = Math.max(1, Math.min(4, getConfigNumber("historical-count", 3)));
  const ccYears = Math.max(1, getConfigNumber("projection-years-cc", 5));
  const tlYears = Math.max(1, getConfigNumber("projection-years-tl", 6));
  const currentYear = new Date().getFullYear();

  const historical = Array.from({ length: historicalCount }, (_, idx) => formatFY(currentYear - historicalCount + idx));
  const provisionalYear = formatFY(currentYear, " (Prov.)");
  const projections = Array.from({ length: Math.max(ccYears, tlYears) }, (_, idx) => {
    const start = currentYear + 1 + idx;
    return {
      label: formatFY(start),
      isCCYear: idx < ccYears,
      isTLYear: idx < tlYears,
      yearNo: idx + 1,
    };
  });

  return {
    historical,
    provisionalYear,
    projections,
    ccYears,
    tlYears,
    allColumns: [...historical, provisionalYear, ...projections.map((p) => p.label)],
  };
}

function buildCmaReport(baseSales, mappedFinancials) {
  const periods = buildPeriods();
  const salesGrowth = getConfigNumber("sales-growth", 10);
  const projectionMode = getConfigText("sales-projection-mode", "growth");
  const grossMargin = getConfigNumber("gross-margin", 28);
  const otherIncomeRatio = getConfigNumber("other-income-ratio", 1.5);
  const employeeRatio = getConfigNumber("employee-ratio", 6);
  const adminRatio = getConfigNumber("admin-ratio", 4);
  const sellingRatio = getConfigNumber("selling-ratio", 3);
  const deprRatio = getConfigNumber("depr-ratio", 10);
  const taxRate = getConfigNumber("tax-rate", 25);
  const receivableDays = getConfigNumber("receivable-days", 60);
  const inventoryDays = getConfigNumber("inventory-days", 45);
  const payableDays = getConfigNumber("payable-days", 30);
  const marginRatio = getConfigNumber("margin-ratio", 25);
  const tlOutstanding = getConfigNumber("tl-outstanding", 12000000);
  const tlRate = getConfigNumber("tl-rate", 11) / 100;
  const ccLimit = getConfigNumber("cc-limit", 5000000);
  const ccUtilization = getConfigNumber("cc-utilization", 85) / 100;
  const moratoriumYears = Math.max(0, getConfigNumber("moratorium-years", 1));

  const nayakSalesAnchor = ccLimit / 0.2;
  const annualPrincipalInstallment = safeDivide(tlOutstanding, Math.max(1, periods.tlYears - moratoriumYears));

  const records = [];
  const allPeriodLabels = periods.allColumns;

  allPeriodLabels.forEach((periodLabel, idx) => {
    const isHistorical = idx < periods.historical.length;
    const isProvisional = idx === periods.historical.length;
    const projectionIdx = idx - periods.historical.length - 1;
    const isProjection = projectionIdx >= 0;
    const salesBase = idx === 0 ? baseSales : records[idx - 1].pl.Sales;

    const growthSales = salesBase * (1 + salesGrowth / 100);
    const nayakProjectedSales = nayakSalesAnchor * Math.pow(1 + salesGrowth / 100, Math.max(0, projectionIdx));
    const sales = isProjection
      ? projectionMode === "nayak"
        ? nayakProjectedSales
        : growthSales
      : isProvisional
        ? salesBase * 1.05
        : salesBase * 0.92;

    const otherIncome = ratio(sales, otherIncomeRatio);

    const openingStock = idx === 0 ? mappedFinancials.inventory : records[idx - 1].pl["Closing Stock"];
    const closingStock = isHistorical
      ? mappedFinancials.inventory * Math.pow(1 + salesGrowth / 100, idx * 0.5)
      : (sales * inventoryDays) / 365;
    const purchases = sales * (1 - grossMargin / 100) * 0.84;
    const directExpenses = sales * (1 - grossMargin / 100) * 0.14;
    const cogs = openingStock + purchases + directExpenses - closingStock;
    const grossProfit = sales - cogs;

    const employeeCost = ratio(sales, employeeRatio);
    const adminExpenses = ratio(sales, adminRatio);
    const sellingExpenses = ratio(sales, sellingRatio);
    const operatingProfit = grossProfit - employeeCost - adminExpenses - sellingExpenses;

    const tlYearNo = projectionIdx + 1;
    const isTLYear = isProjection && tlYearNo <= periods.tlYears;
    const openingTL = isTLYear
      ? Math.max(tlOutstanding - Math.max(0, tlYearNo - moratoriumYears - 1) * annualPrincipalInstallment, 0)
      : 0;
    const installment = isTLYear && tlYearNo > moratoriumYears ? Math.min(annualPrincipalInstallment, openingTL) : 0;
    const interest = openingTL * tlRate;
    const closingTL = Math.max(openingTL - installment, 0);

    const fixedAssets = sales * 0.22;
    const depreciation = ratio(fixedAssets, deprRatio);
    const pbt = operatingProfit + otherIncome - interest - depreciation;
    const tax = Math.max(pbt, 0) * (taxRate / 100);
    const pat = pbt - tax;

    const netWorth = sales * 0.18 + pat * 2;
    const reserves = Math.max(netWorth * 0.35, 0);
    const capital = netWorth - reserves;
    const unsecuredLoans = sales * 0.04;

    const receivables = isHistorical
      ? mappedFinancials.tradeReceivables * Math.pow(1 + salesGrowth / 100, idx * 0.5)
      : (sales * receivableDays) / 365;
    const inventory = closingStock;
    const creditors = isHistorical
      ? mappedFinancials.tradePayables * Math.pow(1 + salesGrowth / 100, idx * 0.5)
      : (purchases * payableDays) / 365;
    const otherCurrentAssets = sales * 0.03;
    const cashBank = Math.max(sales * 0.01 + pat * 0.05, 0);
    const loansAdvances = sales * 0.02;
    const investments = sales * 0.015;
    const otherCurrentLiabilities = sales * 0.02;

    const totalCurrentAssets = receivables + inventory + otherCurrentAssets + cashBank + loansAdvances;
    const totalCurrentLiabilities = creditors + otherCurrentLiabilities;
    const wcGap = totalCurrentAssets - totalCurrentLiabilities;
    const marginContribution = wcGap * (marginRatio / 100);

    const nayakTotalWorkingCapital = sales * 0.25;
    const nayakBorrowerContribution = sales * 0.05;
    const nayakMPBF = nayakTotalWorkingCapital - nayakBorrowerContribution;
    const tandonMethod1 = wcGap * 0.75;
    const tandonMethod2 = totalCurrentAssets * 0.75 - totalCurrentLiabilities;
    const ccRequired = Math.max(0, Math.max(nayakMPBF, tandonMethod1, tandonMethod2, wcGap - marginContribution, ccLimit * ccUtilization));

    const totalLiabilities = capital + reserves + openingTL + unsecuredLoans + creditors + otherCurrentLiabilities;
    const totalAssets = fixedAssets + investments + inventory + receivables + cashBank + loansAdvances + otherCurrentAssets;

    const currentRatio = safeDivide(totalCurrentAssets, totalCurrentLiabilities);
    const debtEquity = safeDivide(openingTL + unsecuredLoans, netWorth);
    const tolTnw = safeDivide(totalLiabilities, netWorth);
    const gpRatio = safeDivide(grossProfit, sales) * 100;
    const npRatio = safeDivide(pat, sales) * 100;
    const opMargin = safeDivide(operatingProfit, sales) * 100;
    const interestCoverage = safeDivide(operatingProfit, interest);

    const grossCashAccrual = pat + depreciation + interest;
    const totalDebtObligation = installment + interest;
    const dscr = safeDivide(grossCashAccrual, totalDebtObligation);

    const ratioChecks = {
      "Current Ratio >= 1.33": currentRatio >= 1.33,
      "Debt Equity <= 2.00": debtEquity <= 2,
      "TOL/TNW <= 4.00": tolTnw <= 4,
      "Interest Coverage >= 1.50": interestCoverage >= 1.5,
      "DSCR >= 1.25": dscr >= 1.25,
    };

    records.push({
      period: periodLabel,
      meta: {
        isHistorical,
        isProvisional,
        projectionNo: tlYearNo,
        isCCYear: isProjection && tlYearNo <= periods.ccYears,
        isTLYear,
      },
      pl: {
        Sales: sales,
        "Opening Stock": openingStock,
        Purchases: purchases,
        "Direct Expenses": directExpenses,
        "Closing Stock": closingStock,
        COGS: cogs,
        "Gross Profit": grossProfit,
        "Employee Cost": employeeCost,
        "Administrative Expenses": adminExpenses,
        "Selling Expenses": sellingExpenses,
        "Operating Profit": operatingProfit,
        Interest: interest,
        Depreciation: depreciation,
        "Profit Before Tax": pbt,
        Tax: tax,
        "Profit After Tax": pat,
      },
      bs: {
        "CAPITAL & LIABILITIES": null,
        Capital: capital,
        Reserves: reserves,
        "Term Loan": openingTL,
        "Unsecured Loans": unsecuredLoans,
        "Trade Payables": creditors,
        "Other Current Liabilities": otherCurrentLiabilities,
        "Total Capital & Liabilities": totalLiabilities,
        ASSETS: null,
        "Fixed Assets": fixedAssets,
        Investments: investments,
        Inventory: inventory,
        "Trade Receivables": receivables,
        "Cash & Bank": cashBank,
        "Loans & Advances": loansAdvances,
        "Other Current Assets": otherCurrentAssets,
        "Total Assets": totalAssets,
      },
      wc: {
        Sales: sales,
        "Receivable Days": receivableDays,
        "Inventory Days": inventoryDays,
        "Creditor Days": payableDays,
        Receivables: receivables,
        Inventory: inventory,
        Creditors: creditors,
        "Working Capital Requirement": wcGap,
        "Nayak MPBF": nayakMPBF,
        "CC / MPBF Considered": ccRequired,
      },
      ratios: {
        "Current Ratio": currentRatio,
        "Debt Equity Ratio": debtEquity,
        "TOL/TNW": tolTnw,
        "Gross Profit Ratio (%)": gpRatio,
        "Net Profit Ratio (%)": npRatio,
        "Operating Margin (%)": opMargin,
        "Interest Coverage Ratio": interestCoverage,
        DSCR: dscr,
      },
      ratioChecks,
      termLoan: {
        "Opening Balance": openingTL,
        Installment: installment,
        Interest: interest,
        "Total Debt Servicing": totalDebtObligation,
        "Closing Balance": closingTL,
      },
      dscr: {
        PAT: pat,
        Depreciation: depreciation,
        "Interest on Term Loan": interest,
        "Gross Cash Accrual": grossCashAccrual,
        Installment: installment,
        "Total Debt Obligation": totalDebtObligation,
        DSCR: dscr,
        "Average DSCR": 0,
      },
    });
  });

  const tlYears = records.filter((year) => year.meta.isTLYear);
  const avgDSCR = safeDivide(
    tlYears.reduce((acc, y) => acc + y.dscr.DSCR, 0),
    Math.max(1, tlYears.length),
  );
  tlYears.forEach((year) => {
    year.dscr["Average DSCR"] = avgDSCR;
  });

  const borrowerSummary = {
    "Borrower Name": getConfigText("borrower-name"),
    Constitution: getConfigText("constitution"),
    PAN: getConfigText("pan"),
    GSTIN: getConfigText("gstin"),
    "Bank / Branch": `${getConfigText("bank-name")} / ${getConfigText("branch-name")}`,
    "Existing Limit": getConfigNumber("existing-limit", 0),
    "Proposed Limit": getConfigNumber("proposed-limit", 0),
    "Facility Type": getConfigText("facility-type"),
    "CMA Prepared Date": getConfigText("prepared-date"),
    "Projection Basis": projectionMode === "nayak" ? "Nayak Committee (CC/20%)" : "Growth %",
  };

  return {
    periods,
    borrowerSummary,
    allYears: records,
    mappedFinancials,
    averages: {
      dscr: avgDSCR,
      acceptableRepayment: avgDSCR >= 1.25,
    },
    assumptions: {
      "Historical years": periods.historical.length,
      "Current year provisional": periods.provisionalYear,
      "Projected years for Cash Credit": periods.ccYears,
      "Projected years for Term Loan": periods.tlYears,
      "Moratorium (years)": moratoriumYears,
      "Sales growth %": salesGrowth,
      "CC based sales (CC/20%)": nayakSalesAnchor,
      "Receivable days": receivableDays,
      "Inventory days": inventoryDays,
      "Creditor days": payableDays,
    },
  };
}

function getRowsByMap(yearRecords, mapKey) {
  const keys = Object.keys(yearRecords[0]?.[mapKey] || {});
  return keys.map((key) => ({
    label: key,
    values: yearRecords.map((year) => year[mapKey][key]),
  }));
}

function buildSectionHtml(title, headerPeriods, rows, formatter = fmtCurrency) {
  const head = `<tr><th>Particulars</th>${headerPeriods.map((p) => `<th>${p}</th>`).join("")}</tr>`;
  const body = rows
    .map(({ label, values }) => {
      const isHeader = values.every((v) => v === null || v === undefined);
      const cells = values.map((v) => `<td>${formatter(v, label)}</td>`).join("");
      return `<tr class="${isHeader ? "section-break" : ""}"><td>${label}</td>${cells}</tr>`;
    })
    .join("");

  return `<section class="report-section"><h3>${title}</h3><table><thead>${head}</thead><tbody>${body}</tbody></table></section>`;
}

function renderRatioValidation(report) {
  const periods = report.allYears.map((y) => y.period);
  const checks = Object.keys(report.allYears[0]?.ratioChecks || {});
  const header = `<tr><th>Norm</th>${periods.map((p) => `<th>${p}</th>`).join("")}</tr>`;
  const rows = checks
    .map((check) => {
      const cells = report.allYears
        .map((y) => `<td class="${y.ratioChecks[check] ? "good" : "bad"}">${y.ratioChecks[check] ? "Within" : "Breach"}</td>`)
        .join("");
      return `<tr><td>${check}</td>${cells}</tr>`;
    })
    .join("");
  return `<section class="report-section"><h3>Ratio Norm Validation</h3><table><thead>${header}</thead><tbody>${rows}</tbody></table></section>`;
}

function renderReport(report) {
  const periods = report.periods.allColumns;
  const ccSeries = report.allYears.filter((y) => y.meta.isHistorical || y.meta.isProvisional || y.meta.isCCYear);
  const tlSeries = report.allYears.filter((y) => y.meta.isTLYear);

  const summaryRows = Object.entries(report.borrowerSummary)
    .map(([k, v]) => `<tr><td>${k}</td><td>${typeof v === "number" ? fmtCurrency(v) : v}</td></tr>`)
    .join("");
  const mappingRows = [
    ["Trade Receivables", report.mappedFinancials.tradeReceivables],
    ["Inventory", report.mappedFinancials.inventory],
    ["Trade Payables", report.mappedFinancials.tradePayables],
  ]
    .map(([k, v]) => `<tr><td>${k}</td><td>${fmtCurrency(v)}</td></tr>`)
    .join("");
  const assumptionsRows = Object.entries(report.assumptions)
    .map(([k, v]) => `<tr><td>${k}</td><td>${typeof v === "number" ? fmtNumber(v) : v}</td></tr>`)
    .join("");

  output.innerHTML = `
    <article class="cma-report">
      <header class="report-header">
        <h2>Banker-Grade CMA Report</h2>
        <p class="good">Average DSCR: ${report.averages.dscr.toFixed(2)} (${report.averages.acceptableRepayment ? "Acceptable banking norm (>= 1.25)" : "Below acceptable banking norm (< 1.25)"})</p>
      </header>
      <section class="report-section"><h3>Summary</h3><table><tbody>${summaryRows}</tbody></table></section>
      <section class="report-section"><h3>Mapped Financial Heads (Uploaded Data)</h3><table><tbody>${mappingRows}</tbody></table></section>
      <section class="report-section"><h3>Assumptions</h3><table><tbody>${assumptionsRows}</tbody></table></section>
      ${buildSectionHtml("Profit & Loss (Vertical Format)", periods, getRowsByMap(report.allYears, "pl"))}
      ${buildSectionHtml("Balance Sheet", periods, getRowsByMap(report.allYears, "bs"))}
      ${buildSectionHtml("Working Capital Assessment", ccSeries.map((y) => y.period), getRowsByMap(ccSeries, "wc"))}
      ${buildSectionHtml("Ratio Analysis", periods, getRowsByMap(report.allYears, "ratios"), (value, label) => label.includes("(%)") ? fmtPct(value) : fmtNumber(value))}
      ${renderRatioValidation(report)}
      ${buildSectionHtml("Term Loan Schedule", tlSeries.map((y) => y.period), getRowsByMap(tlSeries, "termLoan"))}
      ${buildSectionHtml("DSCR", tlSeries.map((y) => y.period), getRowsByMap(tlSeries, "dscr"), (value, label) => label.includes("DSCR") ? fmtNumber(value) : fmtCurrency(value))}
    </article>
  `;
}

function saveFile(content, filename, mime = "text/plain") {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function generateReport() {
  const heads = workbookState.detectedHeads;
  const mappedFinancials = {
    tradeReceivables: lookupMappedAmount(heads, workbookState.mapping.tradeReceivables, 0),
    inventory: lookupMappedAmount(heads, workbookState.mapping.inventory, 0),
    tradePayables: lookupMappedAmount(heads, workbookState.mapping.tradePayables, 0),
  };
  const baseSales = guessYearlySales(workbookState.sourceRows);
  const data = buildCmaReport(baseSales, mappedFinancials);

  workbookState.generated = {
    meta: {
      sheet: workbookState.selectedSheet,
      sourceRecords: workbookState.sourceRows.length,
      generatedAt: new Date().toISOString(),
    },
    ...data,
  };

  renderReport(workbookState.generated);
  downloadReportBtn.disabled = false;
  downloadExcelBtn.disabled = false;
  downloadJsonBtn.disabled = false;
}

function downloadReport() {
  if (!workbookState.generated) return;
  const title = `CMA Report - ${workbookState.generated.meta.sheet}`;
  const styleContent = Array.from(document.styleSheets)
    .map((sheet) => {
      try {
        return Array.from(sheet.cssRules || []).map((rule) => rule.cssText).join("\n");
      } catch (_err) {
        return "";
      }
    })
    .join("\n");
  const html = `<!doctype html><html><head><meta charset="UTF-8"><title>${title}</title><style>${styleContent}</style></head><body>${output.innerHTML}</body></html>`;
  saveFile(html, "cma-report.html", "text/html");
}

function formatSheetWithHeader(ws, title, columnCount) {
  ws["!cols"] = Array.from({ length: columnCount }, (_, i) => ({ wch: i === 0 ? 44 : 18 }));
  ws["!merges"] = ws["!merges"] || [];
  ws["!merges"].push({ s: { r: 0, c: 0 }, e: { r: 0, c: Math.max(columnCount - 1, 1) } });
  XLSX.utils.sheet_add_aoa(ws, [[title]], { origin: "A1" });
}

function applyNumberFormat(ws, startRow, endRow, periodCount, format) {
  for (let r = startRow; r <= endRow; r += 1) {
    for (let c = 2; c <= periodCount + 1; c += 1) {
      const ref = XLSX.utils.encode_cell({ r: r - 1, c: c - 1 });
      if (ws[ref] && (typeof ws[ref].v === "number" || ws[ref].f)) ws[ref].z = format;
    }
  }
}

function makeSectionSheet(title, periods, rows) {
  const aoa = [[title], ["Particulars", ...periods]];
  rows.forEach((row) => {
    aoa.push([row.label, ...row.values]);
  });
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  formatSheetWithHeader(ws, title, periods.length + 1);
  return ws;
}

function toCell(rowIndex1Based, colIndex1Based) {
  return XLSX.utils.encode_cell({ r: rowIndex1Based - 1, c: colIndex1Based - 1 });
}

function setFormula(ws, r, c, formula, format) {
  const ref = toCell(r, c);
  ws[ref] = { t: "n", f: formula };
  if (format) ws[ref].z = format;
}

function downloadExcel() {
  if (!workbookState.generated) return;

  const wb = XLSX.utils.book_new();
  const report = workbookState.generated;
  const periods = report.periods.allColumns;
  const ccSeries = report.allYears.filter((y) => y.meta.isHistorical || y.meta.isProvisional || y.meta.isCCYear);
  const tlSeries = report.allYears.filter((y) => y.meta.isTLYear);

  const summaryData = [["Summary"], ["Particular", "Details"]];
  Object.entries(report.borrowerSummary).forEach(([k, v]) => summaryData.push([k, v]));
  const summaryWs = XLSX.utils.aoa_to_sheet(summaryData);
  formatSheetWithHeader(summaryWs, "Summary", 2);

  const plRows = getRowsByMap(report.allYears, "pl");
  const plWs = makeSectionSheet("Profit & Loss", periods, plRows);
  periods.forEach((_, i) => {
    const col = i + 2;
    setFormula(plWs, 8, col, `${toCell(4, col)}+${toCell(5, col)}+${toCell(6, col)}-${toCell(7, col)}`, "₹#,##0");
    setFormula(plWs, 9, col, `${toCell(3, col)}-${toCell(8, col)}`, "₹#,##0");
    setFormula(plWs, 13, col, `${toCell(9, col)}-${toCell(10, col)}-${toCell(11, col)}-${toCell(12, col)}`, "₹#,##0");
    setFormula(plWs, 16, col, `${toCell(13, col)}-${toCell(14, col)}-${toCell(15, col)}`, "₹#,##0");
    setFormula(plWs, 18, col, `${toCell(16, col)}-${toCell(17, col)}`, "₹#,##0");
  });
  applyNumberFormat(plWs, 3, plRows.length + 2, periods.length, "₹#,##0");

  const bsRows = getRowsByMap(report.allYears, "bs");
  const bsWs = makeSectionSheet("Balance Sheet", periods, bsRows);
  periods.forEach((_, i) => {
    const col = i + 2;
    setFormula(bsWs, 10, col, `${toCell(4, col)}+${toCell(5, col)}+${toCell(6, col)}+${toCell(7, col)}+${toCell(8, col)}+${toCell(9, col)}`, "₹#,##0");
    setFormula(bsWs, 19, col, `${toCell(12, col)}+${toCell(13, col)}+${toCell(14, col)}+${toCell(15, col)}+${toCell(16, col)}+${toCell(17, col)}+${toCell(18, col)}`, "₹#,##0");
  });

  const wcPeriods = ccSeries.map((y) => y.period);
  const wcRows = getRowsByMap(ccSeries, "wc");
  const wcWs = makeSectionSheet("Working Capital", wcPeriods, wcRows);
  wcPeriods.forEach((_, i) => {
    const col = i + 2;
    setFormula(wcWs, 10, col, `MAX(${toCell(9, col)},${toCell(3, col)}*0.20)`, "₹#,##0");
  });

  const ratioWs = makeSectionSheet("Ratio Analysis", periods, getRowsByMap(report.allYears, "ratios"));
  periods.forEach((_, i) => {
    const col = i + 2;
    setFormula(ratioWs, 3, col, `'Working Capital'!${toCell(8, col)}/('Balance Sheet'!${toCell(7, col)}+'Balance Sheet'!${toCell(8, col)})`, "0.00");
    setFormula(ratioWs, 10, col, `'DSCR'!${toCell(9, col)}`, "0.00");
  });

  const tlPeriods = tlSeries.map((y) => y.period);
  const tlRows = getRowsByMap(tlSeries, "termLoan");
  const tlWs = makeSectionSheet("Term Loan", tlPeriods, tlRows);
  tlPeriods.forEach((_, i) => {
    const col = i + 2;
    setFormula(tlWs, 6, col, `${toCell(4, col)}+${toCell(5, col)}`, "₹#,##0");
  });

  const dscrRows = getRowsByMap(tlSeries, "dscr");
  const dscrWs = makeSectionSheet("DSCR", tlPeriods, dscrRows);
  tlPeriods.forEach((_, i) => {
    const col = i + 2;
    setFormula(dscrWs, 6, col, `${toCell(3, col)}+${toCell(4, col)}+${toCell(5, col)}`, "₹#,##0");
    setFormula(dscrWs, 8, col, `${toCell(7, col)}+${toCell(5, col)}`, "₹#,##0");
    setFormula(dscrWs, 9, col, `${toCell(6, col)}/${toCell(8, col)}`, "0.00");
    setFormula(dscrWs, 10, col, `AVERAGE(${toCell(9, 2)}:${toCell(9, tlPeriods.length + 1)})`, "0.00");
  });

  XLSX.utils.book_append_sheet(wb, summaryWs, "Summary");
  XLSX.utils.book_append_sheet(wb, plWs, "Profit & Loss");
  XLSX.utils.book_append_sheet(wb, bsWs, "Balance Sheet");
  XLSX.utils.book_append_sheet(wb, wcWs, "Working Capital");
  XLSX.utils.book_append_sheet(wb, ratioWs, "Ratio Analysis");
  XLSX.utils.book_append_sheet(wb, tlWs, "Term Loan");
  XLSX.utils.book_append_sheet(wb, dscrWs, "DSCR");

  XLSX.writeFile(wb, "cma-report.xlsx");
}

function downloadJson() {
  if (!workbookState.generated) return;
  saveFile(JSON.stringify(workbookState.generated, null, 2), "cma-data.json", "application/json");
}
