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
  "prepared-date", "historical-count", "projection-years-cc", "projection-years-tl", "moratorium-years", "sales-growth", "sales-projection-mode",
  "other-income-ratio", "gross-margin", "employee-ratio", "admin-ratio", "selling-ratio", "depr-ratio", "tax-rate", "receivable-days",
  "inventory-days", "payable-days", "margin-ratio", "cc-limit", "cc-utilization", "tl-outstanding", "tl-rate",
];
const configInputs = configIds.reduce((acc, id) => ({ ...acc, [id]: document.getElementById(id) }), {});

if (!configInputs["prepared-date"].value) configInputs["prepared-date"].value = new Date().toISOString().slice(0, 10);

const MAP_FIELDS = [
  { key: "sales", label: "Sales", aliases: ["salesaccount", "sale", "sales", "turnover"] },
  { key: "purchases", label: "Purchases", aliases: ["purchaseaccount", "purchase", "purchases"] },
  { key: "closingStock", label: "Closing Stock / Inventory", aliases: ["closingstock", "stockinhand", "inventory", "stock"] },
  { key: "tradeReceivables", label: "Trade Receivables", aliases: ["sundrydebtors", "tradereceivables", "debtors"] },
  { key: "tradeCreditors", label: "Trade Creditors", aliases: ["sundrycreditors", "tradecreditors", "creditors"] },
  { key: "workingCapitalBorrowings", label: "Working Capital Borrowings / CC", aliases: ["bankod", "cashcredit", "workingcapital", "cc"] },
  { key: "capital", label: "Capital", aliases: ["capitalaccount", "partnerscapital", "capital"] },
  { key: "cashBank", label: "Cash & Bank", aliases: ["cashatbank", "cashinhand", "bankbalance", "cashbank"] },
];

fileInput.addEventListener("change", handleWorkbookUpload);
generateBtn.addEventListener("click", generateReport);
downloadReportBtn.addEventListener("click", downloadReport);
downloadExcelBtn.addEventListener("click", downloadExcel);
downloadJsonBtn.addEventListener("click", downloadJson);

function getNumeric(value, fallback = 0) {
  if (value === null || value === undefined || value === "") return fallback;
  if (typeof value === "number") return Number.isFinite(value) ? value : fallback;
  const cleaned = String(value).replace(/[,\s]/g, "").replace(/\(([^)]+)\)/, "-$1");
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : fallback;
}

function getConfigNumber(id, fallback = 0) { return getNumeric(configInputs[id]?.value, fallback); }
function getConfigText(id, fallback = "") { return (configInputs[id]?.value || fallback).toString().trim(); }
function ratio(value, percent) { return value * (percent / 100); }
function safeDivide(a, b) { return b ? a / b : 0; }
function formatFY(startYear, suffix = "") { return `FY ${startYear}-${String(startYear + 1).slice(-2)}${suffix}`; }

function fmtCurrency(amount) {
  if (amount === null || amount === undefined) return "";
  return new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", maximumFractionDigits: 0 }).format(amount || 0);
}
function fmtNumber(value) { return value === null || value === undefined ? "" : (value || 0).toFixed(2); }
function fmtPct(value) { return value === null || value === undefined ? "" : `${(value || 0).toFixed(2)}%`; }

function normalize(value) { return (value || "").toString().toLowerCase().replace(/[^a-z0-9]/g, ""); }

function detectSheetByContent(workbook, type) {
  const candidates = workbook.SheetNames.map((name) => {
    const norm = normalize(name);
    const ws = workbook.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    const flat = rows.flat().map((c) => normalize(c)).join("|");
    const nameScore = type === "pl"
      ? (norm.includes("profit") || norm.includes("pl") || norm.includes("trading") ? 3 : 0)
      : (norm.includes("balance") || norm === "bs" ? 3 : 0);
    const textScore = type === "pl"
      ? ((flat.includes("sales") ? 2 : 0) + (flat.includes("purchase") ? 2 : 0) + (flat.includes("grossprofit") ? 2 : 0))
      : ((flat.includes("sundrydebtors") ? 2 : 0) + (flat.includes("sundrycreditors") ? 2 : 0) + (flat.includes("capital") ? 2 : 0));
    return { name, score: nameScore + textScore };
  }).sort((a, b) => b.score - a.score);
  return candidates[0]?.score > 0 ? candidates[0].name : null;
}

function extractHeadsFromTwoSidedSheet(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const heads = [];
  rows.forEach((row, rowIdx) => {
    const cells = row.map((value, colIdx) => ({ value, colIdx }));
    const textCells = cells.filter((c) => typeof c.value === "string" && c.value.trim().length > 2);
    const numberCells = cells.filter((c) => Number.isFinite(getNumeric(c.value, NaN)));

    textCells.forEach((textCell) => {
      const text = textCell.value.trim();
      const norm = normalize(text);
      if (!norm || norm.includes("total") || norm.includes("opening") && !norm.includes("stock")) return;
      const rightNumber = numberCells
        .filter((n) => n.colIdx > textCell.colIdx)
        .sort((a, b) => a.colIdx - b.colIdx)[0];
      if (!rightNumber) return;
      const amount = getNumeric(rightNumber.value, NaN);
      if (!Number.isFinite(amount) || amount === 0) return;
      heads.push({ head: text, norm, amount: Math.abs(amount), row: rowIdx + 1, col: rightNumber.colIdx + 1 });
    });
  });
  const dedup = new Map();
  heads.forEach((h) => { if (!dedup.has(h.norm)) dedup.set(h.norm, h); });
  return Array.from(dedup.values());
}

function autoMapHeads(heads) {
  const mapping = {};
  const confidence = {};
  MAP_FIELDS.forEach((field) => {
    let match = null;
    field.aliases.forEach((alias) => {
      if (match) return;
      match = heads.find((h) => h.norm === alias || h.norm.includes(alias));
      if (match) confidence[field.key] = hScore(match.norm, alias);
    });
    mapping[field.key] = match?.head || "";
  });
  return { mapping, confidence };
}

function hScore(normHead, alias) { return normHead === alias ? 1 : 0.6; }

function buildMappingReview(heads, mapping, confidence) {
  mappingGrid.innerHTML = "";
  const options = ['<option value="">(Not mapped)</option>']
    .concat(heads.map((h) => `<option value="${h.head}">${h.head} (₹${Math.round(h.amount).toLocaleString("en-IN")})</option>`))
    .join("");

  let needsReview = false;
  MAP_FIELDS.forEach((field) => {
    if (!mapping[field.key] || (confidence[field.key] || 0) < 1) needsReview = true;
    const wrap = document.createElement("label");
    wrap.innerHTML = `${field.label}<select data-map-key="${field.key}">${options}</select>`;
    const select = wrap.querySelector("select");
    select.value = mapping[field.key] || "";
    select.addEventListener("change", () => { workbookState.mapping[field.key] = select.value; });
    mappingGrid.appendChild(wrap);
  });

  mappingReview.classList.toggle("hidden", !needsReview);
  workbookState.mapping = { ...mapping };
  return needsReview;
}

function parseWorkbook(workbook) {
  const plSheetName = detectSheetByContent(workbook, "pl");
  const bsSheetName = detectSheetByContent(workbook, "bs");
  if (!plSheetName || !bsSheetName) throw new Error("Could not detect Profit & Loss and Balance Sheet sheets.");

  const plHeads = extractHeadsFromTwoSidedSheet(workbook.Sheets[plSheetName]);
  const bsHeads = extractHeadsFromTwoSidedSheet(workbook.Sheets[bsSheetName]);
  if (!plHeads.length || !bsHeads.length) throw new Error("Could not parse two-sided ledger format from uploaded workbook.");

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
      sheetStatus.innerHTML = `<span class="good">Detected P&L sheet: <strong>${workbookState.parsed.plSheetName}</strong>, Balance Sheet: <strong>${workbookState.parsed.bsSheetName}</strong>.</span>`;
      output.innerHTML = `<p class="good">Workbook parsed successfully. ${workbookState.parsed.needsReview ? "Complete mapping review and generate report." : "Mappings look confident; ready to generate."}</p>`;
      generateBtn.disabled = false;
    } catch (error) {
      workbookState.parsed = null;
      workbookState.parseError = error.message;
      generateBtn.disabled = true;
      mappingReview.classList.add("hidden");
      sheetStatus.innerHTML = `<span class="bad">Upload parsing failed: ${error.message}</span>`;
      output.innerHTML = `<p class="bad">Unable to generate CMA. Please upload a valid Tally final accounts workbook.</p>`;
    }
  };
  reader.readAsArrayBuffer(file);
}

function getMappedAmount(heads, selectedHead) {
  if (!selectedHead) return 0;
  return heads.find((h) => h.head === selectedHead)?.amount || 0;
}

function mappedFinancialsFromParse() {
  if (!workbookState.parsed) throw new Error("No parsed workbook.");
  const heads = workbookState.parsed.allHeads;
  const mapped = {};
  MAP_FIELDS.forEach((field) => { mapped[field.key] = getMappedAmount(heads, workbookState.mapping[field.key]); });
  mapped.inventory = mapped.closingStock;
  mapped.termLoan = getConfigNumber("tl-outstanding", 0);
  return mapped;
}

function buildPeriods() {
  const historicalCount = Math.max(1, Math.min(4, getConfigNumber("historical-count", 3)));
  const ccYears = Math.max(1, getConfigNumber("projection-years-cc", 5));
  const tlYears = Math.max(1, getConfigNumber("projection-years-tl", 6));
  const currentYear = new Date().getFullYear();
  const historical = Array.from({ length: historicalCount }, (_, idx) => formatFY(currentYear - historicalCount + idx));
  const provisionalYear = formatFY(currentYear, " (Prov.)");
  const projections = Array.from({ length: Math.max(ccYears, tlYears) }, (_, idx) => ({ label: formatFY(currentYear + 1 + idx), isCCYear: idx < ccYears, isTLYear: idx < tlYears }));
  return { historical, provisionalYear, projections, ccYears, tlYears, allColumns: [...historical, provisionalYear, ...projections.map((p) => p.label)] };
}

function buildCmaReport(mapped) {
  const periods = buildPeriods();
  const salesGrowth = getConfigNumber("sales-growth", 10) / 100;
  const otherIncomeRatio = getConfigNumber("other-income-ratio", 1.5);
  const employeeRatio = getConfigNumber("employee-ratio", 6);
  const adminRatio = getConfigNumber("admin-ratio", 4);
  const sellingRatio = getConfigNumber("selling-ratio", 3);
  const deprRatio = getConfigNumber("depr-ratio", 10);
  const taxRate = getConfigNumber("tax-rate", 25);
  const receivableDays = getConfigNumber("receivable-days", 60);
  const inventoryDays = getConfigNumber("inventory-days", 45);
  const payableDays = getConfigNumber("payable-days", 30);
  const years = [];

  periods.allColumns.forEach((period, idx) => {
    const prev = years[idx - 1];
    const sales = idx === 0 ? mapped.sales : (prev.pl.Sales * (1 + salesGrowth));
    const openingStock = idx === 0 ? Math.max(mapped.closingStock * 0.9, 0) : prev.pl["Closing Stock"];
    const purchases = idx === 0 ? mapped.purchases : prev.pl.Purchases * (1 + salesGrowth * 0.9);
    const directExpenses = idx === 0 ? sales * 0.05 : prev.pl["Direct Expenses"] * (1 + salesGrowth * 0.8);
    const closingStock = idx === 0 ? mapped.closingStock : (sales * inventoryDays) / 365;
    const cogs = openingStock + purchases + directExpenses - closingStock;
    const grossProfit = sales - cogs;
    const employee = ratio(sales, employeeRatio);
    const admin = ratio(sales, adminRatio);
    const selling = ratio(sales, sellingRatio);
    const ebitda = grossProfit - employee - admin - selling;
    const interest = idx === 0 ? Math.max(mapped.workingCapitalBorrowings * 0.11, 0) : prev.pl.Interest * 0.95;
    const depreciation = ratio(Math.max(mapped.capital * 0.5, sales * 0.1), deprRatio);
    const pbt = ebitda + ratio(sales, otherIncomeRatio) - interest - depreciation;
    const tax = Math.max(pbt, 0) * (taxRate / 100);
    const pat = pbt - tax;

    const capital = idx === 0 ? mapped.capital : prev.bs.Capital + prev.pl["Profit After Tax"] * 0.7;
    const reserves = Math.max(capital * 0.2, 0);
    const netWorth = capital + reserves;
    const termLoan = Math.max(mapped.termLoan - idx * (mapped.termLoan / Math.max(1, periods.tlYears)), 0);
    const wcBorrowings = idx === 0 ? mapped.workingCapitalBorrowings : prev.bs["Working Capital Borrowings / CC"] * (1 + salesGrowth * 0.5);
    const unsecured = sales * 0.03;
    const tradeCreditors = idx === 0 ? mapped.tradeCreditors : (purchases * payableDays) / 365;
    const ocl = sales * 0.015;

    const fixedAssets = sales * 0.2;
    const investments = sales * 0.02;
    const inventory = closingStock;
    const debtors = idx === 0 ? mapped.tradeReceivables : (sales * receivableDays) / 365;
    const cashBank = idx === 0 ? mapped.cashBank : Math.max(prev.assetsCash + pat * 0.1, 0);
    const loansAdv = sales * 0.02;
    const oca = sales * 0.01;

    const totalLiabilities = netWorth + termLoan + wcBorrowings + unsecured + tradeCreditors + ocl;
    const totalAssets = fixedAssets + investments + inventory + debtors + cashBank + loansAdv + oca;

    years.push({
      period,
      assetsCash: cashBank,
      pl: {
        Sales: sales,
        "Opening Stock": openingStock,
        Purchases: purchases,
        "Direct Expenses": directExpenses,
        "Closing Stock": closingStock,
        "Cost of Goods Sold": cogs,
        "Gross Profit": grossProfit,
        "Employee Cost": employee,
        "Administrative Expenses": admin,
        "Selling Expenses": selling,
        "EBITDA / Operating Profit": ebitda,
        Interest: interest,
        Depreciation: depreciation,
        "Profit Before Tax": pbt,
        Tax: tax,
        "Profit After Tax": pat,
      },
      bs: {
        "-- Capital & Liabilities --": null,
        Capital: capital,
        "Reserves / Surplus": reserves,
        "Net Worth": netWorth,
        "Term Loan": termLoan,
        "Working Capital Borrowings / CC": wcBorrowings,
        "Unsecured Loans": unsecured,
        "Trade Creditors": tradeCreditors,
        "Other Current Liabilities": ocl,
        "Total Liabilities": totalLiabilities,
        "-- Assets --": null,
        "Fixed Assets": fixedAssets,
        Investments: investments,
        Inventory: inventory,
        "Trade Receivables": debtors,
        "Cash & Bank": cashBank,
        "Loans & Advances": loansAdv,
        "Other Current Assets": oca,
        "Total Assets": totalAssets,
      },
      ratios: {
        "Current Ratio": safeDivide(inventory + debtors + cashBank + loansAdv + oca, tradeCreditors + ocl),
        "Debt Equity": safeDivide(termLoan + wcBorrowings + unsecured, netWorth),
        "Net Profit Margin (%)": safeDivide(pat, sales) * 100,
      },
    });
  });

  return {
    periods,
    mappedFinancials: mapped,
    allYears: years,
    borrowerSummary: {
      "Borrower Name": getConfigText("borrower-name"),
      Constitution: getConfigText("constitution"),
      PAN: getConfigText("pan"),
      GSTIN: getConfigText("gstin"),
      "Bank / Branch": `${getConfigText("bank-name")} / ${getConfigText("branch-name")}`,
      "Existing Limit": getConfigNumber("existing-limit", 0),
      "Proposed Limit": getConfigNumber("proposed-limit", 0),
      "Facility Type": getConfigText("facility-type"),
      "CMA Prepared Date": getConfigText("prepared-date"),
    },
  };
}

function getRowsByMap(yearRecords, mapKey) {
  const keys = Object.keys(yearRecords[0]?.[mapKey] || {});
  return keys.map((key) => ({ label: key, values: yearRecords.map((year) => year[mapKey][key]) }));
}

function buildSectionHtml(title, headerPeriods, rows, formatter = fmtCurrency) {
  const head = `<tr><th>Particulars</th>${headerPeriods.map((p) => `<th>${p}</th>`).join("")}</tr>`;
  const body = rows.map(({ label, values }) => {
    const isHeader = values.every((v) => v === null || v === undefined);
    return `<tr class="${isHeader ? "section-break" : ""}"><td>${label}</td>${values.map((v) => `<td>${formatter(v, label)}</td>`).join("")}</tr>`;
  }).join("");
  return `<section class="report-section"><h3>${title}</h3><table><thead>${head}</thead><tbody>${body}</tbody></table></section>`;
}

function renderReport(report) {
  const periods = report.periods.allColumns;
  const summaryRows = Object.entries(report.borrowerSummary)
    .map(([k, v]) => `<tr><td>${k}</td><td>${typeof v === "number" ? fmtCurrency(v) : v}</td></tr>`).join("");
  const mapRows = MAP_FIELDS.map((field) => `<tr><td>${field.label}</td><td>${fmtCurrency(report.mappedFinancials[field.key])}</td></tr>`).join("");

  output.innerHTML = `
    <article class="cma-report">
      <header class="report-header"><h2>Banker-Style CMA Report</h2></header>
      <section class="report-section"><h3>Summary</h3><table><tbody>${summaryRows}</tbody></table></section>
      <section class="report-section"><h3>Mapped Upload Heads</h3><table><tbody>${mapRows}</tbody></table></section>
      ${buildSectionHtml("Profit & Loss (Vertical)", periods, getRowsByMap(report.allYears, "pl"))}
      ${buildSectionHtml("Balance Sheet (Vertical)", periods, getRowsByMap(report.allYears, "bs"))}
      ${buildSectionHtml("Ratio Analysis", periods, getRowsByMap(report.allYears, "ratios"), (v, l) => l.includes("%") ? fmtPct(v) : fmtNumber(v))}
    </article>`;
}

function generateReport() {
  try {
    if (!workbookState.parsed) throw new Error(workbookState.parseError || "Upload workbook first.");
    const mapped = mappedFinancialsFromParse();
    if (!mapped.sales) throw new Error("Sales mapping is missing or zero. Complete mapping review.");
    if (!mapped.purchases) throw new Error("Purchases mapping is missing or zero. Complete mapping review.");
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
  const html = `<!doctype html><html><head><meta charset="UTF-8"><title>CMA Report</title></head><body>${output.innerHTML}</body></html>`;
  saveFile(html, "cma-report.html", "text/html");
}

function toCell(row, col) { return XLSX.utils.encode_cell({ r: row - 1, c: col - 1 }); }
function setFormula(ws, row, col, formula, format) {
  const ref = toCell(row, col);
  ws[ref] = { t: "n", f: formula };
  if (format) ws[ref].z = format;
}

function makeSectionSheet(title, periods, rows) {
  const ws = XLSX.utils.aoa_to_sheet([[title], ["Particulars", ...periods], ...rows.map((r) => [r.label, ...r.values])]);
  ws["!cols"] = Array.from({ length: periods.length + 1 }, (_, i) => ({ wch: i === 0 ? 40 : 16 }));
  return ws;
}

function downloadExcel() {
  if (!workbookState.generated) return;
  const wb = XLSX.utils.book_new();
  const periods = workbookState.generated.periods.allColumns;
  const plRows = getRowsByMap(workbookState.generated.allYears, "pl");
  const bsRows = getRowsByMap(workbookState.generated.allYears, "bs");
  const ratioRows = getRowsByMap(workbookState.generated.allYears, "ratios");

  const plWs = makeSectionSheet("Profit & Loss", periods, plRows);
  const bsWs = makeSectionSheet("Balance Sheet", periods, bsRows);
  const ratioWs = makeSectionSheet("Ratios", periods, ratioRows);

  periods.forEach((_, i) => {
    const c = i + 2;
    setFormula(plWs, 8, c, `${toCell(4, c)}+${toCell(5, c)}+${toCell(6, c)}-${toCell(7, c)}`, "₹#,##0");
    setFormula(plWs, 9, c, `${toCell(3, c)}-${toCell(8, c)}`, "₹#,##0");
    setFormula(plWs, 13, c, `${toCell(9, c)}-${toCell(10, c)}-${toCell(11, c)}-${toCell(12, c)}`, "₹#,##0");
    setFormula(plWs, 16, c, `${toCell(13, c)}-${toCell(14, c)}-${toCell(15, c)}`, "₹#,##0");
    setFormula(plWs, 18, c, `${toCell(16, c)}-${toCell(17, c)}`, "₹#,##0");

    setFormula(bsWs, 6, c, `${toCell(4, c)}+${toCell(5, c)}`, "₹#,##0");
    setFormula(bsWs, 11, c, `${toCell(6, c)}+${toCell(7, c)}+${toCell(8, c)}+${toCell(9, c)}+${toCell(10, c)}`, "₹#,##0");
    setFormula(bsWs, 20, c, `${toCell(14, c)}+${toCell(15, c)}+${toCell(16, c)}+${toCell(17, c)}+${toCell(18, c)}+${toCell(19, c)}`, "₹#,##0");

    setFormula(ratioWs, 3, c, `('Balance Sheet'!${toCell(16, c)}+'Balance Sheet'!${toCell(17, c)}+'Balance Sheet'!${toCell(18, c)}+'Balance Sheet'!${toCell(19, c)})/('Balance Sheet'!${toCell(9, c)}+'Balance Sheet'!${toCell(10, c)})`, "0.00");
    setFormula(ratioWs, 4, c, `('Balance Sheet'!${toCell(7, c)}+'Balance Sheet'!${toCell(8, c)}+'Balance Sheet'!${toCell(9, c)})/'Balance Sheet'!${toCell(6, c)}`, "0.00");
    setFormula(ratioWs, 5, c, `'Profit & Loss'!${toCell(18, c)}/'Profit & Loss'!${toCell(3, c)}*100`, "0.00");
  });

  XLSX.utils.book_append_sheet(wb, plWs, "Profit & Loss");
  XLSX.utils.book_append_sheet(wb, bsWs, "Balance Sheet");
  XLSX.utils.book_append_sheet(wb, ratioWs, "Ratios");
  XLSX.writeFile(wb, "cma-report.xlsx");
}

function downloadJson() {
  if (!workbookState.generated) return;
  saveFile(JSON.stringify(workbookState.generated, null, 2), "cma-data.json", "application/json");
}
