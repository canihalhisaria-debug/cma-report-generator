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
const generateCurrentBtn = document.getElementById("generate-current-btn");
const output = document.getElementById("output");
const downloadReportBtn = document.getElementById("download-report");
const downloadExcelBtn = document.getElementById("download-excel");
const downloadJsonBtn = document.getElementById("download-json");
const mappingReview = document.getElementById("mapping-review");
const mappingGrid = document.getElementById("mapping-grid");
const mappingWarningPanel = document.getElementById("mapping-warning-panel");
const historicalLockModeInput = document.getElementById("historical-lock-mode");
const downloadTraceBtn = document.getElementById("download-trace");

const configIds = [
  "borrower-name", "constitution", "pan", "gstin", "bank-name", "branch-name", "existing-limit", "proposed-limit", "facility-type",
  "prepared-date", "projection-years-cc", "sales-growth", "other-income-ratio", "gross-margin", "employee-ratio", "admin-ratio",
  "selling-ratio", "depr-ratio", "tax-rate", "receivable-days", "inventory-days", "payable-days", "cc-limit", "proposed-cc-limit",
  "wc-method", "tl-amount", "tl-rate", "tl-years", "moratorium-years",
];
const configInputs = configIds.reduce((acc, id) => ({ ...acc, [id]: document.getElementById(id) }), {});

if (!configInputs["prepared-date"].value) configInputs["prepared-date"].value = new Date().toISOString().slice(0, 10);

const MAP_FIELDS = [
  { key: "sales", label: "Sales", section: "pl", aliases: ["salesaccounts", "sales", "turnover", "revenuefromoperations"] },
  { key: "otherIncome", label: "Other Income", section: "pl", aliases: ["otherincome", "nonoperatingincome", "miscincome"] },
  { key: "openingStock", label: "Opening Stock", section: "pl", aliases: ["openingstock", "openinginventory"] },
  { key: "purchases", label: "Purchases", section: "pl", aliases: ["purchaseaccounts", "purchases", "purchaseaccount"] },
  { key: "directExpenses", label: "Direct Expenses", section: "pl", aliases: ["directexpenses", "manufacturingexpenses", "carriageinward"] },
  { key: "closingStock", label: "Closing Stock", section: "pl", aliases: ["closingstock", "inventory", "stockinhand"] },
  { key: "employeeCost", label: "Employee Cost", section: "pl", aliases: ["salary", "wages", "employeebenefits"] },
  { key: "adminExpenses", label: "Administrative Expenses", section: "pl", aliases: ["administrativeexpenses", "officeexpenses", "generalandadmin"] },
  { key: "sellingExpenses", label: "Selling Expenses", section: "pl", aliases: ["sellingexpenses", "marketing", "distributionexpenses"] },
  { key: "otherOperatingExpenses", label: "Other Operating Expenses", section: "pl", aliases: ["otheroperatingexpenses", "miscoperatingexpenses"] },
  { key: "interest", label: "Interest", section: "pl", aliases: ["interest", "financecost", "bankinterest"] },
  { key: "depreciation", label: "Depreciation", section: "pl", aliases: ["depreciation"] },
  { key: "tax", label: "Tax", section: "pl", aliases: ["provisionfortax", "incometax", "tax"] },

  { key: "capital", label: "Capital", section: "bs", aliases: ["capitalaccount", "capital", "partnerscapital"] },
  { key: "reserves", label: "Reserves", section: "bs", aliases: ["reserve", "surplus", "retainedearnings", "placcount"] },
  { key: "termLoan", label: "Term Loan", section: "bs", aliases: ["termloan", "securedloan", "vehicleloan"] },
  { key: "unsecuredLoans", label: "Unsecured Loans", section: "bs", aliases: ["unsecuredloan", "loanfromdirectors"] },
  { key: "otherLongTermLiabilities", label: "Other Long Term Liabilities", section: "bs", aliases: ["otherlongtermliabilities", "deferredliability"] },
  { key: "ccBorrowing", label: "CC / Bank OD", section: "bs", aliases: ["bankod", "cashcredit", "ccborrowing", "workingcapitalborrowing"] },
  { key: "tradeCreditors", label: "Trade Creditors", section: "bs", aliases: ["sundrycreditors", "tradecreditors", "creditors"] },
  { key: "otherCurrentLiabilities", label: "Other Current Liabilities", section: "bs", aliases: ["othercurrentliabilities", "dutiesandtaxes", "statutorydues"] },

  { key: "fixedAssets", label: "Fixed Assets", section: "bs", aliases: ["fixedassets", "plantandmachinery", "tangibleassets"] },
  { key: "investments", label: "Investments", section: "bs", aliases: ["investments"] },
  { key: "otherNonCurrentAssets", label: "Other Non Current Assets", section: "bs", aliases: ["othernoncurrentassets", "capitalwip", "intangibleassets"] },
  { key: "inventory", label: "Inventory", section: "bs", aliases: ["closingstock", "inventory", "stockinhand"] },
  { key: "tradeReceivables", label: "Trade Receivables", section: "bs", aliases: ["sundrydebtors", "tradereceivables", "debtors"] },
  { key: "cashBank", label: "Cash & Bank", section: "bs", aliases: ["cashatbank", "cashinhand", "bankbalance", "cashbank"] },
  { key: "loansAdvances", label: "Loans & Advances", section: "bs", aliases: ["loansandadvances", "advancesrecoverable"] },
  { key: "otherCurrentAssets", label: "Other Current Assets", section: "bs", aliases: ["othercurrentassets", "prepaidexpenses"] },
];

const CORE_REQUIRED_FIELDS = ["sales", "purchases", "closingStock", "tradeReceivables", "tradeCreditors", "capital", "fixedAssets", "cashBank"];

const STANDARD_CMA_HEADS = {
  netSales: "Net Sales",
  openingStock: "Opening Stock",
  purchases: "Purchases",
  closingStock: "Closing Stock",
  directExpenses: "Direct Expenses",
  otherIncome: "Other Income",
  operatingExpenses: "Operating Expenses",
  interest: "Interest",
  debtors: "Debtors",
  inventory: "Inventory",
  cashBank: "Cash & Bank",
  otherCurrentAssets: "Other Current Assets",
  creditors: "Creditors",
  otherCurrentLiabilities: "Other Current Liabilities",
  shortTermBorrowings: "Short Term Borrowings",
  capital: "Capital",
  fixedAssets: "Fixed Assets",
  nonCurrentOtherAssets: "Non-current / Other Assets",
};

const CMA_HEAD_TO_FIELD_KEY = {
  [STANDARD_CMA_HEADS.netSales]: "sales",
  [STANDARD_CMA_HEADS.openingStock]: "openingStock",
  [STANDARD_CMA_HEADS.purchases]: "purchases",
  [STANDARD_CMA_HEADS.closingStock]: "closingStock",
  [STANDARD_CMA_HEADS.directExpenses]: "directExpenses",
  [STANDARD_CMA_HEADS.otherIncome]: "otherIncome",
  [STANDARD_CMA_HEADS.interest]: "interest",
  [STANDARD_CMA_HEADS.debtors]: "tradeReceivables",
  [STANDARD_CMA_HEADS.inventory]: "inventory",
  [STANDARD_CMA_HEADS.cashBank]: "cashBank",
  [STANDARD_CMA_HEADS.otherCurrentAssets]: "otherCurrentAssets",
  [STANDARD_CMA_HEADS.creditors]: "tradeCreditors",
  [STANDARD_CMA_HEADS.otherCurrentLiabilities]: "otherCurrentLiabilities",
  [STANDARD_CMA_HEADS.shortTermBorrowings]: "ccBorrowing",
  [STANDARD_CMA_HEADS.capital]: "capital",
  [STANDARD_CMA_HEADS.fixedAssets]: "fixedAssets",
  [STANDARD_CMA_HEADS.nonCurrentOtherAssets]: "otherNonCurrentAssets",
};

const DEFAULT_FALLBACKS = {
  otherIncome: { value: 0, note: "Default 0" },
  openingStock: { value: 0, note: "Default 0" },
  otherOperatingExpenses: { value: 0, note: "Default 0" },
  depreciation: { value: 0, note: "Default 0" },
  tax: { value: 0, note: "Default 0" },
  reserves: { value: 0, note: "Default 0" },
  termLoan: { value: 0, note: "Default 0 (Not Applicable)" },
  unsecuredLoans: { value: 0, note: "Default 0" },
  otherLongTermLiabilities: { value: 0, note: "Default 0" },
  investments: { value: 0, note: "Default 0" },
  otherNonCurrentAssets: { value: 0, note: "Default 0" },
  loansAdvances: { value: 0, note: "Default 0" },
};

fileInput.addEventListener("change", handleWorkbookUpload);
generateBtn.addEventListener("click", () => generateReport({ useCurrentMapping: false }));
generateCurrentBtn.addEventListener("click", () => generateReport({ useCurrentMapping: true }));
downloadReportBtn.addEventListener("click", downloadReport);
downloadExcelBtn.addEventListener("click", downloadExcel);
downloadJsonBtn.addEventListener("click", downloadJson);
downloadTraceBtn?.addEventListener("click", downloadHistoricalTrace);

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
function normalizeConcept(value) {
  return normalize(value)
    .replace(/accounts?|ledger|group|head|main|under/g, "")
    .replace(/[^a-z0-9]/g, "");
}

function isBankOdConcept(text) {
  return /bankod|cashcredit|workingcapital|cc\b|overdraft/.test(text);
}


function classifyAdvanceByIntent(text) {
  if (/advancetosupplier|supplieradvance|advanceforsupplier|advancevendor/.test(text)) return "otherCurrentAssets";
  if (/loanrelatedparty|relatedpartyloan|loanpartner|advancetopartner|loanassociate|loanrelated/.test(text)) return "loansAdvances";
  if (/securitydeposit|rentdeposit|electricitydeposit|tenderdeposit|deposit/.test(text)) return "otherNonCurrentAssets";
  return null;
}

function detectMandatoryCmaHead(head) {
  const text = normalize(head.head);
  if (head.section === "pl") {
    if (/salesaccounts?|sales?|sale\b/.test(text)) return STANDARD_CMA_HEADS.netSales;
    if (/openingstock/.test(text)) return STANDARD_CMA_HEADS.openingStock;
    if (/purchaseaccounts?|purchases?/.test(text)) return STANDARD_CMA_HEADS.purchases;
    if (/closingstock/.test(text)) return STANDARD_CMA_HEADS.closingStock;
    if (/indirectincome/.test(text)) return STANDARD_CMA_HEADS.otherIncome;
    if (/interestoncc|ccinterest|bankinterest|financecost|interest/.test(text)) return STANDARD_CMA_HEADS.interest;
    if (/directexpense|manufacturingexpense|carriageinward|jobwork|power|fuel|packing/.test(text)) return STANDARD_CMA_HEADS.directExpenses;
    if (/salary|wage|employee|admin|office|selling|marketing|repair|maintenance|travelling|miscexpense|indirectexpense/.test(text)) return STANDARD_CMA_HEADS.operatingExpenses;
    return null;
  }

  if (/sundrydebtors?|debtors?|tradereceivables?/.test(text)) return STANDARD_CMA_HEADS.debtors;
  if (/closingstock|inventory|stockinhand/.test(text)) return STANDARD_CMA_HEADS.inventory;
  if (/cashinhand|cashatbank|bankbalance|cashbank/.test(text)) return STANDARD_CMA_HEADS.cashBank;
  if (/gstitc|tdsreceivable|othercurrentassets?/.test(text)) return STANDARD_CMA_HEADS.otherCurrentAssets;
  if (/sundrycreditors?|tradecreditors?|creditors?/.test(text)) return STANDARD_CMA_HEADS.creditors;
  if (/advancefromcustomer|kmr|salarypayable|reimbursementdue|tdspayable/.test(text)) return STANDARD_CMA_HEADS.otherCurrentLiabilities;
  if (/bankod|overdraft|cashcredit|cc\b/.test(text)) return STANDARD_CMA_HEADS.shortTermBorrowings;
  if (/partnerscapital|capitalaccount|capital\b/.test(text)) return STANDARD_CMA_HEADS.capital;
  if (/fixedassets?/.test(text)) return STANDARD_CMA_HEADS.fixedAssets;
  if (/advances?anddeposits?|securitydeposit|deposit/.test(text)) return STANDARD_CMA_HEADS.nonCurrentOtherAssets;
  return null;
}

function cmaHeadForFieldKey(key) {
  return Object.keys(CMA_HEAD_TO_FIELD_KEY).find((head) => CMA_HEAD_TO_FIELD_KEY[head] === key) || "";
}

function buildLayerData(parsed, historicalSources, mapped) {
  const sourceExtract = parsed.allHeads.map((head) => ({
    sourceHead: head.head,
    groupType: head.section === "pl" ? "P&L" : "Balance Sheet",
    amount: head.amount,
    rowNumber: head.rowNumber || null,
    sheet: head.section === "pl" ? parsed.plSheetName : parsed.bsSheetName,
  }));

  const mappedRows = [];
  Object.entries(historicalSources || {}).forEach(([key, rows]) => {
    if (!rows.length) {
      if ((mapped[key] || 0) !== 0) {
        mappedRows.push({
          sourceHead: "-",
          groupType: MAP_FIELDS.find((f) => f.key === key)?.section === "pl" ? "P&L" : "Balance Sheet",
          mappedCmaHead: cmaHeadForFieldKey(key) || MAP_FIELDS.find((f) => f.key === key)?.label || key,
          fy1Amount: mapped[key] || 0,
          includeFlag: "Include",
        });
      }
      return;
    }

    rows.forEach((row) => {
      mappedRows.push({
        sourceHead: row.head,
        groupType: row.section === "pl" ? "P&L" : "Balance Sheet",
        mappedCmaHead: cmaHeadForFieldKey(key) || MAP_FIELDS.find((f) => f.key === key)?.label || key,
        fy1Amount: row.amount,
        includeFlag: row.suppressedReason ? "Exclude" : "Include",
      });
    });
  });

  return { sourceExtract, cmaInputMap: mappedRows };
}
function fmtCurrency(v) { return v === null || v === undefined ? "" : new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", maximumFractionDigits: 0 }).format(v); }
function fmtPct(v) { return `${(v || 0).toFixed(2)}%`; }
function fmtNumber(v) { return v === null || v === undefined ? "" : (v || 0).toFixed(2); }

function mappedValueByHead(rows, cmaHead) {
  return (rows || [])
    .filter((row) => row.mappedCmaHead === cmaHead && row.includeFlag === "Include")
    .reduce((sum, row) => sum + getNumeric(row.fy1Amount, 0), 0);
}

function detectSheetByContent(workbook, type) {
  const hintMap = {
    pl: ["profitandloss", "trading", "sales", "purchase", "grossprofit", "expenses"],
    bs: ["balancesheet", "assets", "liabilities", "capital", "debtors", "creditors"],
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

      records.push({
        head,
        normalized,
        amount: Math.abs(amountCell.value),
        rowNumber: rowIndex + 1,
      });
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
      if (field.key === "directExpenses" && /indirectexpenses?/.test(headRow.normalized)) return;
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
    const scopedHeads = heads.filter((h) => h.section === field.section);
    select.innerHTML = `<option value="">-- Not mapped --</option>${scopedHeads.map((h) => `<option value="${h.head}">${h.head} (${fmtCurrency(h.amount)})</option>`).join("")}`;
    select.value = mapping[field.key] || "";
    select.addEventListener("change", () => { workbookState.mapping[field.key] = select.value; });

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
  if (!plHeads.length || !bsHeads.length) throw new Error("Could not parse Tally style statements.");

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
      output.innerHTML = "<p class='good'>Workbook parsed successfully. Validate mapping (if shown) and generate report.</p>";
      generateBtn.disabled = false;
      generateCurrentBtn.disabled = false;
      renderMappingWarnings();
    } catch (error) {
      workbookState.parsed = null;
      workbookState.parseError = error.message;
      generateBtn.disabled = true;
      generateCurrentBtn.disabled = true;
      mappingReview.classList.add("hidden");
      renderMappingWarnings();
      sheetStatus.innerHTML = `<span class="bad">Upload parsing failed: ${error.message}</span>`;
      output.innerHTML = "<p class='bad'>Unable to parse workbook.</p>";
    }
  };

  reader.readAsArrayBuffer(file);
}

function getMappedAmount(selectedHead, section) {
  if (!selectedHead || !workbookState.parsed) return 0;
  return workbookState.parsed.allHeads.find((h) => h.head === selectedHead && h.section === section)?.amount || 0;
}

function strictHistoricalBucket(head) {
  const mandatoryHead = detectMandatoryCmaHead(head);
  if (mandatoryHead && CMA_HEAD_TO_FIELD_KEY[mandatoryHead]) return CMA_HEAD_TO_FIELD_KEY[mandatoryHead];
  const text = normalize(head.head);
  if (head.section === "pl") {
    if (/indirectexpenses?|administrative|officeexpense|generalandadmin/.test(text)) return "adminExpenses";
    if (/otheroperatingexpenses?|miscoperatingexpenses?/.test(text)) return "otherOperatingExpenses";
    if (/openingstock|openinginventory/.test(text)) return "openingStock";
    if (/purchaseaccounts?|purchases?/.test(text)) return "purchases";
    return null;
  }

  if (/sundrydebtors?|tradereceivables?|debtors?/.test(text)) return "tradeReceivables";
  if (/sundrycreditors?|tradecreditors?|creditors?/.test(text)) return "tradeCreditors";
  if (isBankOdConcept(text)) return "ccBorrowing";
  if (/capitalaccount|partnerscapital|capital\b/.test(text)) return "capital";
  if (/cashinhand|cashatbank|bankbalance|currentbank|currentaccount|bankaccount|cashbank|bankof|hdfc|icici|axis|sbi/.test(text)) return "cashBank";
  const advanceClass = classifyAdvanceByIntent(text);
  if (advanceClass) return advanceClass;
  if (/advancesrecoverable|advancepaid|loans?andadvances?/.test(text)) return "loansAdvances";
  if (/gstitc|tdsreceivable|othercurrentassets?|inputcredit|prepaid|receivable|statutoryreceivable/.test(text)) return "otherCurrentAssets";
  if (/advancefromcustomer|salarypayable|tdspayable|reimbursementpayable|reimbursementdue|outstanding|expensepayable|statutorydues|gstpayable|payable/.test(text)) return "otherCurrentLiabilities";
  if (/loans?andadvances?|advanceagainstloan/.test(text)) return "loansAdvances";
  if (/loan|debenture|longterm|deferred|mortgage/.test(text) && !isBankOdConcept(text)) return "otherLongTermLiabilities";
  if (/stock|invent/.test(text)) return "inventory";
  return null;
}

function mappedFinancialsFromParse({ historicalLockMode = true } = {}) {
  if (!workbookState.parsed) throw new Error("Upload and parse workbook first.");

  const mapped = {};
  const historicalSources = {};
  MAP_FIELDS.forEach((field) => {
    mapped[field.key] = 0;
    historicalSources[field.key] = [];
  });

  const pushSource = (key, head, mode, suppressedReason = "") => {
    if (!suppressedReason) mapped[key] = (mapped[key] || 0) + head.amount;
    historicalSources[key].push({
      head: head.head,
      section: head.section,
      sheet: head.section === "pl" ? workbookState.parsed.plSheetName : workbookState.parsed.bsSheetName,
      amount: head.amount,
      rowNumber: head.rowNumber || null,
      mode,
      suppressedReason,
    });
  };

  const candidates = [];
  const duplicateWarnings = [];
  const parentChildWarnings = [];

  const queueCandidate = (key, head, mode) => {
    candidates.push({
      key,
      head,
      mode,
      normalized: normalizeConcept(head.head).replace(/accounts?|account|ledger$/, ""),
      amount: head.amount,
      rowNumber: head.rowNumber || 999999,
    });
  };

  MAP_FIELDS.forEach((field) => {
    const selectedHead = workbookState.mapping[field.key];
    if (!selectedHead) return;
    const head = workbookState.parsed.allHeads.find((h) => h.head === selectedHead && h.section === field.section);
    if (!head) return;
    queueCandidate(field.key, head, "explicit-mapping");
  });

  const fallbackDefaults = [];
  if (!historicalLockMode) {
    Object.entries(DEFAULT_FALLBACKS).forEach(([key, fallback]) => {
      if ((mapped[key] || 0) !== 0) return;
      mapped[key] = fallback.value;
      const field = MAP_FIELDS.find((f) => f.key === key);
      fallbackDefaults.push({ key, label: field?.label || key, value: fallback.value, note: fallback.note });
    });
  }

  const fallbackWarnings = [];
  const mappedHeadsBySection = {
    pl: new Set(MAP_FIELDS.filter((f) => f.section === "pl").map((f) => workbookState.mapping[f.key]).filter(Boolean)),
    bs: new Set(MAP_FIELDS.filter((f) => f.section === "bs").map((f) => workbookState.mapping[f.key]).filter(Boolean)),
  };

  workbookState.parsed.allHeads
    .filter((head) => !mappedHeadsBySection[head.section].has(head.head))
    .forEach((head) => {
      const fallbackKey = historicalLockMode ? strictHistoricalBucket(head) : classifyFallbackBucket(head);
      if (!fallbackKey) return;
      if (historicalLockMode && ["termLoan"].includes(fallbackKey)) return;
      queueCandidate(fallbackKey, head, historicalLockMode ? "historical-lock-rule" : "auto-fallback");
      fallbackWarnings.push({
        head: head.head,
        section: head.section,
        amount: head.amount,
        fallbackKey,
        fallbackLabel: MAP_FIELDS.find((f) => f.key === fallbackKey)?.label || fallbackKey,
      });
    });

  const specificityScore = (candidate) => {
    const text = normalize(candidate.head.head);
    let score = 0;
    if (candidate.mode === "explicit-mapping") score += 20;
    if (!/accounts?|group|total/.test(text)) score += 5;
    score += Math.max(0, 15 - text.length * 0.05);
    return score;
  };

  const selected = [];
  const singleBucketCandidates = [];
  const bySourceRow = new Map();
  candidates.forEach((candidate) => {
    const sourceId = `${candidate.head.section}::${candidate.head.head}::${candidate.rowNumber}`;
    if (!bySourceRow.has(sourceId)) {
      bySourceRow.set(sourceId, candidate);
      return;
    }
    const existing = bySourceRow.get(sourceId);
    const better = specificityScore(candidate) > specificityScore(existing) ? candidate : existing;
    const dropped = better === candidate ? existing : candidate;
    bySourceRow.set(sourceId, better);
    duplicateWarnings.push({
      head: dropped.head.head,
      amount: dropped.amount,
      key: dropped.key,
      reason: `Single-bucket guard: source row already assigned to ${better.key}`,
    });
    pushSource(dropped.key, dropped.head, dropped.mode, "source-row-multi-bucket-suppressed");
  });
  singleBucketCandidates.push(...Array.from(bySourceRow.values()));

  const byBucketAndConcept = new Map();
  singleBucketCandidates.forEach((candidate) => {
    const concept = `${candidate.key}::${candidate.normalized}`;
    if (!byBucketAndConcept.has(concept)) {
      byBucketAndConcept.set(concept, candidate);
      return;
    }
    const existing = byBucketAndConcept.get(concept);
    const better = specificityScore(candidate) > specificityScore(existing) ? candidate : existing;
    const dropped = better === candidate ? existing : candidate;
    byBucketAndConcept.set(concept, better);
    duplicateWarnings.push({
      head: dropped.head.head,
      amount: dropped.amount,
      key: dropped.key,
      reason: `Duplicate mapping: same concept as ${better.head.head}`,
    });
    pushSource(dropped.key, dropped.head, dropped.mode, "duplicate-concept-suppressed");
  });

  const conceptSelected = Array.from(byBucketAndConcept.values()).sort((a, b) => a.rowNumber - b.rowNumber);
  const groupedByKey = conceptSelected.reduce((acc, item) => {
    if (!acc[item.key]) acc[item.key] = [];
    acc[item.key].push(item);
    return acc;
  }, {});

  Object.values(groupedByKey).forEach((rows) => {
    const suppressed = new Set();
    rows.forEach((row, idx) => {
      rows.forEach((other, jdx) => {
        if (idx === jdx || suppressed.has(row) || suppressed.has(other)) return;
        const rowNorm = normalize(row.head.head);
        const otherNorm = normalize(other.head.head);
        const contains = rowNorm.includes(otherNorm) || otherNorm.includes(rowNorm);
        const nearEqual = Math.abs(row.amount - other.amount) <= Math.max(row.amount, other.amount) * 0.02;
        const closeRows = Math.abs((row.rowNumber || 0) - (other.rowNumber || 0)) <= 3;
        if (!contains || !nearEqual || !closeRows) return;
        const likelyParent = /accounts?|group|total|liabilit|asset/.test(rowNorm) ? row : other;
        const likelyChild = likelyParent === row ? other : row;
        suppressed.add(likelyParent);
        parentChildWarnings.push({
          parent: likelyParent.head.head,
          child: likelyChild.head.head,
          key: likelyParent.key,
        });
      });
    });

    rows.forEach((row) => {
      if (suppressed.has(row)) {
        pushSource(row.key, row.head, row.mode, "parent-child-suppressed");
        return;
      }
      selected.push(row);
    });
  });

  selected.forEach((candidate) => pushSource(candidate.key, candidate.head, candidate.mode));

  const classificationReview = workbookState.parsed.allHeads.map((head) => {
    const explicit = MAP_FIELDS.find((field) => workbookState.mapping[field.key] === head.head && field.section === head.section)?.key;
    const suggestedKey = explicit || (historicalLockMode ? strictHistoricalBucket(head) : classifyFallbackBucket(head));
    const dupRisk = duplicateWarnings.some((d) => d.head === head.head);
    const parentChildFlag = parentChildWarnings.some((p) => p.parent === head.head || p.child === head.head);
    return {
      sourceRow: `${head.head} (Row ${head.rowNumber || "-"})`,
      sourceValue: head.amount,
      suggestedBucket: MAP_FIELDS.find((f) => f.key === suggestedKey)?.label || "Unclassified",
      duplicateRisk: dupRisk ? "Yes" : "No",
      parentChildFlag: parentChildFlag ? "Yes" : "No",
    };
  });

  const missingMandatory = CORE_REQUIRED_FIELDS
    .filter((key) => !workbookState.mapping[key] && !(key === "closingStock" && workbookState.mapping.inventory))
    .map((key) => MAP_FIELDS.find((f) => f.key === key)?.label || key);

  return {
    mapped,
    fallbackWarnings,
    fallbackDefaults,
    missingMandatory,
    historicalSources,
    duplicateWarnings,
    parentChildWarnings,
    classificationReview,
  };
}

function classifyFallbackBucket(head) {
  const mandatoryHead = detectMandatoryCmaHead(head);
  if (mandatoryHead && CMA_HEAD_TO_FIELD_KEY[mandatoryHead]) return CMA_HEAD_TO_FIELD_KEY[mandatoryHead];
  const text = normalize(head.head);
  if (head.section === "pl") {
    if (/income|commission|discountreceived|incentive|rentreceived|interestreceived/.test(text)) return "otherIncome";
    if (/openingstock|openinginventory/.test(text)) return "openingStock";
    if (/purchase|purchaseaccount|purchaseaccounts/.test(text)) return "purchases";
    if (/indirectexpenses|indirectexpense/.test(text)) return "adminExpenses";
    if (/raw|material|consum|manufact|factory|production|jobwork|carriageinward|power|fuel|packing/.test(text)) return "directExpenses";
    if (/salary|wage|staff|employee|labou?r|pf|esi|bonus|gratuity/.test(text)) return "employeeCost";
    if (/admin|office|audit|legal|professional|telephone|internet|printing|stationery|insurance|repair|maintenance/.test(text)) return "adminExpenses";
    if (/selling|market|advert|salespromo|freightout|distribution|commissionpaid|travelling/.test(text)) return "sellingExpenses";
    return "otherOperatingExpenses";
  }

  if (/sundrydebtor|tradereceivable|debtor/.test(text)) return "tradeReceivables";
  if (/sundrycreditor|tradecreditor|creditor/.test(text)) return "tradeCreditors";
  if (isBankOdConcept(text)) return "ccBorrowing";
  if (/cashinhand|cashatbank|bankbalance|currentaccount|bankaccount|overdraftaccount|cashbank/.test(text)) return "cashBank";
  const advanceClass = classifyAdvanceByIntent(text);
  if (advanceClass) return advanceClass;
  if (/advancetohardik|advancerent|advancesalary|loanandadvance|advancepaid/.test(text)) return "loansAdvances";
  if (/gstitc|tdsreceivable|othercurrentasset|inputcredit|prepaid/.test(text)) return "otherCurrentAssets";
  if (/advancefromcustomer|salarypayable|reimbursementdue|tdspayable|kmr/.test(text)) return "otherCurrentLiabilities";
  if (/creditor|payable|dut|tax|gst|expensepayable|provision|outstanding|accrued|liabilit/.test(text)) return "otherCurrentLiabilities";
  if (/loan|debenture|longterm|deferred|borrow|mortgage/.test(text) && !isBankOdConcept(text)) return "otherLongTermLiabilities";
  if (/stock|invent/.test(text)) return "inventory";
  if (/advance|currentasset/.test(text)) return "otherCurrentAssets";
  return "otherNonCurrentAssets";
}

function renderMappingWarnings({
  fallbackWarnings = [],
  fallbackDefaults = [],
  missingMandatory = [],
  duplicateWarnings = [],
  parentChildWarnings = [],
} = {}) {
  if (!fallbackWarnings.length && !fallbackDefaults.length && !missingMandatory.length && !duplicateWarnings.length && !parentChildWarnings.length) {
    mappingWarningPanel.classList.add("hidden");
    mappingWarningPanel.innerHTML = "";
    return;
  }

  const fallbackItems = fallbackWarnings
    .map((item) => `<li><strong>${item.head}</strong> (${item.section.toUpperCase()}, ${fmtCurrency(item.amount)}) → <strong>${item.fallbackLabel}</strong></li>`)
    .join("");
  const defaultItems = fallbackDefaults
    .map((item) => `<li><strong>${item.label}</strong> → ${fmtCurrency(item.value)} <em>(${item.note})</em></li>`)
    .join("");
  const mandatoryItems = missingMandatory.map((item) => `<li>${item}</li>`).join("");
  const duplicateItems = duplicateWarnings.map((item) => `<li><strong>${item.head}</strong> (${fmtCurrency(item.amount)}) → ${item.reason}</li>`).join("");
  const parentChildItems = parentChildWarnings.map((item) => `<li><strong>${item.parent}</strong> suppressed; child retained: <strong>${item.child}</strong></li>`).join("");

  mappingWarningPanel.innerHTML = `
    <h3>⚠️ Mapping Warnings</h3>
    ${missingMandatory.length ? `<p><strong>Missing core heads (generation blocked):</strong></p><ul>${mandatoryItems}</ul>` : ""}
    ${fallbackDefaults.length ? `<p><strong>Fallback defaults applied:</strong></p><ul>${defaultItems}</ul>` : ""}
    ${fallbackWarnings.length ? `<p><strong>Unmapped heads assigned to fallback buckets:</strong></p><ul>${fallbackItems}</ul>` : ""}
    ${duplicateWarnings.length ? `<p><strong>Duplicate concept suppression:</strong></p><ul>${duplicateItems}</ul>` : ""}
    ${parentChildWarnings.length ? `<p><strong>Parent / child group suppression:</strong></p><ul>${parentChildItems}</ul>` : ""}
  `;
  mappingWarningPanel.classList.remove("hidden");

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

  const hasInstallment = rows.some((row) => row.installment > 0);
  const warning = !hasInstallment ? "Installment is zero for all projected years. Repayment structure is incomplete." : "";

  return { applicable: true, rows, warning };
}

function buildCmaReport(mapped, { historicalLockMode = true } = {}) {
  const periods = projectionPeriods();
  const salesGrowth = getConfigNumber("sales-growth", 10) / 100;
  const otherIncomePct = getConfigNumber("other-income-ratio", 1.5) / 100;
  const employeePct = getConfigNumber("employee-ratio", 6) / 100;
  const adminPct = getConfigNumber("admin-ratio", 4) / 100;
  const sellingPct = getConfigNumber("selling-ratio", 3) / 100;
  const deprPct = getConfigNumber("depr-ratio", 10) / 100;
  const taxPct = getConfigNumber("tax-rate", 25) / 100;

  const debtorDays = getConfigNumber("receivable-days", 60);
  const inventoryDays = getConfigNumber("inventory-days", 45);
  const creditorDays = getConfigNumber("payable-days", 30);

  const existingCCLimit = getConfigNumber("cc-limit", mapped.ccBorrowing || 0);
  const historicalCcOutstanding = mapped.ccBorrowing || 0;
  const termLoan = buildTermLoanSchedule(periods, mapped.termLoan || 0);
  const years = [];

  periods.forEach((period, idx) => {
    const isHistoricalYear = idx === 0;
    const prev = years[idx - 1];
    const sales = isHistoricalYear ? mapped.sales : prev.pl.Sales * (1 + salesGrowth);
    const otherIncome = isHistoricalYear ? mapped.otherIncome : sales * otherIncomePct;

    const openingStock = isHistoricalYear ? mapped.openingStock : prev.pl["Less Closing Stock"];
    const purchases = isHistoricalYear ? mapped.purchases : prev.pl.Purchases * (1 + salesGrowth * 0.95);
    const directExpenses = isHistoricalYear ? mapped.directExpenses : prev.pl["Direct Expenses"] * (1 + salesGrowth * 0.9);
    const closingStock = isHistoricalYear ? mapped.closingStock : (sales * inventoryDays) / 365;

    const goodsAvailable = openingStock + purchases + directExpenses;
    const cogs = goodsAvailable - closingStock;
    const totalIncome = sales + otherIncome;
    const grossProfit = sales - cogs;

    const employeeCost = isHistoricalYear ? mapped.employeeCost : sales * employeePct;
    const administrativeExpenses = isHistoricalYear ? mapped.adminExpenses : sales * adminPct;
    const sellingExpenses = isHistoricalYear ? mapped.sellingExpenses : sales * sellingPct;
    const otherOperatingExpenses = isHistoricalYear ? mapped.otherOperatingExpenses : sales * 0.01;
    const totalOperatingExpenses = employeeCost + administrativeExpenses + sellingExpenses + otherOperatingExpenses;
    const ebit = grossProfit + otherIncome - totalOperatingExpenses;
    const ebitda = ebit;

    const tlInterest = termLoan.applicable ? termLoan.rows[idx].interest : 0;
    const wcInterest = existingCCLimit * 0.11;
    const interest = isHistoricalYear
      ? mapped.interest
      : Math.max(mapped.interest || 0, wcInterest * 0.5) + tlInterest;
    const depreciation = isHistoricalYear
      ? mapped.depreciation
      : Math.max(prev.bs["Fixed Assets"] * deprPct, sales * 0.01);

    const pbt = ebit - interest;
    const tax = isHistoricalYear ? 0 : Math.max(pbt, 0) * taxPct;
    const pat = isHistoricalYear ? pbt : (pbt - tax);

    const capital = isHistoricalYear ? mapped.capital : prev.bs.Capital;
    const reserves = isHistoricalYear ? mapped.reserves : prev.bs.Reserves + prev.pl["Profit After Tax"] * 0.7;
    const netWorth = capital + reserves;

    const termLoanOutstanding = isHistoricalYear ? mapped.termLoan : (termLoan.applicable ? termLoan.rows[idx].opening : 0);
    const unsecuredLoans = isHistoricalYear ? mapped.unsecuredLoans : prev.bs["Unsecured Loans"];
    const otherLongTermLiabilities = isHistoricalYear ? mapped.otherLongTermLiabilities : prev.bs["Other Long Term Liabilities"];
    const totalNonCurrentLiabilities = termLoanOutstanding + unsecuredLoans + otherLongTermLiabilities;

    const ccOd = isHistoricalYear ? historicalCcOutstanding : existingCCLimit;
    const tradeCreditors = isHistoricalYear ? mapped.tradeCreditors : (purchases * creditorDays) / 365;
    const otherCurrentLiabilities = isHistoricalYear ? mapped.otherCurrentLiabilities : sales * 0.015;
    const totalCurrentLiabilities = ccOd + tradeCreditors + otherCurrentLiabilities;

    const fixedAssets = isHistoricalYear ? mapped.fixedAssets : Math.max(prev.bs["Fixed Assets"] * 0.95, sales * 0.15);
    const investments = isHistoricalYear ? mapped.investments : prev.bs.Investments;
    const otherNonCurrentAssets = isHistoricalYear ? mapped.otherNonCurrentAssets : prev.bs["Other Non Current Assets"];
    const totalNonCurrentAssets = fixedAssets + investments + otherNonCurrentAssets;

    const inventory = isHistoricalYear ? mapped.inventory : closingStock;
    const tradeReceivables = isHistoricalYear ? mapped.tradeReceivables : (sales * debtorDays) / 365;
    const cashBank = isHistoricalYear ? mapped.cashBank : Math.max(prev.bs["Cash & Bank"] + pat * 0.12, 0);
    const loansAdvances = isHistoricalYear ? mapped.loansAdvances : prev.bs["Loans & Advances"];
    const otherCurrentAssets = isHistoricalYear ? mapped.otherCurrentAssets : sales * 0.01;
    const totalCurrentAssets = inventory + tradeReceivables + cashBank + otherCurrentAssets;

    const totalLiabilities = netWorth + totalNonCurrentLiabilities + totalCurrentLiabilities;
    const totalAssets = totalNonCurrentAssets + totalCurrentAssets;

    const netWorkingCapital = totalCurrentAssets - totalCurrentLiabilities;
    const otherCurrentLiabilitiesOnly = tradeCreditors + otherCurrentLiabilities;
    const workingCapitalGap = Math.max(totalCurrentAssets - otherCurrentLiabilitiesOnly, 0);
    const mpbf = Math.max((0.75 * totalCurrentAssets) - otherCurrentLiabilitiesOnly, 0);
    const borrowerContribution = Math.max(workingCapitalGap - mpbf, 0);
    const shortfall = Math.max(mpbf - existingCCLimit, 0);
    const proposedCCLimit = isHistoricalYear ? historicalCcOutstanding : existingCCLimit + shortfall;

    const installment = termLoan.applicable ? termLoan.rows[idx].installment : 0;
    const dscr = termLoan.applicable ? safeDivide(pat + depreciation + tlInterest, installment + tlInterest) : null;

    years.push({
      year: period,
      pl: {
        Sales: sales,
        "Other Income": otherIncome,
        "Total Income": totalIncome,
        "Opening Stock": openingStock,
        Purchases: purchases,
        "Direct Expenses": directExpenses,
        "Goods Available for Sale": goodsAvailable,
        "Less Closing Stock": closingStock,
        "Cost of Goods Sold": cogs,
        "Gross Profit": grossProfit,
        "Employee Cost": employeeCost,
        "Administrative Expenses": administrativeExpenses,
        "Selling Expenses": sellingExpenses,
        "Other Operating Expenses": otherOperatingExpenses,
        "Total Operating Expenses": totalOperatingExpenses,
        EBITDA: ebitda,
        EBIT: ebit,
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
        "Unsecured Loans": unsecuredLoans,
        "Other Long Term Liabilities": otherLongTermLiabilities,
        "Total Non Current Liabilities": totalNonCurrentLiabilities,
        "CC / Bank OD": ccOd,
        "Trade Creditors": tradeCreditors,
        "Other Current Liabilities": otherCurrentLiabilities,
        "Total Current Liabilities": totalCurrentLiabilities,
        "Total Liabilities": totalLiabilities,
        "Fixed Assets": fixedAssets,
        Investments: investments,
        "Other Non Current Assets": otherNonCurrentAssets,
        "Total Non Current Assets": totalNonCurrentAssets,
        Inventory: inventory,
        "Trade Receivables": tradeReceivables,
        "Cash & Bank": cashBank,
        "Loans & Advances": loansAdvances,
        "Other Current Assets": otherCurrentAssets,
        "Total Current Assets": totalCurrentAssets,
        "Total Assets": totalAssets,
      },
      workingCapital: {
        projectedSales: sales,
        totalCurrentAssets,
        totalCurrentLiabilities,
        netWorkingCapital,
        workingCapitalGap,
        borrowerContribution,
        mpbf,
        bankFinanceRequired: mpbf,
        currentRatio: safeDivide(totalCurrentAssets, totalCurrentLiabilities),
        requiredWorkingCapital: mpbf,
        existingLimit: isHistoricalYear ? historicalCcOutstanding : existingCCLimit,
        proposedLimit: proposedCCLimit,
        shortfall,
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
      "Facility Type": getConfigText("facility-type"),
      "Prepared Date": getConfigText("prepared-date"),
      "Historical Lock Mode": historicalLockMode ? "Enabled" : "Disabled",
    },
  };
}

function buildStatementRows(years, key, rowLabels) {
  return rowLabels.map((label) => ({ label, values: years.map((y) => y[key][label]) }));
}

function buildSectionHtml(title, periods, rows, formatter = fmtCurrency) {
  const head = `<tr><th>Particulars</th>${periods.map((period) => `<th>${period}</th>`).join("")}</tr>`;
  const body = rows.map((row) => {
    const trClass = /total|profit|ebitda|net worth/i.test(row.label) ? " class='total-row'" : (/income|liabilities|assets|expenses|charges|capital/i.test(row.label) ? " class='section-row'" : "");
    return `<tr${trClass}><td>${row.label}</td>${row.values.map((value) => `<td>${formatter(value, row.label)}</td>`).join("")}</tr>`;
  }).join("");
  return `<section class="report-section"><h3>${title}</h3><table><thead>${head}</thead><tbody>${body}</tbody></table></section>`;
}

function buildHistoricalDebugHtml(report) {
  const fieldLabel = Object.fromEntries(MAP_FIELDS.map((f) => [f.key, f.label]));
  const debugRows = Object.entries(report.meta?.historicalDebug?.sources || {}).flatMap(([key, sources]) => {
    if (!sources.length) {
      return [`<tr><td>${fieldLabel[key] || key}</td><td>-</td><td>-</td><td>${fmtCurrency(report.meta?.historicalDebug?.values?.[key] || 0)}</td><td>-</td></tr>`];
    }
    return sources.map((src) => `<tr><td>${fieldLabel[key] || key}</td><td>${src.sheet || src.section.toUpperCase()}</td><td>${src.head} (Row ${src.rowNumber || "-"})</td><td>${fmtCurrency(src.amount)}</td><td>${src.suppressedReason || "-"}</td></tr>`);
  }).join("");

  const reviewRows = (report.meta?.classificationReview || []).map((row) => `
    <tr>
      <td>${row.sourceRow}</td>
      <td>${fmtCurrency(row.sourceValue)}</td>
      <td>${row.suggestedBucket}</td>
      <td>${row.duplicateRisk}</td>
      <td>${row.parentChildFlag}</td>
    </tr>
  `).join("");

  return `
    <section class="report-section">
      <h3>FY-1 Historical Mapping Debug Review</h3>
      <table>
        <thead><tr><th>Final CMA Line Item</th><th>Source Sheet</th><th>Source Row</th><th>Source Value</th><th>Suppression Flag</th></tr></thead>
        <tbody>${debugRows}</tbody>
      </table>
      <h3>Classification Review Panel</h3>
      <table>
        <thead><tr><th>Source Row</th><th>Source Value</th><th>Suggested Bucket</th><th>Duplicate Risk</th><th>Parent/Child Group Flag</th></tr></thead>
        <tbody>${reviewRows}</tbody>
      </table>
    </section>
  `;
}

function renderReport(report) {
  const periods = report.periods;
  const summaryRows = Object.entries(report.borrowerSummary)
    .map(([key, value]) => `<tr><td>${key}</td><td>${typeof value === "number" ? fmtCurrency(value) : value}</td></tr>`)
    .join("");

  const plRows = buildStatementRows(report.years, "pl", [
    "INCOME", "Sales", "Other Income", "Total Income",
    "COST OF GOODS SOLD", "Opening Stock", "Purchases", "Direct Expenses", "Goods Available for Sale", "Less Closing Stock", "Cost of Goods Sold",
    "Gross Profit",
    "OPERATING EXPENSES", "Employee Cost", "Administrative Expenses", "Selling Expenses", "Other Operating Expenses", "Total Operating Expenses",
    "EBITDA", "FINANCIAL CHARGES", "Interest", "Depreciation", "Profit Before Tax", "Tax", "Profit After Tax",
  ]);

  const bsRows = buildStatementRows(report.years, "bs", [
    "CAPITAL & LIABILITIES", "Capital", "Reserves", "Net Worth",
    "Non Current Liabilities", "Term Loan", "Unsecured Loans", "Other Long Term Liabilities", "Total Non Current Liabilities",
    "Current Liabilities", "CC / Bank OD", "Trade Creditors", "Other Current Liabilities", "Total Current Liabilities", "Total Liabilities",
    "ASSETS", "Non Current Assets", "Fixed Assets", "Investments", "Other Non Current Assets", "Total Non Current Assets",
    "Current Assets", "Inventory", "Trade Receivables", "Cash & Bank", "Loans & Advances", "Other Current Assets", "Total Current Assets", "Total Assets",
  ]);

  const wcRows = [
    { label: "Total Current Assets", values: report.years.map((y) => y.workingCapital.totalCurrentAssets) },
    { label: "Total Current Liabilities", values: report.years.map((y) => y.workingCapital.totalCurrentLiabilities) },
    { label: "Net Working Capital", values: report.years.map((y) => y.workingCapital.netWorkingCapital) },
    { label: "Working Capital Gap", values: report.years.map((y) => y.workingCapital.workingCapitalGap) },
    { label: "Borrower Contribution", values: report.years.map((y) => y.workingCapital.borrowerContribution) },
    { label: "MPBF (Tandon Method II)", values: report.years.map((y) => y.workingCapital.mpbf) },
    { label: "Bank Finance Required", values: report.years.map((y) => y.workingCapital.bankFinanceRequired) },
    { label: "Current Ratio", values: report.years.map((y) => y.workingCapital.currentRatio) },
  ];

  const ccRows = [
    { label: "Projected Sales", values: report.years.map((y) => y.workingCapital.projectedSales) },
    { label: "MPBF", values: report.years.map((y) => y.workingCapital.requiredWorkingCapital) },
    { label: "Existing Limit", values: report.years.map((y) => y.workingCapital.existingLimit) },
    { label: "Shortfall", values: report.years.map((y) => y.workingCapital.shortfall) },
    { label: "Proposed CC Limit", values: report.years.map((y) => y.workingCapital.proposedLimit) },
  ];

  const ratioRows = report.years.map((y) => ({
    "Current Ratio": safeDivide(y.bs["Total Current Assets"], y.bs["Total Current Liabilities"]),
    "Quick Ratio": safeDivide(y.bs["Total Current Assets"] - y.bs.Inventory, y.bs["Total Current Liabilities"]),
    "GP Ratio": safeDivide(y.pl["Gross Profit"], y.pl.Sales) * 100,
    "NP Ratio": safeDivide(y.pl["Profit After Tax"], y.pl.Sales) * 100,
    "TOL / TNW": safeDivide(y.bs["Total Liabilities"] - y.bs["Net Worth"], y.bs["Net Worth"]),
    "Debtor Days": safeDivide(y.bs["Trade Receivables"], y.pl.Sales) * 365,
    "Creditor Days": safeDivide(y.bs["Trade Creditors"], y.pl.Purchases) * 365,
    "Inventory Days": safeDivide(y.bs.Inventory, y.pl["Cost of Goods Sold"]) * 365,
    "Interest Coverage": safeDivide(y.pl.EBIT, y.pl.Interest),
    DSCR: y.dscr,
  }));
  const ratioLabels = Object.keys(ratioRows[0]);
  const ratioTableRows = ratioLabels.map((label) => ({ label, values: ratioRows.map((row) => row[label]) }));

  const termLoanHtml = report.termLoan.applicable
    ? buildSectionHtml("Term Loan Schedule", ["Opening", "Principal Repayment", "Interest", "Closing"], report.termLoan.rows.map((r) => ({ label: r.year, values: [r.opening, r.installment, r.interest, r.closing] })))
    : "";
  const historicalDebugHtml = buildHistoricalDebugHtml(report);

  output.innerHTML = `
    <article class="cma-report">
      <header class="report-header"><h2>Banker Grade CMA Report</h2></header>
      <section class="report-section"><h3>Summary</h3><table><tbody>${summaryRows}</tbody></table></section>
      ${historicalDebugHtml}
      ${buildSectionHtml("Profit & Loss", periods, plRows, (value, label) => (label.includes("Ratio") ? fmtNumber(value) : fmtCurrency(value)))}
      ${buildSectionHtml("Balance Sheet", periods, bsRows)}
      ${buildSectionHtml("Working Capital", periods, wcRows, (value, label) => (label.includes("Ratio") ? fmtNumber(value) : fmtCurrency(value)))}
      ${buildSectionHtml("CC Limit Assessment (Tandon Method II)", periods, ccRows)}
      ${buildSectionHtml("Ratio Analysis", periods, ratioTableRows, (value, label) => {
    if (label.includes("Days")) return fmtNumber(value);
    if (label === "GP Ratio" || label === "NP Ratio") return fmtPct(value);
    return fmtNumber(value);
  })}
      ${termLoanHtml}
    </article>`;
}

function generateReport({ useCurrentMapping = false } = {}) {
  try {
    if (!workbookState.parsed) throw new Error(workbookState.parseError || "Upload workbook first.");

    const historicalLockMode = historicalLockModeInput?.checked !== false;
    const {
      mapped,
      fallbackWarnings,
      fallbackDefaults,
      missingMandatory,
      historicalSources,
      duplicateWarnings,
      parentChildWarnings,
      classificationReview,
    } = mappedFinancialsFromParse({ historicalLockMode });
    if (missingMandatory.length) throw new Error(`Please map core heads: ${missingMandatory.join(", ")}`);

    const layerData = buildLayerData(workbookState.parsed, historicalSources, mapped);

    workbookState.generated = {
      meta: {
        sourcePL: workbookState.parsed.plSheetName,
        sourceBS: workbookState.parsed.bsSheetName,
        generatedAt: new Date().toISOString(),
        generationMode: useCurrentMapping ? "current-mapping" : "standard",
        historicalLockMode,
        layers: layerData,
        warnings: {
          missingMandatory,
          fallbackAssignments: fallbackWarnings,
          fallbackDefaults,
          duplicateAssignments: duplicateWarnings,
          parentChildSuppression: parentChildWarnings,
        },
        historicalDebug: {
          values: mapped,
          sources: historicalSources,
        },
        classificationReview,
      },
      ...buildCmaReport(mapped, { historicalLockMode }),
    };

    // Insert display-only headers for professional grouping.
    workbookState.generated.years.forEach((year) => {
      year.pl.INCOME = null;
      year.pl["COST OF GOODS SOLD"] = null;
      year.pl["OPERATING EXPENSES"] = null;
      year.pl["FINANCIAL CHARGES"] = null;
      year.bs["CAPITAL & LIABILITIES"] = null;
      year.bs["Non Current Liabilities"] = null;
      year.bs["Current Liabilities"] = null;
      year.bs.ASSETS = null;
      year.bs["Non Current Assets"] = null;
      year.bs["Current Assets"] = null;
    });

    renderReport(workbookState.generated);
    if (fallbackWarnings.length || fallbackDefaults.length) {
      output.insertAdjacentHTML("afterbegin", `<p class="hint"><strong>Report generated with fallback assumptions for unmapped heads.</strong></p>`);
    }
    if (workbookState.generated.termLoan.warning) {
      output.insertAdjacentHTML("afterbegin", `<p class="bad"><strong>Term Loan Warning:</strong> ${workbookState.generated.termLoan.warning}</p>`);
    }
    renderMappingWarnings({
      fallbackWarnings,
      fallbackDefaults,
      missingMandatory,
      duplicateWarnings,
      parentChildWarnings,
    });
    downloadReportBtn.disabled = false;
    downloadExcelBtn.disabled = false;
    downloadJsonBtn.disabled = false;
    if (downloadTraceBtn) downloadTraceBtn.disabled = false;
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
function setFormula(ws, row, col, formula, format) { ws[toCell(row, col)] = { t: "n", f: formula, z: format || undefined }; }
function buildRowLookup(rowLabels, startRow = 3) {
  return rowLabels.reduce((acc, label, idx) => ({ ...acc, [label]: startRow + idx }), {});
}

function makeSheet(title, periods, rowsByYear, rowLabels) {
  const ws = XLSX.utils.aoa_to_sheet([
    [title],
    ["Particulars", ...periods],
    ...rowLabels.map((label) => [label, ...rowsByYear.map((row) => row[label] ?? "")]),
  ]);
  ws["!cols"] = Array.from({ length: periods.length + 1 }, (_, idx) => ({ wch: idx === 0 ? 36 : 14 }));
  ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: periods.length } }];
  return ws;
}

function applyCurrencyFormat(ws, startRow = 3) {
  const range = XLSX.utils.decode_range(ws["!ref"]);
  for (let r = startRow - 1; r <= range.e.r; r += 1) {
    for (let c = 1; c <= range.e.c; c += 1) {
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
    ["Other Income (% of Sales)", getConfigNumber("other-income-ratio", 1.5)],
    ["Employee Cost (% of Sales)", getConfigNumber("employee-ratio", 6)],
    ["Administrative Expenses (% of Sales)", getConfigNumber("admin-ratio", 4)],
    ["Selling Expenses (% of Sales)", getConfigNumber("selling-ratio", 3)],
    ["Depreciation (% of Fixed Assets)", getConfigNumber("depr-ratio", 10)],
    ["Tax Rate (%)", getConfigNumber("tax-rate", 25)],
    ["Receivable Days", getConfigNumber("receivable-days", 60)],
    ["Inventory Days", getConfigNumber("inventory-days", 45)],
    ["Creditor Days", getConfigNumber("payable-days", 30)],
    ["Existing CC Limit", getConfigNumber("cc-limit", 0)],
  ]);


  const sourceExtractRows = workbookState.generated.meta?.layers?.sourceExtract || [];
  const sourceExtractWs = XLSX.utils.aoa_to_sheet([
    ["Source_Extract"],
    ["Source Head", "Group Type", "FY1 Amount", "Source Sheet", "Row No"],
    ...sourceExtractRows.map((row) => [row.sourceHead, row.groupType, row.amount, row.sheet, row.rowNumber || ""]),
  ]);

  const cmaInputMapRows = workbookState.generated.meta?.layers?.cmaInputMap || [];
  const cmaInputMapWs = XLSX.utils.aoa_to_sheet([
    ["CMA_Input_Map"],
    ["Source Head", "Group Type", "Mapped CMA Head", "FY1 Amount", "Include/Exclude Flag"],
    ...cmaInputMapRows.map((row) => [row.sourceHead, row.groupType, row.mappedCmaHead, row.fy1Amount, row.includeFlag]),
  ]);

  const fy1 = workbookState.generated.years?.[0] || { pl: {}, bs: {}, workingCapital: {} };
  const validationRows = [
    ["Validation"],
    ["CMA Head", "FY1 CMA Value", "FY1 Source Included Value", "Status"],
    ["Net Sales", fy1.pl.Sales || 0, mappedValueByHead(cmaInputMapRows, STANDARD_CMA_HEADS.netSales), ""],
    ["Opening Stock", fy1.pl["Opening Stock"] || 0, mappedValueByHead(cmaInputMapRows, STANDARD_CMA_HEADS.openingStock), ""],
    ["Purchases", fy1.pl.Purchases || 0, mappedValueByHead(cmaInputMapRows, STANDARD_CMA_HEADS.purchases), ""],
    ["Closing Stock", fy1.pl["Less Closing Stock"] || 0, mappedValueByHead(cmaInputMapRows, STANDARD_CMA_HEADS.closingStock), ""],
    ["Other Income", fy1.pl["Other Income"] || 0, mappedValueByHead(cmaInputMapRows, STANDARD_CMA_HEADS.otherIncome), ""],
    ["Interest", fy1.pl.Interest || 0, mappedValueByHead(cmaInputMapRows, STANDARD_CMA_HEADS.interest), ""],
    ["Debtors", fy1.bs["Trade Receivables"] || 0, mappedValueByHead(cmaInputMapRows, STANDARD_CMA_HEADS.debtors), ""],
    ["Inventory", fy1.bs.Inventory || 0, mappedValueByHead(cmaInputMapRows, STANDARD_CMA_HEADS.inventory), ""],
    ["Cash & Bank", fy1.bs["Cash & Bank"] || 0, mappedValueByHead(cmaInputMapRows, STANDARD_CMA_HEADS.cashBank), ""],
    ["Other Current Assets", fy1.bs["Other Current Assets"] || 0, mappedValueByHead(cmaInputMapRows, STANDARD_CMA_HEADS.otherCurrentAssets), ""],
    ["Creditors", fy1.bs["Trade Creditors"] || 0, mappedValueByHead(cmaInputMapRows, STANDARD_CMA_HEADS.creditors), ""],
    ["Other Current Liabilities", fy1.bs["Other Current Liabilities"] || 0, mappedValueByHead(cmaInputMapRows, STANDARD_CMA_HEADS.otherCurrentLiabilities), ""],
    ["Short Term Borrowings", fy1.bs["CC / Bank OD"] || 0, mappedValueByHead(cmaInputMapRows, STANDARD_CMA_HEADS.shortTermBorrowings), ""],
    ["Capital", fy1.bs.Capital || 0, mappedValueByHead(cmaInputMapRows, STANDARD_CMA_HEADS.capital), ""],
    ["Fixed Assets", fy1.bs["Fixed Assets"] || 0, mappedValueByHead(cmaInputMapRows, STANDARD_CMA_HEADS.fixedAssets), ""],
    ["Non-current / Other Assets", fy1.bs["Other Non Current Assets"] || 0, mappedValueByHead(cmaInputMapRows, STANDARD_CMA_HEADS.nonCurrentOtherAssets), ""],
  ];
  const validationWs = XLSX.utils.aoa_to_sheet(validationRows);

  const plLabels = [
    "Sales",
    "Other Income",
    "Total Income",
    "Opening Stock",
    "Purchases",
    "Direct Expenses",
    "Closing Stock",
    "COGS",
    "Gross Profit",
    "Operating Expenses",
    "EBITDA",
    "Depreciation",
    "EBIT",
    "Interest",
    "PBT",
    "Tax",
    "PAT",
  ];

  const bsLabels = [
    "Capital / Net Worth",
    "Term Loan",
    "Unsecured Loans",
    "Other Long Term Liabilities",
    "CC / Bank OD",
    "Trade Creditors",
    "Other Current Liabilities",
    "Total Liabilities",
    "Fixed Assets",
    "Investments",
    "Other Non Current Assets",
    "Inventory",
    "Trade Receivables",
    "Cash / Bank",
    "Advances / Deposits / Other Current Assets",
    "Total Assets",
  ];

  const plRows = years.map((y) => ({
    Sales: y.pl.Sales,
    "Other Income": y.pl["Other Income"],
    "Total Income": "",
    "Opening Stock": y.pl["Opening Stock"],
    Purchases: y.pl.Purchases,
    "Direct Expenses": y.pl["Direct Expenses"],
    "Closing Stock": y.pl["Less Closing Stock"],
    COGS: "",
    "Gross Profit": "",
    "Operating Expenses": y.pl["Total Operating Expenses"],
    EBITDA: "",
    Depreciation: y.pl.Depreciation,
    EBIT: "",
    Interest: y.pl.Interest,
    PBT: "",
    Tax: y.pl.Tax,
    PAT: "",
  }));

  const bsRows = years.map((y) => ({
    "Capital / Net Worth": y.bs["Net Worth"],
    "Term Loan": y.bs["Term Loan"],
    "Unsecured Loans": y.bs["Unsecured Loans"],
    "Other Long Term Liabilities": y.bs["Other Long Term Liabilities"],
    "CC / Bank OD": y.bs["CC / Bank OD"],
    "Trade Creditors": y.bs["Trade Creditors"],
    "Other Current Liabilities": y.bs["Other Current Liabilities"],
    "Total Liabilities": "",
    "Fixed Assets": y.bs["Fixed Assets"],
    Investments: y.bs.Investments,
    "Other Non Current Assets": y.bs["Other Non Current Assets"],
    Inventory: y.bs.Inventory,
    "Trade Receivables": y.bs["Trade Receivables"],
    "Cash / Bank": y.bs["Cash & Bank"],
    "Advances / Deposits / Other Current Assets": (y.bs["Loans & Advances"] || 0) + (y.bs["Other Current Assets"] || 0),
    "Total Assets": "",
  }));

  const plWs = makeSheet("Profit & Loss", periods, plRows, plLabels);
  const bsWs = makeSheet("Balance Sheet", periods, bsRows, bsLabels);
  const plRow = buildRowLookup(plLabels);
  const bsRow = buildRowLookup(bsLabels);

  periods.forEach((_, idx) => {
    const col = idx + 2;
    setFormula(plWs, plRow["Total Income"], col, `${toCell(plRow.Sales, col)}+${toCell(plRow["Other Income"], col)}`, "₹#,##0");
    setFormula(plWs, plRow.COGS, col, `${toCell(plRow["Opening Stock"], col)}+${toCell(plRow.Purchases, col)}+${toCell(plRow["Direct Expenses"], col)}-${toCell(plRow["Closing Stock"], col)}`, "₹#,##0");
    setFormula(plWs, plRow["Gross Profit"], col, `${toCell(plRow.Sales, col)}-${toCell(plRow.COGS, col)}`, "₹#,##0");
    setFormula(plWs, plRow.EBITDA, col, `${toCell(plRow["Gross Profit"], col)}+${toCell(plRow["Other Income"], col)}-${toCell(plRow["Operating Expenses"], col)}`, "₹#,##0");
    setFormula(plWs, plRow.EBIT, col, `${toCell(plRow.EBITDA, col)}-${toCell(plRow.Depreciation, col)}`, "₹#,##0");
    setFormula(plWs, plRow.PBT, col, `${toCell(plRow.EBIT, col)}-${toCell(plRow.Interest, col)}`, "₹#,##0");
    setFormula(plWs, plRow.PAT, col, `${toCell(plRow.PBT, col)}-${toCell(plRow.Tax, col)}`, "₹#,##0");

    setFormula(bsWs, bsRow["Total Liabilities"], col, `${toCell(bsRow["Capital / Net Worth"], col)}+${toCell(bsRow["Term Loan"], col)}+${toCell(bsRow["Unsecured Loans"], col)}+${toCell(bsRow["Other Long Term Liabilities"], col)}+${toCell(bsRow["CC / Bank OD"], col)}+${toCell(bsRow["Trade Creditors"], col)}+${toCell(bsRow["Other Current Liabilities"], col)}`, "₹#,##0");
    setFormula(bsWs, bsRow["Total Assets"], col, `${toCell(bsRow["Fixed Assets"], col)}+${toCell(bsRow.Investments, col)}+${toCell(bsRow["Other Non Current Assets"], col)}+${toCell(bsRow.Inventory, col)}+${toCell(bsRow["Trade Receivables"], col)}+${toCell(bsRow["Cash / Bank"], col)}+${toCell(bsRow["Advances / Deposits / Other Current Assets"], col)}`, "₹#,##0");
  });

  applyCurrencyFormat(plWs);
  applyCurrencyFormat(bsWs);

  const caWs = XLSX.utils.aoa_to_sheet([
    ["Current Assets"],
    ["Particulars", ...periods],
    ["Inventory", ...years.map((y) => y.bs.Inventory)],
    ["Trade Receivables", ...years.map((y) => y.bs["Trade Receivables"])],
    ["Cash / Bank", ...years.map((y) => y.bs["Cash & Bank"])],
    ["Loans & Advances", ...years.map((y) => y.bs["Loans & Advances"])],
    ["Other Current Assets", ...years.map((y) => y.bs["Other Current Assets"])],
    ["Total Current Assets", ...periods.map(() => "")],
  ]);
  const caLabels = ["Inventory", "Trade Receivables", "Cash / Bank", "Loans & Advances", "Other Current Assets", "Total Current Assets"];
  const caRow = buildRowLookup(caLabels);
  periods.forEach((_, idx) => {
    const c = idx + 2;
    setFormula(caWs, caRow["Total Current Assets"], c, `${toCell(caRow.Inventory, c)}+${toCell(caRow["Trade Receivables"], c)}+${toCell(caRow["Cash / Bank"], c)}+${toCell(caRow["Loans & Advances"], c)}+${toCell(caRow["Other Current Assets"], c)}`, "₹#,##0");
  });

  const clWs = XLSX.utils.aoa_to_sheet([
    ["Current Liabilities"],
    ["Particulars", ...periods],
    ["CC / Bank OD", ...years.map((y) => y.bs["CC / Bank OD"])],
    ["Trade Creditors", ...years.map((y) => y.bs["Trade Creditors"])],
    ["Other Current Liabilities", ...years.map((y) => y.bs["Other Current Liabilities"])],
    ["Total Current Liabilities", ...periods.map(() => "")],
  ]);
  const clLabels = ["CC / Bank OD", "Trade Creditors", "Other Current Liabilities", "Total Current Liabilities"];
  const clRow = buildRowLookup(clLabels);
  periods.forEach((_, idx) => {
    const c = idx + 2;
    setFormula(clWs, clRow["Total Current Liabilities"], c, `${toCell(clRow["CC / Bank OD"], c)}+${toCell(clRow["Trade Creditors"], c)}+${toCell(clRow["Other Current Liabilities"], c)}`, "₹#,##0");
  });

  const wcWs = XLSX.utils.aoa_to_sheet([
    ["Working Capital"],
    ["Particulars", ...periods],
    ["Total Current Assets", ...periods.map(() => "")],
    ["Total Current Liabilities", ...periods.map(() => "")],
    ["Working Capital", ...periods.map(() => "")],
  ]);
  const wcLabels = ["Total Current Assets", "Total Current Liabilities", "Working Capital"];
  const wcRow = buildRowLookup(wcLabels);
  periods.forEach((_, idx) => {
    const c = idx + 2;
    setFormula(wcWs, wcRow["Total Current Assets"], c, `'Current Assets'!${toCell(caRow["Total Current Assets"], c)}`, "₹#,##0");
    setFormula(wcWs, wcRow["Total Current Liabilities"], c, `'Current Liabilities'!${toCell(clRow["Total Current Liabilities"], c)}`, "₹#,##0");
    setFormula(wcWs, wcRow["Working Capital"], c, `${toCell(wcRow["Total Current Assets"], c)}-${toCell(wcRow["Total Current Liabilities"], c)}`, "₹#,##0");
  });

  const ratioLabels = ["Current Ratio", "Quick Ratio", "GP Ratio", "NP Ratio", "Debtor Days", "Creditor Days", "Inventory Days", "Interest Coverage"];
  const ratioWs = XLSX.utils.aoa_to_sheet([["Ratios"], ["Particulars", ...periods], ...ratioLabels.map((l) => [l, ...periods.map(() => "")])]);
  const ratioRow = buildRowLookup(ratioLabels);

  periods.forEach((_, idx) => {
    const c = idx + 2;
    setFormula(ratioWs, ratioRow["Current Ratio"], c, `'Current Assets'!${toCell(caRow["Total Current Assets"], c)}/'Current Liabilities'!${toCell(clRow["Total Current Liabilities"], c)}`, "0.00");
    setFormula(ratioWs, ratioRow["Quick Ratio"], c, `('Current Assets'!${toCell(caRow["Total Current Assets"], c)}-'Current Assets'!${toCell(caRow.Inventory, c)})/'Current Liabilities'!${toCell(clRow["Total Current Liabilities"], c)}`, "0.00");
    setFormula(ratioWs, ratioRow["GP Ratio"], c, `'Profit & Loss'!${toCell(plRow["Gross Profit"], c)}/'Profit & Loss'!${toCell(plRow.Sales, c)}*100`, "0.00");
    setFormula(ratioWs, ratioRow["NP Ratio"], c, `'Profit & Loss'!${toCell(plRow.PAT, c)}/'Profit & Loss'!${toCell(plRow.Sales, c)}*100`, "0.00");
    setFormula(ratioWs, ratioRow["Debtor Days"], c, `'Current Assets'!${toCell(caRow["Trade Receivables"], c)}/'Profit & Loss'!${toCell(plRow.Sales, c)}*365`, "0.00");
    setFormula(ratioWs, ratioRow["Creditor Days"], c, `'Current Liabilities'!${toCell(clRow["Trade Creditors"], c)}/'Profit & Loss'!${toCell(plRow.Purchases, c)}*365`, "0.00");
    setFormula(ratioWs, ratioRow["Inventory Days"], c, `'Current Assets'!${toCell(caRow.Inventory, c)}/'Profit & Loss'!${toCell(plRow.COGS, c)}*365`, "0.00");
    setFormula(ratioWs, ratioRow["Interest Coverage"], c, `'Profit & Loss'!${toCell(plRow.EBIT, c)}/'Profit & Loss'!${toCell(plRow.Interest, c)}`, "0.00");
  });

  const tlWs = workbookState.generated.termLoan.applicable
    ? XLSX.utils.aoa_to_sheet([
      ["Term Loan"],
      ["Year", "Opening", "Principal Repayment", "Interest", "Closing"],
      ...workbookState.generated.termLoan.rows.map((r) => [r.year, r.opening, r.installment, r.interest, r.closing]),
    ])
    : XLSX.utils.aoa_to_sheet([["Term Loan"], ["Not Applicable"]]);



  const validationRange = XLSX.utils.decode_range(validationWs["!ref"]);
  for (let r = 2; r <= validationRange.e.r; r += 1) {
    const rowNo = r + 1;
    validationWs[`D${rowNo}`] = { t: "s", f: `IF(ABS(B${rowNo}-C${rowNo})<1,"OK","Mismatch")` };
  }
  const dscrWs = XLSX.utils.aoa_to_sheet([
    ["DSCR"],
    ["Year", "PAT", "Depreciation", "Interest", "Installment", "DSCR", "Status"],
    ...years.map((y, idx) => {
      const tlInterest = workbookState.generated.termLoan.applicable ? workbookState.generated.termLoan.rows[idx].interest : 0;
      const installment = workbookState.generated.termLoan.applicable ? workbookState.generated.termLoan.rows[idx].installment : 0;
      return [
        y.year,
        y.pl["Profit After Tax"],
        y.pl.Depreciation,
        tlInterest,
        installment,
        y.dscr || "",
        y.dscr === null ? "N/A" : (y.dscr >= 1.25 ? "Acceptable" : "Warning"),
      ];
    }),
    ["Average DSCR", "", "", "", "", workbookState.generated.avgDscr || "", workbookState.generated.avgDscr === null ? "N/A" : (workbookState.generated.avgDscr >= 1.25 ? "Acceptable" : "Warning")],
  ]);

  applyCurrencyFormat(caWs);
  applyCurrencyFormat(clWs);
  applyCurrencyFormat(wcWs);
  applyCurrencyFormat(tlWs);
  applyCurrencyFormat(dscrWs);

  XLSX.utils.book_append_sheet(wb, assumptions, "Assumptions");
  XLSX.utils.book_append_sheet(wb, sourceExtractWs, "Source_Extract");
  XLSX.utils.book_append_sheet(wb, cmaInputMapWs, "CMA_Input_Map");
  XLSX.utils.book_append_sheet(wb, validationWs, "Validation");
  XLSX.utils.book_append_sheet(wb, plWs, "Profit & Loss");
  XLSX.utils.book_append_sheet(wb, bsWs, "Balance Sheet");
  XLSX.utils.book_append_sheet(wb, caWs, "Current Assets");
  XLSX.utils.book_append_sheet(wb, clWs, "Current Liabilities");
  XLSX.utils.book_append_sheet(wb, wcWs, "Working Capital");
  XLSX.utils.book_append_sheet(wb, ratioWs, "Ratios");
  XLSX.utils.book_append_sheet(wb, tlWs, "Term Loan");
  XLSX.utils.book_append_sheet(wb, dscrWs, "DSCR");

  XLSX.writeFile(wb, "banker-grade-cma-report.xlsx");
}

function downloadHistoricalTrace() {
  if (!workbookState.generated?.meta?.historicalDebug) return;
  const fieldLabel = Object.fromEntries(MAP_FIELDS.map((f) => [f.key, f.label]));
  const rows = [["Final CMA Line Item", "Source Sheet", "Source Row", "Source Value"]];
  Object.entries(workbookState.generated.meta.historicalDebug.sources || {}).forEach(([key, sources]) => {
    if (!sources.length) {
      rows.push([fieldLabel[key] || key, "-", "-", workbookState.generated.meta.historicalDebug.values?.[key] || 0]);
      return;
    }
    sources.forEach((src) => {
      rows.push([fieldLabel[key] || key, src.sheet || src.section.toUpperCase(), `${src.head} (Row ${src.rowNumber || "-"})`, src.amount]);
    });
  });
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Historical Trace");
  XLSX.writeFile(wb, "historical-trace.xlsx");
}

function downloadJson() {
  if (!workbookState.generated) return;
  saveFile(JSON.stringify(workbookState.generated, null, 2), "cma-data.json", "application/json");
}
