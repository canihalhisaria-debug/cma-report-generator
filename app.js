const workbookState = {
  workbook: null,
  selectedSheet: null,
  sourceRows: [],
  generated: null,
};

const fileInput = document.getElementById("excel-file");
const sheetStatus = document.getElementById("sheet-status");
const sheetSelector = document.getElementById("sheet-selector");
const generateBtn = document.getElementById("generate-btn");
const output = document.getElementById("output");
const downloadReportBtn = document.getElementById("download-report");
const downloadJsonBtn = document.getElementById("download-json");

const configInputs = [
  "historical-count",
  "projection-years",
  "sales-growth",
  "gross-margin",
  "receivable-days",
  "inventory-days",
  "payable-days",
  "cc-limit",
  "cc-utilization",
  "tl-outstanding",
  "tl-rate",
].reduce((acc, id) => {
  acc[id] = document.getElementById(id);
  return acc;
}, {});

fileInput.addEventListener("change", handleWorkbookUpload);
sheetSelector.addEventListener("change", handleSheetSelection);
generateBtn.addEventListener("click", generateReport);
downloadReportBtn.addEventListener("click", downloadReport);
downloadJsonBtn.addEventListener("click", downloadJson);

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
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
  workbookState.sourceRows = rows;

  generateBtn.disabled = false;
  output.innerHTML = `<p class="good">Sheet <strong>${selectedSheet}</strong> ready with ${rows.length} records.</p>`;
}

function getNumeric(value, fallback = 0) {
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

function guessYearlySales(rows, fallback = 10000000) {
  if (!rows.length) return fallback;

  const candidateKeys = ["sales", "turnover", "revenue", "net_sales", "gross_sales"];
  const sample = rows[0];
  const key = Object.keys(sample).find((k) =>
    candidateKeys.some((needle) => k.toLowerCase().includes(needle)),
  );

  if (!key) return fallback;

  const values = rows.map((r) => getNumeric(r[key])).filter((n) => n > 0);
  if (!values.length) return fallback;

  const sum = values.reduce((acc, n) => acc + n, 0);
  return sum / values.length;
}

function fmtCurrency(amount) {
  return new Intl.NumberFormat("en-IN", {
    style: "currency",
    currency: "INR",
    maximumFractionDigits: 0,
  }).format(amount);
}

function buildPeriods() {
  const historicalCount = getNumeric(configInputs["historical-count"].value, 4);
  const projectionYears = getNumeric(configInputs["projection-years"].value, 5);
  const currentYear = new Date().getFullYear();

  const historical = Array.from({ length: historicalCount }, (_, i) => `FY ${currentYear - historicalCount + i}-${String(currentYear - historicalCount + i + 1).slice(-2)}`);
  const provisional = `FY ${currentYear}-${String(currentYear + 1).slice(-2)} (Prov.)`;
  const projections = Array.from({ length: projectionYears }, (_, i) => `FY ${currentYear + 1 + i}-${String(currentYear + 2 + i).slice(-2)}`);

  return { historical, provisional, projections };
}

function generateProjectionRows(baseSales) {
  const growth = getNumeric(configInputs["sales-growth"].value, 10) / 100;
  const grossMargin = getNumeric(configInputs["gross-margin"].value, 28) / 100;

  const receivableDays = getNumeric(configInputs["receivable-days"].value, 60);
  const inventoryDays = getNumeric(configInputs["inventory-days"].value, 45);
  const payableDays = getNumeric(configInputs["payable-days"].value, 30);

  const ccUtilization = getNumeric(configInputs["cc-utilization"].value, 85) / 100;
  const ccLimit = getNumeric(configInputs["cc-limit"].value, 5000000);

  const outstandingTL = getNumeric(configInputs["tl-outstanding"].value, 12000000);
  const tlRate = getNumeric(configInputs["tl-rate"].value, 11) / 100;

  const { historical, provisional, projections } = buildPeriods();
  const allPeriods = [...historical, provisional, ...projections];

  const rows = allPeriods.map((period, index) => {
    const yearOffset = index - historical.length;
    const sales = baseSales * Math.pow(1 + growth, Math.max(0, yearOffset));
    const grossProfit = sales * grossMargin;
    const cogs = sales - grossProfit;
    const receivables = (sales * receivableDays) / 365;
    const inventory = (cogs * inventoryDays) / 365;
    const payables = (cogs * payableDays) / 365;
    const workingCapitalGap = receivables + inventory - payables;

    const ccRequired = Math.max(workingCapitalGap * ccUtilization, ccLimit * 0.65);

    const repaymentYears = projections.length || 1;
    const principalRepayment = index <= historical.length ? 0 : outstandingTL / repaymentYears;
    const openingOutstanding = Math.max(outstandingTL - principalRepayment * Math.max(0, index - historical.length - 1), 0);
    const interest = openingOutstanding * tlRate;
    const closingOutstanding = Math.max(openingOutstanding - principalRepayment, 0);

    return {
      period,
      sales,
      grossProfit,
      receivables,
      inventory,
      payables,
      workingCapitalGap,
      ccRequired,
      openingOutstanding,
      principalRepayment,
      interest,
      closingOutstanding,
    };
  });

  return { periods: allPeriods, rows, projectionsCount: projections.length };
}

function generateReport() {
  const baseSales = guessYearlySales(workbookState.sourceRows);
  const generated = generateProjectionRows(baseSales);

  workbookState.generated = {
    meta: {
      sheet: workbookState.selectedSheet,
      sourceRecords: workbookState.sourceRows.length,
      generatedAt: new Date().toISOString(),
    },
    ...generated,
  };

  renderReport(workbookState.generated);
  downloadReportBtn.disabled = false;
  downloadJsonBtn.disabled = false;
}

function tableHtml(title, headers, rowMapper, data) {
  const head = headers.map((h) => `<th>${h}</th>`).join("");
  const body = data
    .map((item) => {
      const cells = rowMapper(item).map((c) => `<td>${c}</td>`).join("");
      return `<tr>${cells}</tr>`;
    })
    .join("");

  return `
    <div class="table-wrap">
      <h3>${title}</h3>
      <table>
        <thead><tr>${head}</tr></thead>
        <tbody>${body}</tbody>
      </table>
    </div>
  `;
}

function renderReport(report) {
  const wcTable = tableHtml(
    "Statement of Working Capital Assessment",
    ["Period", "Sales", "Receivables", "Inventory", "Payables", "WCG", "CC Required"],
    (r) => [
      r.period,
      fmtCurrency(r.sales),
      fmtCurrency(r.receivables),
      fmtCurrency(r.inventory),
      fmtCurrency(r.payables),
      fmtCurrency(r.workingCapitalGap),
      fmtCurrency(r.ccRequired),
    ],
    report.rows,
  );

  const tlTable = tableHtml(
    "Term Loan Repayment & Interest Projection",
    ["Period", "Opening", "Principal", "Interest", "Closing"],
    (r) => [
      r.period,
      fmtCurrency(r.openingOutstanding),
      fmtCurrency(r.principalRepayment),
      fmtCurrency(r.interest),
      fmtCurrency(r.closingOutstanding),
    ],
    report.rows,
  );

  output.innerHTML = `
    <p class="good">
      Generated banker-ready CMA with ${report.rows.length} periods
      (${report.meta.sourceRecords} source rows, ${report.projectionsCount} repayment years).
    </p>
    ${wcTable}
    ${tlTable}
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

function downloadReport() {
  if (!workbookState.generated) return;
  const title = `CMA Report - ${workbookState.generated.meta.sheet}`;
  const html = `<!doctype html><html><head><meta charset="UTF-8"><title>${title}</title></head><body>${output.innerHTML}</body></html>`;
  saveFile(html, "cma-report.html", "text/html");
}

function downloadJson() {
  if (!workbookState.generated) return;
  saveFile(JSON.stringify(workbookState.generated, null, 2), "cma-data.json", "application/json");
}
