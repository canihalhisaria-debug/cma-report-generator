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
const downloadExcelBtn = document.getElementById("download-excel");
const downloadJsonBtn = document.getElementById("download-json");

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

function formatFY(startYear, suffix = "") {
  return `FY ${startYear}-${String(startYear + 1).slice(-2)}${suffix}`;
}

function fmtCurrency(amount) {
  return new Intl.NumberFormat("en-IN", {
    style: "currency",
    currency: "INR",
    maximumFractionDigits: 0,
  }).format(amount || 0);
}

function fmtPct(value) {
  return `${(value || 0).toFixed(2)}%`;
}

function safeDivide(a, b) {
  if (!b) return 0;
  return a / b;
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
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null });
  workbookState.sourceRows = rows;

  generateBtn.disabled = false;
  output.innerHTML = `<p class="good">Sheet <strong>${selectedSheet}</strong> ready with ${rows.length} records.</p>`;
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

  const historical = Array.from({ length: historicalCount }, (_, idx) => {
    const start = currentYear - historicalCount + idx;
    return formatFY(start);
  });

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

function buildCmaReport(baseSales) {
  const periods = buildPeriods();
  const salesGrowth = getConfigNumber("sales-growth", 10);
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

  const ccLimit = getConfigNumber("cc-limit", 5000000);
  const ccUtilization = getConfigNumber("cc-utilization", 85);
  const tlOutstanding = getConfigNumber("tl-outstanding", 12000000);
  const tlRate = getConfigNumber("tl-rate", 11) / 100;
  const moratoriumYears = Math.max(0, getConfigNumber("moratorium-years", 0));

  const allYears = [];
  const principalYears = Math.max(1, periods.tlYears - moratoriumYears);
  const annualPrincipalInstallment = tlOutstanding / principalYears;

  periods.allColumns.forEach((periodLabel, idx) => {
    const relativeYear = idx - periods.historical.length;
    const factor = relativeYear <= 0 ? Math.pow(1 + salesGrowth / 100, relativeYear) : Math.pow(1 + salesGrowth / 100, relativeYear);
    const sales = baseSales * factor;
    const otherIncome = ratio(sales, otherIncomeRatio);
    const totalIncome = sales + otherIncome;

    const purchases = sales * (1 - grossMargin / 100) * 0.88;
    const directExpenses = sales * (1 - grossMargin / 100) * 0.12;
    const grossProfit = totalIncome - purchases - directExpenses;

    const employeeCost = ratio(sales, employeeRatio);
    const adminExpenses = ratio(sales, adminRatio);
    const sellingExpenses = ratio(sales, sellingRatio);
    const ebitda = grossProfit - employeeCost - adminExpenses - sellingExpenses;

    const projectionIdx = idx - periods.historical.length - 1;
    const tlYearNo = projectionIdx + 1;
    const shouldApplyTL = tlYearNo > 0 && tlYearNo <= periods.tlYears;
    const openingTL = shouldApplyTL
      ? Math.max(
          tlOutstanding - Math.max(0, tlYearNo - moratoriumYears - 1) * annualPrincipalInstallment,
          0,
        )
      : 0;
    const installment = shouldApplyTL && tlYearNo > moratoriumYears ? Math.min(annualPrincipalInstallment, openingTL) : 0;
    const interest = openingTL * tlRate;
    const closingTL = Math.max(openingTL - installment, 0);

    const fixedAssets = sales * 0.22;
    const depreciation = ratio(fixedAssets, deprRatio);
    const pbt = ebitda - interest - depreciation;
    const tax = Math.max(pbt, 0) * (taxRate / 100);
    const pat = pbt - tax;

    const netWorth = sales * 0.18 + pat * 2;
    const reserves = Math.max(netWorth * 0.35, 0);
    const capital = netWorth - reserves;

    const unsecuredLoans = sales * 0.04;
    const receivables = (sales * receivableDays) / 365;
    const inventory = ((purchases + directExpenses) * inventoryDays) / 365;
    const creditors = ((purchases + directExpenses) * payableDays) / 365;
    const otherCurrentAssets = sales * 0.03;
    const cashBank = Math.max(sales * 0.01 + pat * 0.05, 0);
    const loansAdvances = sales * 0.02;
    const investments = sales * 0.015;
    const otherCurrentLiabilities = sales * 0.02;

    const totalCurrentAssets = receivables + inventory + otherCurrentAssets + cashBank + loansAdvances;
    const totalCurrentLiabilities = creditors + otherCurrentLiabilities;
    const wcGap = totalCurrentAssets - totalCurrentLiabilities;
    const marginContribution = wcGap * (marginRatio / 100);
    const ccRequired = Math.max(wcGap - marginContribution, ccLimit * (ccUtilization / 100));

    const totalLiabilities = capital + reserves + openingTL + unsecuredLoans + creditors + otherCurrentLiabilities;
    const totalAssets = fixedAssets + investments + inventory + receivables + cashBank + loansAdvances + otherCurrentAssets;

    const currentRatio = safeDivide(totalCurrentAssets, totalCurrentLiabilities);
    const debtEquity = safeDivide(openingTL + unsecuredLoans, netWorth);
    const tolTnw = safeDivide(totalLiabilities, netWorth);
    const gpRatio = safeDivide(grossProfit, sales) * 100;
    const npRatio = safeDivide(pat, sales) * 100;
    const ebitdaMargin = safeDivide(ebitda, sales) * 100;
    const interestCoverage = safeDivide(ebitda, interest);
    const debtorDays = safeDivide(receivables, sales) * 365;
    const creditorDays = safeDivide(creditors, purchases + directExpenses) * 365;
    const inventoryHoldingDays = safeDivide(inventory, purchases + directExpenses) * 365;
    const grossCashAccrual = pat + depreciation + interest;
    const totalDebtObligation = installment + interest;
    const dscr = safeDivide(grossCashAccrual, totalDebtObligation);

    allYears.push({
      period: periodLabel,
      meta: {
        isHistorical: idx < periods.historical.length,
        isProvisional: idx === periods.historical.length,
        projectionNo: projectionIdx + 1,
        isCCYear: projectionIdx >= 0 && projectionIdx < periods.ccYears,
        isTLYear: projectionIdx >= 0 && projectionIdx < periods.tlYears,
      },
      pl: {
        Sales: sales,
        "Other income": otherIncome,
        "Total income": totalIncome,
        "Purchases / raw material": purchases,
        "Direct expenses": directExpenses,
        "Gross profit": grossProfit,
        "Employee cost": employeeCost,
        "Administrative expenses": adminExpenses,
        "Selling and distribution expenses": sellingExpenses,
        "EBITDA / Operating profit": ebitda,
        Interest: interest,
        Depreciation: depreciation,
        "Profit before tax": pbt,
        Tax: tax,
        "Profit after tax": pat,
      },
      bs: {
        "Capital / share capital / partners capital": capital,
        Reserves: reserves,
        "Net worth": netWorth,
        "Term loan": openingTL,
        "Unsecured loans": unsecuredLoans,
        "Trade creditors": creditors,
        "Other current liabilities": otherCurrentLiabilities,
        "Total liabilities": totalLiabilities,
        "Fixed assets": fixedAssets,
        Investments: investments,
        Inventory: inventory,
        "Trade receivables": receivables,
        "Cash and bank": cashBank,
        "Loans and advances": loansAdvances,
        "Other current assets": otherCurrentAssets,
        "Total assets": totalAssets,
      },
      wc: {
        Sales: sales,
        Receivables: receivables,
        Inventory: inventory,
        "Other current assets": otherCurrentAssets,
        "Total current assets": totalCurrentAssets,
        Creditors: creditors,
        "Other current liabilities": otherCurrentLiabilities,
        "Working capital gap": wcGap,
        "Margin contribution": marginContribution,
        "MPBF / CC required": ccRequired,
      },
      ratios: {
        "Current Ratio": currentRatio,
        "Debt Equity Ratio": debtEquity,
        "TOL/TNW": tolTnw,
        "GP Ratio": gpRatio,
        "NP Ratio": npRatio,
        "EBITDA Margin": ebitdaMargin,
        "Interest Coverage Ratio": interestCoverage,
        "Debtor Days": debtorDays,
        "Creditor Days": creditorDays,
        "Inventory Holding Days": inventoryHoldingDays,
        DSCR: dscr,
      },
      termLoan: {
        "Opening balance": openingTL,
        Installment: installment,
        Interest: interest,
        "Total debt servicing": totalDebtObligation,
        "Closing balance": closingTL,
      },
      dscr: {
        PAT: pat,
        Depreciation: depreciation,
        "Interest on term loan": interest,
        "Gross cash accrual": grossCashAccrual,
        Installment: installment,
        "Total debt obligation": totalDebtObligation,
        DSCR: dscr,
      },
    });
  });

  const borrowerSummary = {
    "Borrower name": getConfigText("borrower-name"),
    Constitution: getConfigText("constitution"),
    PAN: getConfigText("pan"),
    GSTIN: getConfigText("gstin"),
    "Bank / branch": `${getConfigText("bank-name")} / ${getConfigText("branch-name")}`,
    "Existing limit": getConfigNumber("existing-limit", 0),
    "Proposed limit": getConfigNumber("proposed-limit", 0),
    "Facility type": getConfigText("facility-type"),
    "CMA prepared date": getConfigText("prepared-date"),
  };

  return {
    periods,
    borrowerSummary,
    allYears,
    assumptions: {
      "Historical years": periods.historical.length,
      "Current year provisional": periods.provisionalYear,
      "Projected years for Cash Credit": periods.ccYears,
      "Projected years for Term Loan": periods.tlYears,
      "Moratorium (years)": moratoriumYears,
      "Sales growth %": salesGrowth,
      "Gross margin %": grossMargin,
      "Tax rate %": taxRate,
      "Receivable days": receivableDays,
      "Inventory days": inventoryDays,
      "Creditor days": payableDays,
      "Margin contribution %": marginRatio,
      "Term loan rate %": tlRate * 100,
    },
  };
}

function buildSectionHtml(title, headerPeriods, rows, formatter = fmtCurrency) {
  const head = `<tr><th>Particulars</th>${headerPeriods.map((p) => `<th>${p}</th>`).join("")}</tr>`;
  const body = rows
    .map(({ label, values }) => {
      const cells = values.map((v) => `<td>${formatter(v, label)}</td>`).join("");
      return `<tr><td>${label}</td>${cells}</tr>`;
    })
    .join("");

  return `<section class="report-section"><h3>${title}</h3><table><thead>${head}</thead><tbody>${body}</tbody></table></section>`;
}

function getRowsByMap(yearRecords, mapKey) {
  const keys = Object.keys(yearRecords[0]?.[mapKey] || {});
  return keys.map((key) => ({
    label: key,
    values: yearRecords.map((year) => year[mapKey][key]),
  }));
}

function buildNarrative(report) {
  const series = report.allYears;
  const salesStart = series[0].pl.Sales;
  const salesEnd = series[series.length - 1].pl.Sales;
  const patStart = series[0].pl["Profit after tax"];
  const patEnd = series[series.length - 1].pl["Profit after tax"];
  const avgDSCR = safeDivide(
    series.filter((y) => y.meta.isTLYear).reduce((acc, y) => acc + y.ratios.DSCR, 0),
    Math.max(1, series.filter((y) => y.meta.isTLYear).length),
  );
  const maxCC = Math.max(...series.filter((y) => y.meta.isCCYear).map((y) => y.wc["MPBF / CC required"]), 0);

  return [
    `Historical performance trend indicates sales movement from ${fmtCurrency(salesStart)} to ${fmtCurrency(salesEnd)} and PAT from ${fmtCurrency(patStart)} to ${fmtCurrency(patEnd)}.`,
    `Projected growth assumptions are aligned with annual sales growth of ${getConfigNumber("sales-growth", 10).toFixed(1)}% and gross margin of ${getConfigNumber("gross-margin", 28).toFixed(1)}%.`,
    `Working capital assessment shows peak MPBF / cash credit requirement of ${fmtCurrency(maxCC)} based on receivable, inventory and creditor cycle assumptions.`,
    `Repayment comfort appears ${avgDSCR >= 1.25 ? "adequate" : "stretched"} with average DSCR of ${avgDSCR.toFixed(2)} over the term-loan horizon including moratorium support.`,
    "Major observations: monitor creditor discipline, maintain margin contribution, and align disbursement with projected operating cash accrual.",
  ];
}

function renderReport(report) {
  const periods = report.periods.allColumns;
  const plRows = getRowsByMap(report.allYears, "pl");
  const bsRows = getRowsByMap(report.allYears, "bs");
  const wcRows = getRowsByMap(report.allYears.filter((y) => y.meta.isHistorical || y.meta.isProvisional || y.meta.isCCYear), "wc");
  const ratioRows = getRowsByMap(report.allYears, "ratios");
  const tlRows = getRowsByMap(report.allYears.filter((y) => y.meta.isTLYear), "termLoan");
  const dscrRows = getRowsByMap(report.allYears.filter((y) => y.meta.isTLYear), "dscr");
  const notes = buildNarrative(report);

  const summaryRows = Object.entries(report.borrowerSummary)
    .map(([k, v]) => `<tr><td>${k}</td><td>${typeof v === "number" ? fmtCurrency(v) : v}</td></tr>`)
    .join("");

  output.innerHTML = `
    <article class="cma-report">
      <header class="report-header">
        <h2>Credit Monitoring Arrangement (CMA) Report</h2>
        <p class="good">Generated with ${periods.length} year-columns (${report.periods.historical.length} historical + provisional + projected).</p>
      </header>

      <section class="report-section">
        <h3>A. Borrower / Proposal Summary</h3>
        <table>
          <tbody>${summaryRows}</tbody>
        </table>
      </section>

      ${buildSectionHtml("B. Profit & Loss Statement", periods, plRows)}
      ${buildSectionHtml("C. Balance Sheet", periods, bsRows)}
      ${buildSectionHtml(
        "D. Working Capital Assessment",
        report.allYears.filter((y) => y.meta.isHistorical || y.meta.isProvisional || y.meta.isCCYear).map((y) => y.period),
        wcRows,
      )}
      ${buildSectionHtml(
        "E. Ratio Analysis",
        periods,
        ratioRows,
        (value, label) => (label.includes("Days") || label.includes("Ratio") || label === "DSCR" ? value.toFixed(2) : fmtPct(value)),
      )}
      ${buildSectionHtml(
        "F. Term Loan Repayment Schedule",
        report.allYears.filter((y) => y.meta.isTLYear).map((y) => y.period),
        tlRows,
      )}
      ${buildSectionHtml(
        "G. DSCR Sheet",
        report.allYears.filter((y) => y.meta.isTLYear).map((y) => y.period),
        dscrRows,
        (value, label) => (label === "DSCR" ? value.toFixed(2) : fmtCurrency(value)),
      )}

      <section class="report-section">
        <h3>H. Banker Summary / Key Notes</h3>
        <ol>${notes.map((line) => `<li>${line}</li>`).join("")}</ol>
      </section>
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
  const baseSales = guessYearlySales(workbookState.sourceRows);
  const data = buildCmaReport(baseSales);

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
  const html = `<!doctype html><html><head><meta charset="UTF-8"><title>${title}</title><style>${document.querySelector("style")?.innerHTML || ""}</style></head><body>${output.innerHTML}</body></html>`;
  saveFile(html, "cma-report.html", "text/html");
}

function formatSheetWithHeader(ws, title, columnCount) {
  ws["!cols"] = Array.from({ length: columnCount }, (_, i) => ({ wch: i === 0 ? 42 : 16 }));
  ws["!merges"] = ws["!merges"] || [];
  ws["!merges"].push({ s: { r: 0, c: 0 }, e: { r: 0, c: Math.max(columnCount - 1, 1) } });
  XLSX.utils.sheet_add_aoa(ws, [[title]], { origin: "A1" });
}

function makeDataSheet(title, periods, rows, valueFormatter) {
  const aoa = [[title], ["Particulars", ...periods]];
  rows.forEach((row) => {
    aoa.push([row.label, ...row.values]);
  });
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  formatSheetWithHeader(ws, title, periods.length + 1);

  for (let r = 3; r <= aoa.length; r += 1) {
    for (let c = 2; c <= periods.length + 1; c += 1) {
      const ref = XLSX.utils.encode_cell({ r: r - 1, c: c - 1 });
      if (ws[ref] && typeof ws[ref].v === "number") {
        ws[ref].z = valueFormatter;
      }
    }
  }

  return ws;
}

function downloadExcel() {
  if (!workbookState.generated) return;

  const wb = XLSX.utils.book_new();
  const report = workbookState.generated;
  const periods = report.periods.allColumns;
  const ccPeriods = report.allYears.filter((y) => y.meta.isHistorical || y.meta.isProvisional || y.meta.isCCYear).map((y) => y.period);
  const tlPeriods = report.allYears.filter((y) => y.meta.isTLYear).map((y) => y.period);

  const summaryData = [["Borrower / Proposal Summary"], ["Particular", "Details"]];
  Object.entries(report.borrowerSummary).forEach(([k, v]) => summaryData.push([k, v]));
  const summaryWs = XLSX.utils.aoa_to_sheet(summaryData);
  formatSheetWithHeader(summaryWs, "Borrower / Proposal Summary", 2);
  summaryWs["!cols"] = [{ wch: 34 }, { wch: 40 }];

  const plWs = makeDataSheet("Profit & Loss Statement", periods, getRowsByMap(report.allYears, "pl"), "₹#,##0");
  const bsWs = makeDataSheet("Balance Sheet", periods, getRowsByMap(report.allYears, "bs"), "₹#,##0");
  const wcWs = makeDataSheet("Working Capital Assessment", ccPeriods, getRowsByMap(report.allYears.filter((y) => y.meta.isHistorical || y.meta.isProvisional || y.meta.isCCYear), "wc"), "₹#,##0");
  const ratioWs = makeDataSheet("Ratio Analysis", periods, getRowsByMap(report.allYears, "ratios"), "0.00");
  const tlWs = makeDataSheet("Term Loan Repayment Schedule", tlPeriods, getRowsByMap(report.allYears.filter((y) => y.meta.isTLYear), "termLoan"), "₹#,##0");
  const dscrWs = makeDataSheet("DSCR Sheet", tlPeriods, getRowsByMap(report.allYears.filter((y) => y.meta.isTLYear), "dscr"), "₹#,##0");

  const assumptions = [["Assumptions"], ["Parameter", "Value"]];
  Object.entries(report.assumptions).forEach(([k, v]) => assumptions.push([k, v]));
  const assumptionsWs = XLSX.utils.aoa_to_sheet(assumptions);
  formatSheetWithHeader(assumptionsWs, "Assumptions", 2);
  assumptionsWs["!cols"] = [{ wch: 36 }, { wch: 20 }];

  XLSX.utils.book_append_sheet(wb, summaryWs, "Summary");
  XLSX.utils.book_append_sheet(wb, plWs, "P&L");
  XLSX.utils.book_append_sheet(wb, bsWs, "Balance Sheet");
  XLSX.utils.book_append_sheet(wb, wcWs, "Working Capital");
  XLSX.utils.book_append_sheet(wb, ratioWs, "Ratios");
  XLSX.utils.book_append_sheet(wb, tlWs, "Term Loan");
  XLSX.utils.book_append_sheet(wb, dscrWs, "DSCR");
  XLSX.utils.book_append_sheet(wb, assumptionsWs, "Assumptions");

  XLSX.writeFile(wb, "cma-report.xlsx");
}

function downloadJson() {
  if (!workbookState.generated) return;
  saveFile(JSON.stringify(workbookState.generated, null, 2), "cma-data.json", "application/json");
}
