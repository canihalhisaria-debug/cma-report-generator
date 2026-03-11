const num = (id) => Number(document.getElementById(id).value || 0);

const INPUT_IDS = ["cc-amount", "roi", "projection-years"];

function safeDiv(a, b) {
  return b === 0 ? 0 : a / b;
}

function money(v) {
  return Number(v || 0).toLocaleString("en-IN", { maximumFractionDigits: 0 });
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
  const row = (label, vals, kind = "money") => `<tr><td>${label}</td>${vals.map((v) => `<td>${kind === "pct" ? pct(v) : kind === "num" ? Number(v).toFixed(2) : money(v)}</td>`).join("")}</tr>`;
  const table = (title, body) => `<h3>${title}</h3><table><thead><tr><th>Head</th>${periods.map((p) => `<th>${p}</th>`).join("")}</tr></thead><tbody>${body}</tbody></table>`;

  const plRows = [
    row("Sales", data.map((d) => d.totalIncome)),
    row("Opening Stock", data.map((d) => d.os)),
    row("Purchases", data.map((d) => d.pur)),
    row("Closing Stock", data.map((d) => d.cs)),
    row("Material Cost", data.map((d) => d.materialCost)),
    row("Operating Expenses", data.map((d) => d.totalOpex)),
    row("Interest on CC", data.map((d) => d.ccInt)),
    row("Interest on TL", data.map((d) => d.tlInt)),
    row("EBIT", data.map((d) => d.ebit)),
    row("EBT", data.map((d) => d.ebt)),
    row("PAT", data.map((d) => d.pat)),
  ].join("");

  const bsRows = [
    row("Inventory", data.map((d) => d.inv)),
    row("Debtors", data.map((d) => d.debt)),
    row("Creditors", data.map((d) => d.cred)),
    row("Current Assets", data.map((d) => d.ca)),
    row("Current Liabilities", data.map((d) => d.cl)),
    row("Working Capital", data.map((d) => d.wc)),
  ].join("");

  const ratioRows = [
    row("Current Ratio", data.map((d) => d.currentRatio), "num"),
    row("Quick Ratio", data.map((d) => d.quickRatio), "num"),
    row("EBIT Ratio", data.map((d) => d.ebitRatio), "pct"),
    row("Net Profit Ratio", data.map((d) => d.netProfitRatio), "pct"),
    row("DSCR", data.map((d) => d.dscr), "num"),
  ].join("");

  document.getElementById("output").innerHTML =
    table("Projected P&L", plRows) +
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
