# Banker-Grade CMA Dashboard

Static browser app for Credit Monitoring Arrangement (CMA) projection with a structured banker-grade dashboard.

## What this build includes

- **Dashboard input sheet behavior**: actual current-year values are entered and treated as locked baseline.
- **Projection controls**:
  - Domestic Sale Growth %
  - Export Sale Growth %
  - Other Income Growth %
  - Purchase / Material Cost Growth %
  - Operating Expense Growth %
  - Debtor Days
  - Creditor Days
  - Inventory Holding Days
  - Interest Rate on CC
  - Interest Rate on TL
  - Tax Rate
- **Auto-generation output statements**:
  - Profit & Loss
  - Balance Sheet
  - Working Capital
  - Ratios
  - DSCR
  - Validation
- **Formula trace** is visible in the generated tables for key calculated heads.
- **Excel export** with required output sheets:
  - Dashboard
  - Profit & Loss
  - Balance Sheet
  - Working Capital
  - Ratios
  - DSCR
  - Validation

## Run

```bash
python3 -m http.server 4173
```

Open `http://localhost:4173`.
