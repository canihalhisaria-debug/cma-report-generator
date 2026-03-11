# CMA Report Generator (2-Input Version)

This project now uses **only two user inputs**:

- **CC Amount**
- **ROI (%)**

All major CMA outputs are auto-calculated:

- Sales
- Interest
- Stock
- Debtors
- Creditors
- Projected P&L
- Projected Balance Sheet
- Working Capital Analysis
- Financial Ratios

Also enforced:

- **First year closing stock = second year opening stock**.
- New workbook is generated as `FINAL_CMA_REPORT_FINAL_FIXED.xlsx` (original is not overwritten).

## Web app run

```bash
python3 -m http.server 4173
```

Open `http://localhost:4173`.

## Script run

```bash
python3 cma_fix.py --cc-amount 10000000 --roi 11
```
