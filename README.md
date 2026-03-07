# CMA Report Generator (GitHub Pages Ready)

A static, browser-only CMA Report Generator tailored for CA firms. It reads uploaded Excel files directly in the browser, supports Cash Credit and Term Loan projections, handles 3-4 historical years plus one provisional year, and generates dynamic repayment-year banker-ready outputs.

## Key Features

- **No backend, no binaries**: Pure HTML/CSS/JavaScript app suitable for GitHub Pages.
- **Excel in browser**: Upload `.xlsx` / `.xls`; parsing is done locally using SheetJS CDN, including Tally-style two-sided P&L and Balance Sheet layouts.
- **CMA period modeling**:
  - 3 or 4 historical years
  - 1 provisional year
  - user-defined projection/repayment years
- **Credit analysis outputs**:
  - Working Capital Assessment with Cash Credit requirement
  - Term Loan repayment and interest projection schedule
- **Head mapping controls**:
  - Auto-maps common Tally ledger heads to CMA structure
  - Shows mapping review when confidence is low (no silent sample fallback)
- **Export options**:
  - Download banker-ready report as HTML
  - Download generated data as JSON

## Run Locally

Any static server works.

```bash
python3 -m http.server 4173
```

Open: `http://localhost:4173`

## Deploy to GitHub Pages

1. Push this repository to GitHub.
2. In **Settings → Pages**:
   - Source: **Deploy from a branch**
   - Branch: `main` (or your branch), folder: `/ (root)`
3. Save and wait for deployment.

Because this app is fully static and browser-based, it is fully compatible with GitHub Pages.
