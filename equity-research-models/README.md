# Equity research models

This repo contains integrated financial models and equity research notes that I built as part of a self-directed project after finishing my economics degree. The goal was to replicate the kind of work a junior analyst would do at an investment bank or Big 4 advisory firm.

Each company gets the same treatment: a full 3-statement model in Excel, a short equity research note in Word, and a Python script that validates the input data. Everything is built from scratch using public filings, not templates.

## What's in here

### Tesco PLC (TSCO LN)

Tesco is the UK's largest grocer with about 27% market share. I picked it because it has clean financials, plenty of public data, and interesting modelling challenges around IFRS 16 lease accounting.

**Files:**

- `Tesco_Equity_Research_Model_v2.xlsx` - 12-tab integrated model (933 formulas, 0 errors)
- `Tesco_Equity_Research_Note_v2.docx` - 3-page equity research note with BUY recommendation at 310p
- `tesco_data_validation.py` - Python script that cross-checks model inputs against source data

**Model structure:**

The Excel model has a Cover page, Assumptions tab (all inputs in blue, all formulas in black, cross-sheet links in green), Income Statement, four supporting schedules (Working Capital with DIO/DSO/DPO, Capex & D&A with PP&E and ROU rollforwards, Debt & Interest with bond and IFRS 16 lease rollforwards, Dividend & Buyback with share count roll), Balance Sheet with audit check, Cash Flow Statement, DCF Valuation (mid-year convention, Gordon Growth terminal value, WACC/TGR sensitivity table), Trading Comps (6 European grocery peers), and a Charts tab.

**Valuation summary:**

DCF implies roughly 280p per share. Comps point to 280-310p. Current price is 260p. I set a BUY target at 310p, about 19% upside.

### Diageo PLC (DGE LN)

Diageo is the world's largest spirits company, with brands like Johnnie Walker, Tanqueray, Smirnoff, Guinness, and Don Julio. I chose it because it operates in a completely different sector from Tesco, with much higher margins (~35% EBITDA vs ~9.5%), massive brand intangibles on the balance sheet, and an unusual working capital profile where inventory days run at 160-170 (spirits need to age in barrels for years before they can be sold).

**Files:**

- `Diageo_Equity_Research_Model_v2.xlsx` - 12-tab integrated model (933 formulas, 0 errors)
- `Diageo_Equity_Research_Note_v2.docx` - 3-page equity research note with BUY recommendation at 2,800p
- `diageo_data_validation.py` - Python script that cross-checks model inputs (84 checks, all pass)

**Model structure:**

Same 12-sheet architecture as Tesco. Revenue segmented into North America and International. The Capex & DA tab includes a Goodwill & Brands line (about £17.8bn of indefinite-life intangibles). The Debt schedule reflects Diageo's heavier leverage versus Tesco (2.7x ND/EBITDA vs 0.3x), and the forecast models a gradual deleveraging over FY25-29.

**Valuation summary:**

DCF (WACC 6.2%, TGR 2.5%) implies an enterprise value of roughly £70bn. After deducting £16.9bn of debt and £750m of leases, the equity value works out to about 2,800p per share. Trading comps against Pernod Ricard, Brown-Forman, Campari, Remy Cointreau, and Constellation Brands support a similar range. Current price is 2,500p, so I set a BUY target at 2,800p, about 12% upside.

## Technical details

**Excel model:** Built with Python (openpyxl). Every cell contains a proper Excel formula, not a hardcoded number. You can change any blue input and the whole model recalculates. Source notes in column C cite specific pages from Tesco's Annual Report.

**Word documents:** Generated with Node.js (docx-js). Professional formatting with headers, footers, page numbers, and tables.

**Data validation:** Each company has a Python script that sanity-checks the historical inputs. The scripts verify that balance sheets balance, that growth rates are mathematically consistent, and that key ratios fall within reasonable ranges.

## Methodology

**Revenue:** Segmental build (UK & ROI + Central Europe). Historical from annual reports, forecast using consensus-anchored growth rates.

**EBITDA:** Margin-driven. Historical margins calculated, forecast margins ramped conservatively toward management guidance.

**D&A:** Bottom-up from PP&E and ROU asset rollforwards, plus intangible amortisation.

**Working capital:** Days-based (DIO, DSO, DPO applied to revenue).

**Debt:** Bond rollforward with scheduled repayments. IFRS 16 lease liability rollforward separate.

**DCF:** WACC from CAPM (risk-free = 10yr gilt, beta from Bloomberg, ERP from Damodaran). Mid-year convention. Gordon Growth terminal value. Sensitivity table varies WACC and terminal growth.

**Comps:** EV/EBITDA and P/E multiples for 5-6 sector peers. Median applied to subject company to get implied equity value.

## Tools used

- Python 3 (openpyxl for Excel generation, pandas/numpy for data validation)
- Node.js (docx-js for Word document generation)
- LibreOffice (formula recalculation and error checking via recalc.py)
- Data sources: Tesco Annual Reports FY2020-FY2024, Diageo Annual Reports FY2020-FY2024, Bloomberg, FactSet consensus, Damodaran

## About me

Zak Whatmough. Recent economics graduate looking for junior analyst roles in investment banking, private equity, or Big 4 advisory. This project is meant to show that I can build a real model from scratch, write a coherent investment case, and use Python to automate and validate financial data.

If you have questions about the models or want to discuss anything, feel free to open an issue or reach out.
