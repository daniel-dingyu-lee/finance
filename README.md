# finance

## LBO Model Builder

`build_lbo_model.py` generates a formula-driven LBO (leveraged buyout) model as an Excel workbook — Assumptions, Sources & Uses, Operating Model, Free Cash Flow, Debt Schedule, Returns (MOIC/IRR), and a Sensitivity grid, all linked with live formulas rather than static numbers. `LBO_Model.xlsx` is the sample output with default assumptions.

The Sensitivity tab shows MOIC and IRR across a 5x5 grid of Entry and Exit EV/EBITDA multiples, holding leverage, growth, margins, and rates fixed at their Assumptions-tab values.

### How to run

```
pip install openpyxl
python3 build_lbo_model.py
```

This writes `LBO_Model.xlsx` next to the script. Edit the yellow input cells on the Assumptions tab to change entry/exit multiples, leverage, growth, margins, etc. — the rest of the model recalculates automatically.

## Private Credit Model Builder

`build_private_credit_model.py` generates a formula-driven private credit (unitranche) model as an Excel workbook — Summary dashboard, Assumptions (Tranche A "First Out" / Tranche B "Last Out"), SOFR forward curve, quarterly Debt Schedule, Covenant testing (leverage & interest coverage), and lender Returns (IRR/MOIC) by tranche and blended. `Private_Credit_Model.xlsx` is the sample output with default assumptions.

The Debt Schedule includes a cash sweep: excess free cash flow (EBITDA × FCF conversion % less cash interest and mandatory amortization) is swept sequentially — First Out paid down before Last Out — at par, accelerating principal repayment beyond the scheduled amortization.

### How to run

```
pip install openpyxl
python3 build_private_credit_model.py
```

This writes `Private_Credit_Model.xlsx` next to the script. Edit the yellow input cells on the Assumptions tab (spread, OID, leverage covenants, sweep %, exit year, etc.) — the rest of the model recalculates automatically.

## Three-Statement + DCF Model Builder

`build_three_statement_dcf_model.py` generates a linked 3-statement operating model (Income Statement, Balance Sheet, Cash Flow Statement) plus a DCF that derives Unlevered Free Cash Flow directly from the operating model, rather than projecting FCF independently. `Three_Statement_DCF_Model.xlsx` is the sample output with default assumptions.

Years run horizontally across columns (2025A actual, 2026E–2030E projected). The Balance Sheet includes a Check row that ties Assets to Liabilities & Equity to zero every year, confirming the statements are fully linked (AR/Inventory/AP off DSO/DIO/DPO assumptions, PP&E rolling forward with CapEx and D&A, Retained Earnings rolling forward with Net Income).

### How to run

```
pip install openpyxl
python3 build_three_statement_dcf_model.py
```

This writes `Three_Statement_DCF_Model.xlsx` next to the script. Edit the blue input cells on the Assumptions tab (tax rate, WACC, terminal growth, margins, growth rates, working capital days) — the rest of the model recalculates automatically.
