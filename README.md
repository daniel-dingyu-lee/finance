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
