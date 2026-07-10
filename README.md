# finance

## LBO Model Builder

`build_lbo_model.py` generates a formula-driven LBO (leveraged buyout) model as an Excel workbook — Assumptions, Sources & Uses, Operating Model, Free Cash Flow, Debt Schedule, and Returns (MOIC/IRR), all linked with live formulas rather than static numbers. `LBO_Model.xlsx` is the sample output with default assumptions.

### How to run

```
pip install openpyxl
python3 build_lbo_model.py
```

This writes `LBO_Model.xlsx` next to the script. Edit the yellow input cells on the Assumptions tab to change entry/exit multiples, leverage, growth, margins, etc. — the rest of the model recalculates automatically.
