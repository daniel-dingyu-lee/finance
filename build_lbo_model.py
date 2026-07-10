import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# ---------- Styles ----------
NAVY = "1F3864"
LIGHT_BLUE = "D9E2F3"
GREY = "F2F2F2"
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")   # yellow = hardcoded input
CALC_FILL = PatternFill("solid", fgColor="FFFFFF")     # white = formula
HEADER_FILL = PatternFill("solid", fgColor=NAVY)
SECTION_FILL = PatternFill("solid", fgColor=LIGHT_BLUE)
TOTAL_FILL = PatternFill("solid", fgColor=GREY)

HEADER_FONT = Font(bold=True, color="FFFFFF", size=12)
SECTION_FONT = Font(bold=True, color=NAVY, size=11)
TOTAL_FONT = Font(bold=True)
INPUT_FONT = Font(color="1F4E78")
TITLE_FONT = Font(bold=True, size=14, color=NAVY)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

USD0 = '#,##0'
USD1 = '#,##0.0'
PCT1 = '0.0%'
MULT = '0.00"x"'

def style_header(ws, row, col1, col2):
    for c in range(col1, col2 + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT

def title(ws, text, row=1, span=8):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = TITLE_FONT

def section(ws, text, row, span=8):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    cell = ws.cell(row=row, column=1, value=text)
    cell.fill = SECTION_FILL
    cell.font = SECTION_FONT

def label(ws, row, col, text, bold=False, indent=0):
    cell = ws.cell(row=row, column=col, value=text)
    cell.font = Font(bold=bold)
    cell.alignment = Alignment(indent=indent)
    return cell

def inp(ws, row, col, value, fmt=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.fill = INPUT_FILL
    cell.font = INPUT_FONT
    if fmt:
        cell.number_format = fmt
    cell.border = BORDER
    return cell

def calc(ws, row, col, formula, fmt=None, bold=False):
    cell = ws.cell(row=row, column=col, value=formula)
    if fmt:
        cell.number_format = fmt
    if bold:
        cell.font = TOTAL_FONT
    cell.border = BORDER
    return cell

def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

# =========================================================
# SHEET 1: ASSUMPTIONS
# =========================================================
ws = wb.active
ws.title = "Assumptions"
set_widths(ws, [34, 14, 14, 14, 14, 14, 14, 14])
title(ws, "LBO Model — Assumptions", 1)
ws.cell(row=2, column=1, value="Yellow cells = hardcoded inputs. Change these to drive the whole model.").font = Font(italic=True, size=9, color="808080")

r = 4
section(ws, "Transaction", r); r += 1
label(ws, r, 1, "Transaction / Entry Year"); inp(ws, r, 2, 2026); ROW_ENTRY_YEAR = r; r += 1
label(ws, r, 1, "Hold Period (years)"); inp(ws, r, 2, 5); ROW_HOLD = r; r += 1
label(ws, r, 1, "Entry EV / EBITDA Multiple"); inp(ws, r, 2, 9.0, MULT); ROW_ENTRY_MULT = r; r += 1
label(ws, r, 1, "Exit EV / EBITDA Multiple"); inp(ws, r, 2, 9.0, MULT); ROW_EXIT_MULT = r; r += 1
label(ws, r, 1, "Transaction Fees (% of EV)"); inp(ws, r, 2, 0.02, PCT1); ROW_FEES = r; r += 1
label(ws, r, 1, "Financing Fees (% of Total Debt)"); inp(ws, r, 2, 0.02, PCT1); ROW_FIN_FEES = r; r += 2

section(ws, "Operating Assumptions (Entry Year, Year 0)", r); r += 1
label(ws, r, 1, "Entry Revenue ($mm)"); inp(ws, r, 2, 500.0, USD1); ROW_REV0 = r; r += 1
label(ws, r, 1, "Revenue Growth Rate (annual)"); inp(ws, r, 2, 0.06, PCT1); ROW_GROWTH = r; r += 1
label(ws, r, 1, "Entry EBITDA Margin"); inp(ws, r, 2, 0.22, PCT1); ROW_MARGIN0 = r; r += 1
label(ws, r, 1, "EBITDA Margin Expansion (bps/yr)"); inp(ws, r, 2, 0.0025, PCT1); ROW_MARGIN_EXP = r; r += 1
label(ws, r, 1, "D&A (% of Revenue)"); inp(ws, r, 2, 0.035, PCT1); ROW_DA = r; r += 1
label(ws, r, 1, "Capex (% of Revenue)"); inp(ws, r, 2, 0.03, PCT1); ROW_CAPEX = r; r += 1
label(ws, r, 1, "Change in NWC (% of Rev. Growth $)"); inp(ws, r, 2, 0.15, PCT1); ROW_NWC = r; r += 1
label(ws, r, 1, "Tax Rate"); inp(ws, r, 2, 0.25, PCT1); ROW_TAX = r; r += 2

section(ws, "Financing Assumptions", r); r += 1
label(ws, r, 1, "Term Loan Leverage (x EBITDA)"); inp(ws, r, 2, 4.5, MULT); ROW_TL_LEV = r; r += 1
label(ws, r, 1, "Term Loan Interest Rate"); inp(ws, r, 2, 0.085, PCT1); ROW_TL_RATE = r; r += 1
label(ws, r, 1, "Term Loan Mandatory Amort (%/yr of orig.)"); inp(ws, r, 2, 0.01, PCT1); ROW_TL_AMORT = r; r += 1
label(ws, r, 1, "Subordinated Notes Leverage (x EBITDA)"); inp(ws, r, 2, 1.5, MULT); ROW_SUB_LEV = r; r += 1
label(ws, r, 1, "Subordinated Notes Interest Rate"); inp(ws, r, 2, 0.11, PCT1); ROW_SUB_RATE = r; r += 1
label(ws, r, 1, "Revolver Interest Rate"); inp(ws, r, 2, 0.075, PCT1); ROW_REV_RATE = r; r += 1
label(ws, r, 1, "Minimum Cash Balance ($mm)"); inp(ws, r, 2, 10.0, USD1); ROW_MIN_CASH = r; r += 1
label(ws, r, 1, "Cash Sweep (% of Excess FCF to Debt)"); inp(ws, r, 2, 1.00, PCT1); ROW_SWEEP = r; r += 1
label(ws, r, 1, "Management Rollover (% of Sponsor Equity)"); inp(ws, r, 2, 0.00, PCT1); ROW_MGMT = r; r += 2

A = "Assumptions"
def aref(row):
    return f"'{A}'!$B${row}"

# =========================================================
# SHEET 2: SOURCES & USES
# =========================================================
su = wb.create_sheet("Sources & Uses")
set_widths(su, [34, 16, 14, 16])
title(su, "Sources & Uses", 1)
r = 3
label(su, r, 1, "Entry EBITDA ($mm)", bold=True)
calc(su, r, 2, f"='{A}'!$B${ROW_REV0}*'{A}'!$B${ROW_MARGIN0}", USD1, bold=True)
ROW_SU_EBITDA = r; r += 2

section(su, "Uses", r); r += 1
label(su, r, 1, "Purchase Enterprise Value")
calc(su, r, 2, f"=B{ROW_SU_EBITDA}*{aref(ROW_ENTRY_MULT)}", USD1)
ROW_USES_EV = r; r += 1
label(su, r, 1, "Transaction Fees")
calc(su, r, 2, f"=B{ROW_USES_EV}*{aref(ROW_FEES)}", USD1)
ROW_USES_FEES = r; r += 1
label(su, r, 1, "Financing Fees")
ROW_USES_FINFEES = r
r += 1
label(su, r, 1, "Total Uses", bold=True)
ROW_USES_TOTAL = r
for c in range(1, 3):
    su.cell(row=r, column=c).fill = TOTAL_FILL
r += 2

section(su, "Sources", r); r += 1
label(su, r, 1, "Term Loan")
calc(su, r, 2, f"=B{ROW_SU_EBITDA}*{aref(ROW_TL_LEV)}", USD1)
ROW_SRC_TL = r; r += 1
label(su, r, 1, "Subordinated Notes")
calc(su, r, 2, f"=B{ROW_SU_EBITDA}*{aref(ROW_SUB_LEV)}", USD1)
ROW_SRC_SUB = r; r += 1
label(su, r, 1, "Revolver Draw at Close")
calc(su, r, 2, 0, USD1)
ROW_SRC_REV = r; r += 1
label(su, r, 1, "Total Debt")
calc(su, r, 2, f"=SUM(B{ROW_SRC_TL}:B{ROW_SRC_REV})", USD1)
ROW_SRC_DEBT_TOTAL = r; r += 1
label(su, r, 1, "Management Rollover Equity")
ROW_SRC_MGMT = r
r += 1
label(su, r, 1, "Sponsor Equity (plug)", bold=True)
ROW_SRC_SPONSOR = r
r += 1
label(su, r, 1, "Total Sources", bold=True)
ROW_SRC_TOTAL = r
for c in range(1, 3):
    su.cell(row=r, column=c).fill = TOTAL_FILL
r += 1
label(su, r, 1, "Check (Sources - Uses)")
ROW_CHECK = r
r += 2

# fill financing fees now that TL/Sub known
calc(su, ROW_USES_FINFEES, 2, f"=B{ROW_SRC_DEBT_TOTAL}*{aref(ROW_FIN_FEES)}", USD1)
calc(su, ROW_USES_TOTAL, 2, f"=SUM(B{ROW_USES_EV}:B{ROW_USES_FINFEES})", USD1, bold=True)

calc(su, ROW_SRC_MGMT, 2, f"=(B{ROW_USES_TOTAL}-B{ROW_SRC_DEBT_TOTAL})*{aref(ROW_MGMT)}", USD1)
calc(su, ROW_SRC_SPONSOR, 2, f"=B{ROW_USES_TOTAL}-B{ROW_SRC_DEBT_TOTAL}-B{ROW_SRC_MGMT}", USD1, bold=True)
calc(su, ROW_SRC_TOTAL, 2, f"=B{ROW_SRC_DEBT_TOTAL}+B{ROW_SRC_MGMT}+B{ROW_SRC_SPONSOR}", USD1, bold=True)
calc(su, ROW_CHECK, 2, f"=B{ROW_SRC_TOTAL}-B{ROW_USES_TOTAL}", USD1)

label(su, r, 1, "Opening Leverage (Total Debt / EBITDA)")
calc(su, r, 2, f"=B{ROW_SRC_DEBT_TOTAL}/B{ROW_SU_EBITDA}", MULT)
r += 1

S = "Sources & Uses"

# =========================================================
# SHEET 3: OPERATING MODEL
# =========================================================
om = wb.create_sheet("Operating Model")
set_widths(om, [30] + [14]*6)
title(om, "Operating Model", 1)

r = 3
label(om, r, 1, "Year", bold=True)
for i in range(0, 6):
    c = om.cell(row=r, column=2+i, value=(f"=Year 0" if i==0 else None))
YEAR_ROW = r
for i in range(6):
    if i == 0:
        calc(om, r, 2+i, f"={A}!$B${ROW_ENTRY_YEAR}", '0')
    else:
        calc(om, r, 2+i, f"={get_column_letter(2+i-1)}{r}+1", '0')
    om.cell(row=r, column=2+i).font = TOTAL_FONT
r += 1

label(om, r, 1, "Revenue", bold=True)
REV_ROW = r
calc(om, r, 2, f"={A}!$B${ROW_REV0}", USD1, bold=True)
for i in range(1, 6):
    col = get_column_letter(2+i); prev = get_column_letter(2+i-1)
    calc(om, r, 2+i, f"={prev}{r}*(1+{A}!$B${ROW_GROWTH})", USD1, bold=True)
r += 1

label(om, r, 1, "  % growth")
GROWTH_ROW = r
calc(om, r, 2, "", PCT1)
for i in range(1, 6):
    col = get_column_letter(2+i); prev = get_column_letter(2+i-1)
    calc(om, r, 2+i, f"={col}{REV_ROW}/{prev}{REV_ROW}-1", PCT1)
r += 1

label(om, r, 1, "EBITDA Margin")
MARGIN_ROW = r
calc(om, r, 2, f"={A}!$B${ROW_MARGIN0}", PCT1)
for i in range(1, 6):
    col = get_column_letter(2+i); prev = get_column_letter(2+i-1)
    calc(om, r, 2+i, f"={prev}{r}+{A}!$B${ROW_MARGIN_EXP}", PCT1)
r += 1

label(om, r, 1, "EBITDA", bold=True)
EBITDA_ROW = r
for i in range(6):
    col = get_column_letter(2+i)
    calc(om, r, 2+i, f"={col}{REV_ROW}*{col}{MARGIN_ROW}", USD1, bold=True)
r += 1

label(om, r, 1, "Less: D&A")
DA_ROW = r
for i in range(6):
    col = get_column_letter(2+i)
    calc(om, r, 2+i, f"=-{col}{REV_ROW}*{A}!$B${ROW_DA}", USD1)
r += 1

label(om, r, 1, "EBIT")
EBIT_ROW = r
for i in range(6):
    col = get_column_letter(2+i)
    calc(om, r, 2+i, f"={col}{EBITDA_ROW}+{col}{DA_ROW}", USD1)
r += 1

label(om, r, 1, "Less: Cash Interest Expense, net")
INT_ROW = r
for i in range(6):
    col = get_column_letter(2+i)
    if i == 0:
        calc(om, r, 2+i, 0, USD1)
    else:
        calc(om, r, 2+i, f"=-'Debt Schedule'!{col}{{INT_TOTAL_ROW}}", USD1)
r += 1

label(om, r, 1, "EBT")
EBT_ROW = r
for i in range(6):
    col = get_column_letter(2+i)
    calc(om, r, 2+i, f"={col}{EBIT_ROW}+{col}{INT_ROW}", USD1)
r += 1

label(om, r, 1, "Less: Cash Taxes")
TAX_ROW = r
for i in range(6):
    col = get_column_letter(2+i)
    calc(om, r, 2+i, f"=-MAX({col}{EBT_ROW},0)*{A}!$B${ROW_TAX}", USD1)
r += 1

label(om, r, 1, "Net Income", bold=True)
NI_ROW = r
for i in range(6):
    col = get_column_letter(2+i)
    calc(om, r, 2+i, f"={col}{EBT_ROW}+{col}{TAX_ROW}", USD1, bold=True)
r += 2

O = "Operating Model"

# =========================================================
# SHEET 4: FREE CASH FLOW
# =========================================================
fc = wb.create_sheet("Free Cash Flow")
set_widths(fc, [32] + [14]*6)
title(fc, "Free Cash Flow", 1)
r = 3
label(fc, r, 1, "Year", bold=True)
for i in range(6):
    col = get_column_letter(2+i)
    calc(fc, r, 2+i, f"='{O}'!{col}{YEAR_ROW}", '0', bold=True)
FCF_YEAR_ROW = r; r += 1

label(fc, r, 1, "Net Income")
FCF_NI_ROW = r
for i in range(6):
    col = get_column_letter(2+i)
    calc(fc, r, 2+i, f"='{O}'!{col}{NI_ROW}" if i>0 else 0, USD1)
r += 1

label(fc, r, 1, "Plus: D&A")
FCF_DA_ROW = r
for i in range(6):
    col = get_column_letter(2+i)
    calc(fc, r, 2+i, f"=-'{O}'!{col}{DA_ROW}" if i>0 else 0, USD1)
r += 1

label(fc, r, 1, "Less: Capex")
FCF_CAPEX_ROW = r
for i in range(6):
    col = get_column_letter(2+i)
    calc(fc, r, 2+i, f"=-'{O}'!{col}{REV_ROW}*{A}!$B${ROW_CAPEX}" if i>0 else 0, USD1)
r += 1

label(fc, r, 1, "Less: Increase in NWC")
FCF_NWC_ROW = r
for i in range(6):
    col = get_column_letter(2+i); prev = get_column_letter(2+i-1)
    if i == 0:
        calc(fc, r, 2+i, 0, USD1)
    else:
        calc(fc, r, 2+i, f"=-('{O}'!{col}{REV_ROW}-'{O}'!{prev}{REV_ROW})*{A}!$B${ROW_NWC}", USD1)
r += 1

label(fc, r, 1, "Unlevered Free Cash Flow", bold=True)
FCF_UNLEV_ROW = r
for i in range(6):
    col = get_column_letter(2+i)
    calc(fc, r, 2+i, f"=SUM({col}{FCF_NI_ROW}:{col}{FCF_NWC_ROW})+{col}{FCF_DA_ROW}*0" , USD1, bold=True)
    # NI already nets interest & taxes below; recompute properly below
r += 2

F = "Free Cash Flow"

# Simpler/clean cash-available-for-debt-paydown build (avoid double count):
# CFADS = Net Income + D&A - Capex - Increase in NWC  (interest & taxes already in NI)
for i in range(6):
    col = get_column_letter(2+i)
    formula = f"={col}{FCF_NI_ROW}+{col}{FCF_DA_ROW}+{col}{FCF_CAPEX_ROW}+{col}{FCF_NWC_ROW}"
    fc.cell(row=FCF_UNLEV_ROW, column=2+i, value=formula)
    fc.cell(row=FCF_UNLEV_ROW, column=2+i).number_format = USD1
    fc.cell(row=FCF_UNLEV_ROW, column=2+i).font = TOTAL_FONT

fc.cell(row=3, column=1, value="Year")
label(fc, FCF_UNLEV_ROW, 1, "Cash Flow Available for Debt Service (CFADS)", bold=True)

# =========================================================
# SHEET 5: DEBT SCHEDULE
# =========================================================
ds = wb.create_sheet("Debt Schedule")
set_widths(ds, [32] + [14]*6)
title(ds, "Debt Schedule", 1)
r = 3
label(ds, r, 1, "Year", bold=True)
for i in range(6):
    col = get_column_letter(2+i)
    calc(ds, r, 2+i, f"='{O}'!{col}{YEAR_ROW}", '0', bold=True)
DS_YEAR_ROW = r; r += 2

label(ds, r, 1, "CFADS (from FCF tab)")
DS_CFADS_ROW = r
for i in range(6):
    col = get_column_letter(2+i)
    calc(ds, r, 2+i, f"='{F}'!{col}{FCF_UNLEV_ROW}" if i>0 else 0, USD1)
r += 1

label(ds, r, 1, "Less: Mandatory Term Loan Amort.")
DS_MAND_ROW = r
for i in range(6):
    col = get_column_letter(2+i)
    if i == 0:
        calc(ds, r, 2+i, 0, USD1)
    else:
        calc(ds, r, 2+i, f"=-MIN('{S}'!$B${ROW_SRC_TL}*{A}!$B${ROW_TL_AMORT}, {{TL_BEG}})", USD1)
r += 1

label(ds, r, 1, "Cash Flow Available for Optional Sweep")
DS_SWEEP_AVAIL_ROW = r
for i in range(6):
    col = get_column_letter(2+i)
    calc(ds, r, 2+i, f"=MAX({col}{DS_CFADS_ROW}+{col}{DS_MAND_ROW},0)*{A}!$B${ROW_SWEEP}" if i>0 else 0, USD1)
r += 2

section(ds, "Term Loan", r); r += 1
label(ds, r, 1, "Beginning Balance")
TL_BEG_ROW = r; r += 1
label(ds, r, 1, "Mandatory Amortization")
TL_MAND_ROW = r; r += 1
label(ds, r, 1, "Optional Sweep")
TL_SWEEP_ROW = r; r += 1
label(ds, r, 1, "Ending Balance", bold=True)
TL_END_ROW = r; r += 1
label(ds, r, 1, "Interest Expense (on beg. balance)")
TL_INT_ROW = r; r += 2

section(ds, "Subordinated Notes", r); r += 1
label(ds, r, 1, "Beginning Balance")
SUB_BEG_ROW = r; r += 1
label(ds, r, 1, "Optional Sweep (after Term Loan repaid)")
SUB_SWEEP_ROW = r; r += 1
label(ds, r, 1, "Ending Balance", bold=True)
SUB_END_ROW = r; r += 1
label(ds, r, 1, "Interest Expense (on beg. balance)")
SUB_INT_ROW = r; r += 2

section(ds, "Revolver", r); r += 1
label(ds, r, 1, "Beginning Balance")
REV_BEG_ROW = r; r += 1
label(ds, r, 1, "(Draw) / Paydown")
REV_DRAW_ROW = r; r += 1
label(ds, r, 1, "Ending Balance", bold=True)
REV_END_ROW = r; r += 1
label(ds, r, 1, "Interest Expense (on beg. balance)")
REV_INT_ROW = r; r += 2

label(ds, r, 1, "Total Interest Expense", bold=True)
INT_TOTAL_ROW = r
for c in range(1, 8):
    ds.cell(row=r, column=c).fill = TOTAL_FILL
r += 1
label(ds, r, 1, "Total Debt (Ending)", bold=True)
TOTAL_DEBT_ROW = r
for c in range(1, 8):
    ds.cell(row=r, column=c).fill = TOTAL_FILL
r += 1
label(ds, r, 1, "Cash Balance")
CASH_ROW = r; r += 1
label(ds, r, 1, "Net Debt", bold=True)
NET_DEBT_ROW = r
for c in range(1, 8):
    ds.cell(row=r, column=c).fill = TOTAL_FILL
r += 1
label(ds, r, 1, "Leverage (Net Debt / EBITDA)")
LEV_ROW = r; r += 1

# now populate Term Loan
for i in range(6):
    col = get_column_letter(2+i); prev = get_column_letter(2+i-1)
    if i == 0:
        calc(ds, TL_BEG_ROW, 2+i, 0, USD1)
        calc(ds, TL_MAND_ROW, 2+i, 0, USD1)
        calc(ds, TL_SWEEP_ROW, 2+i, 0, USD1)
        calc(ds, TL_END_ROW, 2+i, f"='{S}'!$B${ROW_SRC_TL}", USD1, bold=True)
        calc(ds, TL_INT_ROW, 2+i, 0, USD1)
    else:
        calc(ds, TL_BEG_ROW, 2+i, f"={prev}{TL_END_ROW}", USD1)
        calc(ds, TL_MAND_ROW, 2+i, f"=-MIN('{S}'!$B${ROW_SRC_TL}*{A}!$B${ROW_TL_AMORT},{col}{TL_BEG_ROW})", USD1)
        calc(ds, TL_SWEEP_ROW, 2+i, f"=-MIN(MAX({col}{DS_SWEEP_AVAIL_ROW},0),{col}{TL_BEG_ROW}+{col}{TL_MAND_ROW})", USD1)
        calc(ds, TL_END_ROW, 2+i, f"={col}{TL_BEG_ROW}+{col}{TL_MAND_ROW}+{col}{TL_SWEEP_ROW}", USD1, bold=True)
        calc(ds, TL_INT_ROW, 2+i, f"={col}{TL_BEG_ROW}*{A}!$B${ROW_TL_RATE}", USD1)

# fix the mandatory amort formula reference in DS_MAND_ROW (placeholder {TL_BEG} was unused; recompute cleanly)
for i in range(6):
    col = get_column_letter(2+i)
    if i == 0:
        calc(ds, DS_MAND_ROW, 2+i, 0, USD1)
    else:
        calc(ds, DS_MAND_ROW, 2+i, f"={col}{TL_MAND_ROW}", USD1)

# Subordinated Notes
for i in range(6):
    col = get_column_letter(2+i); prev = get_column_letter(2+i-1)
    if i == 0:
        calc(ds, SUB_BEG_ROW, 2+i, 0, USD1)
        calc(ds, SUB_SWEEP_ROW, 2+i, 0, USD1)
        calc(ds, SUB_END_ROW, 2+i, f"='{S}'!$B${ROW_SRC_SUB}", USD1, bold=True)
        calc(ds, SUB_INT_ROW, 2+i, 0, USD1)
    else:
        calc(ds, SUB_BEG_ROW, 2+i, f"={prev}{SUB_END_ROW}", USD1)
        remaining_sweep = f"MAX({col}{DS_SWEEP_AVAIL_ROW}+{col}{TL_SWEEP_ROW},0)"
        calc(ds, SUB_SWEEP_ROW, 2+i, f"=-MIN({remaining_sweep},{col}{SUB_BEG_ROW})", USD1)
        calc(ds, SUB_END_ROW, 2+i, f"={col}{SUB_BEG_ROW}+{col}{SUB_SWEEP_ROW}", USD1, bold=True)
        calc(ds, SUB_INT_ROW, 2+i, f"={col}{SUB_BEG_ROW}*{A}!$B${ROW_SUB_RATE}", USD1)

# Revolver (funds shortfall if CFADS+beginning cash < 0, else stays 0)
for i in range(6):
    col = get_column_letter(2+i); prev = get_column_letter(2+i-1)
    if i == 0:
        calc(ds, REV_BEG_ROW, 2+i, 0, USD1)
        calc(ds, REV_DRAW_ROW, 2+i, f"='{S}'!$B${ROW_SRC_REV}", USD1)
        calc(ds, REV_END_ROW, 2+i, f"={col}{REV_BEG_ROW}+{col}{REV_DRAW_ROW}", USD1, bold=True)
        calc(ds, REV_INT_ROW, 2+i, 0, USD1)
    else:
        calc(ds, REV_BEG_ROW, 2+i, f"={prev}{REV_END_ROW}", USD1)
        shortfall = f"MIN({col}{DS_CFADS_ROW}+{col}{DS_MAND_ROW}+{col}{TL_SWEEP_ROW}+{col}{SUB_SWEEP_ROW},0)"
        calc(ds, REV_DRAW_ROW, 2+i, f"=-({shortfall})", USD1)
        calc(ds, REV_END_ROW, 2+i, f"=MAX({col}{REV_BEG_ROW}+{col}{REV_DRAW_ROW},0)", USD1, bold=True)
        calc(ds, REV_INT_ROW, 2+i, f"={col}{REV_BEG_ROW}*{A}!$B${ROW_REV_RATE}", USD1)

# Totals
for i in range(6):
    col = get_column_letter(2+i)
    calc(ds, INT_TOTAL_ROW, 2+i, f"={col}{TL_INT_ROW}+{col}{SUB_INT_ROW}+{col}{REV_INT_ROW}", USD1, bold=True)
    calc(ds, TOTAL_DEBT_ROW, 2+i, f"={col}{TL_END_ROW}+{col}{SUB_END_ROW}+{col}{REV_END_ROW}", USD1, bold=True)
    if i == 0:
        calc(ds, CASH_ROW, 2+i, f"={A}!$B${ROW_MIN_CASH}", USD1)
    else:
        prev = get_column_letter(2+i-1)
        excess_cash = f"MAX({col}{DS_CFADS_ROW}+{col}{DS_MAND_ROW}+{col}{TL_SWEEP_ROW}+{col}{SUB_SWEEP_ROW},0)*(1-{A}!$B${ROW_SWEEP})"
        calc(ds, CASH_ROW, 2+i, f"={prev}{CASH_ROW}+{excess_cash}", USD1)
    calc(ds, NET_DEBT_ROW, 2+i, f"={col}{TOTAL_DEBT_ROW}-{col}{CASH_ROW}", USD1, bold=True)
    calc(ds, LEV_ROW, 2+i, f"={col}{NET_DEBT_ROW}/'{O}'!{col}{EBITDA_ROW}", MULT)

D = "Debt Schedule"

# now go back and link Operating Model interest row to Debt Schedule total interest row
for i in range(1, 6):
    col = get_column_letter(2+i)
    om.cell(row=INT_ROW, column=2+i, value=f"=-'{D}'!{col}{INT_TOTAL_ROW}")
    om.cell(row=INT_ROW, column=2+i).number_format = USD1

# =========================================================
# SHEET 6: RETURNS
# =========================================================
rt = wb.create_sheet("Returns")
set_widths(rt, [32, 16, 16])
title(rt, "Returns Analysis", 1)
r = 3
hold_col_letter = None  # exit year column = 2 + hold (row offset in Operating Model columns, col index 2..7 => B..G, hold=5 -> col index 7 -> G)

label(rt, r, 1, "Exit Year", bold=True)
calc(rt, r, 2, f"='{O}'!G{YEAR_ROW}", '0', bold=True)
r += 1
label(rt, r, 1, "Exit EBITDA")
ROW_EXIT_EBITDA = r
calc(rt, r, 2, f"='{O}'!G{EBITDA_ROW}", USD1)
r += 1
label(rt, r, 1, "Exit EV/EBITDA Multiple")
calc(rt, r, 2, f"={A}!$B${ROW_EXIT_MULT}", MULT)
r += 1
label(rt, r, 1, "Exit Enterprise Value", bold=True)
ROW_EXIT_EV = r
calc(rt, r, 2, f"=B{ROW_EXIT_EBITDA}*{A}!$B${ROW_EXIT_MULT}", USD1, bold=True)
r += 1
label(rt, r, 1, "Less: Net Debt at Exit")
ROW_EXIT_NETDEBT = r
calc(rt, r, 2, f"=-'{D}'!G{NET_DEBT_ROW}", USD1)
r += 1
label(rt, r, 1, "Exit Equity Value", bold=True)
ROW_EXIT_EQUITY = r
calc(rt, r, 2, f"=B{ROW_EXIT_EV}+B{ROW_EXIT_NETDEBT}", USD1, bold=True)
for c in (1,2):
    rt.cell(row=r, column=c).fill = TOTAL_FILL
r += 2

section(rt, "Sponsor Returns", r); r += 1
label(rt, r, 1, "Initial Sponsor Equity")
ROW_INIT_EQ = r
calc(rt, r, 2, f"='{S}'!B${ROW_SRC_SPONSOR}", USD1)
r += 1
label(rt, r, 1, "Exit Sponsor Equity Proceeds")
ROW_EXIT_EQ2 = r
calc(rt, r, 2, f"=B{ROW_EXIT_EQUITY}*(1-{A}!$B${ROW_MGMT})", USD1)
r += 1
label(rt, r, 1, "Hold Period (years)")
ROW_HOLD2 = r
calc(rt, r, 2, f"={A}!$B${ROW_HOLD}", '0')
r += 1
label(rt, r, 1, "MOIC", bold=True)
ROW_MOIC = r
calc(rt, r, 2, f"=B{ROW_EXIT_EQ2}/B{ROW_INIT_EQ}", MULT, bold=True)
r += 1
label(rt, r, 1, "IRR", bold=True)
ROW_IRR = r
calc(rt, r, 2, f"=(B{ROW_MOIC})^(1/B{ROW_HOLD2})-1", PCT1, bold=True)
for c in (1,2):
    rt.cell(row=r, column=c).fill = TOTAL_FILL
r += 2

section(rt, "Cash Flow Check (for XIRR)", r); r += 1
label(rt, r, 1, "Year 0 (Equity Out)")
label(rt, r+1, 1, "Years 1-4")
label(rt, r+2, 1, "Exit Year (Equity In)")
calc(rt, r, 2, f"=-B{ROW_INIT_EQ}", USD1)
calc(rt, r+1, 2, 0, USD1)
calc(rt, r+2, 2, f"=B{ROW_EXIT_EQ2}", USD1)
r += 4

label(rt, r, 1, "Note: exit assumed at end of Year 5 (col G). Adjust Hold Period + column refs on this tab if you change the hold period.", )
rt.cell(row=r, column=1).font = Font(italic=True, size=9, color="808080")
rt.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)

# =========================================================
# Freeze panes / gridlines cleanup
# =========================================================
for sheet in [ws, su, om, fc, ds, rt]:
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "B4" if sheet != ws else "A1"

import os
out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "LBO_Model.xlsx")
wb.save(out_path)
print(f"saved to {out_path}")
