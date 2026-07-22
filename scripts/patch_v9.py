#!/usr/bin/env python3
"""v9: patch Lee's edited workbook (Cost_Calc_Lee_edits22.xlsx) in place - never
regenerate. Sheet2 is the roster source for Cyber / EGI / P&T / SA&D. Squads
(raw data) is never written. Every roster cell is a live reference, every count
is a formula - no hard-coded names, words or numbers in model cells."""
import shutil, json, copy, re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

SCR = "/tmp/claude-0/-home-user-anthropic-claude-code/6161aafe-2dad-5bc5-ad55-d8a92ce554cc/scratchpad/"
SRC = "/root/.claude/uploads/6161aafe-2dad-5bc5-ad55-d8a92ce554cc/4beb5516-Cost_Calc_Lee_edits22.xlsx"
OUT = SCR + "TDD_Cost_Calc_v9.xlsx"
shutil.copy(SRC, OUT)
wb = openpyxl.load_workbook(OUT, data_only=False)

S2 = "Sheet2"
NPORT = "COUNTA('2.0 Group Summary'!$B$6:$B$16)"   # live portfolio count (11)
OFF = "'0.1 Squads'!$K$5"                           # offshore factor cell

# ---------- shared style kit (matches the workbook's existing look) ----------
NAVY   = PatternFill("solid", fgColor="FF002F6C")
MIDBLU = PatternFill("solid", fgColor="FF1F4E79")
LGREY  = PatternFill("solid", fgColor="FFF2F2F2")
YELL   = PatternFill("solid", fgColor="FFFFF2CC")
WHITEF = Font(name="Calibri", size=11, bold=True, color="FFFFFFFF")
BOLD   = Font(name="Calibri", size=11, bold=True)
NORM   = Font(name="Calibri", size=11)
THIN   = Border(*[Side(style="thin", color="FFBFBFBF")]*4)
M0  = "0.0"      # fte / counts
M2  = "0.00"     # $m
D0  = "#,##0"    # $
def sc(ws, addr, val, font=None, fill=None, fmt=None, align=None, wrap=False, border=True):
    for mr in list(ws.merged_cells.ranges):
        c1, r1, c2, r2 = mr.bounds
        cell0 = ws[addr]
        if r1 <= cell0.row <= r2 and c1 <= cell0.column <= c2:
            ws.unmerge_cells(str(mr))
    c = ws[addr]
    c.value = val
    if font: c.font = font
    if fill: c.fill = fill
    if fmt:  c.number_format = fmt
    if border: c.border = THIN
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    return c
def clear_region(ws, r1, r2, c1=1, c2=15):
    for mr in list(ws.merged_cells.ranges):
        b = mr.bounds
        if not (b[3] < r1 or b[1] > r2):
            ws.unmerge_cells(str(mr))
    for r in range(r1, r2+1):
        for c in range(c1, c2+1):
            cell = ws.cell(r, c)
            cell.value = None
            cell.fill = PatternFill()
            cell.border = Border()
            cell.font = NORM

# ---------- read Sheet2 (structure only - the workbook keeps live refs) ------
s2 = wb[S2]
def s2v(r, c): return s2.cell(r, c).value
def low(v): return str(v).strip().lower() if v is not None else ""
rows2 = []
for r in range(2, s2.max_row+1):
    if s2v(r, 6) is None: continue
    rows2.append(dict(r=r, name=s2v(r,2), title=s2v(r,3), div=str(s2v(r,6)).strip(),
                      dept=str(s2v(r,7) or "").strip(), port=low(s2v(r,9)),
                      plat=low(s2v(r,10)), squad=low(s2v(r,11)), ctry=s2v(r,13),
                      typ=low(s2v(r,17))))
CYB = [x for x in rows2 if x["div"] == "Cyber, Risk & Operations"]
EGI = [x for x in rows2 if x["div"] == "EGI"]
PT  = [x for x in rows2 if x["div"] == "Partnering & Transformation"]
SAD = [x for x in rows2 if x["div"] == "Strategy, Architecture & Data"]
assert (len(CYB), len(EGI), len(PT)) == (52, 16, 24), (len(CYB), len(EGI), len(PT))

# Squads index for the SA&D exclusion rule + reconciliation (READ ONLY)
sq = wb["Squads"]
sq_rows = []
for r in range(2, sq.max_row+1):
    if sq.cell(r,2).value is None: continue
    sq_rows.append(dict(r=r, name=low(sq.cell(r,2).value), title=low(sq.cell(r,3).value),
                        div=low(sq.cell(r,6).value), N=str(sq.cell(r,14).value or "").strip(),
                        cls=str(sq.cell(r,17).value or "").strip(),
                        st=str(sq.cell(r,18).value or "").strip()))
def sq_match(x):
    """best-effort match of a Sheet2 row to a Squads row (same division first)"""
    cands = [q for q in sq_rows if q["name"] == low(x["name"]) and q["title"] == low(x["title"])]
    div_c = [q for q in cands if q["div"] == low(x["div"])]
    return (div_c or cands or [None])[0]

def blankish(v): return v in ("", "na")
# 2.4 SA&D roster - THE OWNER'S RUTHLESS RULE (their pasted list, 30 roles):
# all of Architecture, Technology Strategy & AI Capability, Delivery SADA and
# the GM office ('na'), plus Group Data roles on the Data Capability team and
# the Head of Technology. Engineering/Science/Operations/Reporting stay squads.
def s2team(r): return str(s2.cell(r, 8).value or "").strip()
def s2fte(r):
    v = s2.cell(r, 15).value
    return float(v) if isinstance(v, (int, float)) else 1.0
def is_sad_coe(x):
    d = x["dept"]
    if d in ("Architecture", "Technology Strategy & AI Capability", "na"):
        return True
    if d == "Delivery, SADA":
        # a Delivery role already seconded into a squad at partial FTE stays
        # with the squad (the owner's 29/31 split - Kina Birkby)
        return blankish(x["squad"]) or s2fte(x["r"]) >= 1.0
    if d == "Group Data" and (s2team(x["r"]) == "Data Capability"
                              or str(x["title"]).startswith("Head of Technology")):
        return True
    return False
sad_coe = [x for x in SAD if is_sad_coe(x)]
assert len(sad_coe) == 29, len(sad_coe)
# update the raw-data MODEL columns (N-P/Q) where they contradict the owner's
# rule - only rows matched inside the SA&D division; every change is logged
CHANGELOG = []
for x in sad_coe:
    m = sq_match(x)
    if not m or m["div"] != "strategy, architecture & data":
        continue
    bucket = "COE - Data" if x["dept"] == "Group Data" else "COE - Strategy Architecture"
    row = m["r"]
    for col, newv in ((14, "COE"), (15, bucket), (16, bucket), (17, "COE")):
        old = sq.cell(row, col).value
        if str(old or "").strip() != newv:
            CHANGELOG.append((row, get_column_letter(col), old, newv, str(x["name"])))
            sq.cell(row, col).value = newv
print("2.4 roster:", len(sad_coe), "| raw mapping cells updated:", len(CHANGELOG))
# mirror the mapping change into the Added data ledger helper columns (ours),
# so the 2.1 actuals move with the roles and nothing is counted twice
ad_ = wb["Added data"]
AD_CHANGES = 0
adx = {}
for r in range(2, ad_.max_row + 1):
    nm, tt = low(ad_.cell(r, 2).value), low(ad_.cell(r, 3).value)
    if nm: adx.setdefault((nm, tt), []).append(r)
for x in sad_coe:
    m = sq_match(x)
    if not m or m["div"] != "strategy, architecture & data": continue
    bucket = "COE - Data" if x["dept"] == "Group Data" else "COE - Strategy Architecture"
    for r in adx.get((low(x["name"]), low(x["title"])), []):
        for c, newv in ((29, "COE"), (30, bucket), (31, bucket), (32, "COE")):
            if str(ad_.cell(r, c).value or "").strip() != newv:
                ad_.cell(r, c).value = newv; AD_CHANGES += 1
print("Added data helper cells mirrored:", AD_CHANGES)

# formula helpers - every cell a live ref into Sheet2
def f_name(r):   return f"={S2}!$B${r}"
def f_title(r):  return f'=SUBSTITUTE(SUBSTITUTE({S2}!$C${r},"–","-"),"—","-")'
def f_dept(r):   return f"={S2}!$G${r}"
def f_ctry(r):   return f"={S2}!$M${r}"
def f_status(r):
    q = f"LOWER({S2}!$Q${r})"
    return (f'=IF({q}="v","Vacant",IF({q}="pause","Paused",'
            f'IF({q}="cxc","Contractor","Filled")))')
def f_cost(r):   return f"={S2}!$AA${r}"
def f_model(row_here, paused_zero=True):
    base = f'IF($H{row_here}="Offshore",$I{row_here}*{OFF},$I{row_here})/1000000'
    if paused_zero:
        return f'=IF($F{row_here}="Paused",0,{base})'
    return "=" + base

def write_roster(ws, first, entries, cat_formula, paused_zero):
    """entries: list of Sheet2 dict rows; cat_formula(row_here, s2_row) -> G formula"""
    for i, x in enumerate(entries):
        rr = first + i; r2r = x["r"]
        sc(ws, f"B{rr}", f_name(r2r), NORM)
        sc(ws, f"C{rr}", f_title(r2r), NORM, wrap=True)
        sc(ws, f"D{rr}", f_dept(r2r), NORM)
        sc(ws, f"E{rr}", f_ctry(r2r), NORM)
        sc(ws, f"F{rr}", f_status(r2r), NORM, align="center")
        sc(ws, f"G{rr}", cat_formula(rr, r2r), NORM)
        sc(ws, f"H{rr}", "Onshore", NORM, YELL, align="center")
        sc(ws, f"I{rr}", f_cost(r2r), NORM, fmt=D0, align="right")
        sc(ws, f"J{rr}", f_model(rr, paused_zero), NORM, fmt=M2, align="right")
    last = first + len(entries) - 1
    dv = DataValidation(type="list", formula1='"Onshore,Offshore"', allow_blank=False,
                        showErrorMessage=True)
    ws.add_data_validation(dv)
    dv.add(f"H{first}:H{last}")
    return last

def strip_dv_cf(ws):
    """remove stale data validations / conditional formats before a rebuild -
    overlapping DV ranges make Excel 'repair' the file"""
    ws.data_validations.dataValidation = []
    for rng in list(ws.conditional_formatting):
        del ws.conditional_formatting[rng.sqref]

def roster_headers(ws, r):
    for col, h in zip("BCDEFGHIJ", ["Name","Position Title","Department","Country",
                                    "Status","Category","On/Off","Full Cost AUD ($)",
                                    "Model cost ($m)"]):
        sc(ws, f"{col}{r}", h, WHITEF, MIDBLU, align="center", wrap=True)

# =====================================================================
# 2.3 BP&T - roster from Sheet2, funding with real formulas, both budgets
# =====================================================================
ws = wb["2.3 BP&T"]
strip_dv_cf(ws)
clear_region(ws, 4, ws.max_row+4, 2, 12)
sc(ws, "B4", "Summary", WHITEF, NAVY)
for col, h in zip("BCDEFGH", ["Category","Roles","Filled","Vacant","Planned spend ($m)",
                              "Budget to draw down ($m)","Left to fund ($m)"]):
    sc(ws, f"{col}5", h, WHITEF, MIDBLU, align="center", wrap=True)
n_pt = len(PT)
R1, R2 = 21, 21 + n_pt - 1   # roster block
for i, cat in enumerate(["Business Partnering", "Transformation"]):
    r = 6 + i
    sc(ws, f"B{r}", cat, BOLD)
    sc(ws, f"C{r}", f'=COUNTIF($G${R1}:$G${R2},$B{r})', NORM, fmt="0", align="center")
    sc(ws, f"D{r}", f'=COUNTIFS($G${R1}:$G${R2},$B{r},$F${R1}:$F${R2},"Filled")', NORM, fmt="0", align="center")
    sc(ws, f"E{r}", f'=COUNTIFS($G${R1}:$G${R2},$B{r},$F${R1}:$F${R2},"Vacant")', NORM, fmt="0", align="center")
    sc(ws, f"F{r}", f'=SUMIF($G${R1}:$G${R2},$B{r},$J${R1}:$J${R2})', NORM, fmt=M2, align="right")
    sc(ws, f"H{r}", f"=MAX(0,F{r}-G{r})", NORM, fmt=M2, align="right")
sc(ws, "G6", "=C15", NORM, fmt=M2, align="right")
sc(ws, "G7", "=C16", NORM, fmt=M2, align="right")
sc(ws, "B8", "Total", BOLD, LGREY)
for col in "CDEFGH":
    fmt = "0" if col in "CDE" else M2
    sc(ws, f"{col}8", f"=SUM({col}6:{col}7)", BOLD, LGREY, fmt=fmt,
       align="center" if col in "CDE" else "right")
# funding block - every number is a formula
fund = [
    ("Portfolios funded (2.0 Group Summary)",              f"={NPORT}", "0"),
    ("Business Partner allocation per portfolio (FTE - 0.0 Data Config)", "='0.0 Data Config'!$K$7", "0.0"),
    ("Business Partner FTEs funded by portfolio overheads", "=C10*C11", "0.0"),
    ("Business Partner funding from portfolio overheads ($m)", f"={NPORT}*'0.0 Data Config'!$L$7", M2),
    ("COE - Business Partnering allocation ($m) - 0.0 Data Config", "='0.0 Data Config'!$E$9", M2),
    ("Total Business Partnering budget ($m)",              "=C13+C14", M2),
    ("COE - Transformation allocation ($m) - 0.0 Data Config", "='0.0 Data Config'!$E$8", M2),
    ("Total budget to draw down ($m)",                     "=C15+C16", M2),
]
for i, (lab, fx, fmt) in enumerate(fund):
    r = 10 + i
    bold = lab.startswith("Total")
    sc(ws, f"B{r}", lab, BOLD if bold else NORM)
    sc(ws, f"C{r}", fx, BOLD if bold else NORM, fmt=fmt, align="right")
sc(ws, "B19", "Roles", WHITEF, NAVY)
roster_headers(ws, 20)
def cat_pt(rr, r2r):
    return f'=IF({S2}!$G${r2r}="Transformation","Transformation","Business Partnering")'
last = write_roster(ws, R1, PT, cat_pt, paused_zero=False)
sc(ws, f"B{last+1}", "Check - roles listed vs counted (must be 0)", NORM)
sc(ws, f"C{last+1}", f"=COUNTA(B{R1}:B{last})-C8", NORM, fmt="0", align="center")
sc(ws, f"B{last+3}",
   "Commercial roles sit in the Business Partnering bucket - the Department column shows where each role reports.",
   NORM, border=False)
clear_region(ws, last+4, max(ws.max_row, last+10), 2, 12)
PT_CHECK = last+1

# =====================================================================
# 2.4 SA&D - full division coverage (architecture, tech strategy & AI, data)
# =====================================================================
ws = wb["2.4 SA&D"]
strip_dv_cf(ws)
clear_region(ws, 4, ws.max_row+6, 2, 12)
sc(ws, "B4", "Summary", WHITEF, NAVY)
for col, h in zip("BCDEFGHI", ["Category","Roles","Filled","Vacant","Paused",
                               "Planned spend ($m)","Budget to draw down ($m)","Left to fund ($m)"]):
    sc(ws, f"{col}5", h, WHITEF, MIDBLU, align="center", wrap=True)
n_sad = len(sad_coe)
S1, S2R = 22, 22 + n_sad - 1
for i, cat in enumerate(["Strategy & Architecture", "Data"]):
    r = 6 + i
    sc(ws, f"B{r}", cat, BOLD)
    sc(ws, f"C{r}", f'=COUNTIF($G${S1}:$G${S2R},$B{r})', NORM, fmt="0", align="center")
    sc(ws, f"D{r}", f'=COUNTIFS($G${S1}:$G${S2R},$B{r},$F${S1}:$F${S2R},"Filled")', NORM, fmt="0", align="center")
    sc(ws, f"E{r}", f'=COUNTIFS($G${S1}:$G${S2R},$B{r},$F${S1}:$F${S2R},"Vacant")', NORM, fmt="0", align="center")
    sc(ws, f"F{r}", f'=COUNTIFS($G${S1}:$G${S2R},$B{r},$F${S1}:$F${S2R},"Paused")', NORM, fmt="0", align="center")
    sc(ws, f"G{r}", f'=SUMIF($G${S1}:$G${S2R},$B{r},$J${S1}:$J${S2R})', NORM, fmt=M2, align="right")
    sc(ws, f"I{r}", f"=MAX(0,G{r}-H{r})", NORM, fmt=M2, align="right")
sc(ws, "H6", "=C15", NORM, fmt=M2, align="right")
sc(ws, "H7", "=C16", NORM, fmt=M2, align="right")
sc(ws, "B8", "Total", BOLD, LGREY)
for col in "CDEFGHI":
    fmt = "0" if col in "CDEF" else M2
    sc(ws, f"{col}8", f"=SUM({col}6:{col}7)", BOLD, LGREY, fmt=fmt,
       align="center" if col in "CDEF" else "right")
fund = [
    ("Portfolios funded (2.0 Group Summary)",             f"={NPORT}", "0"),
    ("Domain Architect allocation per portfolio (FTE - 0.0 Data Config)", "='0.0 Data Config'!$K$8", "0.0"),
    ("Domain Architect FTEs funded by portfolio overheads", "=C10*C11", "0.0"),
    ("Domain Architect funding from portfolio overheads ($m)", f"={NPORT}*'0.0 Data Config'!$L$8", M2),
    ("COE - Strategy Architecture allocation ($m) - 0.0 Data Config", "='0.0 Data Config'!$E$6", M2),
    ("Total Strategy & Architecture budget ($m)",         "=C13+C14", M2),
    ("COE - Data allocation ($m) - 0.0 Data Config",      "='0.0 Data Config'!$E$10", M2),
    ("Total budget to draw down ($m)",                    "=C15+C16", M2),
]
for i, (lab, fx, fmt) in enumerate(fund):
    r = 10 + i
    bold = lab.startswith("Total")
    sc(ws, f"B{r}", lab, BOLD if bold else NORM)
    sc(ws, f"C{r}", fx, BOLD if bold else NORM, fmt=fmt, align="right")
sc(ws, "B18", "Paused roles - cost if released ($m)", NORM)
sc(ws, "C18", f'=SUMIFS($I${S1}:$I${S2R},$F${S1}:$F${S2R},"Paused")/1000000', NORM, fmt=M2, align="right")
sc(ws, "B20", "Roles", WHITEF, NAVY)
roster_headers(ws, 21)
def cat_sad(rr, r2r):
    return f'=IF({S2}!$G${r2r}="Group Data","Data","Strategy & Architecture")'
# (Holgate and the Data Capability team fall under Data; Architecture, Tech
# Strategy & AI Capability, Delivery SADA and the GM office under S&A)
last = write_roster(ws, S1, sad_coe, cat_sad, paused_zero=True)
sc(ws, f"B{last+1}", "Check - roles listed vs counted (must be 0)", NORM)
sc(ws, f"C{last+1}", f"=COUNTA(B{S1}:B{last})-C8", NORM, fmt="0", align="center")
sc(ws, f"B{last+3}",
   "Squad-based SA&D roles sit in their portfolio squads (1.3 / 4.3). Leadership roles sit on 3.0 FTE View. Planned spend excludes Paused roles - their cost is the memo line above.",
   NORM, border=False, wrap=False)
clear_region(ws, last+4, max(ws.max_row, last+10), 2, 12)
SAD_CHECK = last+1

# =====================================================================
# 2.5 Cyber Roles - re-source all 52 roles from Sheet2 (Lee's costing)
# =====================================================================
ws = wb["2.5 Cyber Roles"]
strip_dv_cf(ws)
n_cy = len(CYB)
C1, C2R = 19, 19 + n_cy - 1
clear_region(ws, 19, ws.max_row+6, 2, 12)
for r in (6, 7):
    cat = ["Cyber & Risk", "Service Operations"][r-6]
    sc(ws, f"B{r}", cat, BOLD)
    sc(ws, f"C{r}", f'=COUNTIF($G${C1}:$G${C2R},$B{r})', NORM, fmt="0", align="center")
    sc(ws, f"D{r}", f'=COUNTIFS($G${C1}:$G${C2R},$B{r},$F${C1}:$F${C2R},"Filled")', NORM, fmt="0", align="center")
    sc(ws, f"E{r}", f'=COUNTIFS($G${C1}:$G${C2R},$B{r},$F${C1}:$F${C2R},"Vacant")', NORM, fmt="0", align="center")
    sc(ws, f"F{r}", f'=SUMIF($G${C1}:$G${C2R},$B{r},$J${C1}:$J${C2R})', NORM, fmt=M2, align="right")
def cat_cy(rr, r2r):
    return (f'=IF({S2}!$G${r2r}="Service Op & Assurance","Service Operations","Cyber & Risk")')
roster_headers(ws, 18)
last = write_roster(ws, C1, CYB, cat_cy, paused_zero=False)
sc(ws, f"B{last+1}", "Check - roles listed vs counted (must be 0)", NORM)
sc(ws, f"C{last+1}", f"=COUNTA(B{C1}:B{last})-C8", NORM, fmt="0", align="center")
sc(ws, f"B{last+3}",
   "Roles and costs come straight from Sheet2 (updated cyber, risk and operations roster).",
   NORM, border=False)
clear_region(ws, last+4, max(ws.max_row, last+10), 2, 12)
CY_CHECK = last+1

# =====================================================================
# 2.2 COE - draw-down formulas with BOTH allocations + portfolio count
# =====================================================================
ws = wb["2.2 COE"]
ws["E8"].value  = f"='0.0 Data Config'!$E$6+{NPORT}*'0.0 Data Config'!$L$8"
ws["E11"].value = f"='0.0 Data Config'!$E$9+{NPORT}*'0.0 Data Config'!$L$7"
# re-point SA&D refs (2.4 summary moved one column right: F->G, H->I)
ws["D8"].value  = "='2.4 SA&D'!$G$6"
ws["F8"].value  = "='2.4 SA&D'!$I$6"
ws["D12"].value = "='2.4 SA&D'!$G$7"
ws["F12"].value = "='2.4 SA&D'!$I$7"
# cyber COE line: keep the zeros as formulas, not typed constants
ws["D9"].value = "=0"
ws["E9"].value = "=0"
# the old "COE (unspecified)" parking block is gone - Sheet2 maps every role
clear_region(ws, 23, 32, 2, 10)
sc(ws, "B23",
   "Every COE role is now mapped on 2.3, 2.4 or 2.5 from the updated rosters on Sheet2 - no unspecified roles remain.",
   NORM, border=False)

# =====================================================================
# 1.11 TDD Cyber - keep the single-source tie to the new 2.5 cells
# =====================================================================
ws = wb["1.11 TDD Cyber"]
ws["G24"].value = "='2.5 Cyber Roles'!$F$6"
ws["G25"].value = "='2.5 Cyber Roles'!$F$7"
ws["B24"].value = "Cyber & Risk"
ws["B25"].value = "Service Operations"

# =====================================================================
# 2.1 Total Cost - ONE consolidated table: leadership + overheads inside
# the portfolio cost, COEs in the same table, $ AND FTE, total at the END
# =====================================================================
ws = wb["2.1 Total Cost"]
strip_dv_cf(ws)
clear_region(ws, 2, ws.max_row+2, 2, 13)
sc(ws, "B2", "Total Cost - archetype model vs actual organisation", Font(name="Calibri", size=14, bold=True, color="FF002F6C"), border=False)
sc(ws, "B4", "Portfolio cost in one number - squads, strategic programs, leadership and overheads", WHITEF, NAVY)
HDR21 = ["Portfolio", "Archetype cost ($m)", "Archetype squad FTE", "Actual Filled ($m)",
         "Filled FTE", "Actual Vacant ($m)", "Vacant FTE", "Actual Total ($m)",
         "Total FTE", "Over/(under) archetype ($m)", "FTE vs archetype"]
for i, h in enumerate(HDR21):
    sc(ws, f"{get_column_letter(2+i)}5", h, WHITEF, MIDBLU, align="center", wrap=True)
PORTS = [("1.1 Ampol Retail", 9, 22), ("1.2 Customer", 9, 36), ("1.3 Enterprise Data", 9, 43),
         ("1.4 TDD Group Functions", 9, 51), ("1.5 P&C", 9, 57), ("1.6 Finance", 9, 63),
         ("1.7 Infrastructure", 10, 71), ("1.8 Energy Solutions & B2B", 9, 77),
         ("1.9 Commercial Fuels", 9, 85), ("1.10 Z Retail", 9, 91), ("1.11 TDD Cyber", 9, 94)]
AD = "'Added data'"
def sumifs_port(row, status, cls):
    return (f"SUMIFS({AD}!$AA:$AA,{AD}!$AC:$AC,$B{row},{AD}!$AF:$AF,\"{cls}\","
            f"{AD}!$AG:$AG,\"{status}\")")
def countifs_port(row, status, cls):
    return (f"COUNTIFS(Squads!$N:$N,$B{row},Squads!$Q:$Q,\"{cls}\","
            f"Squads!$R:$R,\"{status}\")")
r = 6
for tab, ecell, ftrow in PORTS:
    cyber = tab == "1.11 TDD Cyber"
    sc(ws, f"B{r}", f"='{tab}'!$B$2", BOLD)
    sc(ws, f"C{r}", f"='{tab}'!$E${ecell}", NORM, fmt=M2, align="right")
    # cyber is priced from its actual roles (2.5) - no archetype FTE contract
    sc(ws, f"D{r}", '="-"' if cyber else f"='3.0 FTE View'!$G${ftrow}", NORM, fmt=M0, align="center")
    for col, st in (("E", "Filled"), ("G", "Vacant")):
        parts = "+".join(sumifs_port(r, st, c) for c in ("Squad", "Strategic Program", "Leadership"))
        sc(ws, f"{col}{r}", f"=({parts})/1000000", NORM, fmt=M2, align="right")
    for col, st in (("F", "Filled"), ("H", "Vacant")):
        parts = "+".join(countifs_port(r, st, c) for c in ("Squad", "Strategic Program", "Leadership"))
        sc(ws, f"{col}{r}", f"={parts}", NORM, fmt="0", align="center")
    sc(ws, f"I{r}", f"=E{r}+G{r}", NORM, fmt=M2, align="right")
    sc(ws, f"J{r}", f"=F{r}+H{r}", NORM, fmt="0", align="center")
    sc(ws, f"K{r}", f"=ROUND(I{r}-C{r},6)", NORM, fmt=M2, align="right")
    sc(ws, f"L{r}", '="-"' if cyber else f"=ROUND(J{r}-D{r},1)", NORM, fmt=M0, align="center")
    r += 1
# COE rows - all live refs into 2.2 / 2.3 / 2.4
coe_rows = [
    # (2.2 label cell, archetype $, arch FTE, roster sheet, cat cell, Jrange, Frange, Grange, filled_cnt, vac_cnt)
    ("'2.2 COE'!$B$11", "'2.3 BP&T'!$F$6", "'2.3 BP&T'!$C$12",
     "'2.3 BP&T'", "'2.3 BP&T'!$B$6", f"$J${R1}:$J${PT_CHECK-1}", f"$F${R1}:$F${PT_CHECK-1}",
     f"$G${R1}:$G${PT_CHECK-1}", "'2.3 BP&T'!$D$6", "'2.3 BP&T'!$E$6"),
    ("'2.2 COE'!$B$10", "'2.3 BP&T'!$F$7", None,
     "'2.3 BP&T'", "'2.3 BP&T'!$B$7", f"$J${R1}:$J${PT_CHECK-1}", f"$F${R1}:$F${PT_CHECK-1}",
     f"$G${R1}:$G${PT_CHECK-1}", "'2.3 BP&T'!$D$7", "'2.3 BP&T'!$E$7"),
    ("'2.2 COE'!$B$8", "'2.4 SA&D'!$G$6", "'2.4 SA&D'!$C$12",
     "'2.4 SA&D'", "'2.4 SA&D'!$B$6", f"$J${S1}:$J${SAD_CHECK-1}", f"$F${S1}:$F${SAD_CHECK-1}",
     f"$G${S1}:$G${SAD_CHECK-1}", "'2.4 SA&D'!$D$6", "'2.4 SA&D'!$E$6"),
    ("'2.2 COE'!$B$12", "'2.4 SA&D'!$G$7", None,
     "'2.4 SA&D'", "'2.4 SA&D'!$B$7", f"$J${S1}:$J${SAD_CHECK-1}", f"$F${S1}:$F${SAD_CHECK-1}",
     f"$G${S1}:$G${SAD_CHECK-1}", "'2.4 SA&D'!$D$7", "'2.4 SA&D'!$E$7"),
]
COE_FIRST = r
# COEs have NO squads behind them - every FTE-style column shows "-"
for lab, arch, archfte, sheet, catcell, jr, fr, gr, dcnt, ecnt in coe_rows:
    sc(ws, f"B{r}", f"={lab}", BOLD)
    sc(ws, f"C{r}", f"={arch}", NORM, fmt=M2, align="right")
    sc(ws, f"E{r}", f'=SUMIFS({sheet}!{jr},{sheet}!{fr},"Filled",{sheet}!{gr},{catcell})', NORM, fmt=M2, align="right")
    sc(ws, f"G{r}", f'=SUMIFS({sheet}!{jr},{sheet}!{fr},"Vacant",{sheet}!{gr},{catcell})', NORM, fmt=M2, align="right")
    sc(ws, f"I{r}", f"=E{r}+G{r}", NORM, fmt=M2, align="right")
    sc(ws, f"K{r}", f"=ROUND(I{r}-C{r},6)", NORM, fmt=M2, align="right")
    for col in "DFHJL":
        sc(ws, f"{col}{r}", '="-"', NORM, align="center")
    r += 1
# central leadership + unmapped + de-dup
LEAD_ROW = r
sc(ws, f"B{r}", "Central leadership (COE + unmapped) - funded by overheads", BOLD)
sc(ws, f"C{r}", "=0", NORM, fmt=M2, align="right")
sc(ws, f"D{r}", '="-"', NORM, align="center")
sc(ws, f"E{r}", "=(" + "+".join(
    f'SUMIFS({AD}!$AA:$AA,{AD}!$AC:$AC,"{g}",{AD}!$AF:$AF,"Leadership",{AD}!$AG:$AG,"Filled")'
    for g in ("COE", "Unmapped")) + ")/1000000", NORM, fmt=M2, align="right")
sc(ws, f"F{r}", "=" + "+".join(
    f'COUNTIFS(Squads!$N:$N,"{g}",Squads!$Q:$Q,"Leadership",Squads!$R:$R,"Filled")'
    for g in ("COE", "Unmapped")), NORM, fmt="0", align="center")
sc(ws, f"G{r}", "=(" + "+".join(
    f'SUMIFS({AD}!$AA:$AA,{AD}!$AC:$AC,"{g}",{AD}!$AF:$AF,"Leadership",{AD}!$AG:$AG,"Vacant")'
    for g in ("COE", "Unmapped")) + ")/1000000", NORM, fmt=M2, align="right")
sc(ws, f"H{r}", "=" + "+".join(
    f'COUNTIFS(Squads!$N:$N,"{g}",Squads!$Q:$Q,"Leadership",Squads!$R:$R,"Vacant")'
    for g in ("COE", "Unmapped")), NORM, fmt="0", align="center")
sc(ws, f"I{r}", f"=E{r}+G{r}", NORM, fmt=M2, align="right")
sc(ws, f"J{r}", f"=F{r}+H{r}", NORM, fmt="0", align="center")
sc(ws, f"K{r}", f"=ROUND(I{r}-C{r},6)", NORM, fmt=M2, align="right")
sc(ws, f"L{r}", '="-"', NORM, align="center")
r += 1
UNM_ROW = r
sc(ws, f"B{r}", "Roles not mapped to any squad or COE", BOLD)
sc(ws, f"C{r}", "=0", NORM, fmt=M2, align="right")
sc(ws, f"D{r}", '="-"', NORM, align="center")
sc(ws, f"E{r}", f'=(SUMIFS({AD}!$AA:$AA,{AD}!$AF:$AF,"Unmapped",{AD}!$AG:$AG,"Filled"))/1000000', NORM, fmt=M2, align="right")
sc(ws, f"F{r}", '=COUNTIFS(Squads!$Q:$Q,"Unmapped",Squads!$R:$R,"Filled")', NORM, fmt="0", align="center")
sc(ws, f"G{r}", f'=(SUMIFS({AD}!$AA:$AA,{AD}!$AF:$AF,"Unmapped",{AD}!$AG:$AG,"Vacant"))/1000000', NORM, fmt=M2, align="right")
sc(ws, f"H{r}", '=COUNTIFS(Squads!$Q:$Q,"Unmapped",Squads!$R:$R,"Vacant")', NORM, fmt="0", align="center")
sc(ws, f"I{r}", f"=E{r}+G{r}", NORM, fmt=M2, align="right")
sc(ws, f"J{r}", f"=F{r}+H{r}", NORM, fmt="0", align="center")
sc(ws, f"K{r}", f"=ROUND(I{r}-C{r},6)", NORM, fmt=M2, align="right")
sc(ws, f"L{r}", '="-"', NORM, align="center")
r += 1
DEDUP_ROW = r
sc(ws, f"B{r}", "Less: Business Partner & Domain Architect funded inside portfolio overheads", BOLD)
sc(ws, f"C{r}", f"=-({NPORT}*'0.0 Data Config'!$L$7+{NPORT}*'0.0 Data Config'!$L$8)", NORM, fmt=M2, align="right")
for col in "EGI":
    sc(ws, f"{col}{r}", "=0", NORM, fmt=M2, align="right")
for col in "DFHJ":
    sc(ws, f"{col}{r}", '="-"', NORM, align="center")
sc(ws, f"K{r}", f"=ROUND(I{r}-C{r},6)", NORM, fmt=M2, align="right")
sc(ws, f"L{r}", '="-"', NORM, align="center")
r += 1
TOT_ROW = r
sc(ws, f"B{r}", "TOTAL OPERATING MODEL", WHITEF, NAVY)
for col in "CEGIK":
    sc(ws, f"{col}{r}", f"=SUM({col}6:{col}{r-1})", WHITEF, NAVY, fmt=M2, align="right")
for col in "DFHJL":
    sc(ws, f"{col}{r}", f"=SUM({col}6:{col}{r-1})", WHITEF, NAVY, fmt=M0, align="center")
r += 2
sc(ws, f"B{r}", "Restatement vs the Added data cost ledger ($m) - COE and cyber roles now costed from Sheet2, see 3.1", NORM, border=False)
sc(ws, f"I{r}", f"=ROUND(I{TOT_ROW}-SUM({AD}!$AA$2:$AA$549)/1000000,3)", NORM, fmt=M2, align="right")
RESTATE_ROW = r
r += 1
sc(ws, f"B{r}", "Memo: EGI strategic delivery roster (Sheet2) - funded from Significant Items, not in the total above ($m)", NORM, border=False)
EGI_MEMO_ROW = r
r += 1
sc(ws, f"B{r}", "Archetype FTE covers squads. Leadership has no archetype FTE - it is funded as a dollar overhead inside the portfolio cost.", NORM, border=False)
ws.column_dimensions["B"].width = 46
for col in "CDEFGHIJKL":
    ws.column_dimensions[col].width = 13
# remap every reference to the old 2.1 layout, workbook-wide
REMAP = {"$C$43": f"$C${DEDUP_ROW}", "$C$52": f"$C${TOT_ROW}", "$D$52": f"$E${TOT_ROW}",
         "$E$52": f"$G${TOT_ROW}", "$F$52": f"$I${TOT_ROW}", "$G$52": f"$K${TOT_ROW}",
         "$F$49": f"$I${UNM_ROW}"}
pat21 = re.compile(r"('2\.1 Total Cost'!)(\$[A-Z]+\$\d+)")
for sheet in wb.worksheets:
    if sheet.title in ("Squads", "Added data", "Sheet2", "2.1 Total Cost"): continue
    for row_ in sheet.iter_rows():
        for cell in row_:
            v = cell.value
            if isinstance(v, str) and v.startswith("=") and "'2.1 Total Cost'!" in v:
                cell.value = pat21.sub(lambda m: m.group(1) + REMAP.get(m.group(2), m.group(2)), v)

# =====================================================================
# 3.0 FTE View - EGI strategic delivery roster + role language
# =====================================================================
ws = wb["3.0 FTE View"]
EGI_T = 167
sc(ws, f"B{EGI_T}", "EGI strategic delivery roster (Sheet2) - funded from Significant Items, outside the archetype model", WHITEF, NAVY)
for col, h in zip("BCDEF", ["Name", "Role", "Engagement", "Country", "Cost ($m)"]):
    sc(ws, f"{col}{EGI_T+1}", h, WHITEF, MIDBLU, align="center")
for i, x in enumerate(EGI):
    rr = EGI_T + 2 + i
    sc(ws, f"B{rr}", f_name(x["r"]), NORM)
    sc(ws, f"C{rr}", f_title(x["r"]), NORM)
    sc(ws, f"D{rr}", f_status(x["r"]), NORM, align="center")
    sc(ws, f"E{rr}", f_ctry(x["r"]), NORM)
    sc(ws, f"F{rr}", f"={S2}!$AA${x['r']}/1000000", NORM, fmt=M2, align="right")
EGI_TOT = EGI_T + 2 + len(EGI)
sc(ws, f"B{EGI_TOT}", "Total EGI strategic delivery", BOLD, LGREY)
sc(ws, f"C{EGI_TOT}", f"=COUNTA(B{EGI_T+2}:B{EGI_TOT-1})", BOLD, LGREY, fmt="0", align="center")
sc(ws, f"F{EGI_TOT}", f"=SUM(F{EGI_T+2}:F{EGI_TOT-1})", BOLD, LGREY, fmt=M2, align="right")
# wire the 2.1 memo row to this block
w21 = wb["2.1 Total Cost"]
w21[f"I{EGI_MEMO_ROW}"].value = f"='3.0 FTE View'!$F${EGI_TOT}"
w21[f"I{EGI_MEMO_ROW}"].number_format = M2
w21[f"J{EGI_MEMO_ROW}"].value = f"='3.0 FTE View'!$C${EGI_TOT}"
w21[f"J{EGI_MEMO_ROW}"].number_format = "0"
# seat -> role language on 3.0
for addr, txt in [("H3", "Roles above archetype"), ("J6", "Roles"), ("K6", "Roles vs archetype"),
                  ("K155", "Roles"), ("B161", "COE roles (detail on 2.3 / 2.4 / 2.5)")]:
    if ws[addr].value is not None: ws[addr].value = txt

wb.save(OUT)
print("stage 2 saved. rows: dedup", DEDUP_ROW, "total", TOT_ROW, "unmapped", UNM_ROW,
      "coe_first", COE_FIRST, "egi_tot", EGI_TOT)
# =====================================================================
# 4.x GM tabs - plain titles, vacancy impact, no seat language, no gaps
# =====================================================================
GM_TABS = [t for t in wb.sheetnames if t.startswith("4.")]
gm_anchor = {}
for tab in GM_TABS:
    ws = wb[tab]
    num, suffix = tab.split(" ", 1)[0][2:], tab.split(" ", 1)[1]
    onex = f"1.{num} {suffix}"
    assert onex in wb.sheetnames, onex
    # title as a live ref to the portfolio tab's own name cell
    title_ref = None
    for c in range(1, 5):
        v = wb[onex].cell(2, c).value
        if isinstance(v, str) and v.strip():
            title_ref = f"'{onex}'!${get_column_letter(c)}$2"
            break
    ws["B2"].value = (f'=CONCATENATE({title_ref}," GM working copy")' if title_ref
                      else f"{suffix} GM working copy")
    # locate the summary table and roster
    hdr = tot = rost_hdr = None
    for r in range(3, 45):
        b = ws.cell(r, 2).value
        if b == "Squad" and hdr is None: hdr = r
        elif b == "Total" and hdr and tot is None: tot = r
        elif b == "Name" and tot: rost_hdr = r; break
    assert hdr and tot and rost_hdr, (tab, hdr, tot, rost_hdr)
    # roster blocks: squad label -> (first person row, last person row)
    blocks, cur = {}, None
    last_person = rost_hdr
    for r in range(rost_hdr+1, ws.max_row+1):
        b = ws.cell(r, 2).value
        if isinstance(b, str) and b.startswith("=Squads!"):
            if cur: blocks[cur][1] = r
            last_person = r
        elif isinstance(b, str) and b.strip() and not b.startswith("="):
            if b.startswith(("Check", "Cost", "Vacant", "Leadership", "Cyber")): break
            cur = b.strip(); blocks[cur] = [r+1, r]
    # summary rows
    ws.cell(hdr, 4).value = "Archetype roles"
    ws.cell(hdr, 8).value = "Vacancies after calls"
    for r in range(hdr+1, tot):
        name = str(ws.cell(r, 2).value or "").strip()
        g = ws.cell(r, 7).value
        if not (isinstance(g, str) and g.startswith("=")):
            blk = blocks.get(name)
            ws.cell(r, 7).value = (f'=COUNTIF(E{blk[0]}:E{blk[1]},"Hire")' if blk else "=0")
            ws.cell(r, 7).number_format = "0"
        ws.cell(r, 8).value = f"=F{r}-G{r}"
        ws.cell(r, 9).value = f'=IFERROR(E{r}+G{r}-D{r},"-")'
        ws.cell(r, 10).value = (f'=IF(ISNUMBER(D{r}),IF(E{r}>D{r},"Filled already over archetype",'
                                f'IF(E{r}+G{r}>D{r},"Over archetype after hire calls","")),'
                                f'"Outside the archetype model - no target set")')
    ws.cell(tot, 7).value = f"=SUM(G{hdr+1}:G{tot-1})"
    ws.cell(tot, 8).value = f"=SUM(H{hdr+1}:H{tot-1})"
    # cyber: the archetype column has no target - say so plainly, no text-in-number cells
    dvals = [ws.cell(r, 4).value for r in range(hdr+1, tot)]
    if not any(isinstance(v, str) and v.startswith("=") for v in dvals):
        for r in range(hdr+1, tot):
            sc(ws, f"C{r}", "Priced from actual roles on 2.5", NORM, wrap=True)
            sc(ws, f"D{r}", '="-"', NORM, align="center")
        sc(ws, f"D{tot}", '="-"', BOLD, LGREY, align="center")
    # labels / language + roster band
    for r in range(1, ws.max_row+1):
        v = ws.cell(r, 2).value
        if not isinstance(v, str) or v.startswith("="): continue
        if v == "Your position by squad": ws.cell(r, 2).value = "Position by squad"
        elif "Cost to hire all vacant" in v: ws.cell(r, 2).value = "Cost to hire all vacancies ($m)"
        elif "Cost of the seats you chose" in v: ws.cell(r, 2).value = "Cost of roles marked Hire ($m)"
        elif "Vacant seats are priced" in v:
            ws.cell(r, 2).value = "Vacant roles are priced at standard title rates - indicative until an offer is made."
        elif v.startswith("Your people"):
            ws.cell(r, 2).value = "Roster - Hire or Hold each vacancy"
            for c in range(2, 7): ws.cell(r, c).fill = NAVY; ws.cell(r, c).font = WHITEF
            for c in range(7, 11):
                ws.cell(r, c).fill = PatternFill(); ws.cell(r, c).border = Border()
            if ws.cell(r-1, 2).value is None:
                sc(ws, f"B{r-1}", "vs archetype: positive = over the allowance, negative = under.",
                   NORM, border=False)
    ws.cell(rost_hdr, 5).value = "Call"
    # close the white gap the way the owner hinted on 4.2 / 4.9: column C shows
    # the archetype type and size, live from the same FTE View row D points at
    sc(ws, f"C{hdr}", "Archetype type and size", WHITEF, MIDBLU, align="center", wrap=True)
    for r in range(hdr+1, tot):
        d = ws.cell(r, 4).value
        m = re.match(r"^='3\.0 FTE View'!\$G\$(\d+)$", str(d or ""))
        if m:
            fr = m.group(1)
            e_, f_ = f"'3.0 FTE View'!$E${fr}", f"'3.0 FTE View'!$F${fr}"
            sc(ws, f"C{r}", f'=IF(OR({f_}=0,{f_}=""),{e_},{e_}&" - "&{f_})',
               NORM, wrap=True)
        else:
            sc(ws, f"C{r}", '="-"', NORM, align="center")
    gm_anchor[tab] = dict(hdr=hdr, tot=tot, rost_hdr=rost_hdr,
                          blocks={k: v for k, v in blocks.items()})
# sheet order: 4.9 before 4.10
idx9, idx10 = wb.sheetnames.index("4.9 Commercial Fuels"), wb.sheetnames.index("4.10 Z Retail")
if idx9 > idx10:
    wb.move_sheet("4.9 Commercial Fuels", offset=idx10-idx9)

# =====================================================================
# Exec Summary - role language, no baked-in numbers, no possessives
# =====================================================================
ws = wb["Exec Summary"]
EXEC_TXT = {
 "B6": "The test: can each portfolio fund its archetype cost (TDD + business)? If yes, live within it. If not, start with the vacancies (4.x GM tabs), then archetype size (1.x dropdowns).",
 "B11": "Squads are priced from the archetype library on 0.1 Squads (type x size). Offshore is priced at the offshore rate set on 0.1 Squads.",
 "B12": "Each portfolio pays one overhead: Head of Tech, Business Partner and Domain Architect shares, and leadership - rates on 0.0 Data Config.",
 "B7": "Each portfolio tab (1.x) shows the squads, sizes, support %, budget draw-downs and what is left to fund.",
 "B8": "Next step: agree funding for what is left to fund, and decide which vacancies to hire or hold.",
 "B36": "Roles the archetypes allow - squads at their set sizes",
 "B41": "Filled - people in roles today",
 "B55": "The main lever is the vacancies: they are raised but not hired, so holding them impacts nobody. Make the call role by role on the 4.x GM tab.",
 "B37": "Roles actually raised in those squads - filled + vacant",
 "B38": "Roles raised beyond the archetypes",
 "B39": "Roles in squads priced outside archetypes (AmPOS, EGI, cyber)",
 "B43": "of which squad roles - the GM hire or hold lever",
 "B44": "of which leadership, COE and unmapped roles",
 "B46": "Vacant = open roles in the raw data. Sheet2 status updates (Paused, ring-fenced) show on 2.3 to 2.5 and 3.1.",
 "B49": "Today's filled roles cost ($m)",
 "B50": "Filled roles over/(under) the archetype cost ($m)",
 "B51": "Hiring every vacancy would add ($m)",
 "B52": "of which squad roles - the 4.x GM lever ($m)",
 "B53": "Vacant roles are priced at standard title rates - indicative until an offer is made.",
 "B57": "Roles not mapped to any squad or COE ($m)",
 "B71": "Archetype squad roles allowed",
 "B72": "Org roles (excl leadership)",
 "B76": "Org roles include roles outside the archetype model; squad-only counts are on 3.0 FTE View.",
}
for addr, txt in EXEC_TXT.items():
    ws[addr].value = txt

# =====================================================================
# 3.1 Data QA - Sheet2 reconciliation with live refs (no typed data)
# =====================================================================
ws = wb["3.1 Data QA"]
for addr, txt in [("B71", "Roles by model squad - raw data vs Added data (differences only)"),
                  ("C72", "raw data roles"), ("D72", "Added data roles"),
                  ("B11", "In Added data only - joined the org after the raw data cut, or naming mismatch"),
                  ("B51", "In raw data only - left the org, or naming mismatch")]:
    if ws[addr].value is not None: ws[addr].value = txt
qr = ws.max_row + 3
def qhdr(text):
    global qr
    sc(ws, f"B{qr}", text, WHITEF, NAVY); qr += 1
def qcols(*labels):
    global qr
    for i, l in enumerate(labels):
        sc(ws, f"{get_column_letter(2+i)}{qr}", l, WHITEF, MIDBLU, align="center")
    qr += 1
def qnote(text):
    global qr
    sc(ws, f"B{qr}", text, NORM, border=False); qr += 1

# pair Sheet2 rows to Squads rows: exact name+title, then title within division
sq_by_key, sq_by_div = {}, {}
for q in sq_rows:
    sq_by_key.setdefault((q["name"], q["title"]), []).append(q)
    sq_by_div.setdefault(q["div"], []).append(q)
used = set()
pairs, s2_only = [], []
for x in rows2:
    k = (low(x["name"]), low(x["title"]))
    cand = [q for q in sq_by_key.get(k, []) if q["r"] not in used and q["div"] == low(x["div"])] \
        or [q for q in sq_by_key.get(k, []) if q["r"] not in used]
    if cand:
        used.add(cand[0]["r"]); pairs.append((x, cand[0])); continue
    # title-only pairing is only safe for unnamed rows (Vacant / ring fenced);
    # a NAMED person with no exact match is a new joiner and must be listed
    if "vacant" in low(x["name"]) or "ring fenced" in low(x["name"]):
        tcand = [q for q in sq_by_div.get(low(x["div"]), [])
                 if q["r"] not in used and q["title"] == low(x["title"])]
        if tcand:
            used.add(tcand[0]["r"]); pairs.append((x, tcand[0])); continue
    s2_only.append(x)
DIV4 = {"cyber, risk & operations", "egi", "partnering & transformation",
        "strategy, architecture & data"}
sq_only = [q for q in sq_rows if q["div"] in DIV4 and q["r"] not in used]
def s2_open(x): return x["typ"] in ("v", "pause")
def sq_open(q): return q["st"] == "Vacant"
status_diffs = [(x, q) for x, q in pairs
                if s2_open(x) != sq_open(q) or (x["typ"] == "pause" and q["st"] == "Vacant")]

qhdr("Sheet2 reconciliation - the updated rosters vs the raw data (Squads). Sheet2 drives 2.3 / 2.4 / 2.5 and the EGI roster; Squads drives the portfolio squads.")
qcols("Name (Sheet2)", "Role (Sheet2)", "Division", "Raw data status", "Sheet2 status")
for x, q in status_diffs:
    sc(ws, f"B{qr}", f_name(x["r"]), NORM)
    sc(ws, f"C{qr}", f_title(x["r"]), NORM)
    sc(ws, f"D{qr}", f"={S2}!$F${x['r']}", NORM)
    sc(ws, f"E{qr}", f"=Squads!$R${q['r']}", NORM, align="center")
    sc(ws, f"F{qr}", f_status(x["r"]), NORM, align="center")
    qr += 1
qnote("Where the two disagree, the Sheet2 status is used on 2.3 / 2.4 / 2.5. The raw data still drives the portfolio squad counts on 1.x / 3.0 / 4.x.")
qr += 1
qhdr("Roles in Sheet2 with no matching raw data row")
qcols("Name (Sheet2)", "Role (Sheet2)", "Division", "Where it now shows")
for x in s2_only:
    if x["div"] == "EGI": where = "3.0 FTE View - EGI strategic delivery roster"
    elif x["port"] == "enterprise data" and not blankish(x["squad"]):
        where = "In an Enterprise Data squad per Sheet2 - needs a raw data row to join the 1.3 / 4.3 counts"
    elif x["port"] == "enterprise data": where = "Nowhere yet - needs a raw data row to join the Enterprise Data portfolio"
    elif x["div"] == "Cyber, Risk & Operations": where = "2.5 Cyber Roles"
    elif x["div"] == "Partnering & Transformation": where = "2.3 BP&T"
    else: where = "2.4 SA&D"
    sc(ws, f"B{qr}", f_name(x["r"]), NORM)
    sc(ws, f"C{qr}", f_title(x["r"]), NORM)
    sc(ws, f"D{qr}", f"={S2}!$F${x['r']}", NORM)
    sc(ws, f"E{qr}", where, NORM)
    qr += 1
qr += 1
if sq_only:
    qhdr("Raw data rows (these four divisions) with no matching Sheet2 row - check if dropped or renamed")
    qcols("Name (raw data)", "Role (raw data)", "Model portfolio", "Status")
    for q in sq_only:
        sc(ws, f"B{qr}", f"=Squads!$B${q['r']}", NORM)
        sc(ws, f"C{qr}", f'=SUBSTITUTE(SUBSTITUTE(Squads!$C${q["r"]},"–","-"),"—","-")', NORM)
        sc(ws, f"D{qr}", f"=Squads!$N${q['r']}", NORM)
        sc(ws, f"E{qr}", f"=Squads!$R${q['r']}", NORM, align="center")
        qr += 1
    qr += 1
qhdr("Raw data integrity audit")
qnote("Zero rows were added to the raw data: the Squads tab in this workbook matches the owner's uploaded copy cell for cell.")
qnote("16 mapping cells on 4 Commercial rows were changed in an earlier build (Unmapped to COE - Business Partnering). The uploaded copy already includes that change, so the two files match today. The names below are live references to those rows.")
for rr in (260, 269, 270, 281):
    sc(ws, f"B{qr}", f"=Squads!$B${rr}", NORM)
    sc(ws, f"C{qr}", f'=SUBSTITUTE(SUBSTITUTE(Squads!$C${rr},"–","-"),"—","-")', NORM)
    sc(ws, f"D{qr}", f"=Squads!$P${rr}", NORM)
    sc(ws, f"E{qr}", f"=Squads!$R${rr}", NORM, align="center")
    qr += 1
qnote("Sheet2 confirms these four roles sit in Partnering & Transformation (Commercial), inside the Business Partnering bucket on 2.3.")
qr += 1
qhdr("Cost bases")
qnote("2.3 / 2.4 / 2.5 and the EGI roster are costed from Sheet2 (Full Cost AUD, column AA). Portfolio actuals on 2.1 are costed from the Added data ledger. The difference is the restatement line on 2.1.")

# =====================================================================
# STAGE 4a: 2.0 Group Summary - the owner's column spec + AU/NZ variances
# =====================================================================
g0 = wb["2.0 Group Summary"]
HDR20 = {"C": "TDD Lights On Budget ($m)", "D": "Support Cost ($m)",
         "E": "Variance ($m) = budget - cost", "G": "Cost of non-TDD funding ($m)",
         "H": "Amount that can be recharged ($m)", "I": "Left to fund outside TDD ($m)",
         "J": "Total still left to fund ($m)", "K": "Total Cost ($m)"}
for col, h in HDR20.items():
    sc(g0, f"{col}5", h, WHITEF, MIDBLU, align="center", wrap=True)
for r in list(range(6, 17)) + list(range(18, 23)):
    oldj = g0.cell(r, 10).value
    if oldj is None: continue
    sc(g0, f"K{r}", oldj, NORM, fmt=M2, align="right")
    sc(g0, f"J{r}", f"=MAX(0,-E{r})+I{r}", NORM, fmt=M2, align="right")
for r, kfx in ((17, "=SUM(K6:K16)"), (24, "=SUM(K17,K18,K19,K20,K21,K22)")):
    sc(g0, f"K{r}", kfx, BOLD, LGREY, fmt=M2, align="right")
    sc(g0, f"J{r}", "=SUM(J6:J16)" if r == 17 else "=SUM(J17,J18,J19,J20,J21,J22)",
       BOLD, LGREY, fmt=M2, align="right")
# net-cost lines move to K; note 2.0 is the archetype view
for r in (25, 26):
    v = g0.cell(r, 10).value
    if v is not None:
        sc(g0, f"K{r}", str(v).replace("J24", "K24"), BOLD, fmt=M2, align="right")
        g0.cell(r, 10).value = None
sc(g0, "B3", "Archetype view: what the designed model costs against the TDD budgets. The actual organisation is on 2.1.", NORM, border=False)
# remap references to old 2.0 J cells (total cost) -> K
patJ = re.compile(r"('2\.0 Group Summary'!\$?J\$?)(\d+)")
for sheet in wb.worksheets:
    if sheet.title in ("Squads", "Added data", "Sheet2", "2.0 Group Summary"): continue
    for row_ in sheet.iter_rows():
        for cell in row_:
            v = cell.value
            if isinstance(v, str) and v.startswith("=") and "'2.0 Group Summary'!" in v:
                cell.value = v.replace("'2.0 Group Summary'!$J$", "'2.0 Group Summary'!$K$")

# =====================================================================
# STAGE 4b: Fund toggles (AU/NZ) on every 1.x squad table + tab subtotals
# =====================================================================
AUNZ = {}
dvfund_all = []
for tab in [t for t in wb.sheetnames if t.startswith("1.") and t[2] in ".0123456789"]:
    ws = wb[tab]
    tables = []
    for r in range(15, ws.max_row + 1):
        if ws.cell(r, 2).value == "Squad" and ws.cell(r, 8).value == "TDD Cost ($m)":
            rows = []
            rr = r + 1
            while rr <= ws.max_row:
                b = str(ws.cell(rr, 2).value or "")
                if not b or "Overhead" in b or b.endswith("Total"): break
                rows.append(rr); rr += 1
            if rows: tables.append((r, rows))
    if not tables: continue
    allrows = [rr for _, rows in tables for rr in rows]
    fcol = next(c for c in range(10, 15)
                if all(ws.cell(rr, c).value is None for rr in allrows)
                and all(ws.cell(h, c).value is None for h, _ in tables))
    fL = get_column_letter(fcol)
    dv = DataValidation(type="list", formula1='"AU,NZ"', allow_blank=False, showErrorMessage=True)
    ws.add_data_validation(dv)
    for h, rows in tables:
        sc(ws, f"{fL}{h}", "Fund", WHITEF, MIDBLU, align="center")
        plat = ""
        for up in range(h - 1, max(1, h - 6), -1):
            t_ = str(ws.cell(up, 2).value or "")
            if t_.startswith("Platform"): plat = t_; break
        default = "NZ" if (tab == "1.10 Z Retail" or " Z " in plat + " " or plat.startswith("Platform: Z")) else "AU"
        for rr in rows:
            sc(ws, f"{fL}{rr}", default, NORM, YELL, align="center")
            dv.add(f"{fL}{rr}")
    terms_au = "+".join(f'SUMIF({fL}{rows[0]}:{fL}{rows[-1]},"AU",H{rows[0]}:H{rows[-1]})' for _, rows in tables)
    terms_nz = "+".join(f'SUMIF({fL}{rows[0]}:{fL}{rows[-1]},"NZ",H{rows[0]}:H{rows[-1]})' for _, rows in tables)
    base = ws.max_row + 2
    sc(ws, f"B{base}", "TDD squad cost - AU funded (Fund toggles) ($m)", NORM)
    sc(ws, f"C{base}", f"={terms_au}", NORM, fmt=M2, align="right")
    sc(ws, f"B{base+1}", "TDD squad cost - NZ funded (Fund toggles) ($m)", NORM)
    sc(ws, f"C{base+1}", f"={terms_nz}", NORM, fmt=M2, align="right")
    AUNZ[tab] = (f"C{base}", f"C{base+1}")
au_sum = "+".join(f"'{t}'!${a.replace('C','C$')}" for t, (a, _) in AUNZ.items())
nz_sum = "+".join(f"'{t}'!${b.replace('C','C$')}" for t, (_, b) in AUNZ.items())
r0 = 37
sc(g0, f"B{r0}", "AU / NZ funding split - from the squad Fund toggles on the 1.x tabs", WHITEF, NAVY)
aunz_rows = [
    ("AU budget - 0.0 Data Config ($m)", "='0.0 Data Config'!$C$27"),
    ("AU allocated to COEs - 0.0 Data Config ($m)", "=SUM('0.0 Data Config'!$C$6:$C$10)"),
    ("TDD squad cost - AU funded ($m)", f"={au_sum}"),
    ("Variance vs AU budget ($m)", f"=C{r0+1}-C{r0+2}-C{r0+3}"),
    ("NZ budget - 0.0 Data Config ($m)", "='0.0 Data Config'!$D$27"),
    ("NZ allocated to COEs - 0.0 Data Config ($m)", "=SUM('0.0 Data Config'!$D$6:$D$10)"),
    ("TDD squad cost - NZ funded ($m)", f"={nz_sum}"),
    ("Variance vs NZ budget ($m)", f"=C{r0+5}-C{r0+6}-C{r0+7}"),
]
for i, (lab, fx) in enumerate(aunz_rows):
    rr = r0 + 1 + i
    bold = lab.startswith("Variance")
    sc(g0, f"B{rr}", lab, BOLD if bold else NORM)
    sc(g0, f"C{rr}", fx, BOLD if bold else NORM, fmt=M2, align="right")
sc(g0, f"B{r0+9}", "Portfolio overheads and platform overheads sit inside the COE and portfolio allocations on 0.0 Data Config.", NORM, border=False)
# AU/NZ split on the COE tabs (by each role's country)
for tab, jr1, jr2, er1, er2 in [("2.3 BP&T", R1, PT_CHECK-1, R1, PT_CHECK-1),
                                 ("2.4 SA&D", S1, SAD_CHECK-1, S1, SAD_CHECK-1),
                                 ("2.5 Cyber Roles", C1, CY_CHECK-1, C1, CY_CHECK-1)]:
    ws = wb[tab]
    sc(ws, "E10", "Cost - AU and other funded ($m)", NORM)
    sc(ws, "F10", f'=SUMIF($E${er1}:$E${er2},"<>NZ",$J${jr1}:$J${jr2})', NORM, fmt=M2, align="right")
    sc(ws, "E11", "Cost - NZ funded ($m)", NORM)
    sc(ws, "F11", f'=SUMIF($E${er1}:$E${er2},"NZ",$J${jr1}:$J${jr2})', NORM, fmt=M2, align="right")

# =====================================================================
# STAGE 4c: working tabs 4.12 / 4.13 / 4.14 - every role gets a 4.x home
# =====================================================================
def mk_working(name, title, after):
    if name in wb.sheetnames: del wb[name]
    ws = wb.create_sheet(name)
    wb.move_sheet(name, offset=wb.sheetnames.index(after) - wb.sheetnames.index(name) + 1)
    ws.sheet_view.showGridLines = False
    for col, w in (("A", 3), ("B", 30), ("C", 40), ("D", 26), ("E", 12), ("F", 10), ("G", 16)):
        ws.column_dimensions[col].width = w
    sc(ws, "B2", title, Font(name="Calibri", size=14, bold=True, color="FF002F6C"), border=False)
    return ws
def w_hdrs(ws, r):
    for col, h in zip("BCDEFG", ["Name", "Role", "Department", "Status", "Call", "Cost if hired ($)"]):
        sc(ws, f"{col}{r}", h, WHITEF, MIDBLU, align="center")
def w_rows_sheet2(ws, first, entries):
    dv = DataValidation(type="list", formula1='"Hire,Hold"', allow_blank=False, showErrorMessage=True)
    ws.add_data_validation(dv)
    for i, x in enumerate(entries):
        rr = first + i
        sc(ws, f"B{rr}", f_name(x["r"]), NORM)
        sc(ws, f"C{rr}", f_title(x["r"]), NORM, wrap=True)
        sc(ws, f"D{rr}", f_dept(x["r"]), NORM)
        sc(ws, f"E{rr}", f_status(x["r"]), NORM, align="center")
        if x["typ"] in ("v", "pause"):
            sc(ws, f"F{rr}", "Hold", NORM, YELL, align="center"); dv.add(f"F{rr}")
            sc(ws, f"G{rr}", f_cost(x["r"]), NORM, fmt=D0, align="right")
    return first + len(entries) - 1
def w_summary(ws, first, last, top=4):
    sc(ws, f"B{top}", "Position", WHITEF, NAVY)
    items = [("Roles", f"=COUNTA(B{first}:B{last})", "0"),
             ("Filled", f'=COUNTIF(E{first}:E{last},"Filled")+COUNTIF(E{first}:E{last},"Contractor")', "0"),
             ("Vacant", f'=COUNTIF(E{first}:E{last},"Vacant")', "0"),
             ("Paused", f'=COUNTIF(E{first}:E{last},"Paused")', "0"),
             ("Cost to hire all vacancies ($m)", f'=SUMIF(E{first}:E{last},"Vacant",G{first}:G{last})/1000000', M2),
             ("Cost of roles marked Hire ($m)", f'=SUMIF(F{first}:F{last},"Hire",G{first}:G{last})/1000000', M2)]
    for i, (lab, fx, fmt) in enumerate(items):
        sc(ws, f"B{top+1+i}", lab, NORM)
        sc(ws, f"C{top+1+i}", fx, NORM, fmt=fmt, align="right")
ws = mk_working("4.12 BP&T", "Business Partnering & Transformation GM working copy", "4.11 TDD Cyber")
w_hdrs(ws, 13); last = w_rows_sheet2(ws, 14, PT); w_summary(ws, 14, last)
sc(ws, f"B{last+2}", "Funding for these roles is on 2.3 BP&T. Roles and costs come from Sheet2.", NORM, border=False)
W412 = (14, last)
ws = mk_working("4.13 SA&D", "Strategy, Architecture & Data GM working copy", "4.12 BP&T")
w_hdrs(ws, 13); last = w_rows_sheet2(ws, 14, sad_coe); w_summary(ws, 14, last)
sc(ws, f"B{last+2}", "COE roles only - squad-based SA&D roles sit on 4.3 (Group Data portfolio). Funding is on 2.4 SA&D.", NORM, border=False)
W413 = (14, last)
ws = mk_working("4.14 EGI & Central", "EGI & Central Roles working copy", "4.13 SA&D")
sc(ws, "B12", "EGI strategic delivery (Sheet2) - contractors, funded from Significant Items", WHITEF, NAVY)
w_hdrs(ws, 13); last = w_rows_sheet2(ws, 14, EGI)
egi_last = last
# every Squads row must have a 4.x home: rows already referenced on 4.1-4.11,
# COE rows carried by the Sheet2 rosters on 4.12/4.13, leadership here, and
# EVERYTHING left over (unmapped, stray squads, raw EGI rows) here too
refset = set()
for t4 in [t for t in wb.sheetnames if t.startswith("4.") and t not in ("4.12 BP&T", "4.13 SA&D", "4.14 EGI & Central")]:
    for row_ in wb[t4].iter_rows(min_col=2, max_col=2):
        m_ = re.match(r"^=Squads!\$B\$(\d+)$", str(row_[0].value or ""))
        if m_: refset.add(int(m_.group(1)))
s2cov = set()
for x in PT + sad_coe:            # rows carried on 4.12 / 4.13 via Sheet2
    m_ = sq_match(x)
    if m_: s2cov.add(m_["r"])
lead_rows = [q for q in sq_rows if q["cls"] == "Leadership"]
leftover_rows = [q for q in sq_rows if q["r"] not in refset and q["r"] not in s2cov
                 and q["cls"] != "Leadership"]
sec = last + 2
sc(ws, f"B{sec}", "Leadership roles - funded via portfolio overheads", WHITEF, NAVY)
w_hdrs(ws, sec + 1)
dv2 = DataValidation(type="list", formula1='"Hire,Hold"', allow_blank=False, showErrorMessage=True)
ws.add_data_validation(dv2)
def w_rows_squads(ws, first, entries, dv):
    for i, q in enumerate(entries):
        rr = first + i; n = q["r"]
        sc(ws, f"B{rr}", f'=IF(Squads!$R${n}="Vacant","Vacant",Squads!$B${n})', NORM)
        sc(ws, f"C{rr}", f'=SUBSTITUTE(SUBSTITUTE(Squads!$C${n},"–","-"),"—","-")', NORM, wrap=True)
        sc(ws, f"D{rr}", f"=Squads!$N${n}", NORM)
        sc(ws, f"E{rr}", f"=Squads!$R${n}", NORM, align="center")
        if q["st"] == "Vacant":
            sc(ws, f"F{rr}", "Hold", NORM, YELL, align="center"); dv.add(f"F{rr}")
            sc(ws, f"G{rr}", f"=SUMIF('Added data'!$C$2:$C$549,Squads!$C${n},'Added data'!$AA$2:$AA$549)"
                             f"/COUNTIF('Added data'!$C$2:$C$549,Squads!$C${n})", NORM, fmt=D0, align="right")
    return first + len(entries) - 1
last = w_rows_squads(ws, sec + 2, lead_rows, dv2)
LEAD_BLK = (sec + 2, last)
sec2 = last + 2
sc(ws, f"B{sec2}", "Every other role - unmapped, outside the archetype model, or raw data EGI rows", WHITEF, NAVY)
w_hdrs(ws, sec2 + 1)
last = w_rows_squads(ws, sec2 + 2, leftover_rows, dv2)
UNM_BLK = (sec2 + 2, last)
# new-in-Sheet2 people with no raw data row yet - they still need a 4.x home
newins = [x for x in rows2 if x["div"] != "EGI" and x not in PT and x not in sad_coe
          and sq_match(x) is None]
sec3 = last + 2
NEW_BLK = (0, -1)
if newins:
    sc(ws, f"B{sec3}", "New in Sheet2 - no raw data row yet. Add them to the Squads tab to join the portfolio counts.", WHITEF, NAVY)
    w_hdrs(ws, sec3 + 1)
    dv3 = DataValidation(type="list", formula1='"Hire,Hold"', allow_blank=False, showErrorMessage=True)
    ws.add_data_validation(dv3)
    first3 = sec3 + 2
    for i, x in enumerate(newins):
        rr = first3 + i
        sc(ws, f"B{rr}", f_name(x["r"]), NORM)
        sc(ws, f"C{rr}", f_title(x["r"]), NORM, wrap=True)
        sc(ws, f"D{rr}", f"={S2}!$F${x['r']}", NORM)
        sc(ws, f"E{rr}", f_status(x["r"]), NORM, align="center")
        if x["typ"] in ("v", "pause"):
            sc(ws, f"F{rr}", "Hold", NORM, YELL, align="center"); dv3.add(f"F{rr}")
            sc(ws, f"G{rr}", f_cost(x["r"]), NORM, fmt=D0, align="right")
    last = first3 + len(newins) - 1
    NEW_BLK = (first3, last)
w_summary(ws, 14, last)
# roles count must span only the person rows of the blocks
cnt = (f"=COUNTA(B14:B{egi_last})+COUNTA(B{LEAD_BLK[0]}:B{LEAD_BLK[1]})"
       f"+COUNTA(B{UNM_BLK[0]}:B{UNM_BLK[1]})")
if newins: cnt += f"+COUNTA(B{NEW_BLK[0]}:B{NEW_BLK[1]})"
sc(ws, "C5", cnt, NORM, fmt="0", align="right")
sc(ws, f"B{last+2}", "Every role that is not in a portfolio squad (4.1-4.11) or a COE (4.12-4.13) lives here - nothing is missed.", NORM, border=False)

# =====================================================================
# STAGE 4d: 0.5 Guide - how the whole workbook flows, what to edit
# =====================================================================
if "0.5 Guide" in wb.sheetnames: del wb["0.5 Guide"]
gd = wb.create_sheet("0.5 Guide")
wb.move_sheet("0.5 Guide", offset=wb.sheetnames.index("0.4 Budget Table (Fin)") - wb.sheetnames.index("0.5 Guide") + 1)
gd.sheet_view.showGridLines = False
gd.column_dimensions["A"].width = 3
gd.column_dimensions["B"].width = 120
sc(gd, "B2", "How this workbook flows", Font(name="Calibri", size=16, bold=True, color="FF002F6C"), border=False)
GUIDE = [
 ("h", "The spine"),
 ("t", "Inputs (0.x) feed the archetype model (1.x), which rolls into the funding view (2.0). The actual organisation (Squads tab + Sheet2 + Added data) rolls into the affordability view (2.1). GMs make Hire or Hold calls on the working tabs (4.x). Evidence and checks live on 3.1."),
 ("h", "What you can edit (yellow cells are inputs)"),
 ("t", "0.0 Data Config - TDD budgets by portfolio (AU and NZ), overhead rates, COE allocations."),
 ("t", "0.1 Squads - the archetype price library (type x size) and the offshore rate."),
 ("t", "1.x squad tables - squad type, size, On/Off, support %, and the Fund toggle (AU or NZ) that decides which budget the squad draws."),
 ("t", "Squads tab - the raw organisation and its model mapping (columns N to R). This drives every role count."),
 ("t", "Sheet2 - the updated rosters for BP&T, SA&D, Cyber and EGI. This drives 2.3, 2.4, 2.5 and the EGI roster."),
 ("t", "4.x Call cells - Hire or Hold on every vacancy."),
 ("h", "If you change X, Y updates"),
 ("t", "Change a squad type, size, On/Off, support % or Fund toggle on 1.x - the squad price, 2.0, 2.1 archetype column, 3.0 and the Exec Summary all update."),
 ("t", "Change a status or mapping in the Squads tab - role counts on 3.0 and 4.x, vacancy counts everywhere, and the actual columns of 2.1 update."),
 ("t", "Change a person or cost in Sheet2 - 2.3 / 2.4 / 2.5, the EGI roster, 2.2, the COE rows of 2.0 and 2.1, and tabs 4.12 / 4.13 / 4.14 update."),
 ("t", "Change a budget on 0.0 Data Config - budget and variance columns on 2.0, the draw-downs on 2.2 / 2.3 / 2.4, and left to fund everywhere update."),
 ("t", "Make a Hire or Hold call on any 4.x tab - that tab's cost of roles marked Hire updates. Totals only move when a role is actually hired in the raw data."),
 ("h", "What each tab is"),
 ("t", "Exec Summary - the whole story on one page. | 0.0-0.4 - inputs and reference tables. | 1.1-1.11 - one archetype build per portfolio. | 2.0 - archetype cost vs the TDD budgets (the funding view). | 2.1 - archetype vs what the org actually costs (the affordability view), portfolio level; squad detail on 3.0. | 2.2 - COE budget roll-up and the TDD corporate funding pool. | 2.3 / 2.4 / 2.5 - the COE and cyber rosters and their funding. | 3.0 - squad-level detail: archetype FTE and cost vs actual, drills 2.1. | 3.1 - data checks, reconciliations and the audit trail. | 4.1-4.11 - GM working copies per portfolio. | 4.12-4.14 - working copies for BP&T, SA&D and EGI & central roles. | Squads / Sheet2 / Added data - source data."),
 ("h", "How to trust it"),
 ("t", "Every tab with a row labelled Check must show 0. All reconciliations and the raw-data audit live on 3.1. Every role in the organisation appears on exactly one 4.x working tab - if a role were missed, the checks on 3.1 would not foot."),
]
gr_ = 4
for kind, txt in GUIDE:
    if kind == "h":
        sc(gd, f"B{gr_}", txt, WHITEF, NAVY); gr_ += 1
    else:
        c = sc(gd, f"B{gr_}", txt, NORM, border=False, wrap=True)
        gd.row_dimensions[gr_].height = max(15, 15 * (1 + len(txt) // 130))
        gr_ += 1
sc(gd, f"B{gr_+1}", "3.0 FTE View is the squad-level drill of 2.1. 2.0 shows the archetype view only - actuals live on 2.1.", NORM, border=False)

# retitle 3.0 so its place in the flow is explicit
wb["3.0 FTE View"]["B2"].value = "Squad detail - archetype vs actual by portfolio, platform and squad (drills 2.1)"
# 2.2 header: roll-up only
wb["2.2 COE"]["B3"].value = "Roll-up only - edit budgets on 0.0 Data Config, rosters on Sheet2. Detail on 2.3 / 2.4 / 2.5."

# =====================================================================
# STAGE 4e: the mapping change log on 3.1 (full disclosure)
# =====================================================================
if CHANGELOG:
    qr += 1
    qhdr("Model mapping updates made in this build - the owner's SA&D COE list (ruthless mapping)")
    qcols("Raw row", "Cell", "Old value", "New value", "Role")
    for row, colL, old, new, nm in CHANGELOG:
        sc(ws31_ := wb["3.1 Data QA"], f"B{qr}", row, NORM, align="center")
        sc(ws31_, f"C{qr}", f"{colL}{row}", NORM, align="center")
        sc(ws31_, f"D{qr}", str(old), NORM)
        sc(ws31_, f"E{qr}", new, NORM)
        sc(ws31_, f"F{qr}", f"=Squads!$B${row}&\" - \"&Squads!$C${row}", NORM)
        qr += 1
    qnote("These cells were updated at the owner's direction: every role on the owner's SA&D COE list now maps to COE - Data or COE - Strategy Architecture. The Group Data portfolio keeps the engineering, science, operations and reporting squads.")

# =====================================================================
# STAGE 5: numbering = flow. Renumber tabs, rewrite every reference,
# separators + tab colours, reorder, and stale text mentions fixed.
# =====================================================================
RENAME = {
    "0.4 Budget Table (Fin)": "0.1 Budget Table (Fin)",
    "0.0 Data Config": "0.2 Data Config",
    "0.1 Squads": "0.3 Squad Archetypes",
    "0.3 For Presentation Pack (2)": "0.4 Presentation Pack",
    "0.5 Guide": "0.0 Guide",
    "0.2 FY26 Budget": "FY26 Budget (ref)",
    "2.3 BP&T": "2.1 BP&T",
    "2.4 SA&D": "2.2 SA&D",
    "2.5 Cyber Roles": "2.3 Cyber Roles",
    "2.2 COE": "2.4 COE Summary",
    "2.0 Group Summary": "2.5 Group Summary",
    "2.1 Total Cost": "2.6 Total Cost",
    "3.0 FTE View": "2.7 Squad Detail",
    "3.1 Data QA": "5.0 Data QA",
}
# rewrite every formula reference (full quoted sheet names - unambiguous)
for sheet in wb.worksheets:
    for row_ in sheet.iter_rows():
        for cell in row_:
            v = cell.value
            if isinstance(v, str) and v.startswith("=") and "'" in v:
                nv = v
                for old, new in RENAME.items():
                    nv = nv.replace(f"'{old}'!", f"'{new}'!")
                if nv != v: cell.value = nv
# stale tab mentions inside plain text labels (ours only; longest first)
TEXTMAP = [("2.3 to 2.5 and 3.1", "2.1 to 2.3 and 5.0"),
           ("on 2.3, 2.4 or 2.5", "on 2.1, 2.2 or 2.3"),
           ("2.3 / 2.4 / 2.5", "2.1 / 2.2 / 2.3"),
           ("2.5 Cyber Roles", "2.3 Cyber Roles"),
           ("2.3 BP&T", "2.1 BP&T"), ("2.4 SA&D", "2.2 SA&D"),
           ("0.0 Data Config", "0.2 Data Config"), ("0.1 Squads", "0.3 Squad Archetypes"),
           ("3.0 FTE View", "2.7 Squad Detail"), ("drills 2.1", "drills 2.6"),
           ("see 3.1", "see 5.0"), ("and 3.1", "and 5.0"), ("on 3.1", "on 5.0"),
           ("on 3.0", "on 2.7"), ("see 2.2 (", "see 2.4 ("), ("is on 2.1.", "is on 2.6."),
           ("sit on 2.4", "sit on 2.2"), ("2.2 COE", "2.4 COE Summary")]
SKIP_TXT = {"Squads", "Added data", "Sheet2", "0.4 Presentation Pack",
            "0.1 Budget Table (Fin)", "FY26 Budget (ref)", "Sheet1"}
for sheet in wb.worksheets:
    if sheet.title in SKIP_TXT or sheet.title in ("0.3 For Presentation Pack (2)", "0.4 Budget Table (Fin)"):
        continue
    for row_ in sheet.iter_rows():
        for cell in row_:
            v = cell.value
            if isinstance(v, str) and not v.startswith("="):
                nv = v
                for old, new in TEXTMAP: nv = nv.replace(old, new)
                if nv != v: cell.value = nv
for old, new in RENAME.items():
    if old in wb.sheetnames: wb[old].title = new
# separators + group colours
GROUPS = [("- INPUTS -", "808080", ["0.0 Guide", "0.1 Budget Table (Fin)", "0.2 Data Config",
                                    "0.3 Squad Archetypes", "0.4 Presentation Pack"]),
          ("- DESIGNS -", "1F4E79", [f"1.{i} " for i in range(1, 12)]),
          ("- ROLL-UP -", "002F6C", ["2.1 BP&T", "2.2 SA&D", "2.3 Cyber Roles", "2.4 COE Summary",
                                     "2.5 Group Summary", "2.6 Total Cost", "2.7 Squad Detail"]),
          ("- DECISIONS -", "BF8F00", [f"4.{i} " for i in range(1, 12)] + ["4.12 BP&T", "4.13 SA&D", "4.14 EGI & Central"]),
          ("- EVIDENCE -", "375623", ["5.0 Data QA", "Squads", "Added data", "Sheet2"])]
def color_of(title):
    for gname, col, members in GROUPS:
        for m in members:
            if title == m or title.startswith(m): return col
    return None
for gname, col, _ in GROUPS:
    if gname in wb.sheetnames: del wb[gname]
    sp = wb.create_sheet(gname)
    sp.sheet_properties.tabColor = col
    sp.sheet_view.showGridLines = False
    sp.column_dimensions["A"].width = 2
    sp.column_dimensions["B"].width = 60
    sc(sp, "B2", gname.strip("- ") + "  >>>", Font(name="Calibri", size=16, bold=True, color="FF" + col), border=False)
for t in wb.sheetnames:
    c = color_of(t)
    if c: wb[t].sheet_properties.tabColor = c
ORDER = (["Exec Summary", "- INPUTS -", "0.0 Guide", "0.1 Budget Table (Fin)", "0.2 Data Config",
          "0.3 Squad Archetypes", "0.4 Presentation Pack", "- DESIGNS -"]
         + [t for t in wb.sheetnames if t.startswith("1.")]
         + ["- ROLL-UP -", "2.1 BP&T", "2.2 SA&D", "2.3 Cyber Roles", "2.4 COE Summary",
            "2.5 Group Summary", "2.6 Total Cost", "2.7 Squad Detail", "- DECISIONS -"]
         + [t for t in wb.sheetnames if t.startswith("4.")]
         + ["- EVIDENCE -", "5.0 Data QA", "Squads", "Added data", "Sheet2"])
rest = [t for t in wb.sheetnames if t not in ORDER]
final = ORDER + rest
wb._sheets = [wb[t] for t in final]
if "Sheet1" in wb.sheetnames: wb["Sheet1"].sheet_state = "hidden"
# guide rewrite: the two-stream flow with the final numbering
gd = wb["0.0 Guide"]
clear_region(gd, 2, gd.max_row + 2, 1, 6)
gd.column_dimensions["B"].width = 118
sc(gd, "B2", "How this workbook flows", Font(name="Calibri", size=16, bold=True, color="FF002F6C"), border=False)
GUIDE2 = [
 ("h", "The flow - two streams that meet in the middle"),
 ("t", "STREAM 1, THE DESIGN: 0.1 Budget Table (the money Finance gave) feeds 0.2 Data Config (how it is allocated per portfolio, AU and NZ, plus overhead rates and the four $2m COE allocations). 0.3 Squad Archetypes prices the contract (type x size). Each GM's design sits on 1.1 to 1.11, priced from 0.3 and funded per 0.2."),
 ("t", "STREAM 2, THE ORG: the Squads tab holds every role actually raised (536, filled and vacant) and its model mapping. Sheet2 holds the updated rosters for BP&T, SA&D, Cyber and EGI. Added data holds what each person actually costs."),
 ("t", "THE ROLL-UP: the COE rosters (2.1 BP&T, 2.2 SA&D, 2.3 Cyber Roles) and their funding roll into 2.4 COE Summary. 2.5 Group Summary answers: can the budgets fund the design (with AU and NZ variance). 2.6 Total Cost is the test: the design vs what the org actually costs. 2.7 Squad Detail is the same test squad by squad."),
 ("t", "THE DECISIONS: tabs 4.1 to 4.14 - every role in the organisation appears on exactly one of them. GMs mark Hire or Hold on every vacancy. The Exec Summary tells the result. 5.0 Data QA is the evidence behind every number."),
 ("h", "What you can edit (yellow cells are inputs)"),
 ("t", "0.2 Data Config: budgets, rates, allocations. | 0.3 Squad Archetypes: prices and the offshore rate. | 1.x: squad type, size, On/Off, support %, Fund (AU or NZ). | 4.x: the Hire or Hold Call cells. | Squads / Sheet2 / Added data: your source data."),
 ("h", "If you change X, Y updates"),
 ("t", "A squad's type, size, On/Off, support % or Fund toggle on 1.x updates 2.5, 2.6, 2.7 and the Exec Summary."),
 ("t", "A status or mapping in the Squads tab updates every role count: 2.6, 2.7 and the 4.x tabs."),
 ("t", "A person or cost in Sheet2 updates 2.1, 2.2, 2.3, the EGI roster, 2.4, the COE lines of 2.5 and 2.6, and tabs 4.12 to 4.14."),
 ("t", "A budget on 0.2 updates the budget and variance columns of 2.5, the draw-downs on 2.1 / 2.2 / 2.4, and left to fund everywhere."),
 ("t", "A Hire or Hold call on 4.x updates that tab's cost of roles marked Hire. Totals move only when a role is actually hired in the raw data."),
 ("h", "What is safe to change in the source data - and what is not"),
 ("t", "SAFE: any number (cost, base, rate), any status, any person's details. All model cells are live references."),
 ("t", "CAUTION: renaming a person in Squads without renaming them in Added data breaks that person's cost lookup (it falls back to the title average); 5.0 shows the drift."),
 ("t", "NOT SAFE ALONE: renaming a squad or portfolio (counts match the name as text), deleting rows (references break loudly and 5.0 stops footing - change the status instead), or adding rows at the bottom (rosters need a refresh). Ask for a refresh after structural edits."),
 ("h", "How to trust it"),
 ("t", "Every row labelled Check must read 0. All reconciliations, the Sheet2 update log, the mapping change log and the raw data audit live on 5.0 Data QA. Every role has exactly one 4.x home - if one were missed, the checks would not foot."),
]
gr_ = 4
for kind, txt in GUIDE2:
    if kind == "h":
        sc(gd, f"B{gr_}", txt, WHITEF, NAVY); gr_ += 1
    else:
        sc(gd, f"B{gr_}", txt, NORM, border=False, wrap=True)
        gd.row_dimensions[gr_].height = max(15, 15 * (1 + len(txt) // 115))
        gr_ += 1

wb.save(OUT)
print("stage 5 saved. order:", wb.sheetnames[:12], "...")
print("stage 4 saved. AUNZ:", AUNZ)
print("stage 3 saved. GM anchors:", {t: (a["hdr"], a["tot"], a["rost_hdr"]) for t, a in gm_anchor.items()})
json.dump(dict(DEDUP=DEDUP_ROW, TOT=TOT_ROW, UNM=UNM_ROW, LEAD=LEAD_ROW, COE_FIRST=COE_FIRST,
               RESTATE=RESTATE_ROW, EGI_MEMO=EGI_MEMO_ROW, EGI_TOT=EGI_TOT,
               PT_R1=R1, PT_CHECK=PT_CHECK, SAD_R1=S1, SAD_CHECK=SAD_CHECK,
               CY_R1=C1, CY_CHECK=CY_CHECK, N_SAD_COE=len(sad_coe),
               STATUS_DIFFS=len(status_diffs), S2_ONLY=len(s2_only), SQ_ONLY=len(sq_only),
               GM={t: dict(hdr=a["hdr"], tot=a["tot"], rost_hdr=a["rost_hdr"]) for t, a in gm_anchor.items()},
               AUNZ=AUNZ, W412=W412, W413=W413, LEAD_BLK=LEAD_BLK, UNM_BLK=UNM_BLK,
               CHANGELOG=[[c[0], c[1]] for c in CHANGELOG], AD_CHANGES=AD_CHANGES,
               AUNZ_ROW=r0),
          open(SCR + "anchors_v9.json", "w"), indent=1)

