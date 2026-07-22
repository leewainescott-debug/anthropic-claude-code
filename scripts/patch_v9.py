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
# 2.4 SA&D roster: not in a squad, and not an Enterprise Data / leadership seat
# already carried in the portfolio model (those live on 1.3 / 3.0 FTE View)
sad_coe, sad_skip = [], []
for x in SAD:
    if not blankish(x["squad"]):
        sad_skip.append((x, "in squad " + x["squad"])); continue
    m = sq_match(x)
    if m and m["cls"] == "Leadership":
        sad_skip.append((x, "leadership - funded via overheads, on 3.0")); continue
    if m and m["N"] == "Enterprise Data":
        sad_skip.append((x, "Enterprise Data portfolio - on 1.3 / 4.3")); continue
    if not m and x["port"] == "enterprise data":
        sad_skip.append((x, "NEW in Sheet2 - Enterprise Data, no raw data row")); continue
    sad_coe.append(x)
print("2.4 roster:", len(sad_coe), "| skipped:", len(sad_skip))
for x, why in sad_skip:
    if "NEW" in why: print("  NEW:", x["name"], "-", x["title"])

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
    cat = ["TDD COE", "TDD Cyber"][r-6]
    sc(ws, f"B{r}", cat, BOLD)
    sc(ws, f"C{r}", f'=COUNTIF($G${C1}:$G${C2R},$B{r})', NORM, fmt="0", align="center")
    sc(ws, f"D{r}", f'=COUNTIFS($G${C1}:$G${C2R},$B{r},$F${C1}:$F${C2R},"Filled")', NORM, fmt="0", align="center")
    sc(ws, f"E{r}", f'=COUNTIFS($G${C1}:$G${C2R},$B{r},$F${C1}:$F${C2R},"Vacant")', NORM, fmt="0", align="center")
    sc(ws, f"F{r}", f'=SUMIF($G${C1}:$G${C2R},$B{r},$J${C1}:$J${C2R})', NORM, fmt=M2, align="right")
def cat_cy(rr, r2r):
    return (f'=IF({S2}!$G${r2r}="Service Op & Assurance","TDD Cyber","TDD COE")')
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
    sc(ws, f"B{r}", f"='{tab}'!$B$2", BOLD)
    sc(ws, f"C{r}", f"='{tab}'!$E${ecell}", NORM, fmt=M2, align="right")
    sc(ws, f"D{r}", f"='3.0 FTE View'!$G${ftrow}", NORM, fmt=M0, align="center")
    for col, st in (("E", "Filled"), ("G", "Vacant")):
        parts = "+".join(sumifs_port(r, st, c) for c in ("Squad", "Strategic Program", "Leadership"))
        sc(ws, f"{col}{r}", f"=({parts})/1000000", NORM, fmt=M2, align="right")
    for col, st in (("F", "Filled"), ("H", "Vacant")):
        parts = "+".join(countifs_port(r, st, c) for c in ("Squad", "Strategic Program", "Leadership"))
        sc(ws, f"{col}{r}", f"={parts}", NORM, fmt="0", align="center")
    sc(ws, f"I{r}", f"=E{r}+G{r}", NORM, fmt=M2, align="right")
    sc(ws, f"J{r}", f"=F{r}+H{r}", NORM, fmt="0", align="center")
    sc(ws, f"K{r}", f"=ROUND(I{r}-C{r},6)", NORM, fmt=M2, align="right")
    sc(ws, f"L{r}", f"=ROUND(J{r}-D{r},1)", NORM, fmt=M0, align="center")
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
for lab, arch, archfte, sheet, catcell, jr, fr, gr, dcnt, ecnt in coe_rows:
    sc(ws, f"B{r}", f"={lab}", BOLD)
    sc(ws, f"C{r}", f"={arch}", NORM, fmt=M2, align="right")
    sc(ws, f"D{r}", f"={archfte}" if archfte else '="-"', NORM, fmt=M0, align="center")
    sc(ws, f"E{r}", f'=SUMIFS({sheet}!{jr},{sheet}!{fr},"Filled",{sheet}!{gr},{catcell})', NORM, fmt=M2, align="right")
    sc(ws, f"F{r}", f"={dcnt}", NORM, fmt="0", align="center")
    sc(ws, f"G{r}", f'=SUMIFS({sheet}!{jr},{sheet}!{fr},"Vacant",{sheet}!{gr},{catcell})', NORM, fmt=M2, align="right")
    sc(ws, f"H{r}", f"={ecnt}", NORM, fmt="0", align="center")
    sc(ws, f"I{r}", f"=E{r}+G{r}", NORM, fmt=M2, align="right")
    sc(ws, f"J{r}", f"=F{r}+H{r}", NORM, fmt="0", align="center")
    sc(ws, f"K{r}", f"=ROUND(I{r}-C{r},6)", NORM, fmt=M2, align="right")
    sc(ws, f"L{r}", f'=IF(ISNUMBER(D{r}),ROUND(J{r}-D{r},1),"-")', NORM, fmt=M0, align="center")
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
                                f'"No archetype - review these roles")')
    ws.cell(tot, 7).value = f"=SUM(G{hdr+1}:G{tot-1})"
    ws.cell(tot, 8).value = f"=SUM(H{hdr+1}:H{tot-1})"
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
    ws.cell(rost_hdr, 5).value = "Call"
    # close the white gap: merge B:C on the summary block rows
    for r in [2] + list(range(hdr-1, tot+1)):
        if ws.cell(r, 3).value is None:
            rng = f"B{r}:C{r}"
            covered = any(rng == str(m) for m in ws.merged_cells.ranges)
            if not covered:
                try: ws.merge_cells(rng)
                except Exception: pass
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
 "B7": "Each portfolio tab (1.x) shows the squads, sizes, support %, budget draw-downs and what is left to fund.",
 "B8": "Next step: agree funding for what is left to fund, and decide which vacancies to hire or hold.",
 "B36": "Roles the archetypes allow - squads at their set sizes",
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
qnote("Zero rows were added to the raw data: the Squads tab in this workbook matches the uploaded copy cell for cell (538 rows, checked on every populated cell).")
qnote("16 mapping cells on 4 Commercial rows were changed in an earlier build (Unmapped to COE - Business Partnering): rows 260, 269, 270, 281, columns N to Q. The names below are live references to those rows.")
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

wb.save(OUT)
print("stage 3 saved. GM anchors:", {t: (a["hdr"], a["tot"], a["rost_hdr"]) for t, a in gm_anchor.items()})
json.dump(dict(DEDUP=DEDUP_ROW, TOT=TOT_ROW, UNM=UNM_ROW, LEAD=LEAD_ROW, COE_FIRST=COE_FIRST,
               RESTATE=RESTATE_ROW, EGI_MEMO=EGI_MEMO_ROW, EGI_TOT=EGI_TOT,
               PT_R1=R1, PT_CHECK=PT_CHECK, SAD_R1=S1, SAD_CHECK=SAD_CHECK,
               CY_R1=C1, CY_CHECK=CY_CHECK, N_SAD_COE=len(sad_coe),
               STATUS_DIFFS=len(status_diffs), S2_ONLY=len(s2_only), SQ_ONLY=len(sq_only),
               GM={t: dict(hdr=a["hdr"], tot=a["tot"], rost_hdr=a["rost_hdr"]) for t, a in gm_anchor.items()}),
          open(SCR + "anchors_v9.json", "w"), indent=1)

