#!/usr/bin/env python3
"""Build the offshore scenario builder workbook: a usable role picker.

Answers from Lee (12 Aug): Excel file; each role its own start month;
three separate baskets to compare; yes to a month-by-month picture.

The math is preserved exactly from the existing model:
  reduction/yr = cost today - 0.4 * full onshore cost  (0 if not moved, or vendor)
  FY split is by calendar year (Ampol Dec year-end): 2026, 2027.
"""
import json, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import LineChart, BarChart, Reference, Series
from openpyxl.utils import get_column_letter

ROLES = json.load(open("/tmp/claude-0/-home-user-anthropic-claude-code/e550b440-3996-5abb-87e5-bafafe598f82/scratchpad/roles.json"))
OUT = "deliverables/TDD_Offshore_Scenario_Builder.xlsx"
PWD = "Tdd123"

# ---- palette / house style (model style: Calibri, navy headers) ----
NAVY   = "0F2E52"
BAND   = "DDE3EC"   # light group band
TOTAL  = "E7E7E7"   # grey total
INPUT  = "FFF2CC"   # yellow input
GREENF = "C6EFCE"   # toggle "on"
GREENT = "006100"
WHITE  = "FFFFFF"
BANDR  = "F2F4F8"   # alt-row banding
FONT   = "Calibri"

MONEY = '#,##0.000;(#,##0.000);-'
MON2  = '#,##0.00;(#,##0.00);-'
FTEF  = '#,##0.0;(#,##0.0);-'
CNT   = '#,##0;(#,##0);-'

thin = Side(style="thin", color="B7BFCB")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

def f(size=11, bold=False, color="000000"):
    return Font(name=FONT, size=size, bold=bold, color=color)
def fill(hexv):
    return PatternFill("solid", fgColor=hexv)

def hdr(cell, text):
    cell.value = text
    cell.font = f(11, True, WHITE)
    cell.fill = fill(NAVY)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = border

def title(ws, cellref, text):
    ws[cellref] = text
    ws[cellref].font = f(16, True, NAVY)

MONTHS = ["Jul-2026","Aug-2026","Sep-2026","Oct-2026","Nov-2026","Dec-2026",
          "Jan-2027","Feb-2027","Mar-2027","Apr-2027","May-2027","Jun-2027",
          "Jul-2027","Aug-2027","Sep-2027","Oct-2027","Nov-2027","Dec-2027"]

wb = openpyxl.Workbook()

# =====================================================================
# 4 Lists  (month list + named range)  -- build first, hidden
# =====================================================================
wl = wb.active
wl.title = "4 Lists"
wl["B1"] = "Month list, used by the dropdowns. Do not change."
wl["B1"].font = f(11, True)
for i, m in enumerate(MONTHS):
    wl.cell(2+i, 2, m).font = f(11)
    wl.cell(2+i, 3, i+1).font = f(11)
wb.defined_names.add(openpyxl.workbook.defined_name.DefinedName(
    "MonthList", attr_text="'4 Lists'!$B$2:$B$19"))
wl.sheet_state = "hidden"
wl.column_dimensions["B"].width = 14

# =====================================================================
# 1 Start here
# =====================================================================
ws = wb.create_sheet("1 Start here")
title(ws, "B2", "Offshore scenario builder")
ws["B3"] = "Pick roles into three baskets, give each role a start month, and watch the cost come out."
ws["B3"].font = f(11)

ws["B5"] = "How to use it"
ws["B5"].font = f(12, True, NAVY)
steps = [
 "1.  Go to the Pick roles tab.",
 "2.  Use the filter arrows on the header row to narrow the 540 roles (by portfolio, squad, AU or NZ, status).",
 "3.  For a role you want to move, set In A, In B or In C to Yes and choose its start month. A role can sit in more than one basket.",
 "4.  A toggled role turns green. The Compare tab updates as you go.",
 "5.  Name your three baskets in the yellow cells below.",
]
for i, s in enumerate(steps):
    ws.cell(6+i, 2, s).font = f(11)

r = 13
ws.cell(r, 2, "Name your baskets").font = f(12, True, NAVY)
names = ["Example basket", "Basket B", "Basket C"]
for i in range(3):
    ws.cell(r+1+i, 2, f"Basket {'ABC'[i]} name").font = f(11)
    c = ws.cell(r+1+i, 3, names[i])
    c.font = f(11, True); c.fill = fill(INPUT); c.border = border
    c.protection = Protection(locked=False)
NAME_CELLS = ["'1 Start here'!$C$14", "'1 Start here'!$C$15", "'1 Start here'!$C$16"]

r = 19
ws.cell(r, 2, "How the money works").font = f(12, True, NAVY)
notes = [
 "A role that moves offshore costs 40% of its full onshore cost. The cost taken out is what it costs today less that 40%.",
 "Roles on vendor day rates are left as they are. A held role costs nothing today, so moving it shows as added cost.",
 "The year is the calendar year (Ampol's December year-end): 2026 and 2027. The horizon runs Jul-2026 to Dec-2027.",
 "NZ roles are already converted to AUD at the planning rate. Costs are in $m.",
 "The locked tabs hold the sums. To unlock them the password is Tdd123.",
 "Built from the role mapping in TDD_FY27_Cost_Calc_0608 (11 August 2026), copied in full on the REVIEW tab.",
]
for i, s in enumerate(notes):
    ws.cell(r+1+i, 2, s).font = f(11)
ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 16   # holds the "Basket A name" labels without clipping
ws.sheet_view.showGridLines = False
ws.column_dimensions["C"].width = 24   # the yellow name inputs

# =====================================================================
# 2 Pick roles
# =====================================================================
pk = wb.create_sheet("2 Pick roles")
title(pk, "B1", "Pick roles for each basket")
pk["B2"] = "Filter to find roles, then set In A / In B / In C to Yes and choose a start month. Only the green-capable cells can be edited."
pk["B2"].font = f(11)

# columns (start at B=2 so there is a left margin like the model)
COLS = [
 ("Role ID", 10, "id"),
 ("Role", 30, "role"),
 ("Portfolio", 20, "portfolio"),
 ("Squad or line", 24, "squad"),
 ("AU / NZ", 8, "country"),
 ("Status", 22, "status"),
 ("FTE", 6, "fte"),
 ("Cost today ($m)", 12, "today"),
 ("In A?", 8, "inA"),
 ("A start month", 12, "amon"),
 ("In B?", 8, "inB"),
 ("B start month", 12, "bmon"),
 ("In C?", 8, "inC"),
 ("C start month", 12, "cmon"),
 # hidden helpers
 ("Full onshore ($m)", 12, "onshore"),
 ("Reduction A ($m/yr)", 12, "redA"),
 ("pos A", 6, "posA"),
 ("2026 A ($m)", 10, "fy26A"),
 ("2027 A ($m)", 10, "fy27A"),
 ("Reduction B ($m/yr)", 12, "redB"),
 ("pos B", 6, "posB"),
 ("2026 B ($m)", 10, "fy26B"),
 ("2027 B ($m)", 10, "fy27B"),
 ("Reduction C ($m/yr)", 12, "redC"),
 ("pos C", 6, "posC"),
 ("2026 C ($m)", 10, "fy26C"),
 ("2027 C ($m)", 10, "fy27C"),
 ("In any basket?", 12, "any"),
]
key2col = {}
HDR_ROW = 3
FIRST = 4
LAST = FIRST + len(ROLES) - 1   # 4 .. 543
for i, (label, width, key) in enumerate(COLS):
    col = 2 + i
    key2col[key] = col
    hdr(pk.cell(HDR_ROW, col), label)
    pk.column_dimensions[get_column_letter(col)].width = width

def L(key):  # column letter for a key
    return get_column_letter(key2col[key])

# worked example preload for basket A
EXAMPLE = {"R0002":"Oct-2026","R0029":"Oct-2026","R0351":"Jan-2027","R0468":"Oct-2026"}

for idx, rec in enumerate(ROLES):
    row = FIRST + idx
    def put(key, val, num=None, inp=False, center=False):
        c = pk.cell(row, key2col[key], val)
        c.font = f(11); c.border = border
        if num: c.number_format = num
        if center: c.alignment = Alignment(horizontal="center")
        if inp:
            c.fill = fill(INPUT); c.protection = Protection(locked=False)
        return c
    put("id", rec["role_id"], center=True)
    put("role", rec["role"])
    put("portfolio", rec["portfolio"])
    put("squad", rec["squad"])
    put("country", rec["country"], center=True)
    put("status", rec["status"])
    put("fte", rec["fte"], num=FTEF, center=True)
    put("today", rec["today"], num=MONEY)
    put("onshore", rec["onshore"], num=MONEY)
    # inputs
    ex = EXAMPLE.get(rec["role_id"])
    put("inA", "Yes" if ex else "No", inp=True, center=True)
    put("amon", ex if ex else None, inp=True, center=True)
    put("inB", "No", inp=True, center=True)
    put("bmon", None, inp=True, center=True)
    put("inC", "No", inp=True, center=True)
    put("cmon", None, inp=True, center=True)
    # helper formulas
    on = f"${L('onshore')}{row}"; td = f"${L('today')}{row}"; st = f"${L('status')}{row}"
    for tog, mon, red, pos, fy26, fy27 in [
        ("inA","amon","redA","posA","fy26A","fy27A"),
        ("inB","bmon","redB","posB","fy26B","fy27B"),
        ("inC","cmon","redC","posC","fy26C","fy27C")]:
        tg=f"${L(tog)}{row}"; mn=f"${L(mon)}{row}"
        pk.cell(row, key2col[red], f'=IF(AND({tg}="Yes",{st}<>"Vendor day rates"),{td}-0.4*{on},0)').number_format=MONEY
        pk.cell(row, key2col[pos], f'=IFERROR(MATCH({mn},MonthList,0),0)')
        p=f"${L(pos)}{row}"; rd=f"${L(red)}{row}"
        pk.cell(row, key2col[fy26], f'=IF(OR({tg}<>"Yes",{mn}=""),0,{rd}*MAX(0,7-{p})/12)').number_format=MONEY
        pk.cell(row, key2col[fy27], f'=IF(OR({tg}<>"Yes",{mn}=""),0,{rd}*IF({p}<=7,12,19-{p})/12)').number_format=MONEY
    pk.cell(row, key2col["any"],
            f'=IF(OR(${L("inA")}{row}="Yes",${L("inB")}{row}="Yes",${L("inC")}{row}="Yes"),"Yes","No")')
    for k in ("redA","posA","fy26A","fy27A","redB","posB","fy26B","fy27B","redC","posC","fy26C","fy27C","any"):
        pk.cell(row, key2col[k]).font=f(11); pk.cell(row, key2col[k]).border=border

# data validations
dv_yn = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False)
dv_mo = DataValidation(type="list", formula1="MonthList", allow_blank=True)
pk.add_data_validation(dv_yn); pk.add_data_validation(dv_mo)
for key in ("inA","inB","inC"):
    dv_yn.add(f"{L(key)}{FIRST}:{L(key)}{LAST}")
for key in ("amon","bmon","cmon"):
    dv_mo.add(f"{L(key)}{FIRST}:{L(key)}{LAST}")

# conditional formatting: toggle green when Yes; row banding
for key in ("inA","inB","inC"):
    rng=f"{L(key)}{FIRST}:{L(key)}{LAST}"
    pk.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Yes"'],
        fill=fill(GREENF), font=Font(name=FONT, size=11, bold=True, color=GREENT)))
# subtle banding on the visible reference block (Role..Cost today) for even rows
band_range=f"{L('id')}{FIRST}:{L('cmon')}{LAST}"
pk.conditional_formatting.add(band_range, FormulaRule(formula=[f"AND(ISEVEN(ROW()),${L('any')}{FIRST}<>\"\")"],
    fill=fill(BANDR)))
# highlight the whole row's Role cell when it is in any basket
pk.conditional_formatting.add(f"{L('role')}{FIRST}:{L('role')}{LAST}",
    FormulaRule(formula=[f'${L("any")}{FIRST}="Yes"'], font=Font(name=FONT,size=11,bold=True,color=NAVY)))

# autofilter over the header + data
pk.auto_filter.ref = f"{L('id')}{HDR_ROW}:{L('any')}{LAST}"
# hide helper columns
for key in ("onshore","redA","posA","fy26A","fy27A","redB","posB","fy26B","fy27B","redC","posC","fy26C","fy27C"):
    pk.column_dimensions[get_column_letter(key2col[key])].hidden = True
# freeze: keep Role ID + Role and the header rows
pk.freeze_panes = f"{L('portfolio')}{FIRST}"
pk.sheet_view.showGridLines = False
pk.column_dimensions["A"].width = 2
# protect
pk.protection = openpyxl.worksheet.protection.SheetProtection(
    sheet=True, password=PWD, autoFilter=False, sort=False,
    formatCells=False, formatColumns=False, formatRows=False,
    selectLockedCells=False, selectUnlockedCells=False)

RD = LAST  # last data row for cross-sheet ranges

# =====================================================================
# 3 Compare
# =====================================================================
cp = wb.create_sheet("3 Compare")
title(cp, "B1", "Compare the three baskets")
cp["B2"] = "Everything here moves as you toggle roles on the Pick roles tab."
cp["B2"].font = f(11)
cp.sheet_view.showGridLines = False

def pkref(key):
    return f"'2 Pick roles'!${L(key)}${FIRST}:${L(key)}${RD}"

# header row of basket names
HR = 4
cp.cell(HR,2,"Measure").font=f(11,True,WHITE); cp.cell(HR,2).fill=fill(NAVY); cp.cell(HR,2).border=border
cp.cell(HR,2).alignment=Alignment(horizontal="left",vertical="center")
for j in range(3):
    c=cp.cell(HR,3+j, f"={NAME_CELLS[j]}")
    c.font=f(11,True,WHITE); c.fill=fill(NAVY); c.border=border
    c.alignment=Alignment(horizontal="center")
cp.column_dimensions["B"].width=30
for col in ("C","D","E"): cp.column_dimensions[col].width=18

# measures
tog=["inA","inB","inC"]; red=["redA","redB","redC"]; fy26=["fy26A","fy26B","fy26C"]; fy27=["fy27A","fy27B","fy27C"]
measures = [
 ("Roles in the basket", lambda j: f'=COUNTIF({pkref(tog[j])},"Yes")', CNT),
 ("People (FTE)",        lambda j: f'=SUMIF({pkref(tog[j])},"Yes",{pkref("fte")})', FTEF),
 ("Cost today ($m)",     lambda j: f'=SUMIF({pkref(tog[j])},"Yes",{pkref("today")})', MON2),
 ("Cost at the offshore rate ($m)", lambda j: f'=R{{cost}}', MON2),  # filled below
 ("Cost taken out, full year ($m)", lambda j: f'=SUM({pkref(red[j])})', MON2),
 ("Taken out in 2026 ($m)", lambda j: f'=SUM({pkref(fy26[j])})', MON2),
 ("Taken out in 2027 ($m)", lambda j: f'=SUM({pkref(fy27[j])})', MON2),
]
MROW0=5
for i,(label,fn,fmt) in enumerate(measures):
    row=MROW0+i
    lc=cp.cell(row,2,label); lc.font=f(11,True if label.startswith("Cost taken") else False); lc.border=border
    lc.fill=fill(BAND) if i in (4,) else PatternFill()
    for j in range(3):
        col=3+j
        if label.startswith("Cost at the offshore"):
            # cost today (row-2) minus cost taken out full year (row+1)
            today_row=MROW0+2; taken_row=MROW0+4
            form=f"={get_column_letter(col)}{today_row}-{get_column_letter(col)}{taken_row}"
        else:
            form=fn(j)
        c=cp.cell(row,col,form); c.font=f(11,True if label.startswith("Cost taken") else False)
        c.number_format=fmt; c.border=border; c.alignment=Alignment(horizontal="center")
        if i==4: c.fill=fill(BAND)

# portfolio breakdown
# build the exhaustive portfolio line list from data (fold RETAIL->Retail, blank->Not mapped)
ports=[]
seen=set()
for rec in ROLES:
    p=rec["portfolio"]
    key = "Retail" if (p or "").upper()=="RETAIL" else (p if p else "Not mapped")
    if key not in seen:
        seen.add(key); ports.append(key)
# stable-ish order: named first as encountered, Not mapped last
if "Not mapped" in ports:
    ports = [p for p in ports if p!="Not mapped"] + ["Not mapped"]

PB0 = MROW0 + len(measures) + 2
cp.cell(PB0,2,"Cost taken out by portfolio, full year ($m)").font=f(12,True,NAVY)
hr2=PB0+1
cp.cell(hr2,2,"Portfolio").font=f(11,True,WHITE); cp.cell(hr2,2).fill=fill(NAVY); cp.cell(hr2,2).border=border
for j in range(3):
    c=cp.cell(hr2,3+j, f"={NAME_CELLS[j]}"); c.font=f(11,True,WHITE); c.fill=fill(NAVY); c.border=border
    c.alignment=Alignment(horizontal="center")
pcol=pkref("portfolio")
for i,p in enumerate(ports):
    row=hr2+1+i
    cp.cell(row,2,p).font=f(11); cp.cell(row,2).border=border
    for j in range(3):
        col=3+j
        if p=="Retail":
            # SUMIFS is case-insensitive, so "Retail" already captures the RETAIL-spelled rows too
            form=(f'=SUMIFS({pkref(red[j])},{pcol},"Retail")')
        elif p=="Not mapped":
            form=(f'=SUM({pkref(red[j])})-SUMIF({pkref("any")},"skip",{pkref(red[j])})')  # placeholder, replaced below
        else:
            form=f'=SUMIFS({pkref(red[j])},{pcol},"{p}")'
        c=cp.cell(row,col,form); c.number_format=MONEY; c.font=f(11); c.border=border
        c.alignment=Alignment(horizontal="center")
# Not mapped = total - sum of named lines (guarantees the check ties)
named_rows=[hr2+1+i for i,p in enumerate(ports) if p!="Not mapped"]
nm_row=hr2+1+ports.index("Not mapped")
for j in range(3):
    col=get_column_letter(3+j)
    named_sum="+".join(f"{col}{rr}" for rr in named_rows)
    cp.cell(nm_row,3+j).value=f"=SUM({pkref(red[j])})-({named_sum})"
# total + check
tot_row=hr2+1+len(ports)
cp.cell(tot_row,2,"Total").font=f(11,True); cp.cell(tot_row,2).fill=fill(TOTAL); cp.cell(tot_row,2).border=border
chk_row=tot_row+1
cp.cell(chk_row,2,"Check against full year (should be 0)").font=f(11); cp.cell(chk_row,2).border=border
for j in range(3):
    col=get_column_letter(3+j)
    c=cp.cell(tot_row,3+j, f"=SUM({col}{hr2+1}:{col}{tot_row-1})"); c.number_format=MONEY
    c.font=f(11,True); c.fill=fill(TOTAL); c.border=border; c.alignment=Alignment(horizontal="center")
    taken_full=MROW0+4
    c2=cp.cell(chk_row,3+j, f"={col}{tot_row}-{col}{taken_full}"); c2.number_format=MONEY
    c2.font=f(11); c2.border=border; c2.alignment=Alignment(horizontal="center")

# ---- month-by-month run-rate block (to the right, feeds the chart) ----
MB_C = 8  # column H
cp.cell(HR, MB_C, "Month").font=f(11,True,WHITE); cp.cell(HR,MB_C).fill=fill(NAVY); cp.cell(HR,MB_C).border=border
for j in range(3):
    c=cp.cell(HR, MB_C+1+j, f"={NAME_CELLS[j]}"); c.font=f(11,True,WHITE); c.fill=fill(NAVY); c.border=border
    c.alignment=Alignment(horizontal="center")
cp.column_dimensions[get_column_letter(MB_C)].width=12
for j in range(3): cp.column_dimensions[get_column_letter(MB_C+1+j)].width=14
posk=["posA","posB","posC"]
for p in range(18):
    row=HR+1+p
    cp.cell(row, MB_C, f"='4 Lists'!$B${2+p}").font=f(11); cp.cell(row,MB_C).border=border
    for j in range(3):
        # run-rate reached by month index (p+1): sum annual reductions of roles started by then
        form=(f'=SUMPRODUCT(({pkref(posk[j])}>=1)*({pkref(posk[j])}<={p+1})*{pkref(red[j])})')
        c=cp.cell(row, MB_C+1+j, form); c.number_format=MONEY; c.font=f(11); c.border=border
        c.alignment=Alignment(horizontal="center")
MB_LAST=HR+18

# ---- charts ----
# line chart: run-rate ramp
lc=LineChart(); lc.title="Cost taken out, run-rate by month ($m/yr)"
lc.height=8; lc.width=18; lc.style=2
data=Reference(cp, min_col=MB_C+1, max_col=MB_C+3, min_row=HR, max_row=MB_LAST)
cats=Reference(cp, min_col=MB_C, max_col=MB_C, min_row=HR+1, max_row=MB_LAST)
lc.add_data(data, titles_from_data=True); lc.set_categories(cats)
for s in lc.series: s.smooth=False
lc.y_axis.title="$m a year"; lc.x_axis.title=None
lc.x_axis.delete=False; lc.y_axis.delete=False
cp.add_chart(lc, f"B{chk_row+3}")

# bar chart: 2026 vs 2027
bc=BarChart(); bc.type="col"; bc.title="Cost taken out by year ($m)"
bc.height=8; bc.width=10; bc.style=2
# tiny helper table for the bar (years down, baskets across)
BY0=chk_row+3
BY_C=8
cp.cell(BY0, BY_C, "Year").font=f(11,True,WHITE); cp.cell(BY0,BY_C).fill=fill(NAVY); cp.cell(BY0,BY_C).border=border
for j in range(3):
    c=cp.cell(BY0,BY_C+1+j, f"={NAME_CELLS[j]}"); c.font=f(11,True,WHITE); c.fill=fill(NAVY); c.border=border
    c.alignment=Alignment(horizontal="center")
for yi,(ylab,mrow) in enumerate([("2026",MROW0+5),("2027",MROW0+6)]):
    row=BY0+1+yi
    cp.cell(row,BY_C,ylab).font=f(11); cp.cell(row,BY_C).border=border
    for j in range(3):
        col=get_column_letter(3+j)
        c=cp.cell(row,BY_C+1+j, f"={col}{mrow}"); c.number_format=MON2; c.font=f(11); c.border=border
        c.alignment=Alignment(horizontal="center")
bdata=Reference(cp, min_col=BY_C+1, max_col=BY_C+3, min_row=BY0, max_row=BY0+2)
bcats=Reference(cp, min_col=BY_C, max_col=BY_C, min_row=BY0+1, max_row=BY0+2)
bc.add_data(bdata, titles_from_data=True); bc.set_categories(bcats)
bc.y_axis.title="$m"; bc.x_axis.delete=False; bc.y_axis.delete=False
cp.add_chart(bc, f"H{BY0+4}")

cp.protection = openpyxl.worksheet.protection.SheetProtection(
    sheet=True, password=PWD, selectLockedCells=False, selectUnlockedCells=False)

# order tabs: Start here, Pick roles, Compare, (Lists hidden)
wb.move_sheet("1 Start here", -(wb.sheetnames.index("1 Start here")))
order=["1 Start here","2 Pick roles","3 Compare","4 Lists"]
wb._sheets.sort(key=lambda s: order.index(s.title))
wb.active = wb.sheetnames.index("1 Start here")

wb.save(OUT)
print("saved", OUT)
print("rows:", FIRST, "..", RD, "= roles:", RD-FIRST+1)
print("portfolio lines:", ports)
