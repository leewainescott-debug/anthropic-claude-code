#!/usr/bin/env python3
"""Faithful real-squad reroute of the 2.x working tabs.

Each tab keeps the owner's decision surface - Squad, Roles, Filled, Vacant,
Planning to hire, Vacancies remaining, Cost to hire vacant, Cost after vacancy
decisions - plus the two cost memos and the pricing notes. Squad rows are the
REAL squads from REVIEW; the roster below references REVIEW rows LIVE (name,
role, status, cost), so nothing is copied as a literal and the source formulas
(cost, status) stay authoritative. Archetype-vs-actual comparison lives at
portfolio level on 3.2 and design-squad level on 3.3, because real squads have
no per-squad archetype target. Squads / Added data stay in place (3.3 and the
COE cost rows still read them)."""
import openpyxl
from openpyxl.utils import column_index_from_string as ci, get_column_letter as gl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.views import Selection
from collections import defaultdict, OrderedDict

SRC="clean_reroute.xlsx"; OUT="clean_v3.xlsx"
wb=openpyxl.load_workbook(SRC)
REV="REVIEW - Complete Role Mapping"; Q="'"+REV+"'"
ws=wb[REV]
def c(r,cc): return ws.cell(r,ci(cc)).value

PMAP={"ampol retail":"Ampol Retail","retail":"Ampol Retail","z":"Z Retail","customer":"Customer",
"ampol customer":"Customer","enterprise data":"Enterprise Data","tdd":"TDD Group Functions",
"infrastructure":"Infrastructure","b2b & energy solutions":"Energy Solutions & B2B",
"commercial fuels":"Commercial Fuels","finance":"Finance","p&c":"P&C","p&c, finance & legal":"P&C",
"egi":"EGI & Central","coe - cyber, risk & operations":"COE Cyber",
"coe - partnering & transformation":"COE BP&T","coe - strategy, architecture, data":"COE SA&D"}
SALIAS={"ampos":"AmPOS","customer, ai":"Customer AI","manuacturing group projects":"Manufacturing Group Projects",
"integration & process automation":"Integration & Process Automation","technology suport":"Technology Support",
"data - au":"Data AU","data - nz":"Data NZ","na":"(all roles)","":"(unassigned)"}
def nsq(s):
    s=" ".join(str(s or "").split()); return SALIAS.get(s.lower(), s)
def is_vac(r):  # matches MStatus vacant branch exactly
    return str(c(r,'Q') or "").strip().lower()=="v"
def mtab(r): return PMAP.get(str(c(r,'I') or "").strip().lower())

# ---- helper MSquad (normalised real squad) in AO; collect rows per tab/squad ----
AO=ci("AO"); ws.cell(1,AO).value="MSquad"
data=defaultdict(lambda:OrderedDict())
for r in range(2,ws.max_row+1):
    if not ws.cell(r,2).value:
        ws.cell(r,AO).value=None; continue
    sq=nsq(c(r,'K')); ws.cell(r,AO).value=sq
    t=mtab(r)
    if not t: continue
    data[t].setdefault(sq,[]).append(dict(dr=r, vac=is_vac(r)))
# clear the stray #N/A in REVIEW R445 (source column 'Unit' is otherwise blank)
if str(ws.cell(445,ci('R')).value or "").strip():
    ws.cell(445,ci('R')).value=None

# ---- styling ----
NAVY=PatternFill("solid",fgColor="1F4E79"); GREY=PatternFill("solid",fgColor="F2F2F2")
TOT=PatternFill("solid",fgColor="D9D9D9"); YEL=PatternFill("solid",fgColor="FFF2CC")
SQH=PatternFill("solid",fgColor="DDEBF7")
WF=Font(color="FFFFFF",bold=True); BF=Font(bold=True); NF=Font(); BL=Font(color="0000FF"); IT=Font(italic=True,size=10)
thin=Side(style="thin",color="BFBFBF"); BOX=Border(thin,thin,thin,thin)
M2='#,##0.00;[Red](#,##0.00);"-"'; D0='#,##0;[Red](#,##0);"-"'; DOL='#,##0'
CEN=Alignment(horizontal="center",vertical="center"); RT=Alignment(horizontal="right")
def sc(w,ref,v,font=NF,fill=None,al=None,fmt=None,box=True):
    x=w[ref]; x.value=v; x.font=font
    if fill:x.fill=fill
    if al:x.alignment=al
    if fmt:x.number_format=fmt
    if box:x.border=BOX
    return x

WT=OrderedDict([("2.1 Ampol Retail","Ampol Retail"),("2.2 Customer","Customer"),
("2.3 Enterprise Data","Enterprise Data"),("2.4 TDD Group Functions","TDD Group Functions"),
("2.5 P&C","P&C"),("2.6 Finance","Finance"),("2.7 Infrastructure","Infrastructure"),
("2.8 Energy Solutions & B2B","Energy Solutions & B2B"),("2.9 Commercial Fuels","Commercial Fuels"),
("2.10 Z Retail","Z Retail"),("2.11 TDD Cyber","COE Cyber"),("2.12 BP&T","COE BP&T"),
("2.13 SA&D","COE SA&D"),("2.14 EGI & Central","EGI & Central")])
SUMHDR=["Squad","Roles","Filled","Vacant","Planning to hire","Vacancies remaining",
        "Cost to hire vacant ($m)","Cost after vacancy decisions ($m)"]  # B..I
itot={}; htot={}
for tabname,mt in WT.items():
    w=wb[tabname]
    for row in list(w.iter_rows(min_row=3)):
        for cl in row:
            cl.value=None; cl.fill=PatternFill(); cl.border=Border()
            cl.alignment=Alignment(); cl.number_format="General"
    for dv in list(w.data_validations.dataValidation): w.data_validations.dataValidation.remove(dv)
    squads=data.get(mt,{})
    sc(w,"B2",f"{mt} - working copy",Font(bold=True,size=15),box=False)
    sc(w,"B4","Position by squad",WF,NAVY,box=False)
    for i,h in enumerate(SUMHDR): sc(w,f"{gl(2+i)}5",h,WF,NAVY,CEN)
    first=6; n=len(squads); total_row=first+n
    # layout: summary(n) + total + 2 memos + 3 notes + blank + roster title + roster header
    rost_hdr=total_row+8
    roster_rows={}; rr=rost_hdr+1
    for sq,roles in squads.items():
        roster_rows[sq]=(rr+1, rr+len(roles)); rr+=1+len(roles)
    dv=DataValidation(type="list",formula1='"Hire,Hold,Offshore"',allow_blank=False); w.add_data_validation(dv)
    r=first
    for sq,roles in squads.items():
        a,b=roster_rows[sq]; sqq=sq.replace('"','""')
        sc(w,f"B{r}",sq,NF,GREY if (r-first)%2 else None)
        sc(w,f"C{r}",f"=COUNTA($B${a}:$B${b})",NF,al=CEN,fmt=D0)
        sc(w,f"D{r}",f'=COUNTIF($D${a}:$D${b},"Filled")',NF,al=CEN,fmt=D0)
        sc(w,f"E{r}",f'=COUNTIF($D${a}:$D${b},"Vacant")',NF,al=CEN,fmt=D0)
        sc(w,f"F{r}",f'=COUNTIF($E${a}:$E${b},"Hire")',NF,al=CEN,fmt=D0)
        sc(w,f"G{r}",f'=E{r}-F{r}-COUNTIF($E${a}:$E${b},"Offshore")',NF,al=CEN,fmt=D0)
        sc(w,f"H{r}",f'=SUMIFS($F${a}:$F${b},$D${a}:$D${b},"Vacant")/1000000',NF,al=RT,fmt=M2)
        base=f'SUMIFS({Q}!$AA:$AA,{Q}!$AJ:$AJ,"{mt}",{Q}!$AO:$AO,"{sqq}",{Q}!$AK:$AK,"Filled")'
        dec=f'SUMIFS($F${a}:$F${b},$E${a}:$E${b},"Hire")+0.4*SUMIFS($F${a}:$F${b},$E${a}:$E${b},"Offshore")'
        sc(w,f"I{r}",f'=({base}+{dec})/1000000',NF,al=RT,fmt=M2)
        r+=1
    last=r-1
    sc(w,f"B{total_row}","Total",BF,TOT)
    for col in "CDEFG": sc(w,f"{col}{total_row}",f"=SUM({col}{first}:{col}{last})",BF,TOT,CEN,D0)
    for col in "HI": sc(w,f"{col}{total_row}",f"=SUM({col}{first}:{col}{last})",BF,TOT,RT,M2)
    itot[tabname]=f"I{total_row}"; htot[tabname]=f"H{total_row}"
    # memos + notes
    ra,rb=rost_hdr+1, rr-1  # full roster body span
    sc(w,f"B{total_row+1}","Cost to hire all vacancies ($m)",NF,box=False)
    sc(w,f"C{total_row+1}",f"=H{total_row}",NF,al=RT,fmt=M2,box=False)
    sc(w,f"B{total_row+2}","Cost of roles marked Hire ($m)",NF,box=False)
    sc(w,f"C{total_row+2}",f'=SUMIF($E${ra}:$E${rb},"Hire",$F${ra}:$F${rb})/1000000',NF,al=RT,fmt=M2,box=False)
    sc(w,f"B{total_row+3}","Vacant roles are priced at standard title rates - indicative until an offer is made.",IT,box=False)
    sc(w,f"B{total_row+4}","Leadership roles are funded via the portfolio overheads and sit on 3.3 FTE View, not here.",IT,box=False)
    sc(w,f"B{total_row+5}","Archetype targets are set per portfolio on 3.2 Total Cost and per design squad on 3.3 FTE View.",IT,box=False)
    # roster
    sc(w,f"B{rost_hdr-1}",f"{mt} FTE",Font(bold=True,size=12),box=False)
    for i,h in enumerate(["Name","Role","Status","Vacancy lever","Cost if hired ($)"]):
        sc(w,f"{gl(2+i)}{rost_hdr}",h,WF,NAVY,CEN)
    rr=rost_hdr+1
    for sq,roles in squads.items():
        sc(w,f"B{rr}",sq,BF,SQH)
        for col in "CDEF": sc(w,f"{col}{rr}","",BF,SQH)
        rr+=1
        for x in roles:
            dr=x["dr"]
            sc(w,f"B{rr}",f"={Q}!$B${dr}",NF)
            sc(w,f"C{rr}",f"={Q}!$C${dr}",NF)
            sc(w,f"D{rr}",f"={Q}!$AK${dr}",NF,al=CEN)
            if x["vac"]:
                cell=sc(w,f"E{rr}","Hold",BL,YEL,CEN); dv.add(cell)
                sc(w,f"F{rr}",f"={Q}!$AA${dr}",NF,al=RT,fmt=DOL)
            else:
                sc(w,f"E{rr}","",NF); sc(w,f"F{rr}","",NF)
            rr+=1
    for col,wd in {"B":30,"C":34,"D":11,"E":16,"F":16,"G":18,"H":20,"I":24}.items(): w.column_dimensions[col].width=wd

# ---- 3.2 F (cost after vacancy decisions) -> rebuilt tab I-total ----
tc=wb["3.2 Total Cost"]
F32={6:"2.1 Ampol Retail",7:"2.2 Customer",8:"2.3 Enterprise Data",9:"2.4 TDD Group Functions",
10:"2.5 P&C",11:"2.6 Finance",12:"2.7 Infrastructure",13:"2.8 Energy Solutions & B2B",
14:"2.9 Commercial Fuels",15:"2.10 Z Retail",20:"2.11 TDD Cyber"}
for r,t in F32.items():
    tc[f"F{r}"].value=f"='{t}'!${itot[t][0]}${itot[t][1:]}"

# ---- Exec C52 -> sum of rebuilt tab cost-to-hire totals ----
picks=list(F32.values())
wb["Exec Summary"]["C52"].value="="+"+".join(f"'{t}'!${htot[t][0]}${htot[t][1:]}" for t in picks)

# ---- 3.3 O/P: keep headers; decisions live on 2.x, so show the held baseline ----
f3=wb["3.3 FTE View"]
for r in range(7, 95):
    if f3.cell(r,9).value is not None:      # I col has a vacant count -> a detail row
        f3.cell(r,15).value=f"=I{r}"        # O vacancies remaining (pre-decision = vacant)
        f3.cell(r,16).value=f"=N{r}"        # P cost after decisions (held = actual)

# ---- 4.0 Data QA C258: status is column D on the rebuilt tabs ----
qa=wb["4.0 Data QA"]
qa["C258"].value=qa["C258"].value.replace("$F$1:$F$500","$D$1:$D$500")

# ---- remove ALL frozen panes cleanly (owner instruction) ----
for w in wb.worksheets:
    w.freeze_panes=None; w.sheet_view.pane=None
    w.sheet_view.selection=[Selection(activeCell="A1",sqref="A1")]

wb.save(OUT)
print("saved",OUT)
print(" I-totals:",{k:itot[k] for k in list(itot)[:4]})
print(" H-totals:",{k:htot[k] for k in list(htot)[:4]})
