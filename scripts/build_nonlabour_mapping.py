#!/usr/bin/env python3
"""Build TDD_NonLabour_Mapping.xlsx v2 on Lee's REAL operating model.

Base file = LEE'S UPLOADED WORKBOOK, loaded whole so his tabs keep their
values, formulas, formatting, column widths and hidden state. My tabs are
added alongside. Structure = docs/OP_MODEL_STRUCTURE.md (his chart, 26/08):
portfolio and platform only, never invented; anything without an honest
home lands in a named "Awaiting ruling" row.

Leaf truth (proven): SW 51,288,134.47 + HW 25,468,998.32 = 76,757,132.79
Parents excluded: SW {HCORP, BULKFUEL, MARKETING, SIG_ITEMS}, HW {HCORP, MARKETING}
"""
import json, warnings, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from collections import defaultdict
warnings.filterwarnings("ignore")

SRC = "/root/.claude/uploads/e550b440-3996-5abb-87e5-bafafe598f82/8a9f57ab-TDD_AU_Consolidated_2027_budget.xlsx"
ROLES = json.load(open("/tmp/claude-0/-home-user-anthropic-claude-code/e550b440-3996-5abb-87e5-bafafe598f82/scratchpad/roles.json"))
OUT = "deliverables/TDD_NonLabour_Mapping.xlsx"
PWD = "Tdd123"

NAVY="0F2E52"; INK="1D2939"; MUTE="667085"; BANDR="F6F8FB"; EMPH="E8EEF6"
INPUT="FFF4CC"; INPUTB="E3B505"; WHITE="FFFFFF"; TOTAL="E7E7E7"; FONT="Calibri"
AMBER="B45309"
MONEY='#,##0.00;(#,##0.00);""'

def f(size=11,bold=False,color=INK): return Font(name=FONT,size=size,bold=bold,color=color)
def fill(h): return PatternFill("solid",fgColor=h)
def hdr(c,t,align="center"):
    c.value=t; c.font=f(11,True,WHITE); c.fill=fill(NAVY)
    c.alignment=Alignment(horizontal=align,vertical="center",wrap_text=True)

OPEN="Awaiting ruling"
SW_PARENTS={"APPLHCORP","APPLBULKFUEL","APPLMARKETING","APPLSIG_ITEMS"}
HW_PARENTS={"APPLHCORP","APPLMARKETING"}

# sig-item cost centres land on REAL platforms, flagged sig funded
SIG_CC={
 "APOLLO":       ("B2B & Energy Solutions","Energy Solutions"),
 "CTRMSIG":      ("Commercial Fuels","CTRM"),
 "CRPOSSIG":     ("Ampol Retail","AmPOS"),
 "CUSTLOYSIG":   ("Ampol Customer","Ampol Loyalty & Martech"),
 "INTEMERSIG":   ("EG Integration","EGI"),
 "HEMERALD":     ("EG Integration","EGI"),
 "TDDEINTSIG":   ("EG Integration","EGI"),
 "PCTECHKCM":    ("P&C","P&C"),
 "PCTECHMYHR":   ("P&C","P&C"),
 "APPLSIG_ITEMS":("TDD","TDD"),   # HW leaf: the device fleet, sig funded
}
DIRECT_CC={
 "APPLB2B":       ("B2B & Energy Solutions","B2B"),
 "APPLAETOTAL":   ("B2B & Energy Solutions","Energy Solutions"),
 "INTEGM":        ("B2B & Energy Solutions","B2B"),
 "APPLFIREFINING":("Fuel Infrastructure","Manufacturing"),
 "APPLTDDIST":    ("Fuel Infrastructure","Future Fuels"),
 "APPLSUPPLYOPS": ("Commercial Fuels","Supply"),
 "APPLSUPFPO":    ("Commercial Fuels","Supply"),
 "APPLFIFINANCE": ("Finance","Finance"),
 "APPLHFIN":      ("Finance","Finance"),
 "APPLHHR":       ("P&C","P&C"),
 "APPLRCALSTORES":("Ampol Retail","Store Operations"),
}
# open questions: portfolio known or unknown, platform never invented
OPEN_CC={
 "APPLLUBRICANTS":("Commercial Fuels",OPEN,"Lubricants: no platform on the chart"),
 "APPLHLSA":      (OPEN,OPEN,"Legal & Secretariat: no home on the chart"),
 "APPLHALT":      (OPEN,OPEN,"Executive (ALT & CEO): no home on the chart"),
}
CC_HOME={"APPLH_INFOTECH":"TDD","APPLBULKFUEL":"Commercial Fuels","APPLMTOTAL":"Ampol Retail"}

EUC_WORDS=("LAPTOP","DESKTOP","MONITOR","DOCKING"," DOCK","SURFACE","IPAD","IPHONE",
           "PRINTER","TONER","MICROSOFT","MSFT","M365","OFFICE 365","BACKPACK","HEADSET","KEYBOARD")
SEC_WORDS=("TRELLIX","CROWDSTRIKE","ZSCALER","PALO ALTO","SPLUNK","NETSKOPE","PROOFPOINT",
           "OKTA","TENABLE","QUALYS","SENTINEL","DEFENDER","MIMECAST","CYBERARK","FORTINET")
CLOUD_WORDS=("AZURE","AWS","AMAZON WEB")
NET_WORDS=("TELSTRA","CISCO","MERAKI","SD-WAN","SDWAN","FIREWALL","SWITCH","ROUTER",
           "WIRELESS","WIFI","WI-FI","NETWORK")
DC_WORDS=("NTT","VMWARE","SERVER","STORAGE","BACKUP","VEEAM","DATA CENTRE","DATACENTRE","DATACENTER")
DATA_WORDS=("SNOWFLAKE","DATABRICKS","POWER BI","POWERBI","TABLEAU","INFORMATICA","COLLIBRA")
SAP_WORDS=("SAP ","SAP-","S/4","ARIBA","SUCCESSFACTOR"," BW ")
POS_WORDS=("AMPOS","POS ","EFTPOS","INGENICO","PINPAD","PIN PAD","VERIFONE")
CCENTRE_WORDS=("GENESYS","NICE LTD","TWILIO","CONTACT CENTRE","CONTACT CENTER")
SITE_WORDS=("GILBARCO","WAYNE","DISPENSER","TANK GAUGE","FORECOURT","CCTV","CAMERA","FUEL")
TOOL_WORDS=("ATLASSIAN","JIRA","CONFLUENCE","GITHUB","GITLAB","DEVOPS")

def classify(cc,ce_code,text,vendor):
    """(class, basis, group, portfolio, platform) - platforms only from the chart."""
    t=(text or "").upper(); v=(vendor or "").upper()
    if cc in SIG_CC:
        p,pl=SIG_CC[cc]
        return ("Sig funded","Sig item cost centre, recharges to its program","",p,pl)
    if cc in DIRECT_CC:
        p,pl=DIRECT_CC[cc]
        return ("Direct","Cost centre lands on one platform","",p,pl)
    if cc in OPEN_CC:
        p,pl,why=OPEN_CC[cc]
        return ("Open question",why,"",p,pl)
    if cc=="APPLCCO":
        return ("Decision","CCO function spans Fuels and B2B","CCO","Commercial Fuels",OPEN)
    if cc=="APPLBRNDCOMS":
        return ("Proposed","Brand and comms: proposed Ampol Customer, veto if wrong","",
                "Ampol Customer","Ampol Loyalty & Martech")
    if cc=="APPLGENERAL":
        return ("Decision","HighRadius / OneStream, Finance systems","GENERAL","Finance","Finance")
    home=CC_HOME.get(cc,"TDD")
    if str(ce_code) in ("801450","801440") or any(w in t for w in EUC_WORDS):
        return ("Decision","End-user device or per-user software","EUC","TDD","TDD")
    if any(w in t for w in SEC_WORDS) or any(w in v for w in SEC_WORDS):
        return ("Rule","Security product","","Cyber","Cyber")
    if "SERVICENOW" in t or "SERVICENOW" in v or "SERVICE NOW" in t:
        return ("Rule","Service management tooling","","TDD","TDD")
    if any(w in t for w in CLOUD_WORDS) or any(w in v for w in CLOUD_WORDS):
        return ("Rule","Cloud consumption","","TDD","TDD")
    if any(w in t for w in DC_WORDS) or "NTT" in v:
        return ("Rule","Data centre and compute","","TDD","TDD")
    if any(w in t for w in NET_WORDS) or "TELSTRA" in v:
        return ("Rule","Network and carriage","","TDD","TDD")
    if any(w in t for w in TOOL_WORDS):
        return ("Rule","Engineering tooling","","TDD","TDD")
    if any(w in t for w in DATA_WORDS):
        return ("Rule","Data platform product","","Enterprise Data","Data")
    if any(w in t for w in SAP_WORDS) or v.startswith("SAP"):
        return ("Open question","SAP estate: no ERP platform on the chart","",OPEN,OPEN)
    if "WAY4" in t:
        return ("Open question","Way4 cards: no cards platform on the chart","",OPEN,OPEN)
    if "SALESFORCE" in t or "SALESFORCE" in v:
        if cc=="APPLMTOTAL":
            return ("Decision","Marketing Salesforce","MKT","Ampol Customer","Ampol Loyalty & Martech")
        return ("Rule","Enterprise CRM","","Ampol Customer","Ampol Loyalty & Martech")
    if "ADOBE" in t or "ADOBE" in v:
        if cc=="APPLMTOTAL":
            return ("Decision","Marketing Adobe stack","MKT","Ampol Customer","Ampol Loyalty & Martech")
        return ("Rule","Digital experience product","","Ampol Customer","Ampol Digital")
    if any(w in t for w in POS_WORDS):
        return ("Rule","Store POS and payments","","Ampol Retail","Store Operations")
    if any(w in t for w in CCENTRE_WORDS):
        return ("Rule","Contact centre product","","Ampol Customer","Ampol Digital")
    if cc=="APPLMTOTAL" and any(w in t for w in SITE_WORDS):
        return ("Rule","Site and store equipment","","Ampol Retail","Store Operations")
    if "LOYALTY" in t and cc=="APPLMTOTAL":
        return ("Decision","Loyalty spend","MKT","Ampol Customer","Ampol Loyalty & Martech")
    return ("Open","No product tell yet, needs digging","",home,OPEN)

# ---------------- load LEE'S workbook whole ----------------
wb=openpyxl.load_workbook(SRC)   # formulas, styles, widths, hidden state preserved
his_sheets=list(wb.sheetnames)
print("his tabs, untouched:", his_sheets)

def leaf_rows(sheet, parents):
    ws=wb[sheet]; out=[]
    for i,r in enumerate(ws.iter_rows(min_row=15, values_only=True), start=15):
        cc=r[2]; amt=r[26]
        if cc in (None,"","Overall Result") or amt is None: continue
        try: a=float(amt)
        except: continue
        if cc in parents: continue
        out.append((i,cc,r[3],r[4],r[5],r[6],r[10],r[12],r[13],a))
    return out
sw=leaf_rows("SW Line Items",SW_PARENTS)
hw=leaf_rows("HW Line Items",HW_PARENTS)
sw_tot=sum(x[-1] for x in sw); hw_tot=sum(x[-1] for x in hw)
assert abs(sw_tot-51288134.47)<0.5, sw_tot
assert abs(hw_tot-25468998.32)<0.5, hw_tot
print(f"leaves: SW {len(sw)} {sw_tot:,.2f} | HW {len(hw)} {hw_tot:,.2f}")

records=[]; stats=defaultdict(float)
for tag,rows in (("S",sw),("H",hw)):
    tab="SW Line Items" if tag=="S" else "HW Line Items"
    for (rowno,cc,ccname,ce,cename,doc,vend,post,text,a) in rows:
        cls,basis,grp,dp,dpl=classify(cc,ce,text,vend)
        records.append([f"{tag}{rowno}",tab,rowno,cc,ccname,ce,cename,doc,
                        (vend if vend not in (None,"","Not assigned","#") else ""),
                        str(post)[:10] if post else "",(text or "")[:80],a,cls,basis,grp,dp,dpl])
        stats[cls]+=a
print("class $m:", {k:round(v/1e6,3) for k,v in sorted(stats.items())})

# ---------------- 2 Decisions ----------------
dec=wb.create_sheet("2 Decisions")
dec.sheet_view.showGridLines=False
dec.column_dimensions["A"].width=2
for col,w in (("B",44),("C",30),("D",14),("E",64)):
    dec.column_dimensions[col].width=w
dec["B2"]="Decisions and open questions"; dec["B2"].font=f(16,True,NAVY)
dec["B3"]="Four dropdowns move money live. Below them, the named questions where your chart has no home for a cost, with the dollars attached. Nothing is invented: those lines sit in Awaiting ruling rows until you answer."
dec["B3"].font=f(11,color=MUTE); dec["B3"].alignment=Alignment(wrap_text=True)
dec.row_dimensions[3].height=30
DEC_ROWS={}
decisions=[
 ("EUC","End-user devices and per-user software (laptops, monitors, Microsoft per-user)",
  ["Central: TDD platform","Follow the cost centre portfolio"],
  "Central puts the whole device and workplace estate on the TDD platform. Follow leaves each device with the portfolio whose cost centre bought it. The sig-funded device fleet stays sig funded either way."),
 ("MKT","Marketing digital spend (Salesforce, Adobe, loyalty bought through the marketing cost centre)",
  ["Ampol Customer","Ampol Retail"],
  "Ampol Customer lands it on Ampol Loyalty & Martech. Ampol Retail keeps it in the Retail portfolio on Above Store."),
 ("CCO","Chief Commercial Officer function software",
  ["Commercial Fuels","B2B & Energy Solutions"],
  "Portfolio call only; the platform stays an Awaiting ruling row either way."),
 ("GENERAL","Non-functional general centre (HighRadius, OneStream)",
  ["Finance","TDD"],
  "Almost all Finance systems by vendor and text."),
]
hdr(dec.cell(4,2),"Decision",align="left"); hdr(dec.cell(4,3),"Your call (dropdown)")
hdr(dec.cell(4,4),"Status"); hdr(dec.cell(4,5),"What it moves",align="left")
r=5
for key,label,options,note in decisions:
    dec.cell(r,2,label).font=f(11); dec.cell(r,2).alignment=Alignment(wrap_text=True,vertical="center")
    c=dec.cell(r,3,options[0]); c.font=f(11,True); c.fill=fill(INPUT)
    c.border=Border(*[Side("thin",color=INPUTB)]*4); c.protection=Protection(locked=False)
    c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    dv=DataValidation(type="list", formula1='"'+",".join(options)+'"', allow_blank=False)
    dec.add_data_validation(dv); dv.add(f"C{r}")
    dec.cell(r,4,"PROPOSED").font=f(10,True,AMBER); dec.cell(r,4).alignment=Alignment(horizontal="center",vertical="center")
    dec.cell(r,5,note).font=f(10,color=MUTE); dec.cell(r,5).alignment=Alignment(wrap_text=True,vertical="center")
    dec.row_dimensions[r].height=40
    DEC_ROWS[key]=r; r+=1

# open questions with dollars
oq=defaultdict(float)
for rec in records:
    if rec[12]=="Open question": oq[rec[13]]+=rec[11]
    if rec[12]=="Open": oq["No product tell yet, needs digging"]+=rec[11]
    if rec[12]=="Proposed": oq["PROPOSED Brand and comms to Ampol Customer, veto if wrong"]+=rec[11]
QR=r+1
dec.cell(QR,2,"Open questions, with the money attached").font=f(12,True,NAVY)
QR+=1
hdr(dec.cell(QR,2),"Question",align="left"); hdr(dec.cell(QR,3),"$ (12 months)")
for q,amt in sorted(oq.items(), key=lambda kv:-abs(kv[1])):
    QR+=1
    dec.cell(QR,2,q).font=f(11); dec.cell(QR,2).alignment=Alignment(wrap_text=True)
    c=dec.cell(QR,3,amt); c.number_format=MONEY; c.font=f(11); c.alignment=Alignment(horizontal="right")
dec.protection=openpyxl.worksheet.protection.SheetProtection(sheet=True,password=PWD,
    selectLockedCells=False,selectUnlockedCells=False)
DEC_SHEET="'2 Decisions'"

# ---------------- 3 Ledger ----------------
lg=wb.create_sheet("3 Ledger")
lg.sheet_view.showGridLines=False
COLS=["Line ID","Raw tab","Raw row","Cost centre","Cost centre name","Cost element",
      "Cost element name","CO document","Vendor","Posting date","Document text",
      "Total $ (12 months)","Class","Basis","Decision","Portfolio","Platform"]
widths=[9,13,8,17,20,9,19,12,20,11,42,13,12,34,9,22,24]
lg["B1"]="Every line, placed on the operating model"; lg["B1"].font=f(16,True,NAVY)
lg["B2"]="One row per leaf line, Jul-2025 to Jun-2026 actuals. Line ID = raw tab row (S15 is row 15 of SW Line Items). Portfolio and Platform come only from the chart; Awaiting ruling marks the open questions. Shows the landing under the proposed decisions; refreshed when your rulings land."
lg["B2"].font=f(10,color=MUTE)
HR=4
for j,(t,w) in enumerate(zip(COLS,widths)):
    col=2+j
    hdr(lg.cell(HR,col),t, align="left" if t in ("Cost centre name","Cost element name","Document text","Basis","Portfolio","Platform","Vendor") else "center")
    lg.column_dimensions[get_column_letter(col)].width=w
lg.row_dimensions[HR].height=30
row=HR+1
for rec in records:
    vals=rec[:15]+[rec[15],rec[16]]
    for j,v in enumerate(vals):
        c=lg.cell(row,2+j,v); c.font=f(10)
        if j==10: c.alignment=Alignment(horizontal="left")
        if j==11: c.number_format='#,##0.00;(#,##0.00)'
        if j==16 and v==OPEN: c.font=f(10,True,AMBER)
        if j==15 and v==OPEN: c.font=f(10,True,AMBER)
    row+=1
LG_LAST=row-1
lg.auto_filter.ref=f"B{HR}:R{LG_LAST}"
lg.freeze_panes=f"E{HR+1}"
lg.protection=openpyxl.worksheet.protection.SheetProtection(sheet=True,password=PWD,
    autoFilter=False,sort=False,selectLockedCells=False,selectUnlockedCells=False)
print("ledger rows:",LG_LAST-HR)

# ---------------- 4 Profile (block-table, live off the 4 decisions) ----------------
base=defaultdict(float)
euc_total=0.0; euc_by_home=defaultdict(float)
mkt_total=0.0; cco_t=0.0; general_t=0.0
for rec in records:
    grp=rec[14]; amt=rec[11]; cc=rec[3]
    if grp=="EUC":
        euc_total+=amt
        home={"APPLMTOTAL":"Ampol Retail","APPLBULKFUEL":"Commercial Fuels"}.get(cc,CC_HOME.get(cc,"TDD"))
        euc_by_home[home]+=amt
    elif grp=="MKT": mkt_total+=amt
    elif grp=="CCO": cco_t+=amt
    elif grp=="GENERAL": general_t+=amt
    else: base[(rec[15],rec[16])]+=amt

DEC=lambda k: f"{DEC_SHEET}!$C${DEC_ROWS[k]}"
terms=defaultdict(list)
terms[("TDD","TDD")].append(f'IF(LEFT({DEC("EUC")},7)="Central",{euc_total!r},{euc_by_home.get("TDD",0.0)!r})')
for home,v in euc_by_home.items():
    if home=="TDD": continue
    terms[(home,"End user devices")].append(f'IF(LEFT({DEC("EUC")},7)="Central",0,{v!r})')
terms[("Ampol Customer","Ampol Loyalty & Martech")].append(f'IF({DEC("MKT")}="Ampol Customer",{mkt_total!r},0)')
terms[("Ampol Retail","Above Store")].append(f'IF({DEC("MKT")}="Ampol Retail",{mkt_total!r},0)')
terms[("Commercial Fuels",OPEN)].append(f'IF({DEC("CCO")}="Commercial Fuels",{cco_t!r},0)')
terms[("B2B & Energy Solutions",OPEN)].append(f'IF({DEC("CCO")}="B2B & Energy Solutions",{cco_t!r},0)')
terms[("Finance","Finance")].append(f'IF({DEC("GENERAL")}="Finance",{general_t!r},0)')
terms[("TDD","Finance systems (if ruled here)")].append(f'IF({DEC("GENERAL")}="TDD",{general_t!r},0)')

pf=wb.create_sheet("4 Profile")
pf.sheet_view.showGridLines=False
pf.column_dimensions["A"].width=2
for col,w in (("B",30),("C",34),("D",15),("E",15),("F",15)):
    pf.column_dimensions[col].width=w
pf["B1"]="Cost profile on the operating model"; pf["B1"].font=f(16,True,NAVY)
pf["B2"]="A$, 12 months of AU hardware and software actuals (Jul-2025 to Jun-2026), labour at the FY27 model price for AU roles. Moves live with the Decisions tab. Awaiting ruling rows are your open questions, never guesses. EG and Z portfolios show nothing because this extract is AU non-labour only."
pf["B2"].font=f(10,color=MUTE)

PORT_ORDER=["EG Integration","EG","Z Retail / Commercial, Supply","Ampol Retail",
            "Z Energy (Digital)","Ampol Customer","Commercial Fuels",
            "B2B & Energy Solutions","Fuel Infrastructure","Finance","P&C",
            "TDD","Enterprise Data","Cyber",OPEN]
allpairs=set(base)|set(terms)
bypf=defaultdict(list)
for p_,pl_ in sorted(allpairs): bypf[p_].append(pl_)

# labour AU by portfolio, folded to the chart's names
lab=defaultdict(float); lab_other=0.0
FOLD={"RETAIL":"Ampol Retail","Retail":"Ampol Retail","Ampol Customer":"Ampol Customer",
 "Commercial Fuels":"Commercial Fuels","B2B & Energy Solutions":"B2B & Energy Solutions",
 "Infrastructure":"Fuel Infrastructure","Enterprise Data":"Enterprise Data",
 "Finance":"Finance","P&C":"P&C","P&C, Finance & Legal":"P&C","TDD":"TDD",
 "EGI":"EG Integration","EGI Integration":"EG Integration",
 "Z":"Z Retail / Commercial, Supply","Z ENERGY (DIGITAL)":"Z Energy (Digital)"}
for rec_r in ROLES:
    if rec_r.get("country")!="AU": continue
    p_=FOLD.get((rec_r.get("portfolio") or "").strip())
    amt_=float(rec_r.get("today") or 0)*1e6
    if p_: lab[p_]+=amt_
    else: lab_other+=amt_

HRP=4
hdr(pf.cell(HRP,2),"Portfolio",align="left"); hdr(pf.cell(HRP,3),"Platform",align="left")
hdr(pf.cell(HRP,4),"Non-labour $"); hdr(pf.cell(HRP,5),"Labour AU $"); hdr(pf.cell(HRP,6),"Total $")
pf.row_dimensions[HRP].height=22
r=HRP+1
port_total_rows=[]
for p_ in PORT_ORDER:
    plats=bypf.get(p_,[])
    if not plats and lab.get(p_,0)==0 and p_!=OPEN:
        plats=["(no cost in this extract)"]
    first=r
    for pl_ in plats:
        c3=pf.cell(r,3,pl_); c3.font=f(10)
        if pl_==OPEN or p_==OPEN: c3.font=f(10,True,AMBER)
        parts=[]
        b=base.get((p_,pl_),0.0)
        if b: parts.append(repr(b))
        parts+=terms.get((p_,pl_),[])
        form="="+("+".join(parts) if parts else "0")
        c=pf.cell(r,4,form); c.number_format=MONEY; c.font=f(10); c.alignment=Alignment(horizontal="right")
        r+=1
    pf.cell(r,2,p_+" total").font=f(11,True)
    c=pf.cell(r,4,f"=SUM(D{first}:D{r-1})"); c.number_format=MONEY; c.font=f(11,True); c.alignment=Alignment(horizontal="right")
    lv=lab.get(p_,0.0)
    c2=pf.cell(r,5,lv); c2.number_format=MONEY; c2.font=f(11,True); c2.alignment=Alignment(horizontal="right")
    c3=pf.cell(r,6,f"=D{r}+E{r}"); c3.number_format=MONEY; c3.font=f(11,True); c3.alignment=Alignment(horizontal="right")
    for cc_ in range(2,7): pf.cell(r,cc_).fill=fill(EMPH)
    port_total_rows.append(r)
    pf.cell(first,2,p_).font=f(11,True,NAVY if p_!=OPEN else AMBER)
    r+=2
gr=r
pf.cell(gr,2,"Total").font=f(12,True)
c=pf.cell(gr,4,"="+"+".join(f"D{x}" for x in port_total_rows)); c.number_format=MONEY; c.font=f(12,True); c.alignment=Alignment(horizontal="right")
c=pf.cell(gr,5,"="+"+".join(f"E{x}" for x in port_total_rows)); c.number_format=MONEY; c.font=f(12,True); c.alignment=Alignment(horizontal="right")
c=pf.cell(gr,6,"="+"+".join(f"F{x}" for x in port_total_rows)); c.number_format=MONEY; c.font=f(12,True); c.alignment=Alignment(horizontal="right")
for cc_ in range(2,7): pf.cell(gr,cc_).fill=fill(TOTAL)
chk=gr+1
pf.cell(chk,2,"Check against the leaf ledger total 76,757,132.79 (0 = ties)").font=f(9,color=MUTE)
c=pf.cell(chk,4,f"=D{gr}-76757132.79"); c.number_format='0.00;(0.00)'; c.font=f(9,color=MUTE); c.alignment=Alignment(horizontal="right")
sig_t=sum(rec[11] for rec in records if rec[12]=="Sig funded")
pf.cell(chk+1,2,"Of which sig funded, recharges to programs").font=f(9,color=MUTE)
c=pf.cell(chk+1,4,sig_t); c.number_format=MONEY; c.font=f(9,color=MUTE); c.alignment=Alignment(horizontal="right")
if lab_other>0:
    pf.cell(chk+2,2,"Labour AU not folded (portfolio blank or NA in the labour tab)").font=f(9,color=MUTE)
    c=pf.cell(chk+2,5,lab_other); c.number_format=MONEY; c.font=f(9,color=MUTE); c.alignment=Alignment(horizontal="right")
pf.protection=openpyxl.worksheet.protection.SheetProtection(sheet=True,password=PWD,
    selectLockedCells=False,selectUnlockedCells=False)

# ---------------- 5 Reconciliation ----------------
rc=wb.create_sheet("5 Reconciliation")
rc.sheet_view.showGridLines=False
rc.column_dimensions["A"].width=2
for col,w in (("B",52),("C",16),("D",58)):
    rc.column_dimensions[col].width=w
rc["B1"]="The bridge to the roughly 302m"; rc["B1"].font=f(16,True,NAVY)
rc["B2"]="Total enterprise technology spend per the FY26 dashboard: Ampol opex 189.1 + Ampol capex 48.1 + Z Energy opex 64.8 = 302.1m budget. What this workbook covers and every named gap. TBC means the dataset is not in this file."
rc["B2"].font=f(10,color=MUTE); rc["B2"].alignment=Alignment(wrap_text=True)
rc.row_dimensions[2].height=28
hdr(rc.cell(4,2),"Block",align="left"); hdr(rc.cell(4,3),"$"); hdr(rc.cell(4,4),"Where it stands",align="left")
rows=[
 ("AU software, 12 months of actuals (this file)",51288134.47,"Placed line by line on the Ledger tab."),
 ("AU hardware, 12 months of actuals (this file)",25468998.32,"Placed line by line on the Ledger tab."),
 ("Labour, AU roles at the FY27 model price",sum(lab.values())+lab_other,"From the labour model, beside non-labour on the Profile."),
 ("Network (Communications-Data, Mobile)","TBC","Not in this extract. 3.1m at B2026 enterprise level."),
 ("Outside services","TBC","Not in this extract. 31.1m at B2026 enterprise level."),
 ("Depreciation","TBC","Not in this extract. 5.8m per the budget detail."),
 ("Ampol capex","TBC","48.1m budget FY26, separate dataset."),
 ("Z Energy and all of NZ (opex 64.8m budget)","TBC","Entire NZ side absent from this AU file."),
 ("NZ labour","TBC","Sits in the labour model at live FX, joins with the Z extract."),
]
r=5
for label,val,note in rows:
    rc.cell(r,2,label).font=f(11)
    c=rc.cell(r,3,val)
    if val=="TBC":
        c.font=f(10,color=MUTE); c.alignment=Alignment(horizontal="center")
    else:
        c.number_format=MONEY; c.font=f(11); c.alignment=Alignment(horizontal="right")
    rc.cell(r,4,note).font=f(10,color=MUTE); rc.cell(r,4).alignment=Alignment(wrap_text=True)
    if r%2==0:
        for cc_ in range(2,5): rc.cell(r,cc_).fill=fill(BANDR)
    r+=1
rc.cell(r,2,"Dashboard benchmark: total enterprise technology spend").font=f(11,True)
c=rc.cell(r,3,302100000); c.number_format=MONEY; c.font=f(11,True); c.alignment=Alignment(horizontal="right")
for cc_ in range(2,5): rc.cell(r,cc_).fill=fill(EMPH)
rc.protection=openpyxl.worksheet.protection.SheetProtection(sheet=True,password=PWD,
    selectLockedCells=False,selectUnlockedCells=False)

# ---------------- 1 Start here ----------------
st=wb.create_sheet("1 Start here")
st.sheet_view.showGridLines=False
st.column_dimensions["A"].width=2
for col in "BCDEFGHIJ": st.column_dimensions[col].width=13
st.merge_cells("B2:J3")
tb=st["B2"]; tb.value="Non-labour cost on the operating model"
tb.font=f(18,True,WHITE); tb.alignment=Alignment(horizontal="left",vertical="center",indent=1)
for rr in (2,3):
    for cc_ in range(2,11): st.cell(rr,cc_).fill=fill(NAVY)
st["B5"]="Every AU hardware and software line placed by portfolio and platform, from your structure chart and nowhere else."
st["B5"].font=f(11,color=MUTE)
st["B7"]="How to use it"; st["B7"].font=f(12,True,NAVY)
steps=[
 "1.  Open 2 Decisions: four dropdowns move money live, and the open questions sit beneath them with the dollars attached.",
 "2.  3 Ledger is every line: 49,910 rows, each traced to its raw tab row, with the class, the basis and the landing.",
 "3.  4 Profile is cost by portfolio and platform, labour beside non-labour, checked against the ledger to the cent.",
 "4.  5 Reconciliation is the bridge to the roughly 302m with every missing block named.",
 "5.  Your original tabs sit at the end exactly as you uploaded them: values, formulas and formatting untouched.",
]
for i,s in enumerate(steps): st.cell(8+i,2,s).font=f(11)
st["B15"]="The rules of this build"; st["B15"].font=f(12,True,NAVY)
notes=[
 "Portfolios and platforms come only from your chart. Where your chart has no home for a cost, the line sits in an Awaiting ruling row and the question is listed with its dollars. Nothing is invented.",
 "Sig-item centres land on their real project platforms (CTRM, AmPOS, EGI) and stay flagged as sig funded, recharging to programs.",
 "The whole book re-adds to 76,757,132.79, the proven leaf total of your export, and the Profile checks itself against it.",
 "Figures are A$ actuals Jul-2025 to Jun-2026. Locked so the sums cannot break; only the yellow cells take typing. Password Tdd123.",
]
for i,s in enumerate(notes):
    st.cell(16+i,2,s).font=f(10); st.cell(16+i,2).alignment=Alignment(wrap_text=True)
    st.row_dimensions[16+i].height=26
st.protection=openpyxl.worksheet.protection.SheetProtection(sheet=True,password=PWD,
    selectLockedCells=False,selectUnlockedCells=False)

# order: my tabs first, his tabs after, untouched
order=["1 Start here","2 Decisions","3 Ledger","4 Profile","5 Reconciliation"]+his_sheets
wb._sheets.sort(key=lambda s: order.index(s.title))
wb.active=wb.sheetnames.index("1 Start here")
wb.calculation.fullCalcOnLoad=True
wb.save(OUT)
print("saved",OUT)
