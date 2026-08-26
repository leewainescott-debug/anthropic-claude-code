#!/usr/bin/env python3
"""Build the final mapping deliverable by ZIP surgery on Lee's own upload.

His five tabs stay BYTE-FOR-BYTE identical (verifiable by hash): every
original zip entry is copied as raw bytes. Only three package parts are
amended (workbook.xml, its rels, [Content_Types].xml) to register four
added sheets, generated as pure values with inline strings, no formulas,
no styles beyond defaults. Nothing else changes.
Added tabs: Mapping (per line), Totals, Reconciliation, Not 100% sure.
"""
import json, re, shutil, zipfile
from collections import defaultdict
from xml.sax.saxutils import escape

SRC="/root/.claude/uploads/e550b440-3996-5abb-87e5-bafafe598f82/8a9f57ab-TDD_AU_Consolidated_2027_budget.xlsx"
A="/tmp/claude-0/-home-user-anthropic-claude-code/e550b440-3996-5abb-87e5-bafafe598f82/scratchpad/nonlabour/agents"
OUT="deliverables/TDD_NonLabour_Mapping.xlsx"

rows=[]
with open(f"{A}/rows_final.jsonl") as fh:
    for line in fh: rows.append(json.loads(line))
assert len(rows)==49910
sw=sum(r["total"] for r in rows if r["tab"]=="SW Line Items")
hw=sum(r["total"] for r in rows if r["tab"]=="HW Line Items")
assert abs(sw-51288134.47)<0.005 and abs(hw-25468998.32)<0.005
dash=json.load(open(f"{A}/dashboard_controls.json"))

def cell(ref, v):
    if v is None or v=="": return ""
    if isinstance(v,(int,float)):
        return f'<c r="{ref}"><v>{repr(round(v,2)) if isinstance(v,float) else v}</v></c>'
    s=escape(str(v))
    return f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{s}</t></is></c>'

def col(i):  # 1 -> A
    s=""
    while i: i,r=divmod(i-1,26); s=chr(65+r)+s
    return s

def sheet_xml(rows_data, widths, freeze=True):
    parts=['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">']
    if freeze:
        parts.append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
    if widths:
        parts.append("<cols>"+"".join(f'<col min="{i+1}" max="{i+1}" width="{w}" customWidth="1"/>' for i,w in enumerate(widths))+"</cols>")
    parts.append("<sheetData>")
    for rn,vals in enumerate(rows_data, start=1):
        cs="".join(cell(f"{col(ci+1)}{rn}", v) for ci,v in enumerate(vals) if v not in (None,""))
        parts.append(f'<row r="{rn}">{cs}</row>')
    parts.append("</sheetData></worksheet>")
    return "\n".join(parts).encode("utf8")

# ---- Mapping tab ----
m_rows=[["Line ID","Tab","Raw row","Cost centre","Total $ (12 months)","Portfolio","Platform",
         "Class","Source","Basis","Open reason"]]
for r in rows:
    m_rows.append([r["id"],r["tab"],r["row"],r.get("cc",""),r["total"],r["portfolio"],r["platform"],
                   r["class"],r.get("source",""),(r.get("basis") or "")[:200],(r.get("open_reason") or "")[:200]])
mapping=sheet_xml(m_rows,[9,13,8,17,14,24,24,11,18,60,45])

# ---- Totals tab ----
agg=defaultdict(float); cnt=defaultdict(int)
for r in rows:
    agg[(r["portfolio"] or "(open)", r["platform"] or "(open)", r["class"])]+=r["total"]
    cnt[(r["portfolio"] or "(open)", r["platform"] or "(open)", r["class"])]+=1
t_rows=[["Totals by portfolio and platform. A$ actuals, the export's periods 007.2025 to 006.2026. Values only, computed from the Mapping tab's lines; re-adds to 76,757,132.79 exactly."],
        [],["Portfolio","Platform","Class","Lines","Total $"]]
gt=0.0
for (p,pl,c),v in sorted(agg.items(), key=lambda kv:(kv[0][0],-abs(kv[1]))):
    t_rows.append([p,pl,c,cnt[(p,pl,c)],v]); gt+=v
t_rows.append([]); t_rows.append(["Grand total","","",len(rows),gt])
t_rows.append(["Check against the proven leaf total","","","",round(gt-76757132.79,2)])
by_class=defaultdict(float)
for r in rows: by_class[r["class"]]+=r["total"]
t_rows.append([])
for c,v in sorted(by_class.items()): t_rows.append([f"Of which {c}","","","",v])
totals=sheet_xml(t_rows,[34,30,12,9,16])

# ---- Reconciliation tab ----
cc_map=defaultdict(float)
for r in rows: cc_map[(r["tab"],r.get("cc",""))]+=r["total"]
r_rows=[["Reconciliation. The dashboard is a BUDGET view of FY26; this file is ACTUALS over the export's periods 007.2025 to 006.2026. The two will never tie to the cent; the dashboard figures sit here as sanity controls only, copied verbatim."],
        [],["Every mapped cost centre re-adds to the raw file to the cent (proven in verification):"],
        ["Tab","Cost centre","Mapped total $"]]
for (tab,cc),v in sorted(cc_map.items()):
    r_rows.append([tab,cc,v])
r_rows.append(["Total","",sum(cc_map.values())])
r_rows.append([]); r_rows.append(["Dashboard sanity controls (budget view, copied verbatim; null where the dashboard hid a value):"])
r_rows.append(["Area","Element","Actual YTD (Jan-Jul 26)","Forecast FY","Budget FY"])
for a in dash.get("areas",[]):
    ces=a.get("cost_elements")
    items= ces.items() if isinstance(ces,dict) else [(e.get("name",""),e) for e in (ces or [])]
    for nm,vals in items:
        if isinstance(vals,dict):
            r_rows.append([a.get("area",""),nm,vals.get("actual_ytd"),vals.get("forecast_fy"),vals.get("budget_fy")])
r_rows.append([]); r_rows.append(["Enterprise benchmark: Ampol opex 189.1 + Ampol capex 48.1 + Z Energy opex 64.8 = 302.1m budget FY26. This file carries the AU hardware and software actuals slice; network, outside services, depreciation, capex and all of Z/NZ are separate datasets."])
recon=sheet_xml(r_rows,[26,26,20,16,16])

# ---- Not 100% sure tab ----
fam=defaultdict(lambda: [0,0.0])
for r in rows:
    if r["class"]=="Open":
        key=(r.get("open_reason") or "(no reason recorded)")[:160]
        fam[key][0]+=1; fam[key][1]+=r["total"]
n_rows=[["Not 100% sure. Every line here is unplaced or part-placed, with its reason and its dollars. Nothing on this tab was guessed. Total 7,640,287.29."],
        [],["Reason","Lines","Total $"]]
for k,(c,v) in sorted(fam.items(), key=lambda kv:-abs(kv[1][1])):
    n_rows.append([k,c,v])
n_rows.append(["Total open",sum(c for c,_ in fam.values()),sum(v for _,v in fam.values())])
nsure=sheet_xml(n_rows,[110,9,16])

# ---- ZIP surgery ----
shutil.copyfile(SRC, OUT)
zin=zipfile.ZipFile(SRC)
names=zin.namelist()
wb=zin.read("xl/workbook.xml").decode("utf8")
rels=zin.read("xl/_rels/workbook.xml.rels").decode("utf8")
ct=zin.read("[Content_Types].xml").decode("utf8")
sheet_ids=[int(m) for m in re.findall(r'sheetId="(\d+)"', wb)]
rids=[int(m) for m in re.findall(r'Id="rId(\d+)"', rels)]
sid=max(sheet_ids); rid=max(rids)
NEW=[("Mapping",mapping,"sheetNL1"),("Totals",totals,"sheetNL2"),
     ("Reconciliation",recon,"sheetNL3"),("Not 100pc sure",nsure,"sheetNL4")]
ins=""
rel_ins=""
ct_ins=""
for i,(nm,_,part) in enumerate(NEW,1):
    ins+=f'<sheet name="{escape(nm)}" sheetId="{sid+i}" r:id="rId{rid+i}"/>'
    rel_ins+=f'<Relationship Id="rId{rid+i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/{part}.xml"/>'
    ct_ins+=f'<Override PartName="/xl/worksheets/{part}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
wb2=wb.replace("</sheets>", ins+"</sheets>")
rels2=rels.replace("</Relationships>", rel_ins+"</Relationships>")
ct2=ct.replace("</Types>", ct_ins+"</Types>")
assert wb2!=wb and rels2!=rels and ct2!=ct

with zipfile.ZipFile(OUT,"w",zipfile.ZIP_DEFLATED) as zout:
    for info in zin.infolist():
        data=zin.read(info.filename)
        if info.filename=="xl/workbook.xml": data=wb2.encode("utf8")
        elif info.filename=="xl/_rels/workbook.xml.rels": data=rels2.encode("utf8")
        elif info.filename=="[Content_Types].xml": data=ct2.encode("utf8")
        zout.writestr(info, data)
    for nm,payload,part in NEW:
        zout.writestr(f"xl/worksheets/{part}.xml", payload)
zin.close()
print("built", OUT)

# ---- hash gate ----
import hashlib
za=zipfile.ZipFile(SRC); zb=zipfile.ZipFile(OUT)
changed=[]; added=[]
a_names=set(za.namelist()); b_names=set(zb.namelist())
for n in sorted(a_names):
    if hashlib.sha256(za.read(n)).hexdigest()!=hashlib.sha256(zb.read(n)).hexdigest():
        changed.append(n)
for n in sorted(b_names-a_names): added.append(n)
print("changed parts:", changed)
print("added parts:", added)
allowed={"xl/workbook.xml","xl/_rels/workbook.xml.rels","[Content_Types].xml"}
assert set(changed)<=allowed, "UNEXPECTED CHANGES"
sheet_parts=[n for n in a_names if n.startswith("xl/worksheets/") or n.startswith("xl/sharedStrings")]
assert not any(n in changed for n in sheet_parts), "A DATA PART CHANGED"
print("HASH GATE PASSED: every original sheet and data part byte-identical")
