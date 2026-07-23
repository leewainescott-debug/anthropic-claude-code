#!/usr/bin/env python3
"""Full recalc of clean_v6: compute every value, dump a value map (JSON) for the
QA agent, print the COE tie-out + footing + error scan, and inject the cached
values into a populated copy so the workbook opens filled in."""
import formulas, logging, re, json, zipfile, shutil, os
import xml.etree.ElementTree as ET
import openpyxl
logging.getLogger().setLevel(logging.ERROR)
SRC="clean_v6.xlsx"; POP="clean_v6_pop.xlsx"

xl=formulas.ExcelModel().loads(SRC).finish(); sol=xl.calculate()
val={}
for k,v in sol.items():
    m=re.match(r"^'?\[[^\]]*\]([^!]*?)'?!([A-Z]+\d+)$", k)
    if not m: continue
    x=v.value
    try: x=x[0,0]
    except: pass
    if not isinstance(x,(int,float,str,bool)) and x is not None: x=str(x)
    val[(m.group(1).strip().upper(), m.group(2))]=x
def gv(t,a): return val.get((t.upper(),a))
def isnum(x): return isinstance(x,(int,float)) and not isinstance(x,bool)
def fm(x): return f"{x:.4f}" if isnum(x) else str(x)

# ---- dump JSON value map for the QA agent (only visible sheets, primitives) ----
wb0=openpyxl.load_workbook(SRC, data_only=False)
vis={ws.title for ws in wb0.worksheets if ws.sheet_state=="visible"}
dump={}
for (t,c),x in val.items():
    # store with original-case sheet title where possible
    dump.setdefault(t,{})[c]=x
json.dump(dump, open("tie_out_values.json","w"))
print("value map dumped: tie_out_values.json  (sheets:",len(dump),")")

ERR=("#REF!","#DIV/0!","#VALUE!","#N/A","#NAME?","#NUM!","#NULL!","#CYCLE!")
errs=[(s,c,v) for (s,c),v in val.items() if isinstance(v,str) and v.strip() in ERR]
from collections import Counter
print("\n=== ERROR CELLS:",len(errs))
for s,n in Counter(s for s,_,_ in errs).most_common(): print("   ",s,n)
for s,c,v in errs[:25]: print("     ",s,"!",c,"=",v)

print("\n=== COE TIE-OUT ===")
# 1.11 BP&T rows 6 (BP) 7 (Transformation) 8 (Total): F planned, G budget, H left
for tab,rows in [("1.11 BP&T",[6,7,8]),("1.12 SA&D",[6,7,8])]:
    print(f"-- {tab} --")
    for r in rows:
        if tab=="1.11 BP&T":
            F,G,H=gv(tab,f"F{r}"),gv(tab,f"G{r}"),gv(tab,f"H{r}")
        else:
            F,G,H=gv(tab,f"G{r}"),gv(tab,f"H{r}"),gv(tab,f"I{r}")
        print(f"   r{r}: planned={fm(F)}  budget={fm(G)}  left-to-fund={fm(H)}")
    print(f"   C13(portfolio-funded)={fm(gv(tab,'C13'))}  C14(COE alloc)={fm(gv(tab,'C14'))}  C15(budget)={fm(gv(tab,'C15'))}")
# 3.1 COE rows 18(SA) 20(Transf) 21(BP) 22(Data) 19(Cyber): C budget, D cost, E variance
print("-- 3.1 Group Summary COE rows (C budget | D support cost | E variance) --")
for r,lbl in [(21,"BP"),(20,"Transformation"),(18,"S&A"),(22,"Data"),(19,"Cyber")]:
    print(f"   r{r} {lbl:14s}: budget={fm(gv('3.1 Group Summary',f'C{r}'))}  cost={fm(gv('3.1 Group Summary',f'D{r}'))}  var={fm(gv('3.1 Group Summary',f'E{r}'))}")
# 3.4 COE Summary
print("-- 3.4 COE Summary (F planned | G budget | H left) --")
for r,lbl in [(6,"BP"),(7,"Transf"),(8,"S&A"),(9,"Data"),(10,"Cyber"),(11,"Total")]:
    print(f"   r{r} {lbl:8s}: planned={fm(gv('3.4 COE Summary',f'F{r}'))}  budget={fm(gv('3.4 COE Summary',f'G{r}'))}  left={fm(gv('3.4 COE Summary',f'H{r}'))}")
# consistency: 1.11 BP left-to-fund vs 3.1 BP -variance
bp_1_11=gv("1.11 BP&T","H6"); bp_3_1=gv("3.1 Group Summary","E21")
print(f"\nCONSISTENCY BP: 1.11 left-to-fund={fm(bp_1_11)}  3.1 variance={fm(bp_3_1)}  (should be equal & opposite): tie={isnum(bp_1_11) and isnum(bp_3_1) and abs(bp_1_11+bp_3_1)<1e-4}")
# 3.2 grand totals
print("\n=== 3.2 grand totals ===", {c:fm(gv('3.2 Total Cost',f'{c}24')) for c in 'CDEFKLIMN'})

# ---- recon tab totals ----
print("\n=== 3.5 Source Reconciliation ===")
for r in range(6,21):
    b=gv("3.5 Source Reconciliation",f"B{r}")
    if b in (None,""): continue
    row=[fm(gv("3.5 Source Reconciliation",f"{c}{r}")) for c in "CDEFGHIJK"]
    print(f"   {str(b)[:36]:36s} Sq[R/F/V]={row[0]}/{row[1]}/{row[2]}  Rev[R/F/V]={row[3]}/{row[4]}/{row[5]}  d[R/F/V]={row[6]}/{row[7]}/{row[8]}")

# ---- inject cached values into a populated copy ----
NS="http://schemas.openxmlformats.org/spreadsheetml/2006/main"; NR="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("",NS); ET.register_namespace("r",NR)
fcells={}
for ws in wb0.worksheets:
    st=ws.title.upper(); d={}
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value,str) and c.value.startswith("="):
                key=(st,c.coordinate)
                if key in val: d[c.coordinate]=val[key]
    fcells[st]=d
tmp="_tie";
if os.path.exists(tmp): shutil.rmtree(tmp)
os.makedirs(tmp)
with zipfile.ZipFile(SRC) as z: z.extractall(tmp)
wbt=ET.parse(f"{tmp}/xl/workbook.xml"); wbr=wbt.getroot()
n2r={sh.get("name"):sh.get(f"{{{NR}}}id") for sh in wbr.iter(f"{{{NS}}}sheet")}
rels=ET.parse(f"{tmp}/xl/_rels/workbook.xml.rels").getroot()
r2t={rel.get("Id"):rel.get("Target") for rel in rels}
n2f={n:r2t[r] for n,r in n2r.items() if r in r2t}
def numstr(x):
    if isinstance(x,bool): return "1" if x else "0"
    return repr(float(x))
inj=0
for name,tgt in n2f.items():
    d=fcells.get(name.upper(),{})
    if not d: continue
    path=f"{tmp}/xl/{tgt}"
    if not os.path.exists(path): continue
    tr=ET.parse(path); rt=tr.getroot()
    for c in rt.iter(f"{{{NS}}}c"):
        co=c.get("r")
        if co not in d: continue
        fe=c.find(f"{{{NS}}}f")
        if fe is None: continue
        x=d[co]
        for old in c.findall(f"{{{NS}}}v"): c.remove(old)
        for old in c.findall(f"{{{NS}}}is"): c.remove(old)
        ve=ET.SubElement(c,f"{{{NS}}}v")
        if isinstance(x,str) and x in ERR: c.set("t","e"); ve.text=x
        elif isinstance(x,bool): c.set("t","b"); ve.text="1" if x else "0"
        elif isinstance(x,str): c.set("t","str"); ve.text=x
        else:
            if c.get("t") in ("str","e","b","s"): del c.attrib["t"]
            try: ve.text=numstr(x)
            except: c.set("t","str"); ve.text=str(x)
        c.remove(fe); c.insert(0,fe); inj+=1
    tr.write(path, xml_declaration=True, encoding="UTF-8")
cp=wbr.find(f"{{{NS}}}calcPr")
if cp is None: cp=ET.SubElement(wbr,f"{{{NS}}}calcPr")
cp.set("fullCalcOnLoad","1"); wbt.write(f"{tmp}/xl/workbook.xml", xml_declaration=True, encoding="UTF-8")
if os.path.exists(POP): os.remove(POP)
with zipfile.ZipFile(POP,"w",zipfile.ZIP_DEFLATED) as z:
    for base,_,files in os.walk(tmp):
        for fn in files: z.write(os.path.join(base,fn), os.path.relpath(os.path.join(base,fn),tmp))
shutil.rmtree(tmp)
print("\ninjected",inj,"cached values ->",POP)
print("TIE_OUT_DONE")
