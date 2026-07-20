#!/usr/bin/env python3
"""Inject verified cached values (from the `formulas` engine) into the xlsx
formula cells, preserving all openpyxl formatting. LibreOffice recalc hangs in
this sandbox, so this is the deterministic alternative to populate cached <v>."""
import formulas, logging, re, zipfile, shutil, os
import xml.etree.ElementTree as ET
import openpyxl

SRC="Ampol_Retail_TDD_Cost_Calculator.xlsx"
logging.getLogger().setLevel(logging.ERROR)

# 1) compute all values
xl = formulas.ExcelModel().loads(SRC).finish()
sol = xl.calculate()
sol_map={}
keyre=re.compile(r"\][^!]*'?!")
for k,v in sol.items():
    # k like "'[FILE.XLSX]SHEET NAME'!H13"
    m=re.match(r"^'?\[[^\]]*\]([^!]*?)'?!([A-Z]+\d+)$", k)
    if not m: continue
    sheet=m.group(1).strip().upper(); coord=m.group(2)
    val=v.value
    try: val=val[0,0]
    except: pass
    sol_map[(sheet,coord)]=val

# 2) enumerate formula cells per sheet via openpyxl
wb=openpyxl.load_workbook(SRC, data_only=False)
formula_cells={}   # sheet_title_upper -> {coord: value}
missing=0; total=0
for ws in wb.worksheets:
    st=ws.title.upper()
    d={}
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value,str) and c.value.startswith("="):
                total+=1
                key=(st,c.coordinate)
                if key in sol_map:
                    d[c.coordinate]=sol_map[key]
                else:
                    missing+=1
    formula_cells[st]=d
print(f"formula cells: {total}, resolved: {total-missing}, missing: {missing}")

# 3) map worksheet xml files -> display names
NS_MAIN="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_R)

tmp="_xltmp"
if os.path.exists(tmp): shutil.rmtree(tmp)
os.makedirs(tmp)
with zipfile.ZipFile(SRC) as z:
    z.extractall(tmp)

# workbook.xml : sheet name -> r:id ; rels : r:id -> target
wbtree=ET.parse(f"{tmp}/xl/workbook.xml"); wbroot=wbtree.getroot()
name_to_rid={}
for sh in wbroot.iter(f"{{{NS_MAIN}}}sheet"):
    name_to_rid[sh.get("name")]=sh.get(f"{{{NS_R}}}id")
relstree=ET.parse(f"{tmp}/xl/_rels/workbook.xml.rels"); relsroot=relstree.getroot()
rid_to_target={}
for rel in relsroot:
    rid_to_target[rel.get("Id")]=rel.get("Target")
name_to_file={n:rid_to_target[r] for n,r in name_to_rid.items() if r in rid_to_target}

def num_to_str(x):
    if isinstance(x,bool): return "1" if x else "0"
    if float(x)==int(float(x)) and abs(float(x))<1e15:
        # keep as float repr but trim
        return repr(float(x))
    return repr(float(x))

ERRORS={"#VALUE!","#DIV/0!","#REF!","#NAME?","#NULL!","#NUM!","#N/A"}
injected=0
for name,target in name_to_file.items():
    st=name.upper()
    vals=formula_cells.get(st,{})
    if not vals: continue
    path=f"{tmp}/xl/{target.lstrip('/')}" if not target.startswith('xl/') else f"{tmp}/{target}"
    # target is like 'worksheets/sheet1.xml' relative to xl/
    path=f"{tmp}/xl/{target}" if not target.startswith('/xl') else f"{tmp}{target}"
    if not os.path.exists(path):
        path=f"{tmp}/xl/{target}"
    tree=ET.parse(path); root=tree.getroot()
    for c in root.iter(f"{{{NS_MAIN}}}c"):
        coord=c.get("r")
        if coord not in vals: continue
        f_el=c.find(f"{{{NS_MAIN}}}f")
        if f_el is None: continue
        val=vals[coord]
        # remove existing v
        for old in c.findall(f"{{{NS_MAIN}}}v"): c.remove(old)
        # remove existing 'is' inline
        for old in c.findall(f"{{{NS_MAIN}}}is"): c.remove(old)
        v_el=ET.SubElement(c,f"{{{NS_MAIN}}}v")
        if isinstance(val,str) and val in ERRORS:
            c.set("t","e"); v_el.text=val
        elif isinstance(val,bool):
            c.set("t","b"); v_el.text="1" if val else "0"
        elif isinstance(val,str):
            c.set("t","str"); v_el.text=val
        else:
            if "t" in c.attrib and c.get("t") in ("str","e","b","s"):
                del c.attrib["t"]
            try: v_el.text=num_to_str(val)
            except: c.set("t","str"); v_el.text=str(val)
        # ensure <f> comes before <v>
        c.remove(f_el); c.insert(0,f_el)
        injected+=1
    tree.write(path, xml_declaration=True, encoding="UTF-8")
print("injected cached values:",injected)

# 4) set calcPr fullCalcOnLoad so Excel refreshes too (belt & braces)
calcPr=wbroot.find(f"{{{NS_MAIN}}}calcPr")
if calcPr is None:
    calcPr=ET.SubElement(wbroot,f"{{{NS_MAIN}}}calcPr")
calcPr.set("fullCalcOnLoad","1")
wbtree.write(f"{tmp}/xl/workbook.xml", xml_declaration=True, encoding="UTF-8")

# 5) repackage
out=SRC
if os.path.exists(out): os.remove(out)
with zipfile.ZipFile(out,"w",zipfile.ZIP_DEFLATED) as z:
    for base,_,files in os.walk(tmp):
        for fn in files:
            full=os.path.join(base,fn)
            arc=os.path.relpath(full,tmp)
            z.write(full,arc)
shutil.rmtree(tmp)
print("REPACKAGED",out)
