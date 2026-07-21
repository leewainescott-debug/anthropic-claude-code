#!/usr/bin/env python3
"""QA harness for TDD_Cost_Calc_v6.xlsx. Exit 0 = PASS, 1 = FAIL (with reasons)."""
import formulas, logging, re, sys, openpyxl
logging.getLogger().setLevel(logging.ERROR)

F = "TDD_Cost_Calc_v6.xlsx"
xl = formulas.ExcelModel().loads(F).finish()
sol = xl.calculate()
vals = {}
for k, v in sol.items():
    m = re.match(r"^'?\[[^\]]*\]([^!]*?)'?!([A-Z]+\d+)$", k)
    if not m: continue
    val = v.value
    try: val = val[0, 0]
    except Exception: pass
    vals[(m.group(1).strip().upper(), m.group(2))] = val
def g(s, c): return vals.get((s.upper(), c))

fails = []
def chk(label, got, want, tol=1e-3):
    try: ok = abs(float(got) - float(want)) < tol
    except Exception: ok = False
    if not ok: fails.append(f"{label}: got {got} want {want}")
    return ok

# 1. zero formula errors anywhere
errs = [(s, c, v) for (s, c), v in vals.items()
        if isinstance(v, str) and re.match(r'^#(VALUE|DIV/0|REF|NAME|NULL|NUM|N/A)', v)]
if errs: fails.append(f"{len(errs)} formula errors, first: {errs[:6]}")

# 2. archetype total = 60.58 (user's stated number)
wbf = openpyxl.load_workbook(F)
ft = wbf["3.0 FTE View"]
grand = None
for r in range(1, ft.max_row + 1):
    if ft.cell(r, 2).value == "TOTAL - all modelled squads": grand = r
chk("3.0 archetype cost total (60.58)", g("3.0 FTE VIEW", f"M{grand}"), 60.58, 0.005)

# 3. per-group role tabs: counts and cost vs independent recompute
wbv = openpyxl.load_workbook(F, data_only=False)
import collections
exp = {"2.2 BP&T": 24, "2.3 SA&D": 59, "2.4 CYBER ROLES": 52}
for tab, n in exp.items():
    ws = wbf[tab if tab != "2.4 CYBER ROLES" else "2.4 Cyber Roles"]
    cnt = 0
    for r in range(6, ws.max_row + 1):
        if ws.cell(r, 2).value in (None, "", "Summary"): break
        cnt += 1
    if cnt != n: fails.append(f"{tab} rows: got {cnt} want {n}")
# cost totals: sum of model cost col (all onshore default = full cost)
adv = openpyxl.load_workbook(F, data_only=True)["Added data"]
def group_cost(depts):
    tot = 0
    for r in range(2, adv.max_row + 1):
        d = str(adv.cell(r, 7).value or adv.cell(r, 6).value or "").strip().lower()
        if d in depts and adv.cell(r, 27).value is not None:
            try: tot += float(adv.cell(r, 27).value)
            except Exception: pass
    return tot / 1e6
BPT = {"tdd business partner", "transformation", "commercial"}
SAD = {"architecture", "technology strategy & ai capability", "delivery, sada", "group data"}
CYB = {"cyber strat & tech", "cyber sec ops", "cyber risk", "cyber grc", "service op & assurance"}
# planned spend totals on each tab's summary Total row (F col)
def tab_total(sheet, col="F"):
    ws = wbf[sheet]
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 2).value == "Total":
            return g(sheet, f"{col}{r}")
    return None
chk("2.2 planned total vs Added data", tab_total("2.2 BP&T"), group_cost(BPT), 0.01)
chk("2.3 planned total vs Added data", tab_total("2.3 SA&D"), group_cost(SAD), 0.01)
chk("2.4 planned total vs Added data", tab_total("2.4 Cyber Roles"), group_cost(CYB), 0.01)

# 4. Total-to-fund blocks exist, aligned, positive
TABS = ["1.1 Ampol Retail","1.2 Customer","1.3 Enterprise Data","1.4 TDD Group Functions",
        "1.5 P&C","1.6 Finance","1.7 Infrastructure","1.8 Energy Solutions & B2B",
        "1.9 Commercial Fuels","1.10 Z Retail","1.11 TDD Cyber"]
for t in TABS:
    ws = wbf[t]
    found = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 2).value == "Total to fund": found = r
    if not found:
        fails.append(f"{t}: Total to fund block missing"); continue
    for off, lbl in [(1, "TDD Variance"), (2, "Other Variance"), (3, "Total")]:
        if ws.cell(found + off, 2).value != lbl:
            fails.append(f"{t}: row {found+off} label {ws.cell(found+off,2).value!r} != {lbl}")
    for off in (1, 2, 3):
        v = g(t, f"C{found+off}")
        try:
            if float(v) < -1e-9: fails.append(f"{t} C{found+off} negative: {v}")
        except Exception: fails.append(f"{t} C{found+off} not numeric: {v}")

# 5. reconciliation intact
chk("2.0 C30 allocations", g("2.0 GROUP SUMMARY", "C30"), 43.5)
chk("2.0 C32 check=0", g("2.0 GROUP SUMMARY", "C32"), 0.0, 1e-9)
# 6. 3.0 cross-check = 0
for r in range(1, ft.max_row + 1):
    if ft.cell(r, 2).value == "Difference (must be 0)":
        chk("3.0 cross-check", g("3.0 FTE VIEW", f"C{r}"), 0.0, 1e-9)
# 7. cyber capex visible
chk("1.11 H16 capex 0.5", g("1.11 TDD CYBER", "H16"), 0.5)
# 8. no comments remain
for ws in wbf.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if c.comment is not None:
                fails.append(f"comment survives at {ws.title}!{c.coordinate}"); break

print("FAILS:", len(fails))
for f_ in fails: print("  -", f_)
print("KEY NUMBERS: archetype total:", g("3.0 FTE VIEW", f"M{grand}"),
      "| 2.2:", tab_total("2.2 BP&T"), "| 2.3:", tab_total("2.3 SA&D"),
      "| 2.4:", tab_total("2.4 Cyber Roles"))
sys.exit(1 if fails else 0)
