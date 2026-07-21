#!/usr/bin/env python3
"""QA harness v7 - passes 1-3 of the quadruple check.
Pass 1: independent Python recompute from the workbook's source data.
Pass 2: full formula-engine evaluation (zero errors).
Pass 3: assertion battery incl. cross-foots and source-tie-outs.
Exit 0 = PASS."""
import formulas, logging, re, sys, openpyxl
from collections import Counter
logging.getLogger().setLevel(logging.ERROR)
F = "TDD_Cost_Calc_v7.xlsx"

fails = []
def chk(label, got, want, tol=1e-3):
    try: ok = abs(float(got) - float(want)) < tol
    except Exception: ok = False
    if not ok: fails.append(f"{label}: got {got} want {want}")

# ---- PASS 1: independent recompute from Added data helper cols ----
wv = openpyxl.load_workbook(F, data_only=True)
adv = wv["Added data"]
by = Counter(); by_status = Counter(); grand_actual = 0.0
for r in range(2, adv.max_row + 1):
    cls = adv.cell(r, 32).value
    if cls is None: continue
    cost = adv.cell(r, 27).value
    cost = float(cost) if isinstance(cost, (int, float)) else 0.0
    grand_actual += cost
    by[cls] += cost
    by_status[adv.cell(r, 33).value] += cost
grand_actual /= 1e6
exp_sq = (by["Squad"] + by["Strategic Program"]) / 1e6
exp_lead = by["Leadership"] / 1e6
exp_coe = by["COE"] / 1e6
exp_un = by["Unmapped"] / 1e6
print(f"PASS1 independent: grand={grand_actual:.3f} squads+strat={exp_sq:.3f} "
      f"lead={exp_lead:.3f} coe={exp_coe:.3f} unmapped={exp_un:.3f} "
      f"filled={by_status['Filled']/1e6:.3f} vacant={by_status['Vacant']/1e6:.3f}")
chk("pass1 classes sum to grand", exp_sq + exp_lead + exp_coe + exp_un, grand_actual, 1e-6)

# ---- PASS 2: engine ----
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
errs = [(s, c, v) for (s, c), v in vals.items()
        if isinstance(v, str) and re.match(r'^#(VALUE|DIV/0|REF|NAME|NULL|NUM|N/A)', v)]
if errs: fails.append(f"{len(errs)} formula errors, first: {errs[:6]}")

# ---- PASS 3: assertions ----
wf = openpyxl.load_workbook(F)
tcs = wf["5.0 Total Cost"]
subrows = {}; grand_row = None; check_row = None
for r in range(1, tcs.max_row + 1):
    b = tcs.cell(r, 2).value
    if isinstance(b, str):
        if b.startswith("Subtotal - squads"): subrows["A"] = r
        elif b.startswith("Subtotal - leadership"): subrows["B"] = r
        elif b.startswith("Subtotal - Centres"): subrows["C"] = r
        elif b.startswith("Subtotal - not"): subrows["D"] = r
        elif b == "TOTAL OPERATING MODEL": grand_row = r
        elif b.startswith("Check: Actual Total"): check_row = r
T = "5.0 TOTAL COST"
chk("5.0 actual grand = Added data total", g(T, f"F{grand_row}"), grand_actual, 0.005)
chk("5.0 check cell = 0", g(T, f"C{check_row}"), 0.0, 0.005)
chk("5.0 subA actual = independent", g(T, f"F{subrows['A']}"), exp_sq, 0.005)
chk("5.0 subB actual = independent", g(T, f"F{subrows['B']}"), exp_lead, 0.005)
chk("5.0 subC actual = independent", g(T, f"F{subrows['C']}"), exp_coe, 0.005)
chk("5.0 subD actual = independent", g(T, f"F{subrows['D']}"), exp_un, 0.005)
sumsub = sum(g(T, f"F{subrows[k]}") or 0 for k in "ABCD")
chk("5.0 sections cross-foot", g(T, f"F{grand_row}"), sumsub, 1e-6)
chk("5.0 model squads = archetype 60.58", g(T, f"C{subrows['A']}"), 60.58, 0.005)
chk("5.0 model grand = 2.0 total cost J24", g(T, f"C{grand_row}"),
    (g("2.0 GROUP SUMMARY","J24") or 0), 0.005)
# 3.0 actual col ties to 5.0 section A
ft = wf["3.0 FTE View"]
grand_ft = None; hdr_ft = None
for r in range(1, ft.max_row + 1):
    if ft.cell(r, 2).value == "TOTAL - all modelled squads": grand_ft = r
    if ft.cell(r, 2).value == "Portfolio" and ft.cell(r, 14).value == "Actual cost ($m)": hdr_ft = r
chk("3.0 actual grand = 5.0 squads subtotal", g("3.0 FTE VIEW", f"N{grand_ft}"), g(T, f"F{subrows['A']}"), 0.005)
chk("3.0 archetype total 60.58", g("3.0 FTE VIEW", f"M{grand_ft}"), 60.58, 0.005)
# v6 carryovers
chk("2.0 C30 allocations", g("2.0 GROUP SUMMARY", "C30"), 43.5)
chk("2.0 C32 check", g("2.0 GROUP SUMMARY", "C32"), 0.0, 1e-9)
chk("1.11 capex", g("1.11 TDD CYBER", "H16"), 0.5)
TABS = ["1.1 Ampol Retail","1.2 Customer","1.3 Enterprise Data","1.4 TDD Group Functions",
        "1.5 P&C","1.6 Finance","1.7 Infrastructure","1.8 Energy Solutions & B2B",
        "1.9 Commercial Fuels","1.10 Z Retail","1.11 TDD Cyber"]
for t in TABS:
    ws = wf[t]
    found = None
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 2).value == "Total to fund": found = r
    if not found: fails.append(f"{t}: TTF missing"); continue
    for off in (1, 2, 3):
        v = g(t, f"C{found+off}")
        try:
            if float(v) < -1e-9: fails.append(f"{t} C{found+off} negative: {v}")
        except Exception: fails.append(f"{t} C{found+off} not numeric: {v}")
# no comments anywhere
for ws in wf.worksheets:
    stop = False
    for row in ws.iter_rows():
        for c in row:
            if c.comment is not None:
                fails.append(f"comment at {ws.title}!{c.coordinate}"); stop = True; break
        if stop: break
# 5.1 headline numbers
qa = wf["5.1 Data QA"]
hd = {}
for r in range(1, 15):
    b = qa.cell(r, 2).value
    if b in ("Records", "Vacant"):
        hd[b] = (qa.cell(r, 3).value, qa.cell(r, 4).value)
chk("5.1 raw records", hd.get("Records", (0, 0))[0], 536)
chk("5.1 added records", hd.get("Records", (0, 0))[1], 548)
chk("5.1 raw vacant", hd.get("Vacant", (0, 0))[0], 166)
chk("5.1 added vacant", hd.get("Vacant", (0, 0))[1], 156)

print("FAILS:", len(fails))
for f_ in fails: print("  -", f_)
_rb, _rc = subrows["B"], subrows["C"]
print(f"KEY: grand actual={g(T, f'F{grand_row}')}, model total={g(T, f'C{grand_row}')}, "
      f"lead actual={g(T, f'F{_rb}')}, coe actual={g(T, f'F{_rc}')}")
sys.exit(1 if fails else 0)
