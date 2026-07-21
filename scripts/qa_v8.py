#!/usr/bin/env python3
"""QA harness v8 - quadruple check, passes 1-3.
Pass 1: independent Python recompute from workbook source data (no formulas).
Pass 2: full formula-engine evaluation - zero formula errors anywhere.
Pass 3: assertion battery - every user complaint has a named assertion,
        plus a FUNCTIONAL offshore-toggle test (second engine run on a copy
        with H cells set to Offshore -> J must equal 40% of full cost).
Exit 0 = PASS."""
import formulas, logging, re, sys, json, shutil, openpyxl
from collections import Counter
logging.getLogger().setLevel(logging.ERROR)
SCR = "/tmp/claude-0/-home-user-anthropic-claude-code/6161aafe-2dad-5bc5-ad55-d8a92ce554cc/scratchpad/"
F = SCR + "TDD_Cost_Calc_v8.xlsx"
A = json.load(open(SCR + "anchors_v8.json"))

fails = []
def chk(label, got, want, tol=1e-3):
    try: ok = abs(float(got) - float(want)) < tol
    except Exception: ok = False
    if not ok: fails.append(f"{label}: got {got} want {want}")

# =====================================================================
# PASS 1 - independent recompute straight from Added data / raw data
# =====================================================================
wv = openpyxl.load_workbook(F, data_only=True)
adv = wv["Added data"]
by = Counter(); by_status = Counter(); grand_actual = 0.0
for r in range(2, 550):                       # junk stub row 550 excluded by design
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
chk("pass1 classes sum to grand", exp_sq + exp_lead + exp_coe + exp_un, grand_actual, 1e-6)

# independent role-group membership from RAW DATA (the source of truth for the model)
# + independent python cost mapping (Added data used ONLY as the cost ledger)
rawv = wv["raw data"]
addn = {}; addt = {}
for r in range(2, 550):
    nm_ = str(adv.cell(r, 2).value or "").strip().lower()
    tl_ = str(adv.cell(r, 3).value or "").strip().lower()
    c_ = adv.cell(r, 27).value
    c_ = float(c_) if isinstance(c_, (int, float)) else 0.0
    addn.setdefault(nm_, []).append(c_)
    addt.setdefault(tl_, []).append(c_)
BPT_P = {"COE - Business Partnering": "BPT", "COE - Transformation": "BPT"}
SAD_P = {"COE - Strategy Architecture": "SAD", "COE - Data": "SAD"}
CYB_D = {"cyber strat & tech", "cyber risk", "cyber grc", "cyber sec ops", "service op & assurance"}
def map_cost(nm_, tl_, st_):
    nm_, tl_ = nm_.strip().lower(), tl_.strip().lower()
    if st_ != "Vacant" and len(addn.get(nm_, [])) == 1:
        return addn[nm_][0]
    lst = addt.get(tl_, [])
    return sum(lst) / len(lst) if lst else 0.0
addt_x = {}
for r in range(2, 550):
    tlx = str(adv.cell(r, 3).value or "").lower()
    cx = adv.cell(r, 27).value
    cx = float(cx) if isinstance(cx, (int, float)) else 0.0
    addt_x.setdefault(tlx, []).append(cx)
def title_rate_exact(tl_raw):
    lst = addt_x.get(str(tl_raw).lower(), [])
    return sum(lst) / len(lst) if lst else 0.0
gcnt = Counter(); gcost = Counter(); gvac = Counter()
vac_map = Counter()
n_unspec_i = 0; unspec_cost_i = 0.0
for r in range(2, rawv.max_row + 1):
    msq = str(rawv.cell(r, 16).value or "").strip()
    port = str(rawv.cell(r, 14).value or "").strip()
    nm_ = str(rawv.cell(r, 2).value or "").strip()
    tl_ = str(rawv.cell(r, 3).value or "").strip()
    dp_ = str(rawv.cell(r, 7).value or "").strip().lower()
    st_ = str(rawv.cell(r, 18).value or "").strip()
    if not nm_ and not tl_: continue
    key = BPT_P.get(msq) or SAD_P.get(msq)
    if key is None and port == "TDD Cyber" and dp_ in CYB_D: key = "CYB"
    if msq == "COE (unspecified)":
        n_unspec_i += 1; unspec_cost_i += map_cost(nm_, tl_, st_); continue
    cls_r = str(rawv.cell(r, 17).value or "")
    if cls_r in ("Squad", "Strategic Program") and st_ == "Vacant":
        vac_map[port] += title_rate_exact(rawv.cell(r, 3).value) / 1e6
    if key:
        gcnt[key] += 1; gcost[key] += map_cost(nm_, tl_, st_)
        if st_ == "Vacant": gvac[key] += 1
print(f"PASS1: grand={grand_actual:.3f} squads+strat={exp_sq:.3f} lead={exp_lead:.3f} "
      f"coe={exp_coe:.3f} unmapped={exp_un:.3f} | BPT {gcnt['BPT']}@{gcost['BPT']/1e6:.3f} "
      f"SAD {gcnt['SAD']}@{gcost['SAD']/1e6:.3f} CYB {gcnt['CYB']}@{gcost['CYB']/1e6:.3f} "
      f"unspec {n_unspec_i}@{unspec_cost_i/1e6:.3f}")
chk("pass1 BPT count = build", gcnt["BPT"], A["n_bpt"])
chk("pass1 SAD count = build", gcnt["SAD"], A["n_sad"])
chk("pass1 CYB count = build", gcnt["CYB"], A["n_cyb"])
chk("pass1 unspecified count = build", n_unspec_i, A["n_unspec"])

# =====================================================================
# PASS 2 - engine run, zero errors
# =====================================================================
print("PASS2: engine run 1 (as shipped)...", flush=True)
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
if errs: fails.append(f"{len(errs)} formula errors, first: {errs[:8]}")

# =====================================================================
# PASS 3 - assertion battery
# =====================================================================
wf = openpyxl.load_workbook(F)

# --- structure: exact tab list & order, exec summary first, no dupes ---
TABS = ["1.1 Ampol Retail","1.2 Customer","1.3 Enterprise Data","1.4 TDD Group Functions",
        "1.5 P&C","1.6 Finance","1.7 Infrastructure","1.8 Energy Solutions & B2B",
        "1.9 Commercial Fuels","1.10 Z Retail","1.11 TDD Cyber"]
GM_TABS = [f"4.{i} GM {t.split(' ', 1)[1]}"[:31] for i, t in enumerate(TABS, 1)]
EXPECT = ["Exec Summary","0.0 Data Config","0.1 Squads","0.2 FY26 Budget",
          "0.3 For Presentation Pack (2)","0.4 Budget Table (Fin)"] + TABS + \
         ["2.0 Group Summary","2.1 Total Cost","2.2 COE","2.3 BP&T","2.4 SA&D","2.5 Cyber Roles",
          "3.0 FTE View","3.1 Data QA"] + GM_TABS + ["squad mapping","raw data","Added data","Lists"]
if wf.sheetnames != EXPECT:
    fails.append(f"tab order/list wrong: {wf.sheetnames}")
if len(set(wf.sheetnames)) != len(wf.sheetnames):
    fails.append("duplicate tab names")

# --- no formula/label references to deleted or renamed tabs (incl bare prose) ---
FORBID = ["'5.0 Total Cost'", "'5.1 Data QA'", "'4.0 Insights'", "2.1 COE",
          "'2.2 BP&T'", "'2.3 SA&D'", "'2.4 Cyber Roles'"]
for ws in wf.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str):
                for fb in FORBID:
                    if fb in c.value:
                        fails.append(f"stale ref {fb} at {ws.title}!{c.coordinate}")

# --- 2.1 Total Cost: ties, cross-foots, de-dup ---
T = "2.1 Total Cost"; GR, CK = A["grand"], A["check"]
chk("2.1 actual grand = Added data total", g(T, f"F{GR}"), grand_actual, 0.005)
chk("2.1 check cell = 0", g(T, f"C{CK}"), 0.0, 0.005)
chk("2.1 subA actual = independent squads+strat", g(T, f"F{A['subA']}"), exp_sq, 0.005)
chk("2.1 subB actual = independent leadership", g(T, f"F{A['subB']}"), exp_lead, 0.005)
chk("2.1 subC actual = independent COE", g(T, f"F{A['subC']}"), exp_coe, 0.005)
chk("2.1 subD actual = independent unmapped", g(T, f"F{A['subD']}"), exp_un, 0.005)
for col in "CDEF":
    ssum = sum(g(T, f"{col}{A[k]}") or 0 for k in ("subA","subB","subC","subD"))
    chk(f"2.1 {col} cross-foot", g(T, f"{col}{GR}"), ssum, 1e-6)
L7 = g("0.0 Data Config", "L7"); L8 = g("0.0 Data Config", "L8")
chk("2.1 de-dup row = -(11*L7+11*L8)", g(T, f"C{A['dedup']}"), -(11*float(L7)+11*float(L8)), 1e-6)
coe_gross = sum(g("2.2 COE", f"D{rr}") or 0 for rr in (8, 10, 11, 12))
chk("2.1 subC model = COE gross - overhead-funded BP/DA",
    g(T, f"C{A['subC']}"), coe_gross - (11*float(L7)+11*float(L8)), 1e-6)

# --- role tabs: counts, costs, summary-on-top, vacancy, info lines ---
ROLE = {"2.3 BP&T": ("BPT", A["bp"]), "2.4 SA&D": ("SAD", A["sad"]), "2.5 Cyber Roles": ("CYB", A["cyb"])}
for tab, (key, an) in ROLE.items():
    ws = wf[tab]
    chk(f"{tab} summary is on top", 1 if an["sum_first"] < an["tbl_first"] else 0, 1, 0.5)
    if ws["B4"].value != "Summary": fails.append(f"{tab} B4 != Summary")
    n = an["tbl_last"] - an["tbl_first"] + 1
    chk(f"{tab} role rows = independent", n, gcnt[key])
    chk(f"{tab} total roles cell", g(tab, f"C{an['tot']}"), gcnt[key])
    chk(f"{tab} total vacant cell", g(tab, f"E{an['tot']}"), gvac[key])
    chk(f"{tab} planned spend = independent cost", g(tab, f"F{an['tot']}"), gcost[key]/1e6, 0.01)
    # every role row must carry a nonzero full cost and the offshore-aware formula
    for rr in range(an["tbl_first"], an["tbl_last"] + 1):
        jf = ws[f"J{rr}"].value
        if not (isinstance(jf, str) and "'0.1 Squads'!$K$5" in jf and f"H{rr}" in jf):
            fails.append(f"{tab} J{rr} formula not offshore-aware: {jf}"); break
    zero_cost = [rr for rr in range(an["tbl_first"], an["tbl_last"] + 1)
                 if not isinstance(g(tab, f"I{rr}"), (int, float)) or g(tab, f"I{rr}") <= 0]
    if zero_cost: fails.append(f"{tab} rows with zero/blank full cost: {zero_cost[:6]}")
chk("0.1 Squads K5 offshore rate", g("0.1 Squads", "K5"), 0.4, 1e-9)

# BP / DA overhead "how many times applied" lines
bpi = A["bp"]["info_first"]
chk("2.3 BP overhead applied 11x", g("2.3 BP&T", f"C{bpi}"), 11)
chk("2.3 BP funding/portfolio = L7", g("2.3 BP&T", f"C{bpi+1}"), float(L7), 1e-6)
chk("2.3 BP total funding = 11*L7", g("2.3 BP&T", f"C{bpi+2}"), 11*float(L7), 1e-6)
chk("2.3 BP FTE-equivalents 4.4", g("2.3 BP&T", f"C{bpi+3}"), 4.4, 1e-6)
sdi = A["sad"]["info_first"]
chk("2.4 DA overhead applied 11x", g("2.4 SA&D", f"C{sdi}"), 11)
chk("2.4 DA funding/portfolio = L8", g("2.4 SA&D", f"C{sdi+1}"), float(L8), 1e-6)
chk("2.4 DA total funding = 11*L8", g("2.4 SA&D", f"C{sdi+2}"), 11*float(L8), 1e-6)
chk("2.4 DA FTE-equivalents 5.5", g("2.4 SA&D", f"C{sdi+3}"), 5.5, 1e-6)

# SA&D must NOT contain portfolio data-squad roles
sad_ws = wf["2.4 SA&D"]
for rr in range(A["sad"]["tbl_first"], A["sad"]["tbl_last"] + 1):
    ttl = str(sad_ws[f"C{rr}"].value or "").lower()
    dep = str(sad_ws[f"D{rr}"].value or "").lower()
    if dep == "group data" and any(s in ttl for s in ("data scientist", "reporting analyst")):
        pass  # allowed only if their squad wasn't a portfolio squad - membership already tied to filter above

# --- cyber single-source: 1.11 squads = live refs to 2.5, values equal ---
coe_r = A["cyb"]["cats"]["TDD COE"]; cyb_r = A["cyb"]["cats"]["TDD Cyber"]
w11 = wf["1.11 TDD Cyber"]
if w11["G24"].value != f"='2.5 Cyber Roles'!$F${coe_r}":
    fails.append(f"1.11 G24 not live ref: {w11['G24'].value}")
if w11["G25"].value != f"='2.5 Cyber Roles'!$F${cyb_r}":
    fails.append(f"1.11 G25 not live ref: {w11['G25'].value}")
chk("1.11 G24 = 2.5 TDD COE planned", g("1.11 TDD Cyber", "G24"), g("2.5 Cyber Roles", f"F{coe_r}"), 1e-6)
chk("1.11 G25 = 2.5 TDD Cyber planned", g("1.11 TDD Cyber", "G25"), g("2.5 Cyber Roles", f"F{cyb_r}"), 1e-6)
chk("1.11 capex 0.5 visible", g("1.11 TDD Cyber", "H16"), 0.5)
cyt = A["cyt"]
chk("2.5 budget = people bucket only (capex separate)", g("2.5 Cyber Roles", f"G{cyt}"),
    g("0.0 Data Config", "E23"), 1e-6)
w11c = wf["1.11 TDD Cyber"]
ttf11 = None
for rr in range(1, w11c.max_row + 1):
    if w11c.cell(rr, 2).value == "Total to fund": ttf11 = rr
if ttf11 is None or w11c.cell(ttf11 + 3, 2).value != "Total":
    fails.append("1.11 total-to-fund block shape unexpected")
else:
    chk("2.5 tie row = 1.11 total to fund (one cyber number)",
        g("2.5 Cyber Roles", f"C{A['cy_tie']}"), g("1.11 TDD Cyber", f"C{ttf11+3}"), 0.001)

# --- 2.2 COE hub wired to the new role tabs ---
chk("2.2 D8 = SA&D Strategy planned", g("2.2 COE", "D8"),
    g("2.4 SA&D", f"F{A['sad']['cats']['Strategy & Architecture']}"), 1e-9)
chk("2.2 D10 = BP&T Transformation planned", g("2.2 COE", "D10"),
    g("2.3 BP&T", f"F{A['bp']['cats']['Transformation']}"), 1e-9)
chk("2.2 D11 = BP&T Business Partnering planned", g("2.2 COE", "D11"),
    g("2.3 BP&T", f"F{A['bp']['cats']['Business Partnering']}"), 1e-9)
chk("2.2 D12 = SA&D Data planned", g("2.2 COE", "D12"),
    g("2.4 SA&D", f"F{A['sad']['cats']['Data - COE']}"), 1e-9)

# --- 3.0 FTE View: rollups visible & tied ---
FT = "3.0 FTE View"; gf = A["ft_grand"]
chk("3.0 actual squad cost = 2.1 squads actual", g(FT, f"N{gf}"), g(T, f"F{A['subA']}"), 0.005)
ftws = wf[FT]
n_oh_plat = sum(1 for rr in range(1, gf + 1)
                if isinstance(ftws.cell(rr, 3).value, str) and "platform total (incl overhead)" in ftws.cell(rr, 3).value)
n_egi_plat = sum(1 for rr in range(1, gf + 1)
                 if isinstance(ftws.cell(rr, 3).value, str) and "platform total (no overhead" in ftws.cell(rr, 3).value)
chk("3.0 EGI platforms carry no overhead (5 exempt)", n_egi_plat, 5)
L16 = float(g("0.0 Data Config", "L16") or 0); L10 = float(g("0.0 Data Config", "L10") or 0)
chk("3.0 model grand = squads model + non-EGI platform OHs + portfolio OHs",
    g(FT, f"M{gf}"), (g(T, f"C{A['subA']}") or 0) + n_oh_plat*L16 + 11*L10, 0.005)
chk("3.0 model grand TIES 2.0 J17 delivery total", g(FT, f"M{gf}"), g("2.0 Group Summary", "J17"), 0.005)
chk("3.0 cross-check = 0", g(FT, f"C{A['ft_xcheck']}"), 0.0, 1e-9)
chk("3.0 KPI org roles", g(FT, "C4"), 536)
chk("3.0 KPI model cost = grand M", g(FT, "I4"), g(FT, f"M{gf}"), 1e-9)
chk("3.0 KPI actual cost = grand N", g(FT, "J4"), g(FT, f"N{gf}"), 1e-9)
chk("3.0 Cost var grand = 2.1 squads variance (like-for-like)", g(FT, f"O{gf}"), g(T, f"G{A['subA']}"), 0.001)
acc_k = 0.0
for rr in range(7, gf):
    cval = ftws.cell(rr, 3).value
    if isinstance(cval, str) and "platform total" in cval: continue
    bval = ftws.cell(rr, 2).value
    if isinstance(bval, str) and "portfolio total" in bval: continue
    gv, jv = g(FT, f"G{rr}"), g(FT, f"J{rr}")
    if isinstance(gv, (int, float)) and isinstance(jv, (int, float)):
        acc_k += jv - gv
chk("3.0 seats-vs-model grand counts archetype squads only", g(FT, f"K{gf}"), acc_k, 0.5)
# merged cyber row: model and actual on one basis, equal at ship state
cyrow = None
for rr in range(1, gf + 1):
    if ftws.cell(rr, 4).value == "All cyber squads (detail on 2.5 Cyber Roles)": cyrow = rr
if cyrow is None: fails.append("3.0 merged cyber row missing")
else:
    chk("3.0 cyber row model single-sourced from 1.11/2.5", g(FT, f"M{cyrow}"),
        (g("1.11 TDD Cyber", "G24") or 0) + (g("1.11 TDD Cyber", "G25") or 0), 1e-6)
# leadership counts are live formulas, not hardcoded
lead_tot_row = None
for rr in range(gf, ftws.max_row + 1):
    if ftws.cell(rr, 2).value == "Total leadership roles": lead_tot_row = rr
if lead_tot_row is None or not str(ftws.cell(lead_tot_row, 3).value).startswith("=COUNTIF"):
    fails.append("3.0 leadership count not live COUNTIF")
else:
    chk("3.0 leadership count = 53", g(FT, f"C{lead_tot_row}"), 53)
# 2.0 net memo ties to 2.1 model grand
chk("2.0 J26 net = 2.1 model grand", g("2.0 Group Summary", f"J{A['gs_net']}"), g(T, f"C{GR}"), 1e-6)
for cc in ("G5", "H5", "I5"):
    if "($m)" not in str(wf["2.0 Group Summary"][cc].value):
        fails.append(f"2.0 {cc} header missing ($m)")
if wf["2.0 Group Summary"]["C35"].number_format != "0%":
    fails.append("2.0 C35 not 0% format")
# negative-zero guards
for s_, c_ in [("1.10 Z Retail", "I19"), ("2.0 Group Summary", "I15"), ("2.1 Total Cost", "G38")]:
    v = g(s_, c_)
    if isinstance(v, (int, float)) and v != 0 and abs(v) < 1e-9:
        fails.append(f"negative-zero residue at {s_}!{c_}: {v}")
# offshore single-source: squad offshore column derives from K5
sqws = wf["0.1 Squads"]
for rr in (5, 12, 23):
    if sqws.cell(rr, 7).value is not None:
        fs = str(sqws.cell(rr, 8).value)
        if not (fs.startswith("=G") and "$K$5" in fs):
            fails.append(f"0.1 H{rr} not derived from K5: {fs}")
        chk(f"0.1 H{rr} = G{rr}*40%", g("0.1 Squads", f"H{rr}"), float(g("0.1 Squads", f"G{rr}") or 0)*0.4, 1e-9)
# SupportPct list covers all in-use values
spn = wf.defined_names["SupportPct"].value if "SupportPct" in wf.defined_names else ""
if "$D$2:$D$10" not in spn: fails.append(f"SupportPct range wrong: {spn}")
lvals = {wf["Lists"].cell(rr, 4).value for rr in range(2, 11)}
for needed in (0, 0.2, 0.3, 0.4, 0.5, 0.7, 0.8, 0.9, 1):
    if needed not in lvals: fails.append(f"SupportPct missing {needed}")
# 2.3/2.4 allocation rates reference Data Config, not hardcodes
if "$K$7" not in str(wf["2.3 BP&T"][f"C{A['bp']['info_first']+3}"].value):
    fails.append("2.3 BP FTE-equivalents not referencing K7")
if "$K$8" not in str(wf["2.4 SA&D"][f"C{A['sad']['info_first']+3}"].value):
    fails.append("2.4 DA FTE-equivalents not referencing K8")
# every role-tab row is sourced from RAW DATA; SA&D contains no delivery-class roles (brief 6)
for tab_ in ("2.3 BP&T", "2.4 SA&D", "2.5 Cyber Roles"):
    an_ = {"2.3 BP&T": A["bp"], "2.4 SA&D": A["sad"], "2.5 Cyber Roles": A["cyb"]}[tab_]
    ws_ = wf[tab_]
    for rr in range(an_["tbl_first"], an_["tbl_last"] + 1):
        fm = str(ws_[f"F{rr}"].value)
        m2 = re.search(r"'raw data'!\$R\$(\d+)", fm)
        if not m2:
            fails.append(f"{tab_} row {rr} status not sourced from raw data"); break
        if tab_ == "2.4 SA&D":
            cls2 = rawv.cell(int(m2.group(1)), 17).value
            if cls2 in ("Squad", "Strategic Program"):
                fails.append(f"2.4 row {rr} is delivery-class ({cls2}) - brief 6 breach")
        im = str(ws_[f"I{rr}"].value)
        if "'Added data'" not in im:
            fails.append(f"{tab_} row {rr} cost not mapped from Added data"); break
# COE (unspecified) surfaced on 2.2 and costed by the same mapping
unspec_rows_actual = sum(1 for rr_ in range(A["unspec_first"], A["unspec_tot"])
                         if wf["2.2 COE"].cell(rr_, 2).value not in (None, ""))
chk("2.2 unspecified roster = 7 (actual cells)", unspec_rows_actual, 7)
chk("2.2 unspecified total = independent mapping", g("2.2 COE", f"F{A['unspec_tot']}"), unspec_cost_i, 5.0)
# plumbing hidden, GM-facing count controlled
for s_ in ("Lists", "squad mapping", "0.2 FY26 Budget"):
    if wf[s_].sheet_state != "hidden": fails.append(f"{s_} not hidden")
visible_n = sum(1 for s_ in wf.sheetnames if wf[s_].sheet_state == "visible")
chk("visible sheets = 37", visible_n, 37)
# fill bleed cleared
for t_, cc in [("1.2 Customer", "B15"), ("1.2 Customer", "C15"), ("1.3 Enterprise Data", "C14")]:
    if wf[t_][cc].fill.patternType == "solid" and getattr(wf[t_][cc].fill.fgColor, "rgb", None) == "FF1F4E79":
        fails.append(f"header-fill bleed remains at {t_}!{cc}")
# gridlines off on stragglers
for s_ in ("0.3 For Presentation Pack (2)", "0.4 Budget Table (Fin)", "squad mapping", "Lists"):
    if wf[s_].sheet_view.showGridLines is not False:
        fails.append(f"gridlines still on: {s_}")
# role-tab header label
for t_ in ("2.3 BP&T", "2.4 SA&D", "2.5 Cyber Roles"):
    if wf[t_]["H5"].value != "Left to fund ($m)":
        fails.append(f"{t_} H5 header not 'Left to fund ($m)'")
for tab_, an_ in (("2.3 BP&T", A["bp"]), ("2.4 SA&D", A["sad"]), ("2.5 Cyber Roles", A["cyb"])):
    b1 = str(wf[tab_].cell(an_["tbl_first"], 2).value)
    if not b1.startswith("=IF('raw data'"):
        fails.append(f"{tab_} roster names not live from raw data")
qa31 = wf["3.1 Data QA"]
live31 = sum(1 for row_ in qa31.iter_rows() for c_ in row_
             if isinstance(c_.value, str) and c_.value.startswith("=COUNTA"))
if live31 < 2: fails.append("3.1 live check cells missing")
if "snapshot" not in str(qa31["B2"].value): fails.append("3.1 not labelled a snapshot")
if wf["2.1 Total Cost"]["G5"].value != "Actual over/(under) archetype ($m)":
    fails.append("2.1 G5 header not in archetype vocabulary")
if wf["2.1 Total Cost"]["C5"].value != "Archetype model ($m)":
    fails.append("2.1 C5 header not in archetype vocabulary")

# --- Exec Summary: first sheet, bridge present, drill-down works ---
ex = wf["Exec Summary"]
if wf.sheetnames[0] != "Exec Summary": fails.append("Exec Summary is not the first sheet")
exrow = {}
for rr in range(1, ex.max_row + 1):
    b = ex.cell(rr, 2).value
    if isinstance(b, str) and b not in exrow: exrow[b] = rr
def exv(label):
    rr = exrow.get(label)
    return g("Exec Summary", f"C{rr}") if rr else None
chk("Exec: budget 53.8", exv("Total TDD people budget ($m)"), 53.8, 0.005)
chk("Exec: allocated 43.5", exv("Allocated to portfolios + COEs ($m)"), 43.5, 0.005)
chk("Exec: TDD Cost = 2.0 D24", exv("TDD Cost - funded by TDD ($m)"), g("2.0 Group Summary", "D24"), 1e-6)
chk("Exec: funded outside = 2.0 G24", exv("Funded outside TDD ($m)"), g("2.0 Group Summary", "G24"), 1e-6)
chk("Exec: left to fund = 2.0 I24", exv("Left to fund - portfolios + COEs ($m)"), g("2.0 Group Summary", "I24"), 1e-6)
money_sum = ((exv("TDD Cost - funded by TDD ($m)") or 0)
             + (exv("Funded outside TDD ($m)") or 0)
             + (exv("Less: BP & Domain Architect already funded in portfolio overheads ($m)") or 0))
chk("Exec money block FOOTS to the archetype total", money_sum, exv("Total archetype model cost ($m)"), 1e-6)
chk("Exec archetype total = 2.1 grand", exv("Total archetype model cost ($m)"), g(T, f"C{A['grand']}"), 1e-6)
# gap identity holds (asserted from 2.1 engine values; no plumbing row on the exec page)
chk("gap identity: vacant - filled-underspend = gap",
    (g(T, f"E{GR}") or 0) - ((g(T, f"C{GR}") or 0) - (g(T, f"D{GR}") or 0)), g(T, f"G{GR}"), 5e-5)
chk("Exec COE line = 2.2 F13", exv("COEs - left to fund after budgets, see 2.2 ($m)"), g("2.2 COE", "F13"), 1e-9)
chk("Exec cyber line = 2.5 tie", exv("TDD Cyber - needs more than its budget, see 1.11 ($m)"),
    g("2.5 Cyber Roles", f"C{A['cy_tie']}"), 1e-9)
# the people block tells the archetype story and reconciles on its face
chk("Exec people: archetype seats = FTE G4", exv("Seats the archetypes allow - your squads at their set sizes"), g(FT, "G4"), 1e-6)
chk("Exec people: raised seats = G+K", exv("Seats actually raised in those squads - filled + vacant"),
    (g(FT, f"G{gf}") or 0) + (g(FT, f"K{gf}") or 0), 1e-6)
chk("Exec people: raised beyond = K", exv("Seats raised beyond the archetypes"), g(FT, f"K{gf}"), 1e-6)
chk("Exec people: outside-archetype seats = J-G-K", exv("Seats in squads priced outside archetypes (AmPOS, EGI, cyber)"),
    (g(FT, f"J{gf}") or 0) - (g(FT, f"G{gf}") or 0) - (g(FT, f"K{gf}") or 0), 1e-6)
chk("Exec people: vacant 166", exv("Vacant - raised, not yet hired"), 166)
chk("Exec people: filled 370", exv("Filled - people in seats today"), 370)
chk("Exec people: squad-lever vacancies 125", exv("of which squad seats - the GM hire or hold lever"), 125)
chk("Exec people: non-lever vacancies 41", exv("of which leadership, COE and unmapped seats"), 41)
chk("Exec lever cost = sum of GM hire-all (mapped, all 125 seats)",
    exv("of which squad seats - your 4.x GM lever ($m)"), sum(vac_map.values()), 0.01)
chk("Exec money: TDD net after double-count", exv("of which TDD pays - after the double-count ($m)"),
    (g("2.0 Group Summary", "D24") or 0) + (g(T, f"C{A['dedup']}") or 0), 1e-6)
chk("Exec B49 = filled minus archetype (positive = over)",
    exv("Filled seats over/(under) the archetype cost ($m)"),
    (g(T, f"D{GR}") or 0) - (g(T, f"C{GR}") or 0), 1e-4)
chk("Exec drill: archetype seats lookup (Ampol default)",
    exv("Archetype squad seats allowed"), g("Lists", "K2"), 1e-6)
if not any(k.startswith("Vacant seats are priced at standard title rates") for k in exrow):
    fails.append("Exec title-rate caveat line missing")
chk("Exec decision: over-archetype = G grand", exv("Taking the design OVER the archetypes by ($m)"),
    g(T, f"G{GR}"), 1e-9)
if not any(k.startswith("Vacant = open seats in the raw data") for k in exrow):
    fails.append("Exec vacancy/ring-fenced basis line missing")
for miss in ("Why this workbook exists", "How the model is built - key decisions",
             "Portfolio drill-down", "The money", "What it means"):
    if miss not in exrow: fails.append(f"Exec section missing: {miss}")
if not any(k.startswith("The people") for k in exrow):
    fails.append("Exec section missing: The people")
# one left-to-fund definition everywhere: 2.2 F pulls role-tab H
chk("2.2 F8 = 2.4 net", g("2.2 COE", "F8"), g("2.4 SA&D", f"H{A['sad']['cats']['Strategy & Architecture']}"), 1e-9)
chk("2.2 F10 = 2.3 net", g("2.2 COE", "F10"), g("2.3 BP&T", f"H{A['bp']['cats']['Transformation']}"), 1e-9)
chk("2.2 F11 = 2.3 net", g("2.2 COE", "F11"), g("2.3 BP&T", f"H{A['bp']['cats']['Business Partnering']}"), 1e-9)
chk("2.2 F12 = 2.4 net", g("2.2 COE", "F12"), g("2.4 SA&D", f"H{A['sad']['cats']['Data - COE']}"), 1e-9)
chk("2.0 I24 = portfolio left-to-fund + net COE", g("2.0 Group Summary", "I24"),
    (g("2.0 Group Summary", "I17") or 0) + (g("2.2 COE", "F13") or 0), 1e-3)
for rr_ in (8, 10, 11, 12):
    chk(f"2.2 row {rr_} foots: MAX(0, planned - drawdown) = left to fund",
        max(0.0, (g("2.2 COE", f"D{rr_}") or 0) - (g("2.2 COE", f"E{rr_}") or 0)),
        g("2.2 COE", f"F{rr_}"), 1e-6)
chk("2.2 E13 sums the draw-downs", g("2.2 COE", "E13"),
    sum(g("2.2 COE", f"E{rr_}") or 0 for rr_ in (8, 9, 10, 11, 12)), 1e-6)
for tab, (key, an) in ROLE.items():
    chk(f"{tab} listed-vs-counted check = 0", g(tab, f"C{an['ck']}"), 0.0, 1e-9)
if wf["2.0 Group Summary"]["E5"].value != "Variance ($m) = budget - cost":
    fails.append("2.0 E5 header not updated")
lightson_ok = any(isinstance(wf["1.1 Ampol Retail"].cell(rr, cc).value, str)
                  and "TDD Lights On" in wf["1.1 Ampol Retail"].cell(rr, cc).value
                  and "Data Config" in wf["1.1 Ampol Retail"].cell(rr, cc).value
                  for rr in (4, 5, 6) for cc in range(2, 9))
if not lightson_ok: fails.append("1.1 TDD Lights On label not disambiguated")
sel = None
for rr in range(1, ex.max_row + 1):
    if ex.cell(rr, 2).value == "Portfolio" and ex.cell(rr, 3).value == "Ampol Retail":
        sel = rr; break
if sel is None: fails.append("Exec drill-down selector not found")
else:
    chk("Exec drill budget = 2.0 C6", g("Exec Summary", f"C{sel+1}"), g("2.0 Group Summary", "C6"), 1e-6)
    chk("Exec drill TDD cost = 2.0 D6", g("Exec Summary", f"C{sel+2}"), g("2.0 Group Summary", "D6"), 1e-6)
    tv = exv("TDD Variance - to fund ($m)")
    try:
        if float(tv) < -1e-9: fails.append(f"Exec drill TDD Variance negative: {tv}")
    except Exception: fails.append(f"Exec drill TDD Variance not numeric: {tv}")
    chk("Exec drill total cost = 2.0 J6", exv("Total cost ($m)"), g("2.0 Group Summary", "J6"), 1e-6)

# --- reconciliation carry-overs ---
chk("2.0 C30 allocations 43.5", g("2.0 Group Summary", "C30"), 43.5)
chk("2.0 C32 check = 0", g("2.0 Group Summary", "C32"), 0.0, 1e-9)
for t in TABS:
    ws = wf[t]
    found = None
    for rr in range(1, ws.max_row + 1):
        if ws.cell(rr, 2).value == "Total to fund": found = rr
    if not found: fails.append(f"{t}: Total to fund missing"); continue
    for off in (1, 2, 3):
        v = g(t, f"C{found+off}")
        try:
            if float(v) < -1e-9: fails.append(f"{t} C{found+off} negative: {v}")
        except Exception: fails.append(f"{t} C{found+off} not numeric: {v}")

# --- GM working-copy tabs: counts tie to raw data; decisions wired; totals reconcile ---
raw_v = wv["raw data"]
raw_port = Counter(); raw_port_v = Counter()
for rr in range(2, raw_v.max_row + 1):
    if str(raw_v.cell(rr, 17).value or "") in ("Squad", "Strategic Program"):
        p_ = str(raw_v.cell(rr, 14).value or "").strip()
        raw_port[p_] += 1
        if raw_v.cell(rr, 18).value == "Vacant": raw_port_v[p_] += 1
gm_seats_total = 0
for i, t in enumerate(TABS, 1):
    gname = GM_TABS[i - 1]
    if gname not in A.get("gm", {}): fails.append(f"{gname}: no anchors"); continue
    ga = A["gm"][gname]; pl = t.split(" ", 1)[1]
    tr_ = ga["tot"]
    fil = g(gname, f"E{tr_}"); vac = g(gname, f"F{tr_}"); plan = g(gname, f"G{tr_}"); after = g(gname, f"H{tr_}")
    chk(f"{gname} filled = raw", fil, raw_port[pl] - raw_port_v[pl])
    chk(f"{gname} vacant = raw", vac, raw_port_v[pl])
    chk(f"{gname} all decisions default Hold -> planning 0", plan, 0)
    chk(f"{gname} seats after calls = filled (at default)", after, fil, 1e-9)
    chk(f"{gname} hire-all = mapped cost of its raw vacant seats", g(gname, f"C{ga['vac_cost']}"),
        vac_map.get(pl, 0.0), 0.005)
    chk(f"{gname} planned-hire cost = 0 at default", g(gname, f"C{ga['plan_cost']}"), 0.0, 1e-9)
    gm_seats_total += (fil or 0) + (vac or 0)
chk("GM tabs cover every squad seat (425)", gm_seats_total, 425)
# every GM tab has hire/hold validation and no formula errors is covered by PASS2
for gname in GM_TABS:
    wsg = wf[gname]
    has_dec = any(dv.formula1 == '"Hire,Hold"' for dv in wsg.data_validations.dataValidation)
    if not has_dec: fails.append(f"{gname}: Hire/Hold validation missing")
# FTE 'Other unmapped' must never be negative (telescoping guard)
for rr_ in range(gf, ftws.max_row + 1):
    if ftws.cell(rr_, 2).value == "Other unmapped":
        for cc_ in ("I", "J"):
            vv = g(FT, f"{cc_}{rr_}")
            if isinstance(vv, (int, float)) and vv < 0:
                fails.append(f"3.0 Other unmapped {cc_}{rr_} negative: {vv}")
# 1.x formatting post-conditions (cols B..I only; owner scratch J+ untouched)
GREYS_Q = {"FF808080", "FF7F7F7F", "FFA6A6A6", "FF999999", "FFBFBFBF"}
for t_ in TABS:
    ws_ = wf[t_]
    for row_ in ws_.iter_rows(min_col=2, max_col=9):
        for c_ in row_:
            if c_.value is None or c_.font is None: continue
            try:
                if c_.font.italic or (c_.font.size or 10) < 10:
                    fails.append(f"italic/tiny at {t_}!{c_.coordinate}")
                if getattr(c_.font.color, "rgb", None) in GREYS_Q:
                    fails.append(f"grey font at {t_}!{c_.coordinate}")
            except AttributeError:
                pass
# no en/em dashes in DISPLAYED (evaluated) values on authored tabs, wherever they come from
AUTHORED_U = {s.upper() for s in ["Exec Summary", "2.0 Group Summary", "2.1 Total Cost", "2.2 COE",
              "2.3 BP&T", "2.4 SA&D", "2.5 Cyber Roles", "3.0 FTE View"] + GM_TABS}
dash_vals = [(s, c) for (s, c), v in vals.items()
             if s in AUTHORED_U and isinstance(v, str) and ("–" in v or "—" in v)]
if dash_vals: fails.append(f"en/em dash in displayed values: {dash_vals[:6]}")
# no en/em dashes, no italics, no sub-10pt fonts on authored model tabs
for tname in ["Exec Summary", "2.0 Group Summary", "2.1 Total Cost", "2.2 COE", "2.3 BP&T",
              "2.4 SA&D", "2.5 Cyber Roles", "3.0 FTE View", "3.1 Data QA"] + GM_TABS:
    wsg = wf[tname]
    for row in wsg.iter_rows():
        for c in row:
            if isinstance(c.value, str) and not c.value.startswith("=") and ("–" in c.value or "—" in c.value):
                fails.append(f"dash in {tname}!{c.coordinate}")
            try:
                if c.value is not None and c.font is not None and (c.font.italic or (c.font.size or 10) < 10):
                    fails.append(f"italic/tiny font at {tname}!{c.coordinate}")
            except AttributeError:
                pass

# --- hygiene: no comments; 3.1 Data QA headline numbers ---
for ws in wf.worksheets:
    stop = False
    for row in ws.iter_rows():
        for c in row:
            if c.comment is not None:
                fails.append(f"comment at {ws.title}!{c.coordinate}"); stop = True; break
        if stop: break
qa = wf["3.1 Data QA"]
hd = {}
for rr in range(1, 15):
    b = qa.cell(rr, 2).value
    if b in ("Records", "Vacant"): hd[b] = (qa.cell(rr, 3).value, qa.cell(rr, 4).value)
chk("3.1 raw records 536", hd.get("Records", (0, 0))[0], 536)
chk("3.1 added records 548", hd.get("Records", (0, 0))[1], 548)

# =====================================================================
# FUNCTIONAL offshore-toggle test: engine run 2 on a toggled copy
# =====================================================================
print("PASS3: engine run 2 (offshore toggled)...", flush=True)
import gc
del xl, sol
gc.collect()
F2 = SCR + "_qa_v8_offshore.xlsx"
shutil.copy(F, F2)
w2 = openpyxl.load_workbook(F2)
tog = {}
for tab, (key, an) in ROLE.items():
    rr = an["tbl_first"]
    w2[tab][f"H{rr}"] = "Offshore"
    tog[tab] = rr
w2.save(F2)
xl2 = formulas.ExcelModel().loads(F2).finish()
sol2 = xl2.calculate()
vals2 = {}
for k, v in sol2.items():
    m = re.match(r"^'?\[[^\]]*\]([^!]*?)'?!([A-Z]+\d+)$", k)
    if not m: continue
    val = v.value
    try: val = val[0, 0]
    except Exception: pass
    vals2[(m.group(1).strip().upper(), m.group(2))] = val
def g2(s, c): return vals2.get((s.upper(), c))
for tab, (key, an) in ROLE.items():
    rr = tog[tab]
    full = g(tab, f"I{rr}")
    chk(f"OFFSHORE {tab} J{rr} = 40% of full cost", g2(tab, f"J{rr}"), float(full)*0.4/1e6, 1e-6)
    # planned total must DROP by exactly 60% of that role's cost
    chk(f"OFFSHORE {tab} total moved by -60%", g2(tab, f"F{an['tot']}"),
        (g(tab, f"F{an['tot']}") or 0) - float(full)*0.6/1e6, 1e-6)

print("FAILS:", len(fails))
for f_ in fails: print("  -", f_)
print(f"KEY: actual grand={g(T, f'F{GR}')} model grand={g(T, f'C{GR}')} "
      f"squads model={g(T, 'C'+str(A['subA']))} "
      f"FTE model grand={g(FT, f'M{gf}')} FTE actual={g(FT, f'N{gf}')}")
sys.exit(1 if fails else 0)
