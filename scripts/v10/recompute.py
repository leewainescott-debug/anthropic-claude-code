"""Recompute every figure a reader sees, from REVIEW, and compare it to the tab.

The other passes test the workbook against itself: a control row, a cross-tab agreement, a
formula error. All of those can pass while the whole model is consistently wrong, and twice
in this build they did - a subtotal that summed row variances and dropped every row carrying
a dash, and a subtotal with no archetype that reported the whole cost as a variance. Both
added up. Both were wrong.

This pass builds the numbers again from the ledger, in Python, and asserts the tab matches.
It knows nothing about the workbook's formulas.
"""
import collections
import json
import sys

import openpyxl

REVIEW = "REVIEW - Complete Role Mapping"
PF = ["Ampol Retail", "Customer", "Enterprise Data", "TDD Group Functions", "P&C",
      "Finance", "Infrastructure", "Energy Solutions & B2B", "Commercial Fuels",
      "Z Retail"]
COE = ["COE Cyber", "COE BP&T", "COE SA&D", "EGI"]
TOL = 1e-6


def last(wv):
    R = wv[REVIEW]
    r = R.max_row
    while r > 1 and not str(R.cell(r, 2).value or '').strip():
        r -= 1
    return r


def ledger(wv):
    R = wv[REVIEW]
    cost, roles, filled, vacant, ohcost, ohroles = (collections.Counter() for _ in range(6))
    for i in range(2, last(wv) + 1):
        if not str(R.cell(i, 2).value or "").strip():
            continue
        pf, grp = str(R.cell(i, 36).value), str(R.cell(i, 46).value)
        c = R.cell(i, 27).value or 0
        cost[(pf, grp)] += c
        roles[(pf, grp)] += 1
        if str(R.cell(i, 37).value) == "Filled":
            filled[(pf, grp)] += 1
        else:
            vacant[(pf, grp)] += 1
        if str(R.cell(i, 44).value) != "Squad":
            # AR == AT selects the roles that sit in a portfolio; a COE role carrying an
            # overhead title has AT set to its COE squad
            key = "pf" if str(R.cell(i, 44).value) == grp else "coe"
            ohcost[key] += c
            ohroles[key] += 1
    return cost, roles, filled, vacant, ohcost, ohroles


def find(ws, label, col=2):
    """Exact match first, then a prefix, so a label that gains a clause still resolves."""
    for exact in (True, False):
        for r in range(1, ws.max_row + 1):
            v = ws.cell(r, col).value
            if not isinstance(v, str):
                continue
            if (v.strip() == label) if exact else v.strip().startswith(label):
                return r
    return None


def col_of(ws, header, limit=12):
    for r in range(1, limit + 1):
        for c in range(2, 24):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip() == header:
                return c
    return None


def check(out, name, got, exp, count=False):
    if got is None or exp is None:
        out.append(f"{name}: tab {got!r}, recomputed {exp!r}")
        return False
    ok = (got == exp) if count else abs(got - exp) < TOL
    if not ok:
        out.append(f"{name}: tab {got!r}, recomputed {exp!r}")
    return ok


def run(path, anchors="anchors_final.json"):
    wv = openpyxl.load_workbook(path, data_only=True)
    a = json.load(open(anchors))
    cost, roles, filled, vacant, ohcost, ohroles = ledger(wv)
    # tabs are renamed after the anchors are written, so map by the portfolio cell
    tab_of = {str(wv[t]["C3"].value): t for t in wv.sheetnames if t.startswith("2.")}
    out = []

    # ---- every squad row on every working tab ----
    S = a[list(a)[0]]["cols"]
    for v in a.values():
        ws = wv[tab_of[v["pf"]]]
        for g in v["squads"] + v["direct"] + v["nofig"] + v["overhead"]:
            r, k = v["srow"][g], (v["pf"], g)
            for key, exp, cnt in (("actual", cost[k] / 1e6, False),
                                  ("roles", roles[k], True),
                                  ("filled", filled[k], True),
                                  ("vacant", vacant[k], True)):
                check(out, f"{tab_of[v['pf']]} {g} {key}", ws.cell(r, S[key]).value,
                      exp, cnt)

    # ---- the bridge on 3.1 ----
    b = next(t for t in wv.sheetnames if t.startswith("3.1 "))
    ws = wv[b]
    ca, cc, cr = (col_of(ws, "Actual cost ($m)"), col_of(ws, "Archetype cost ($m)"),
                  col_of(ws, "Total roles"))
    arch = sum(cost[(v["pf"], g)] for v in a.values() for g in v["squads"])
    narch = sum(roles[(v["pf"], g)] for v in a.values() for g in v["squads"])
    direct = sum(cost[(v["pf"], g)] for v in a.values() if v["pf"] in PF
                 for g in v["direct"])
    ndir = sum(roles[(v["pf"], g)] for v in a.values() if v["pf"] in PF
               for g in v["direct"])
    nofig = sum(cost[(v["pf"], g)] for v in a.values() if v["pf"] in PF
                for g in v["nofig"])
    nnof = sum(roles[(v["pf"], g)] for v in a.values() if v["pf"] in PF
               for g in v["nofig"])
    coe = sum(cost[(v["pf"], g)] for v in a.values() if v["pf"] in COE
              for g in v["direct"] + v["squads"] + v["nofig"] + v["overhead"])
    ncoe = sum(roles[(v["pf"], g)] for v in a.values() if v["pf"] in COE
               for g in v["direct"] + v["squads"] + v["nofig"] + v["overhead"])
    # the directly funded step is two subtotals, split by whether a funded figure is set,
    # so its actual cost is the pair added
    DIRECT = ("Directly funded, where the funded figure is set",
              "Directly funded, where no funded figure is set yet")
    for labs, c, n in ((("Squads priced by an archetype",), arch, narch),
                       (DIRECT, direct, ndir),
                       (("COEs and EGI",), coe, ncoe),
                       (("Groups with no archetype and no funded figure",), nofig, nnof),
                       (("Overhead roles in the portfolios",), ohcost["pf"],
                        ohroles["pf"])):
        rs = [find(ws, x) for x in labs]
        if any(x is None for x in rs):
            out.append(f"3.1 has no row {labs!r}")
            continue
        name = labs[0]
        check(out, f"3.1 {name} cost",
              round(sum(ws.cell(x, ca).value or 0 for x in rs), 9), round(c / 1e6, 9))
        check(out, f"3.1 {name} roles", sum(ws.cell(x, cr).value or 0 for x in rs), n, True)
    total = sum(cost.values())
    r = find(ws, "Cost of the")
    check(out, "3.1 ledger cost", ws.cell(r, ca).value, total / 1e6)
    check(out, "3.1 ledger roles", ws.cell(r, cr).value, sum(roles.values()), True)
    # The archetype side is a design figure, so it is checked for internal consistency:
    # the steps that carry one must add to the comparable subtotal. It used to be checked
    # against the ledger line, which only worked while every step had an archetype - and
    # three of them only had one because the column was restating the actual.
    steps = [find(ws, x) for x in
             ("Squads priced by an archetype",
              "Directly funded, where the funded figure is set",
              "Overhead roles in the portfolios - the allowance is on 3.2")]
    cmp_row = find(ws, "Everything with a figure to compare")
    if None in steps or cmp_row is None:
        out.append("3.1 is missing one of the comparable steps")
    else:
        for col, what in ((cc, "archetype"), (ca, "actual")):
            check(out, f"3.1 comparable steps add up - {what}",
                  round(sum(v for v in (ws.cell(x, col).value for x in steps)
                            if isinstance(v, (int, float))), 6),
                  round(ws.cell(cmp_row, col).value, 6))
    # the totals carry the comparison, so the archetype on the ledger row has to be exactly
    # the steps that have one - no more, and nothing dropped
    priced = [find(ws, x) for x in ("Squads priced by an archetype",
                                    "Directly funded, where the funded figure is set",
                                    "Overhead roles in the portfolios - the allowance")]
    k = find(ws, "Cost of the")
    check(out, "3.1 ledger archetype is the priced steps",
          round(ws.cell(k, cc).value, 6),
          round(sum(ws.cell(x, cc).value for x in priced if x), 6))
    g2 = find(ws, "Total cost of TDD including the GM layer")
    check(out, "3.1 grand total variance is actual less archetype",
          round(ws.cell(g2, 6).value, 6),
          round(ws.cell(g2, ca).value - ws.cell(g2, cc).value, 6))
    gm = wv["Lists"]["AG12"].value
    g = find(ws, "Total cost of TDD including the GM layer")
    check(out, "3.1 grand total", ws.cell(g, ca).value, total / 1e6 + gm)

    # ---- 3.2 overhead, portfolios against the COEs ----
    o = next(t for t in wv.sheetnames if t.startswith("3.2 "))
    ws = wv[o]
    r = find(ws, "Of which sits in the")
    check(out, "3.2 portfolio overhead cost",
          ws.cell(r, col_of(ws, "Cost in the portfolios ($m)")).value,
          ohcost["pf"] / 1e6)
    check(out, "3.2 portfolio overhead roles",
          ws.cell(r, col_of(ws, "Roles in the portfolios")).value, ohroles["pf"], True)
    t = find(ws, "Every overhead line")
    check(out, "3.2 overhead inside the COEs",
          ws.cell(t, col_of(ws, "Cost inside the COEs ($m)")).value, ohcost["coe"] / 1e6)

    # ---- 3.3 group total and 3.4 COE total ----
    d = next(t for t in wv.sheetnames if t.startswith("3.3 "))
    ws = wv[d]
    r = find(ws, "Group total")
    check(out, "3.3 group cost", ws.cell(r, col_of(ws, "Actual cost ($m)")).value,
          total / 1e6)
    check(out, "3.3 group roles", ws.cell(r, col_of(ws, "Total roles")).value,
          sum(roles.values()), True)
    c4 = next(t for t in wv.sheetnames if t.startswith("3.4 "))
    ws = wv[c4]
    r = find(ws, "COEs and EGI total")
    check(out, "3.4 COE cost", ws.cell(r, col_of(ws, "Cost ($m)")).value, coe / 1e6)

    # ---- Exec ----
    ws = wv["Exec Summary"]
    filled_cost = 0.0
    R = wv[REVIEW]
    for i in range(2, last(wv) + 1):
        if str(R.cell(i, 2).value or "").strip() and str(R.cell(i, 37).value) == "Filled":
            filled_cost += R.cell(i, 27).value or 0
    for lab, exp, cnt in (("Roles in the ledger", sum(roles.values()), True),
                          ("Cost of the", total / 1e6,
                           False),
                          ("Of which filled roles ($m)", filled_cost / 1e6, False),
                          ("Cost today including the GM layer ($m)", total / 1e6 + gm,
                           False)):
        r = find(ws, lab)
        if r is None:
            out.append(f"Exec has no row {lab!r}")
            continue
        check(out, f"Exec {lab}", ws.cell(r, 3).value, exp, cnt)
    return out


if __name__ == "__main__":
    bad = run(sys.argv[1] if len(sys.argv) > 1 else "cand.xlsx")
    print("=" * 74)
    print(f"FIGURES DISAGREEING WITH A RECOMPUTATION FROM THE LEDGER: {len(bad)}")
    for x in bad[:40]:
        print("   ", x)
