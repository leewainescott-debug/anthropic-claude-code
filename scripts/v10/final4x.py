"""Rewire everything that read the old 3.x row positions, and rebuild Exec and 4.0.

Rebuilding 3.1, 3.2 and 3.3 moves every row on them. Anchors come from the JSON the
builders write rather than from a label search, because two rows on 3.1 now begin with the
same words - a pale section label and the grey subtotal under it - and a search would find
the wrong one.

4.0 used to test one thing thirty-seven ways: roles and cost against the ledger. It could
not fail on a design number, on an allowance, or on a control cell reading -3 somewhere
else in the file, and one of those was live. It now tests the design side too.
"""
import json
import re

import openpyxl
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

import final2x as f2
import opts

REVIEW = f2.REVIEW
REV = f2.REV
LAST = f2.LAST
S = f2.S
A3 = "'0.3 Squad Archetypes'"
G1, G2, G3 = "'3.1 Group Summary'", "'3.2 Total Cost'", "'3.3 FTE View'"
# 3.3 columns after the "How it is funded" column was added
C33 = dict(pf="B", kind="C", squad="D", aroles="G", roles="H", filled="I", vacant="J",
           rafter="K", acost="L", actual="M", var="N", after="O")
# 3.1 is a bridge now: the line name is in B, the portfolio in C, and the figures start at D
C31 = dict(name="B", pf="C", acost="D", actual="E", var="F", after="G", roles="H",
           filled="I", vacant="J", rafter="K")
# 3.2 gained a Basis column, so every figure on it moved one to the right
C32 = dict(line="B", basis="C", rate="D", times="E", allow="F", pfroles="G", pfcost="H",
           notcov="I", coeroles="J", coecost="K")


def find_row(ws, label, col=2, limit=200, exact=False):
    for r in range(1, min(ws.max_row, limit) + 1):
        v = ws.cell(r, col).value
        if not isinstance(v, str):
            continue
        if (v.strip() == label) if exact else v.strip().startswith(label):
            return r
    raise KeyError(f"{ws.title}: no row matching {label!r}")


def anchors(wb):
    a3 = json.load(open("anchors_final3.json"))
    s2, s3 = wb[G2.strip("'")], wb[G3.strip("'")]
    a = dict(a3["3.1"])
    a["ohtot32"] = find_row(s2, "Every overhead line")
    a["ohpf32"] = find_row(s2, "Of which sits in the 525-role ledger")
    a["g33"] = find_row(s3, "Group total")
    a["first33"] = find_row(s3, "Portfolio") + 1
    return a


def repoint(wb, a):
    """Lists!K prices the archetype roles per portfolio out of 3.3.

    3.3 gained a column, so archetype roles moved from F to G, and the old formula excluded
    total rows by testing the Squad column for '*total'. The total row's label now sits in
    the portfolio column, so that test matched nothing and the totals were being counted
    twice. Selecting on the new 'How it is funded' column is both narrower and honest: only
    archetype rows carry an archetype.
    """
    lo, hi = a["first33"], a["g33"] - 1
    new = (f"=SUMIFS({G3}!${C33['aroles']}${lo}:${C33['aroles']}${hi},"
           f"{G3}!${C33['pf']}${lo}:${C33['pf']}${hi},$J{{r}},"
           f'{G3}!${C33["kind"]}${lo}:${C33["kind"]}${hi},"Archetype")')
    l = wb["Lists"]
    n = 0
    for r in range(2, 20):
        if not str(l.cell(r, 10).value or "").strip():
            continue
        if isinstance(l.cell(r, 11).value, str) and "3.3 FTE View" in l.cell(r, 11).value:
            l.cell(r, 11).value = new.format(r=r)
            n += 1
    return n


# ------------------------------------------------------------------- Exec Summary
def build_exec(wb, a, a2):
    ws = wb["Exec Summary"]
    f2.wipe(ws)
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 62
    ws.column_dimensions["C"].width = 18
    gt = a["total"]
    ws.cell(2, 2).value = "TDD operating model - executive summary"
    ws.cell(2, 2).font = opts.TITLE

    def block(r, heading, lines):
        r = opts.bar(ws, r, 2, 2, heading)
        for lab, f, nf in lines:
            ws.cell(r, 2).value = lab
            ws.cell(r, 2).font = opts.BODY
            ws.cell(r, 2).alignment = opts.LFT
            ws.cell(r, 2).border = opts.BOX
            x = ws.cell(r, 3)
            x.value, x.number_format = f, nf
            x.alignment, x.font, x.border = opts.RGT, opts.BODY, opts.BOX
            r += 1
        return r + 1

    r = block(4, "The organisation today", [
        ("Roles in the ledger", f"={G1}!${C31['roles']}${gt}", opts.CT),
        ("Filled", f"={G1}!${C31['filled']}${gt}", opts.CT),
        ("Vacant", f"={G1}!${C31['vacant']}${gt}", opts.CT),
        ("Cost of the 525 roles in the ledger ($m)", f"={G1}!${C31['actual']}${gt}", opts.M2),
        ("Of which filled roles ($m)",
         f'=SUMIFS({REV}!$AA$2:$AA${LAST},{REV}!$AK$2:$AK${LAST},"Filled")/1000000',
         opts.M2),
        ("Of which vacancies not yet filled ($m)",
         f'=SUMIFS({REV}!$AA$2:$AA${LAST},{REV}!$AK$2:$AK${LAST},"Vacant")/1000000',
         opts.M2),
        ("The 8 GMs, outside the ledger ($m)", "=N(Lists!$AG$12)", opts.M2),
        ("Cost today including the GM layer ($m)", f"={G1}!${C31['actual']}${a['grand']}", opts.M2)])

    r = block(r, "Against the archetype", [
        ("Squads priced by an archetype - archetype cost ($m)",
         f"={G1}!${C31['acost']}${a['arch']}", opts.M2),
        ("Squads priced by an archetype - actual ($m)",
         f"={G1}!${C31['actual']}${a['arch']}", opts.M2),
        ("Squads priced by an archetype - over/(under) ($m)",
         f"={G1}!${C31['var']}${a['arch']}", opts.M2),
        ("Directly funded programmes - over/(under) funded ($m)",
         f"={G1}!${C31['var']}${a['direct']}", opts.M2),
        ("COEs and EGI - over/(under) their 1.x planned spend ($m)",
         f"={G1}!${C31['var']}${a['coe']}", opts.M2),
        ("Overhead roles - not covered by the allowance ($m)",
         f"={G1}!${C31['var']}${a['overhead']}", opts.M2),
        # without this line the four components above summed to 6.378 under a total of
        # 8.478, because the total includes the GM layer and nothing listed it
        ("The 8 GMs - over/(under) their allowance ($m)",
         f"={G1}!${C31['var']}${a['gm']}", opts.M2),
        ("Total over/(under) archetype, everything comparable ($m)",
         f"={G1}!${C31['var']}${a['comparable']}+{G1}!${C31['var']}${a['gm']}",
         opts.M2),
        ("Groups with no archetype and no funded figure ($m)",
         f"={G1}!${C31['actual']}${a['nofig']}", opts.M2)])

    r = block(r, "The vacancy decision", [
        ("Vacant roles", f"={G1}!${C31['vacant']}${gt}", opts.CT),
        ("Vacancies set to hire", "=" + "+".join(
            f"N('{t}'!${L(S['hire'])}${i['total_row']})" for t, i in a2.items()),
         opts.CT),
        ("Roles set to offshore", "=" + "+".join(
            f"N('{t}'!${L(S['offshore'])}${i['total_row']})" for t, i in a2.items()),
         opts.CT),
        ("Roles put on hold", "=" + "+".join(
            f"N('{t}'!${L(S['hold'])}${i['total_row']})" for t, i in a2.items()),
         opts.CT),
        ("Roles after the decisions set today", f"={G1}!${C31['rafter']}${gt}", opts.CT),
        ("Cost after the decisions set today ($m)", f"={G1}!${C31['after']}${gt}", opts.M2),
        ("Impact of those decisions ($m)", f"={G1}!${C31['after']}${gt}-{G1}!${C31['actual']}${gt}", opts.M2)])

    # portfolio drill-down. The name list is read out of 3.1's archetype block, which
    # carries one row per portfolio.
    r = opts.bar(ws, r, 2, 2, "Portfolio drill-down")
    sel = r
    ws.cell(r, 2).value = "Pick a portfolio"
    ws.cell(r, 2).font = opts.BOLD
    ws.cell(r, 2).alignment = opts.LFT
    ws.cell(r, 2).border = opts.BOX
    pick = ws.cell(r, 3)
    pick.value = "Ampol Retail"
    pick.fill, pick.border = opts.fl(opts.YEL), opts.BOX
    pick.font, pick.alignment = opts.BODY, opts.CEN
    l = wb["Lists"]
    names = [str(l.cell(x, 45).value) for x in range(2, 12) if l.cell(x, 45).value]
    dv = DataValidation(type="list", formula1='"' + ",".join(names) + '"',
                        allow_blank=False, showDropDown=False)
    ws.add_data_validation(dv)
    dv.add(pick)
    r += 1
    # Read out of 3.3, which carries every squad of every portfolio tagged by how it is
    # funded. Reading it off 3.1 instead would have lost the portfolio's overhead cost,
    # because the overhead allowance is a group figure and sits on one row for all ten.
    lo3, hi3 = a["first33"], a["g33"] - 1
    def s33(col, kind=None):
        # rounded to six places: summing 3.3's already-rounded per-squad variances gave
        # (0.766475) against (0.766476) for the same fact on 3.1
        f = (f"=ROUND(SUMIFS({G3}!${col}${lo3}:${col}${hi3},"
             f"{G3}!${C33['pf']}${lo3}:${C33['pf']}${hi3},$C${sel}")
        if kind:
            f += f',{G3}!${C33["kind"]}${lo3}:${C33["kind"]}${hi3},"{kind}"'
        return f + "),6)"

    for lab, f, nf in (
            ("Roles", s33(C33["roles"]), opts.CT),
            ("Filled", s33(C33["filled"]), opts.CT),
            ("Vacant", s33(C33["vacant"]), opts.CT),
            ("Roles after decisions", s33(C33["rafter"]), opts.CT),
            ("Actual cost ($m)", s33(C33["actual"]), opts.M2),
            ("Of which overhead roles ($m)", s33(C33["actual"], "Overhead"), opts.M2),
            ("Squads priced by an archetype - over/(under) archetype ($m)",
             s33(C33["var"], "Archetype"), opts.M2),
            ("Directly funded programmes - over/(under) funded ($m)",
             s33(C33["var"], "Directly funded"), opts.M2),
            ("Cost after vacancy decisions ($m)", s33(C33["after"]), opts.M2)):
        ws.cell(r, 2).value = lab
        ws.cell(r, 2).font = opts.BODY
        ws.cell(r, 2).alignment = opts.LFT
        ws.cell(r, 2).border = opts.BOX
        x = ws.cell(r, 3)
        x.value, x.number_format, x.alignment = f, nf, opts.RGT
        x.font, x.border = opts.BODY, opts.BOX
        r += 1
    return r


# ---------------------------------------------------------------------- 4.0 Data QA
def coe_control(wb):
    """The 'must be 0' cells the COE tabs carry, found by their own label."""
    out = []
    for tab in ("1.11 BP&T", "1.12 SA&D", "1.13 Cyber Roles"):
        ws = wb[tab]
        for r in range(1, ws.max_row + 1):
            v = ws.cell(r, 2).value
            if isinstance(v, str) and v.strip().lower().startswith("check"):
                for c in range(3, 12):
                    if isinstance(ws.cell(r, c).value, str) and \
                            ws.cell(r, c).value.startswith("="):
                        out.append((tab, f"{L(c)}{r}"))
                        break
                break
    return out


def build_qa(wb, a, a2):
    ws = wb["4.0 Data QA"]
    f2.wipe(ws)
    ws.column_dimensions["A"].width = 2
    ws.cell(2, 2).value = "Data QA - every difference must read zero"
    ws.cell(2, 2).font = opts.TITLE
    HDR = 4
    opts.head(ws, HDR, 2, ["Check", "Model", "Expected", "Difference"], [72, 16, 16, 14])
    gt, g33 = a["total"], a["g33"]
    lo33, hi33 = a["first33"], g33 - 1
    # "<>Squad" also matches the two empty ledger rows, whose AR is a blank string, so
    # the count came to 65 against 63 real overhead roles. The name column is the guard.
    oh = f'{REV}!$AR$2:$AR${LAST},"<>Squad",{REV}!$B$2:$B${LAST},"<>"'
    checks = [
        # ---- the ledger ----
        ("Roles on 3.1 against the ledger", f"={G1}!${C31['roles']}${gt}",
         f"=COUNTA({REV}!$B$2:$B${LAST})", opts.CT),
        ("Filled on 3.1 against the ledger", f"={G1}!${C31['filled']}${gt}",
         f'=COUNTIFS({REV}!$AK$2:$AK${LAST},"Filled")', opts.CT),
        ("Vacant on 3.1 against the ledger", f"={G1}!${C31['vacant']}${gt}",
         f'=COUNTIFS({REV}!$AK$2:$AK${LAST},"Vacant")', opts.CT),
        ("Filled plus vacant against roles on 3.1", f"={G1}!${C31['filled']}${gt}+{G1}!${C31['vacant']}${gt}",
         f"={G1}!${C31['roles']}${gt}", opts.CT),
        ("Cost on 3.1 against the ledger ($m)", f"={G1}!${C31['actual']}${gt}",
         f"=SUM({REV}!$AA$2:$AA${LAST})/1000000", opts.M2),
        ("Filled plus vacant cost against cost today ($m)",
         f'=(SUMIFS({REV}!$AA$2:$AA${LAST},{REV}!$AK$2:$AK${LAST},"Filled")'
         f'+SUMIFS({REV}!$AA$2:$AA${LAST},{REV}!$AK$2:$AK${LAST},"Vacant"))/1000000',
         f"={G1}!${C31['actual']}${gt}", opts.M2),
        # ---- summary against summary, by a different route ----
        ("Roles on 3.3 against 3.1", f"={G3}!${C33['roles']}${g33}",
         f"={G1}!${C31['roles']}${gt}", opts.CT),
        ("Cost on 3.3 against 3.1 ($m)", f"={G3}!${C33['actual']}${g33}",
         f"={G1}!${C31['actual']}${gt}", opts.M2),
        ("Total including the GM layer against the ledger plus the GM input ($m)",
         f"={G1}!${C31['actual']}${a['grand']}",
         f"=SUM({REV}!$AA$2:$AA${LAST})/1000000+N(Lists!$AG$12)", opts.M2),
        ("Archetype variance - the four steps against the comparable subtotal ($m)",
         "=" + "+".join(f"N({G1}!${C31['var']}${a[k]})"
                        for k in ("arch", "direct", "coe", "overhead")),
         f"={G1}!${C31['var']}${a['comparable']}", opts.M2),
        ("Roles including the GM layer against 525 plus the GM count",
         f"={G1}!${C31['roles']}${a['grand']}",
         f"=COUNTA({REV}!$B$2:$B${LAST})+N(Lists!$AG$11)", opts.CT),
        # ---- the design side ----
        ("Archetype cost, squad by squad on 3.3, against 3.1 ($m)",
         f"=SUMIFS({G3}!${C33['acost']}${lo33}:${C33['acost']}${hi33},"
         f'{G3}!${C33["kind"]}${lo33}:${C33["kind"]}${hi33},"Archetype")',
         f"={G1}!${C31['acost']}${a['arch']}", opts.M2),
        ("Directly funded amount, squad by squad on 3.3, against 3.1 ($m)",
         f"=SUMIFS({G3}!${C33['acost']}${lo33}:${C33['acost']}${hi33},"
         f'{G3}!${C33["kind"]}${lo33}:${C33["kind"]}${hi33},"Directly funded")',
         f"={G1}!${C31['acost']}${a['direct']}+{G1}!${C31['acost']}${a['coe']}", opts.M2),
        ("Archetype roles on 3.3 against the priced-per-portfolio list on Lists",
         f"=SUMIFS({G3}!${C33['aroles']}${lo33}:${C33['aroles']}${hi33},"
         f'{G3}!${C33["kind"]}${lo33}:${C33["kind"]}${hi33},"Archetype")',
         "=SUM(Lists!$K$2:$K$11)", opts.C1),
        ("Offshore archetype against 40% of onshore, first archetype",
         f"=ROUND({A3}!$H$5/{A3}!$G$5,6)", f"=ROUND({A3}!$K$5,6)", opts.C1),
        # ---- the overhead allowance ----
        ("Overhead allowance on 3.2 against Lists ($m)",
         f"={G2}!${C32['allow']}${a['ohtot32']}", "=N(Lists!$AJ$8)", opts.M2),
        ("Allowance drawn in the portfolios against the lines that draw it ($m)",
         "=N(Lists!$AJ$9)", '=SUMIF(Lists!$AM$2:$AM$7,"Yes",Lists!$AJ$2:$AJ$7)',
         opts.M2),
        ("Overhead roles in the portfolios plus those inside the COEs against "
         "every overhead role",
         f"={G2}!${C32['pfroles']}${a['ohpf32']}"
         f"+{G2}!${C32['coeroles']}${a['ohtot32']}",
         f"=COUNTIFS({oh})", opts.CT),
        ("Overhead cost in the portfolios plus inside the COEs against every "
         "overhead role ($m)",
         f"={G2}!${C32['pfcost']}${a['ohpf32']}"
         f"+{G2}!${C32['coecost']}${a['ohtot32']}",
         f"=SUMIFS({REV}!$AA$2:$AA${LAST},{oh})/1000000", opts.M2),
        ("Overhead on 3.1 against the portfolio lines on 3.2 ($m)",
         f"={G1}!${C31['actual']}${a['overhead']}",
         f"={G2}!${C32['pfcost']}${a['ohpf32']}", opts.M2),
        # ---- the lever ----
        ("Roles after decisions against roles less anything on hold",
         f"={G1}!${C31['rafter']}${gt}",
         f"={G1}!${C31['roles']}${gt}-" + "-".join(
             f"N('{t}'!${L(S['hold'])}${i['total_row']})" for t, i in a2.items()),
         opts.CT),
        ("Cost after decisions against cost today, with no lever pulled ($m)",
         f"={G1}!${C31['after']}${gt}", f"={G1}!${C31['actual']}${gt}", opts.M2),
    ]
    # the COE tabs price their own planned spend off their own roles list. Where that list
    # is short of the ledger the planned spend is short too, and nothing else in the file
    # notices: 1.12 was missing three roles worth $747,896.
    for tab, pf, cell in (("1.11 BP&T", "COE BP&T", "$F$6+'1.11 BP&T'!$F$7"),
                          ("1.12 SA&D", "COE SA&D", "$G$6+'1.12 SA&D'!$G$7"),
                          ("1.13 Cyber Roles", "COE Cyber",
                           "$F$6+'1.13 Cyber Roles'!$F$7")):
        checks.append((f"{tab} planned spend against the ledger ($m)",
                       f"='{tab}'!{cell}",
                       f'=SUMIFS({REV}!$AA$2:$AA${LAST},'
                       f'{REV}!$AJ$2:$AJ${LAST},"{pf}")/1000000', opts.M2))
    for tab, cell in coe_control(wb):
        checks.append((f"{tab} own control - roles listed against roles counted",
                       f"='{tab}'!{cell}", "=0", opts.CT))
    for tab, inf in a2.items():
        t = inf["total_row"]
        checks.append((f"{tab} roles against the ledger",
                       f"='{tab}'!${L(S['roles'])}${t}",
                       f"=COUNTIFS({REV}!$AJ$2:$AJ${LAST},\"{inf['pf']}\")", opts.CT))
        checks.append((f"{tab} cost against the ledger ($m)",
                       f"='{tab}'!${L(S['actual'])}${t}",
                       f"=SUMIFS({REV}!$AA$2:$AA${LAST},"
                       f"{REV}!$AJ$2:$AJ${LAST},\"{inf['pf']}\")/1000000", opts.M2))
    r = HDR + 1
    for lab, m, e, nf in checks:
        ws.cell(r, 2).value = lab
        ws.cell(r, 2).font = opts.BODY
        ws.cell(r, 2).alignment = opts.LFT
        ws.cell(r, 2).border = opts.BOX
        for col, f in ((3, m), (4, e), (5, f"=ROUND($C{r}-$D{r},6)")):
            x = ws.cell(r, col)
            x.value = f
            x.number_format = nf if col < 5 else (
                opts.CTL_C if nf in (opts.CT, opts.C1) else opts.CTL_M)
            x.alignment, x.font, x.border = opts.RGT, opts.BODY, opts.BOX
        r += 1
    opts.row(ws, r, 2, ["Checks failing", None, None, None], [None] * 4,
             bg=opts.MID, bold=True, top=True)
    ws.cell(r, 2).alignment = opts.LFT
    x = ws.cell(r, 5)
    x.value = f'=COUNTIF($E{HDR+1}:$E{r-1},"<>0")'
    x.number_format, x.alignment = opts.CTL_C, opts.RGT
    return len(checks)


# The retired raw-data tabs keep whatever marks the owner put on them. Everything in the
# live model uses one input colour.
RETIRED = {"Squads", "Added data", "Sheet2", "FY26 Budget (superseded)",
           "squad mapping (superseded)"}


def cream(wb):
    """One input colour across the live model. The file had bright yellow and cream both
    meaning 'typed input', side by side on the same tabs."""
    n = 0
    for ws in wb.worksheets:
        if ws.title in RETIRED:
            continue
        for row in ws.iter_rows():
            for c in row:
                fl = c.fill
                try:
                    rgb = str(fl.start_color.rgb or "").upper() \
                        if fl and fl.patternType else ""
                except Exception:
                    rgb = ""
                if rgb == "FFFFFF00":
                    c.fill = opts.fl("FFFFF2CC")
                    n += 1
    return n


def run(src, dst):
    wb = openpyxl.load_workbook(src)
    a = anchors(wb)
    a2 = json.load(open("anchors_final.json"))
    n = repoint(wb, a)
    build_exec(wb, a, a2)
    k = build_qa(wb, a, a2)
    nc = cream(wb)
    wb.save(dst)
    return [f"Lists archetype-roles lookup repointed on {n} portfolios",
            f"{nc} inputs recoloured to cream across the live model",
            "Exec Summary rebuilt on design against actual, with a portfolio drill-down",
            f"4.0 Data QA rebuilt: {k} checks, model / expected / difference",
            f"anchors: 3.1 ledger r{a['total']} grand r{a['grand']}, "
            f"3.2 overhead total r{a['ohtot32']}, 3.3 total r{a['g33']}"]


if __name__ == "__main__":
    import sys
    for x in run(sys.argv[1], sys.argv[2]):
        print("  ", x)
