"""3.1, 3.2 and 3.3 on layout 3D, reading the rebuilt 2.x tabs.

Design cost against actual cost, no budget anywhere. The design side covers the whole
organisation, so the comparison is like for like across all 525 roles:

    delivery squads in the portfolios   squad archetype from 0.3
    COEs and EGI                        the budget they draw down on 3.4
    overhead roles                      the allowance on Lists

Overhead is a row rather than a column so both sides foot to the same place. Each tab has
one job and states nothing the other two state.
"""
import json
import re

import openpyxl
from openpyxl.utils import get_column_letter as L

import final2x as f2
import opts

REVIEW = f2.REVIEW
REV = f2.REV
LAST = f2.LAST
S = f2.S

PF_ORDER = ["Ampol Retail", "Customer", "Enterprise Data", "TDD Group Functions", "P&C",
            "Finance", "Infrastructure", "Energy Solutions & B2B", "Commercial Fuels",
            "Z Retail"]
COE_ORDER = ["COE Cyber", "COE BP&T", "COE SA&D", "EGI"]
# A COE's design cost is the planned spend on its own 1.x tab, which is built from the
# real roles rather than from an archetype. It equals actual, so the variance is nil by
# construction - and that is the honest statement. The earlier version used 3.4's
# "budget to draw down" (12.00 against 27.77 actual), which asserted the COEs were
# 15.77m over a design that does not exist.
COE_DESIGN = {"COE BP&T": "='1.11 BP&T'!$F$6+'1.11 BP&T'!$F$7",
              "COE SA&D": "='1.12 SA&D'!$G$6+'1.12 SA&D'!$G$7",
              "COE Cyber": "='1.13 Cyber Roles'!$F$6+'1.13 Cyber Roles'!$F$7"}
# EGI has no COE tab. Its design sits in the strategic-programme rows on the portfolio
# design tabs, where the total cost is a yellow input the owner sets. Those inputs
# currently total 1.52m against 4.94m of real people, which is register item 61 and is
# the owner's number to set, not mine to invent.
EGI_DESIGN_CELLS = [("1.1 Ampol Retail", 66), ("1.2 Customer", 54),
                    ("1.4 TDD Group Functions", 30), ("1.6 Finance", 33)]


def coe_design_ref(pf):
    if pf in COE_DESIGN:
        return COE_DESIGN[pf]
    if pf == "EGI":
        return "=" + "+".join(f"N('{t}'!$H${r})" for t, r in EGI_DESIGN_CELLS)
    return "=0"


def order(anchors):
    by = {a["pf"]: t for t, a in anchors.items()}
    return ([(p, by[p], anchors[by[p]]) for p in PF_ORDER if p in by],
            [(p, by[p], anchors[by[p]]) for p in COE_ORDER if p in by])


H31 = ["Portfolio", "Design cost ($m)", "Actual cost ($m)", "Over/(under) design ($m)",
       "Cost after vacancy decisions ($m)", "Roles", "Filled", "Vacant"]
W31 = [38, 15, 14, 17, 19, 8, 8, 8]
F31 = [None, opts.M2, opts.M2, opts.M2, opts.M2, opts.CT, opts.CT, opts.CT]


def build_31(wb, anchors):
    ws = wb["3.1 Group Summary"]
    f2.wipe(ws)
    ws.column_dimensions["A"].width = 2
    ws.cell(2, 2).value = "TDD Summary - all portfolios"
    ws.cell(2, 2).font = opts.TITLE
    pfs, coes = order(anchors)

    r = 4
    tiles = r
    r += 2
    r += 1
    r = opts.bar(ws, r, 2, len(H31), "Design cost against actual cost, by portfolio")
    r = opts.head(ws, r, 2, H31, W31)

    def block(r, items, label, squad_only):
        st = r
        for pf, tab, a in items:
            t, d = a["total_row"], a["delivery_row"]
            ws.cell(r, 2).value = pf
            ws.cell(r, 2).font = opts.BODY
            ws.cell(r, 2).alignment = opts.LFT
            act = f"'{tab}'!${L(S['actual'])}${d if squad_only else t}"
            aft = f"'{tab}'!${L(S['after'])}${d if squad_only else t}"
            des = (f"='{tab}'!${L(S['acost'])}${d}" if squad_only
                   else coe_design_ref(pf))
            f2._m(ws, r, 3, des)
            f2._m(ws, r, 4, f"={act}")
            f2._m(ws, r, 5, f"=$D{r}-$C{r}")
            f2._m(ws, r, 6, f"={aft}")
            for i, k in enumerate(("roles", "filled", "vacant")):
                f2._m(ws, r, 7 + i,
                      f"='{tab}'!${L(S[k])}${d if squad_only else t}", opts.CT)
            r += 1
        opts.row(ws, r, 2, [label] + [None] * (len(H31) - 1), [None] * len(H31),
                 bg=opts.GREY, bold=True)
        ws.cell(r, 2).alignment = opts.LFT
        for c in range(3, 3 + len(H31) - 1):
            x = ws.cell(r, c)
            x.value = f"=SUM({L(c)}{st}:{L(c)}{r-1})"
            x.number_format, x.alignment = F31[c - 2], opts.RGT
        return r + 1, r

    r, s1 = block(r, pfs, "Delivery squads in the portfolios", True)
    r, s2 = block(r, coes, "COEs and EGI", False)

    # overhead as its own row, against the allowance on Lists
    ohsum = "+".join(f"N('{t}'!${L(S['actual'])}${a['overhead_row']})"
                     for _, t, a in pfs if a["overhead_row"])
    ohaft = "+".join(f"N('{t}'!${L(S['after'])}${a['overhead_row']})"
                     for _, t, a in pfs if a["overhead_row"])
    opts.row(ws, r, 2, ["Overhead roles"] +
             [None] * (len(H31) - 1), [None] * len(H31), bg=opts.GREY, bold=True)
    ws.cell(r, 2).alignment = opts.LFT
    f2._m(ws, r, 3, "=N(Lists!$AJ$8)")
    f2._m(ws, r, 4, "=" + ohsum)
    f2._m(ws, r, 5, f"=$D{r}-$C{r}")
    f2._m(ws, r, 6, "=" + ohaft)
    for i, k in enumerate(("roles", "filled", "vacant")):
        f2._m(ws, r, 7 + i,
              "=" + "+".join(f"N('{t}'!${L(S[k])}${a['overhead_row']})"
                             for _, t, a in pfs if a["overhead_row"]), opts.CT)
    s3 = r
    r += 1

    opts.row(ws, r, 2, ["Group total"] + [None] * (len(H31) - 1), [None] * len(H31),
             bg=opts.MID, bold=True, top=True)
    ws.cell(r, 2).alignment = opts.LFT
    for c in range(3, 3 + len(H31) - 1):
        x = ws.cell(r, c)
        x.value = f"={L(c)}{s1}+{L(c)}{s2}+{L(c)}{s3}"
        x.number_format, x.alignment = F31[c - 2], opts.RGT
    gt = r
    r += 2

    for lab, col, f, nf in (
            ("Control - roles against the ledger, must be 0", 7,
             f"=$G${gt}-COUNTA({REV}!$B$2:$B${LAST})", opts.CTL_C),
            ("Control - cost against the ledger ($m), must be 0", 4,
             f"=ROUND($D${gt}-SUM({REV}!$AA$2:$AA${LAST})/1000000,6)", opts.CTL_M)):
        ws.cell(r, 2).value = lab
        ws.cell(r, 2).font = opts.BODY
        f2._m(ws, r, col, f, nf)
        r += 1

    opts.strip(ws, tiles, 2, [
        ("Roles", f"=$G${gt}", opts.CT), ("Filled", f"=$H${gt}", opts.CT),
        ("Vacant", f"=$I${gt}", opts.CT),
        ("Design cost ($m)", f"=$C${gt}", opts.M2),
        ("Actual cost ($m)", f"=$D${gt}", opts.M2),
        ("Over/(under) design ($m)", f"=$E${gt}", opts.M2)], w=20)
    ws.freeze_panes = "C8"
    return {"total": gt, "first": 9,
            "delivery": s1, "coe": s2, "overhead": s3}


H32 = ["Cost", "Design cost ($m)", "Actual cost ($m)", "Over/(under) design ($m)",
       "Cost after vacancy decisions ($m)", "Roles"]
W32 = [46, 15, 14, 17, 19, 9]
H32B = ["Overhead line", "Roles", "Rate ($m)", "Times applied", "Allowance ($m)",
        "Actual cost ($m)", "Over/(under) allowance ($m)"]
W32B = [30, 9, 11, 13, 14, 14, 18]


def build_32(wb, anchors, a31):
    ws = wb["3.2 Total Cost"]
    f2.wipe(ws)
    ws.column_dimensions["A"].width = 2
    ws.cell(2, 2).value = "Total Cost - design against actual"
    ws.cell(2, 2).font = opts.TITLE
    # the three subtotal rows are not contiguous with the group total - the COE block
    # sits between them - so they come from build_31 rather than being counted back
    s1, s2, s3 = a31["delivery"], a31["coe"], a31["overhead"]
    r = opts.bar(ws, 4, 2, len(H32), "Design against actual")
    r = opts.head(ws, r, 2, H32, W32)
    st = r
    for lab, src in (("Delivery squads in the portfolios", s1),
                     ("COEs and EGI - design from their 1.x tabs", s2),
                     ("Overhead roles - against their allowance", s3)):
        ws.cell(r, 2).value = lab
        ws.cell(r, 2).font = opts.BODY
        ws.cell(r, 2).alignment = opts.LFT
        for i, c in enumerate((3, 4, 5, 6)):
            f2._m(ws, r, c, f"='3.1 Group Summary'!${L(c)}${src}")
        f2._m(ws, r, 7, f"='3.1 Group Summary'!$G${src}", opts.CT)
        r += 1
    opts.row(ws, r, 2, ["Cost of the organisation today"] + [None] * (len(H32) - 1),
             [None] * len(H32), bg=opts.MID, bold=True, top=True)
    ws.cell(r, 2).alignment = opts.LFT
    for c in range(3, 8):
        x = ws.cell(r, c)
        x.value = f"=SUM({L(c)}{st}:{L(c)}{r-1})"
        x.number_format = opts.CT if c == 7 else opts.M2
        x.alignment = opts.RGT
    r += 2

    r = opts.bar(ws, r, 2, len(H32B), "Overhead roles - line by line")
    r = opts.head(ws, r, 2, H32B, W32B)
    st2 = r
    l = wb["Lists"]
    for i in range(2, 8):
        ws.cell(r, 2).value = f"=Lists!$AF${i}"
        ws.cell(r, 2).font = opts.BODY
        ws.cell(r, 2).alignment = opts.LFT
        f2._m(ws, r, 3, f"=COUNTIFS({REV}!$AR$2:$AR${LAST},$B{r})", opts.CT)
        f2._m(ws, r, 4, f"=Lists!$AG${i}")
        f2._m(ws, r, 5, f"=Lists!$AH${i}", opts.CT)
        f2._m(ws, r, 6, f"=Lists!$AJ${i}")
        f2._m(ws, r, 7,
              f'=IF($B{r}="Leadership - 8 GMs",N(Lists!$AG$12),'
              f"SUMIFS({REV}!$AA$2:$AA${LAST},{REV}!$AR$2:$AR${LAST},$B{r})/1000000)")
        f2._m(ws, r, 8, f"=$G{r}-$F{r}")
        r += 1
    opts.row(ws, r, 2, ["Overhead total"] + [None] * (len(H32B) - 1),
             [None] * len(H32B), bg=opts.MID, bold=True, top=True)
    ws.cell(r, 2).alignment = opts.LFT
    for c, nf in ((3, opts.CT), (6, opts.M2), (7, opts.M2), (8, opts.M2)):
        x = ws.cell(r, c)
        x.value = f"=SUM({L(c)}{st2}:{L(c)}{r-1})"
        x.number_format, x.alignment = nf, opts.RGT
    r += 2
    ws.cell(r, 2).value = ("The 8 GMs are the only overhead line with no role in the "
                           "ledger. Their cost is the input on Lists and sits above "
                           "the 525.")
    ws.cell(r, 2).font = opts.BODY
    ws.freeze_panes = "C6"
    return {"total": st + 3}


H33 = ["Portfolio", "Squad", "Archetype Type", "Size", "Archetype roles", "Roles",
       "Filled", "Vacant", "Archetype cost ($m)", "Actual cost ($m)",
       "Variance to archetype ($m)", "Cost after vacancy decisions ($m)"]
W33 = [24, 32, 26, 7, 12, 8, 8, 8, 14, 13, 15, 19]
# one entry per column B..M, indexed F33[c - 2]
F33 = ([None] * 4 + [opts.C1] + [opts.CT] * 3 + [opts.M2] * 4)


def build_33(wb, anchors):
    ws = wb["3.3 FTE View"]
    f2.wipe(ws)
    ws.column_dimensions["A"].width = 2
    ws.cell(2, 2).value = "Squad Detail - roles and cost, squad by squad"
    ws.cell(2, 2).font = opts.TITLE
    pfs, coes = order(anchors)
    r = opts.bar(ws, 4, 2, len(H33), "Every squad on every working tab")
    r = opts.head(ws, r, 2, H33, W33)
    pf_rows = []
    for pf, tab, a in pfs + coes:
        st = r
        for gname in a["squads"] + a["overhead"]:
            s = a["srow"][gname]
            ws.cell(r, 2).value = pf
            ws.cell(r, 2).font = opts.BODY
            ws.cell(r, 2).alignment = opts.LFT
            for i, k in enumerate(("squad", "type", "size", "aroles", "roles",
                                   "filled", "vacant", "acost", "actual", "var",
                                   "after")):
                c = 3 + i
                x = ws.cell(r, c)
                x.value = f"='{tab}'!${L(S[k])}${s}"
                x.font = opts.BODY
                if F33[c - 2]:
                    x.number_format, x.alignment = F33[c - 2], opts.RGT
                else:
                    x.alignment = opts.LFT
            r += 1
        opts.row(ws, r, 2, [f"{pf} total"] + [None] * (len(H33) - 1),
                 [None] * len(H33), bg=opts.GREY, bold=True)
        ws.cell(r, 2).alignment = opts.LFT
        for c in range(6, 2 + len(H33)):
            x = ws.cell(r, c)
            x.value = f"=SUM({L(c)}{st}:{L(c)}{r-1})"
            x.number_format, x.alignment = F33[c - 2], opts.RGT
        pf_rows.append(r)
        r += 1
    opts.row(ws, r, 2, ["Group total"] + [None] * (len(H33) - 1), [None] * len(H33),
             bg=opts.MID, bold=True, top=True)
    ws.cell(r, 2).alignment = opts.LFT
    for c in range(6, 2 + len(H33)):
        x = ws.cell(r, c)
        x.value = "=" + "+".join(f"{L(c)}{p}" for p in pf_rows)
        x.number_format, x.alignment = F33[c - 2], opts.RGT
    gt = r
    r += 2
    for lab, col, f, nf in (
            ("Control - roles against the ledger, must be 0", 7,
             f"=$G${gt}-COUNTA({REV}!$B$2:$B${LAST})", opts.CTL_C),
            ("Control - cost against the ledger ($m), must be 0", 11,
             f"=ROUND($K${gt}-SUM({REV}!$AA$2:$AA${LAST})/1000000,6)", opts.CTL_M)):
        ws.cell(r, 2).value = lab
        ws.cell(r, 2).font = opts.BODY
        f2._m(ws, r, col, f, nf)
        r += 1
    ws.freeze_panes = "D6"
    return {"group_total": gt}


def run(src, dst):
    wb = openpyxl.load_workbook(src)
    anchors = json.load(open("anchors_final.json"))
    a31 = build_31(wb, anchors)
    a32 = build_32(wb, anchors, a31)
    a33 = build_33(wb, anchors)
    json.dump({"3.1": a31, "3.2": a32, "3.3": a33},
              open("anchors_final3.json", "w"), indent=1)
    wb.save(dst)
    return [f"3.1 Group Summary: group total row {a31['total']}",
            "3.2 Total Cost: three-line design against actual, plus the overhead lines",
            f"3.3 Squad Detail: every squad, group total row {a33['group_total']}"]


if __name__ == "__main__":
    import sys
    for x in run(sys.argv[1], sys.argv[2]):
        print("  ", x)
