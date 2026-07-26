"""3.1, 3.2 and 3.3, reading the rebuilt 2.x tabs.

3.1 is layout 3D, which is what the owner picked and what the last build did not deliver: a
cost bridge. It starts at the archetype cost of the squads an archetype prices and walks,
one named line at a time, to what TDD actually costs. Every directly funded programme is on
the page by name - AmPOS, CTRM, the six EGI programmes, the two Leadership groups - because
a single "Directly funded 10.44" line tells a reader nothing about what the 10.44 is.
Per-portfolio and per-squad detail lives on 3.3, which is where 3D puts it.

3.2 was "Total Cost" and restated 3.1's subtotals. It is now overhead and leadership, which
is the one thing no other tab in the workbook states.

Design cost against actual cost, no budget anywhere. The design side is built four ways,
because the organisation is funded four ways, and each block states which one it is:

    squads priced by an archetype    the archetype library on 0.3
    directly funded programmes       the amount funded on the design tab's platform block
    COEs and EGI                     the planned spend on their own 1.x tabs
    overhead roles                   the allowance on Lists

The old build put the archetype cost of the squads that have one against the actual cost of
every squad in one subtotal. The group read +11.488 over design; the squads the archetype
actually prices are 0.559 under it, and the difference was the directly funded programmes
being charged against a zero. They are their own block now, priced against what they are
funded, so nothing is compared to nothing and nothing is left out.
"""
import json

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
# A COE's design cost is the planned spend on its own 1.x tab, built from its real roles
# rather than from an archetype. Reading the 1.x tab rather than the actual keeps it an
# independent number: where the 1.x roles list is short of the ledger the variance shows it
# instead of hiding it behind a cell that points at itself.
COE_DESIGN = {"COE BP&T": "='1.11 BP&T'!$F$6+'1.11 BP&T'!$F$7",
              "COE SA&D": "='1.12 SA&D'!$G$6+'1.12 SA&D'!$G$7",
              "COE Cyber": "='1.13 Cyber Roles'!$F$6+'1.13 Cyber Roles'!$F$7"}

# Lists!AF2:AJ7 is the overhead allowance. Three of the six lines are drawn in the
# portfolios; the Business Partner and Domain Architect allowances are drawn inside the
# COEs, where those people sit, and the 8 GMs sit above the 525-role ledger entirely.
# 3.1's overhead row therefore compares the portfolio-drawn allowance to the portfolio
# overhead cost. Netting all six against the portfolio cost only was what made the same
# allowance read -1.271 on one row of 3.2 and +9.996 seven rows below it.
# Yes where the line's people sit in the portfolios and its allowance is therefore the one
# the portfolio overhead cost should be measured against. The Business Partner and Domain
# Architect allowances are drawn inside the COEs, where all thirteen of those people sit,
# and the 8 GMs sit above the 525-role ledger, so none of the three belongs in the figure
# 3.1 compares the portfolio overhead cost to.
DRAWN = ["Yes", "No", "No", "Yes", "Yes", "No"]


def patch_lists(wb):
    """Say where each overhead line's cost sits, and total the portfolio-drawn allowance."""
    l = wb["Lists"]
    col = next(c for c in range(39, 70)
               if all(l.cell(r, c).value is None for r in range(1, 40)))
    cl = L(col)
    x = l.cell(1, col)
    x.value, x.font, x.fill, x.alignment = ("Allowance drawn in the portfolios",
                                            opts.HDRF, opts.fl(opts.NAVY), opts.CEN)
    l.column_dimensions[cl].width = 28
    for i, w in enumerate(DRAWN):
        c = l.cell(2 + i, col)
        c.value, c.font, c.border, c.alignment = w, opts.BODY, opts.BOX, opts.LFT
    a = l.cell(9, 32)
    a.value, a.font = "Allowance drawn in the portfolios", opts.BOLD
    b = l.cell(9, 36)
    b.value = f'=SUMIF(${cl}$2:${cl}$7,"Yes",$AJ$2:$AJ$7)'
    b.font, b.number_format, b.alignment = opts.BOLD, opts.M2, opts.RGT
    return f"Lists: overhead lines tagged in {cl}, portfolio-drawn allowance at AJ9", cl


H31 = ["Line", "Portfolio", "Archetype cost ($m)", "Actual cost ($m)",
       "Variance to archetype ($m)", "Cost after decisions ($m)", "Total roles",
       "Filled", "Vacant", "Total roles after decisions"]
W31 = [34, 24, 15, 14, 17, 17, 9, 8, 8, 14]
NUM = {4: opts.M2, 5: opts.M2, 6: opts.M2, 7: opts.M2,
       8: opts.CT, 9: opts.CT, 10: opts.CT, 11: opts.CT}
FIRST, LASTC = 4, 11                                     # D .. K


def order(anchors):
    by = {a["pf"]: t for t, a in anchors.items()}
    return ([(p, by[p], anchors[by[p]]) for p in PF_ORDER if p in by],
            [(p, by[p], anchors[by[p]]) for p in COE_ORDER if p in by])


def build_31(wb, anchors):
    """The cost bridge. One line per step, every step named.

    Option 3D, which is what the owner picked: start at the archetype cost of the squads
    the archetype prices and walk to what the organisation actually costs, naming each
    step. Every directly funded programme is on its own line by name - AmPOS, CTRM, the
    EGI programmes - because "Directly funded, 10.44" tells a reader nothing about what
    the 10.44 is. Per-portfolio and per-squad detail is on 3.3.
    """
    ws = wb["3.1 Group Summary"]
    f2.wipe(ws)
    ws.column_dimensions["A"].width = 2
    ws.cell(2, 2).value = "TDD cost bridge - archetype cost to actual cost"
    ws.cell(2, 2).font = opts.TITLE
    pfs, coes = order(anchors)

    r = opts.bar(ws, 4, 2, len(H31), "How the organisation gets from its archetype "
                                     "cost to what it actually costs")
    r = opts.head(ws, r, 2, H31, W31)

    def line(rw, name, pf, tab, src, design):
        ws.cell(rw, 2).value = name
        ws.cell(rw, 3).value = pf
        for c in (2, 3):
            ws.cell(rw, c).font = opts.BODY
            ws.cell(rw, c).alignment = opts.LFT
        f2._m(ws, rw, 4, design)
        f2._m(ws, rw, 5, f"='{tab}'!${L(S['actual'])}${src}")
        f2._m(ws, rw, 6, f'=IFERROR(ROUND($E{rw}-$D{rw},6),"-")')
        f2._m(ws, rw, 7, f"='{tab}'!${L(S['after'])}${src}")
        for i, k in enumerate(("roles", "filled", "vacant", "rafter")):
            f2._m(ws, rw, 8 + i, f"='{tab}'!${L(S[k])}${src}", opts.CT)

    def sub(rw, text, r0, r1, blank=()):
        opts.row(ws, rw, 2, [text] + [None] * (len(H31) - 1), [None] * len(H31),
                 bg=opts.GREY, bold=True)
        ws.cell(rw, 2).alignment = opts.LFT
        for c in range(FIRST, LASTC + 1):
            x = ws.cell(rw, c)
            # the variance on a total is actual less archetype, never the sum of the row
            # variances: a row with no archetype carries "-" in that column and drops out
            # of a SUM, so the directly funded subtotal read 0.13 against a real 1.44 and
            # the group read 5.07 against 6.38
            x.value = ('="-"' if c in blank else
                       f"=ROUND($E{rw}-$D{rw},6)" if c == 6
                       else f"=SUM({L(c)}{r0}:{L(c)}{r1})")
            x.number_format, x.alignment = NUM[c], opts.RGT

    def label(rw, text):
        opts.row(ws, rw, 2, [text] + [None] * (len(H31) - 1), [None] * len(H31),
                 bg=opts.PALE, bold=True)
        ws.cell(rw, 2).alignment = opts.LFT

    # ---- step 1: the squads an archetype prices, one line per portfolio ----
    label(r, "Squads priced by an archetype - detail on 3.3")
    r += 1
    st = r
    for pf, tab, a in pfs:
        line(r, pf, pf, tab, a["delivery_row"],
             f"='{tab}'!${L(S['acost'])}${a['delivery_row']}")
        r += 1
    sub(r, "Squads priced by an archetype", st, r - 1)
    s1 = r
    r += 1

    # ---- step 2: every directly funded programme, by name ----
    label(r, "Directly funded programmes and platforms - no archetype prices them, so "
             "the comparison is the amount funded on the 1.x tab")
    r += 1
    st = r
    for pf, tab, a in pfs:
        for g in a["direct"]:
            line(r, g, pf, tab, a["srow"][g],
                 f"='{tab}'!${L(S['acost'])}${a['srow'][g]}")
            r += 1
    sub(r, "Directly funded programmes and platforms", st, r - 1)
    s2 = r
    r += 1

    # ---- step 4: the COEs and EGI ----
    label(r, "COEs and EGI - priced off the planned spend on their own 1.x tabs")
    r += 1
    st = r
    for pf, tab, a in coes:
        line(r, pf, pf, tab, a["total_row"],
             COE_DESIGN.get(pf, f"='{tab}'!${L(S['actual'])}${a['total_row']}"))
        r += 1
    sub(r, "COEs and EGI", st, r - 1)
    s3 = r
    r += 1

    # ---- step 5: overhead in the portfolios, against the allowance ----
    # One row. The allowance is built per portfolio and per platform at group level, so
    # there is no per-portfolio allowance to set beside a per-portfolio cost - ten rows of
    # dashes under a subtotal carrying 6.325 is a subtotal that is not the sum of its rows.
    # The detail by portfolio is on 3.3 and on each working tab, line by line on 3.2.
    oh = [(p, t, a) for p, t, a in pfs if a["overhead_row"]]
    opts.row(ws, r, 2, ["Overhead roles in the portfolios - the allowance is on 3.2"] +
             [None] * (len(H31) - 1), [None] * len(H31), bg=opts.GREY, bold=True)
    ws.cell(r, 2).alignment = opts.LFT
    f2._m(ws, r, 4, "=N(Lists!$AJ$9)")
    for c, k in ((5, "actual"), (7, "after")):
        f2._m(ws, r, c, "=" + "+".join(
            f"N('{t}'!${L(S[k])}${a['overhead_row']})" for _, t, a in oh))
    f2._m(ws, r, 6, f"=ROUND($E{r}-$D{r},6)")
    for i, k in enumerate(("roles", "filled", "vacant", "rafter")):
        f2._m(ws, r, 8 + i, "=" + "+".join(
            f"N('{t}'!${L(S[k])}${a['overhead_row']})" for _, t, a in oh), opts.CT)
    for c in range(FIRST, LASTC + 1):
        ws.cell(r, c).font = opts.BOLD
    s4 = r
    r += 1
    opts.row(ws, r, 2, ["Everything with a figure to compare"] + [None] * (len(H31) - 1),
             [None] * len(H31), bg=opts.MID, bold=True, top=True)
    ws.cell(r, 2).alignment = opts.LFT
    for c in range(FIRST, LASTC + 1):
        x = ws.cell(r, c)
        x.value = (f"=ROUND($E{r}-$D{r},6)" if c == 6
                   else "=" + "+".join(f"N({L(c)}{p})" for p in (s1, s2, s3, s4)))
        x.number_format, x.alignment = NUM[c], opts.RGT
    cmp_row = r
    r += 1

    # ---- the group with nothing to compare, after the comparable steps ----
    # ---- step 5: groups with nothing to compare against ----
    nofig = [(p, t, a) for p, t, a in pfs if a["nofig"]]
    # the label and the subtotal must not read identically, or anything looking the row
    # up by name finds the label - which carries no figures
    label(r, "Groups with no archetype and no funded figure - nothing to compare to")
    r += 1
    st = r
    for pf, tab, a in nofig:
        for g in a["nofig"]:
            line(r, g, pf, tab, a["srow"][g], '="-"')
            r += 1
    sub(r, "Groups with no archetype and no funded figure", st, r - 1, blank=(4, 6))
    s2b = r
    r += 1

    # ---- the ledger total ----
    # Its variance is a dash, not a number. The archetype column prices four of the five
    # steps above it and the actual column covers all five, so no single figure on this row
    # can be true. The comparable subtotal two rows up is the one that can.
    opts.row(ws, r, 2, ["Cost of the 525 roles in the ledger"] + [None] * (len(H31) - 1),
             [None] * len(H31), bg=opts.MID, bold=True, top=True)
    ws.cell(r, 2).alignment = opts.LFT
    for c in range(FIRST, LASTC + 1):
        x = ws.cell(r, c)
        x.value = ('="-"' if c == 6
                   else "=" + "+".join(f"N({L(c)}{p})" for p in (s1, s2, s2b, s3, s4)))
        x.number_format, x.alignment = NUM[c], opts.RGT
    gt = r
    r += 1

    # ---- the GM layer, which has no role in the ledger to price ----
    opts.row(ws, r, 2, ["Leadership - the 8 GMs, outside the 525-role ledger"] +
             [None] * (len(H31) - 1), [None] * len(H31), bg=opts.PALE)
    ws.cell(r, 2).alignment = opts.LFT
    f2._m(ws, r, 4, "=N(Lists!$AJ$7)")
    f2._m(ws, r, 5, "=N(Lists!$AG$12)")
    f2._m(ws, r, 6, f"=ROUND($E{r}-$D{r},6)")
    f2._m(ws, r, 7, "=N(Lists!$AG$12)")
    for c in (8, 9, 11):
        f2._m(ws, r, c, "=N(Lists!$AG$11)", opts.CT)
    f2._m(ws, r, 10, "=0", opts.CT)
    gm = r
    r += 1
    opts.row(ws, r, 2, ["Total cost of TDD including the GM layer"] +
             [None] * (len(H31) - 1), [None] * len(H31), bg=opts.MID, bold=True, top=True)
    ws.cell(r, 2).alignment = opts.LFT
    for c in range(FIRST, LASTC + 1):
        x = ws.cell(r, c)
        x.value = ('="-"' if c == 6 else f"={L(c)}{gt}+{L(c)}{gm}")
        x.number_format, x.alignment = NUM[c], opts.RGT
    grand = r
    r += 2

    for lab, col, f, nf in (
            ("Control - roles against the ledger, must be 0", 8,
             f"=$H${gt}-COUNTA({REV}!$B$2:$B${LAST})", opts.CTL_C),
            ("Control - cost against the ledger ($m), must be 0", 5,
             f"=ROUND($E${gt}-SUM({REV}!$AA$2:$AA${LAST})/1000000,6)", opts.CTL_M)):
        ws.cell(r, 2).value = lab
        ws.cell(r, 2).font = opts.BODY
        f2._m(ws, r, col, f, nf)
        r += 1
    return {"total": gt, "grand": grand, "gm": gm, "arch": s1, "direct": s2,
            "nofig": s2b, "coe": s3, "overhead": s4, "comparable": cmp_row}


# 3.2 was "Total Cost" and restated 3.1's four subtotals in a second table. It gave a
# reader nothing 3.1 did not already give them, and the owner deleted it. What it did carry
# that nothing else does is the overhead allowance, line by line, and the GM layer - so
# that is what the tab is now, and only that.
H32 = ["Overhead line", "Basis", "Rate ($m)", "Times applied", "Allowance ($m)",
       "Roles in the portfolios", "Cost in the portfolios ($m)",
       "Portfolio cost less allowance ($m)", "Roles inside the COEs",
       "Cost inside the COEs ($m)", "Allowance drawn in the portfolios"]
W32 = [26, 13, 11, 13, 14, 13, 15, 19, 12, 15, 17]


def build_32(wb, anchors, a31, wcol):
    """Overhead and leadership. The one thing no other tab states.

    Every figure is computed. A role carrying an overhead title inside a COE has AT set to
    its COE squad rather than to the overhead line, so AR = AT selects the roles that sit
    in a portfolio. That split is the whole point of the tab: counting every role with an
    overhead title against the portfolio allowance produced 62 roles and $22.9m against a
    real 43 and $11.7m, and the same allowance read 1.271 under on one row and 9.996 over
    seven rows below it.
    """
    ws = wb["3.2 Total Cost"]
    f2.wipe(ws)
    ws.column_dimensions["A"].width = 2
    ws.cell(2, 2).value = "Overhead and leadership - the allowance against what it costs"
    ws.cell(2, 2).font = opts.TITLE
    r = opts.bar(ws, 4, 2, len(H32), "Overhead roles, line by line")
    r = opts.head(ws, r, 2, H32, W32)
    st = r
    for i in range(2, 8):
        ws.cell(r, 2).value = f"=Lists!$AF${i}"
        ws.cell(r, 3).value = f"=Lists!$AI${i}"
        for c in (2, 3):
            ws.cell(r, c).font = opts.BODY
            ws.cell(r, c).alignment = opts.LFT
        f2._m(ws, r, 4, f"=Lists!$AG${i}", opts.M3)
        f2._m(ws, r, 5, f"=Lists!$AH${i}", opts.CT)
        f2._m(ws, r, 6, f"=Lists!$AJ${i}")
        both = f"{REV}!$AR$2:$AR${LAST},$B{r},{REV}!$B$2:$B${LAST},\"<>\""
        pf_only = f"{both},{REV}!$AT$2:$AT${LAST},$B{r}"
        gm = '=IF($B{r}="Leadership - 8 GMs",{v},{e})'
        # the 8 GMs carry no role in the ledger, so their count and cost are the input
        f2._m(ws, r, 7, gm.format(r=r, v="N(Lists!$AG$11)",
                                  e=f"COUNTIFS({pf_only})"), opts.CT)
        f2._m(ws, r, 8, gm.format(
            r=r, v="N(Lists!$AG$12)",
            e=f"SUMIFS({REV}!$AA$2:$AA${LAST},{pf_only})/1000000"))
        f2._m(ws, r, 9, f"=ROUND($H{r}-$F{r},6)")
        f2._m(ws, r, 10, gm.format(r=r, v="0", e=f"COUNTIFS({both})-$G{r}"), opts.CT)
        f2._m(ws, r, 11, gm.format(
            r=r, v="0",
            e=f"SUMIFS({REV}!$AA$2:$AA${LAST},{both})/1000000-$H{r}"))
        ws.cell(r, 12).value = f"=Lists!${wcol}${i}"
        ws.cell(r, 12).font = opts.BODY
        ws.cell(r, 12).alignment = opts.CEN
        r += 1
    opts.row(ws, r, 2, ["Every overhead line, including the GM layer"] +
             [None] * (len(H32) - 1),
             [None] * len(H32), bg=opts.MID, bold=True, top=True)
    ws.cell(r, 2).alignment = opts.LFT
    for c, nf in ((6, opts.M2), (7, opts.CT), (8, opts.M2), (9, opts.M2),
                  (10, opts.CT), (11, opts.M2)):
        x = ws.cell(r, c)
        x.value = f"=SUM({L(c)}{st}:{L(c)}{r-1})"
        x.number_format, x.alignment = nf, opts.RGT
    r += 1
    # the one line that ties this tab to the overhead step on the bridge
    opts.row(ws, r, 2, ["Of which sits in the 525-role ledger - the 3.1 overhead line"] +
             [None] * (len(H32) - 1), [None] * len(H32), bg=opts.GREY, bold=True)
    ws.cell(r, 2).alignment = opts.LFT
    f2._m(ws, r, 6, "=N(Lists!$AJ$9)")
    f2._m(ws, r, 7, f"='3.1 Group Summary'!$H${a31['overhead']}", opts.CT)
    f2._m(ws, r, 8, f"='3.1 Group Summary'!$E${a31['overhead']}")
    f2._m(ws, r, 9, f"=ROUND($H{r}-$F{r},6)")
    for c in (6, 7, 8, 9):
        ws.cell(r, c).font = opts.BOLD
    r += 2

    # ---- what the allowance is built from, in full, on the page ----
    # The rate on the block above is an allocated share: half a Head of Technology per
    # portfolio, 30% of a manager per platform. The actual is whole heads. That is the
    # whole reason 6.33 of allowance sits against 11.65 of cost, and a reader cannot judge
    # it unless the full role cost and the allocation are both stated, so they are.
    r = opts.bar(ws, r, 2, 5, "What the allocated rate is built from")
    # rate, times applied and allowance are already on the block above; repeating them here
    # was half a table restating the other half
    r = opts.head(ws, r, 2,
                  ["Input", "Where it is set on 0.2 Data Config", "Full role cost ($m)",
                   "Allocation %", "Allocated rate ($m)"],
                  [26, 42, 14, 12, 16])
    st3 = r
    for lab, where, full, pct, i in (
            ("Head of Technology", "Portfolio Overhead, Head of Tech", "$J$6", "$K$6", 2),
            ("Business Partner", "Portfolio Overhead, Business Partner", "$J$7", "$K$7", 3),
            ("Domain Architect", "Portfolio Overhead, Domain Architect", "$J$8", "$K$8", 4),
            ("Delivery Manager", "Platform Overhead, Delivery Manager", "$J$14", "$K$14",
             5),
            ("Technology Manager", "Platform Overhead, Tech Manager", "$J$15", "$K$15", 6),
            ("Leadership - 8 GMs", "Portfolio Overhead, Leadership Overhead", "$J$9",
             "$K$9", 7)):
        ws.cell(r, 2).value = lab
        ws.cell(r, 3).value = where
        for c in (2, 3):
            ws.cell(r, c).font = opts.BODY
            ws.cell(r, c).alignment = opts.LFT
        f2._m(ws, r, 4, f"='0.2 Data Config'!{full}", opts.M3)
        f2._m(ws, r, 5, f"='0.2 Data Config'!{pct}", opts.PCT)
        f2._m(ws, r, 6, f"=Lists!$AG${i}", opts.M3)
        r += 1
    r += 1
    r += 1
    ws.cell(r, 2).value = ("The allocated rate is the full role cost times the allocation. "
                           "The cost it is measured against is whole roles.")
    ws.cell(r, 2).font = opts.BODY
    ws.cell(r, 2).alignment = opts.LFT
    return {"total": st}


H33 = ["Portfolio", "How it is funded", "Squad", "Archetype Type", "Squad Size",
       "Archetype roles", "Total roles", "Filled", "Vacant",
       "Total roles after decisions", "Archetype cost ($m)", "Actual cost ($m)",
       "Variance to archetype ($m)", "Cost after decisions ($m)"]
W33 = [22, 17, 30, 24, 7, 11, 7, 7, 8, 13, 13, 13, 15, 19]
# one entry per column B..N, indexed F33[c - 2]
F33 = [None, None, None, None, None, opts.C1, opts.CT, opts.CT, opts.CT, opts.CT,
       opts.M2, opts.M2, opts.M2, opts.M2]
# the totals cover the columns that can be added across every kind of row: roles, filled,
# vacant, actual cost and cost after decisions. Archetype roles, design cost and variance
# are not among them - only some rows carry an archetype, so totalling those three beside a
# total actual for every row would read as numbers that do not compute.
T33 = (8, 9, 10, 11, 13, 15)


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
        for kind, names in (("Archetype", a["squads"]),
                            ("Directly funded", a["direct"]),
                            ("No figure to compare", a["nofig"]),
                            ("Overhead", a["overhead"])):
            for gname in names:
                s = a["srow"][gname]
                ws.cell(r, 2).value = pf
                ws.cell(r, 3).value = kind
                for c in (2, 3):
                    ws.cell(r, c).font = opts.BODY
                    ws.cell(r, c).alignment = opts.LFT
                for i, k in enumerate(("squad", "type", "size", "aroles", "roles",
                                       "filled", "vacant", "rafter", "acost", "actual",
                                       "var", "after")):
                    c = 4 + i
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
            x.value = f"=SUM({L(c)}{st}:{L(c)}{r-1})" if c in T33 else '="-"'
            x.number_format, x.alignment = F33[c - 2] or opts.M2, opts.RGT
        pf_rows.append(r)
        r += 1
    opts.row(ws, r, 2, ["Group total"] + [None] * (len(H33) - 1), [None] * len(H33),
             bg=opts.MID, bold=True, top=True)
    ws.cell(r, 2).alignment = opts.LFT
    for c in range(6, 2 + len(H33)):
        x = ws.cell(r, c)
        x.value = ("=" + "+".join(f"{L(c)}{p}" for p in pf_rows) if c in T33 else '="-"')
        x.number_format, x.alignment = F33[c - 2] or opts.M2, opts.RGT
    gt = r
    r += 2
    for lab, col, f, nf in (
            ("Control - roles against the ledger, must be 0", 8,
             f"=$H${gt}-COUNTA({REV}!$B$2:$B${LAST})", opts.CTL_C),
            ("Control - cost against the ledger ($m), must be 0", 13,
             f"=ROUND($M${gt}-SUM({REV}!$AA$2:$AA${LAST})/1000000,6)", opts.CTL_M)):
        ws.cell(r, 2).value = lab
        ws.cell(r, 2).font = opts.BODY
        f2._m(ws, r, col, f, nf)
        r += 1
    return {"group_total": gt}


def run(src, dst):
    wb = openpyxl.load_workbook(src)
    anchors = json.load(open("anchors_final.json"))
    msg, wcol = patch_lists(wb)
    out = [msg]
    a31 = build_31(wb, anchors)
    a32 = build_32(wb, anchors, a31, wcol)
    a33 = build_33(wb, anchors)
    json.dump({"3.1": a31, "3.2": a32, "3.3": a33},
              open("anchors_final3.json", "w"), indent=1)
    wb.save(dst)
    return out + [
        f"3.1: cost bridge, every directly funded programme named, ledger row "
        f"{a31['total']}, grand total row {a31['grand']}",
        "3.2: overhead and leadership, line by line, plus what the allowance is built from",
        f"3.3: every squad by portfolio, total row {a33['group_total']}"]


if __name__ == "__main__":
    import sys
    for x in run(sys.argv[1], sys.argv[2]):
        print("  ", x)
