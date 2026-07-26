"""3.1, 3.2 and 3.3 on layout 3D, reading the rebuilt 2.x tabs.

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


H31 = ["Portfolio", "Archetype cost ($m)", "Actual cost ($m)", "Over/(under) archetype ($m)",
       "Cost after vacancy decisions ($m)", "Roles", "Filled", "Vacant",
       "Roles after decisions"]
W31 = [38, 15, 14, 17, 19, 8, 8, 8, 13]
NUM = {3: opts.M2, 4: opts.M2, 5: opts.M2, 6: opts.M2,
       7: opts.CT, 8: opts.CT, 9: opts.CT, 10: opts.CT}
LASTC = 10                                               # column J


def order(anchors):
    by = {a["pf"]: t for t, a in anchors.items()}
    return ([(p, by[p], anchors[by[p]]) for p in PF_ORDER if p in by],
            [(p, by[p], anchors[by[p]]) for p in COE_ORDER if p in by])


def build_31(wb, anchors):
    ws = wb["3.1 Group Summary"]
    f2.wipe(ws)
    ws.column_dimensions["A"].width = 2
    ws.cell(2, 2).value = "TDD Summary - all portfolios"
    ws.cell(2, 2).font = opts.TITLE
    pfs, coes = order(anchors)
    tiles = 4
    ws.cell(tiles, 2).value = "Group position"
    ws.cell(tiles, 2).font = opts.BOLD
    ws.cell(tiles, 2).alignment = opts.LFT

    r = opts.bar(ws, 7, 2, len(H31), "Archetype cost against actual cost")
    r = opts.head(ws, r, 2, H31, W31)

    def label(rw, text):
        opts.row(ws, rw, 2, [text] + [None] * (len(H31) - 1), [None] * len(H31),
                 bg=opts.PALE, bold=True)
        ws.cell(rw, 2).alignment = opts.LFT
        return rw + 1

    def sub(rw, text, r0, r1, blank=()):
        opts.row(ws, rw, 2, [text] + [None] * (len(H31) - 1), [None] * len(H31),
                 bg=opts.GREY, bold=True)
        ws.cell(rw, 2).alignment = opts.LFT
        for c in range(3, LASTC + 1):
            x = ws.cell(rw, c)
            x.value = '="-"' if c in blank else f"=SUM({L(c)}{r0}:{L(c)}{r1})"
            x.number_format, x.alignment = NUM[c], opts.RGT
        return rw + 1

    def rows(rw, items, row_of, design):
        for pf, tab, a in items:
            src = row_of(a)
            ws.cell(rw, 2).value = pf
            ws.cell(rw, 2).font = opts.BODY
            ws.cell(rw, 2).alignment = opts.LFT
            f2._m(ws, rw, 3, design(pf, tab, src))
            f2._m(ws, rw, 4, f"='{tab}'!${L(S['actual'])}${src}")
            f2._m(ws, rw, 5, f"=IFERROR($D{rw}-$C{rw},\"-\")")
            f2._m(ws, rw, 6, f"='{tab}'!${L(S['after'])}${src}")
            for i, k in enumerate(("roles", "filled", "vacant", "rafter")):
                f2._m(ws, rw, 7 + i, f"='{tab}'!${L(S[k])}${src}", opts.CT)
            rw += 1
        return rw

    own = lambda col: (lambda pf, tab, src: f"='{tab}'!${L(S[col])}${src}")

    r = label(r, "Squads priced by an archetype")
    r = rows(r, pfs, lambda a: a["delivery_row"], own("acost"))
    r = sub(r, "Priced by an archetype", r - len(pfs), r - 1)
    s1 = r - 1

    direct = [(p, t, a) for p, t, a in pfs if a["direct"]]
    r = label(r, "Directly funded programmes and platforms - funded on the 1.x tab")
    r = rows(r, direct, lambda a: a["direct_row"], own("acost"))
    r = sub(r, "Directly funded", r - len(direct), r - 1)
    s2 = r - 1

    def coe(pf, tab, src):
        return COE_DESIGN.get(pf, f"='{tab}'!${L(S['actual'])}${src}")

    r = label(r, "COEs and EGI - planned spend on their own 1.x tabs")
    r = rows(r, coes, lambda a: a["total_row"], coe)
    r = sub(r, "COEs and EGI", r - len(coes), r - 1)
    s3 = r - 1

    # One row, not ten. The allowance is built per portfolio and per platform at group
    # level, so there is no per-portfolio allowance to put beside a per-portfolio cost; a
    # ten-row block would have carried a subtotal the rows above it did not add to. The
    # overhead detail by portfolio is on 3.3 and on each working tab.
    opts.row(ws, r, 2, ["Overhead roles in the portfolios"] +
             [None] * (len(H31) - 1), [None] * len(H31), bg=opts.GREY, bold=True)
    ws.cell(r, 2).alignment = opts.LFT
    oh = [(p, t, a) for p, t, a in pfs if a["overhead_row"]]
    f2._m(ws, r, 3, "=N(Lists!$AJ$9)")
    for c, k in ((4, "actual"), (6, "after")):
        f2._m(ws, r, c, "=" + "+".join(
            f"N('{t}'!${L(S[k])}${a['overhead_row']})" for _, t, a in oh))
    f2._m(ws, r, 5, f"=$D{r}-$C{r}")
    for i, k in enumerate(("roles", "filled", "vacant", "rafter")):
        f2._m(ws, r, 7 + i, "=" + "+".join(
            f"N('{t}'!${L(S[k])}${a['overhead_row']})" for _, t, a in oh), opts.CT)
    for c in range(3, LASTC + 1):
        ws.cell(r, c).font = opts.BOLD
    s4 = r
    r += 1

    r += 1
    opts.row(ws, r, 2, ["Cost of the organisation today"] + [None] * (len(H31) - 1),
             [None] * len(H31), bg=opts.MID, bold=True, top=True)
    ws.cell(r, 2).alignment = opts.LFT
    for c in range(3, LASTC + 1):
        x = ws.cell(r, c)
        x.value = "=" + "+".join(f"N({L(c)}{p})" for p in (s1, s2, s3, s4))
        x.number_format, x.alignment = NUM[c], opts.RGT
    gt = r
    r += 1
    # The 8 GMs cost $5.1m and carry no role in REVIEW, so they sit outside the 525-role
    # ledger. Leaving them off the headline understated TDD by 5.1 on the one page a GM
    # reads first. The ledger total stays exactly where it was, because every control and
    # every check ties to it; the GM layer is stated beneath it and added in a grand total.
    opts.row(ws, r, 2, ["Leadership - the 8 GMs, outside the 525-role ledger"] +
             [None] * (len(H31) - 1), [None] * len(H31), bg=opts.PALE)
    ws.cell(r, 2).alignment = opts.LFT
    f2._m(ws, r, 3, "=N(Lists!$AJ$7)")
    f2._m(ws, r, 4, "=N(Lists!$AG$12)")
    f2._m(ws, r, 5, f"=$D{r}-$C{r}")
    f2._m(ws, r, 6, "=N(Lists!$AG$12)")
    for c in (7, 8, 10):
        f2._m(ws, r, c, "=N(Lists!$AG$11)", opts.CT)
    f2._m(ws, r, 9, "=0", opts.CT)
    gm = r
    r += 1
    opts.row(ws, r, 2, ["Total including the GM layer"] + [None] * (len(H31) - 1),
             [None] * len(H31), bg=opts.MID, bold=True, top=True)
    ws.cell(r, 2).alignment = opts.LFT
    for c in range(3, LASTC + 1):
        x = ws.cell(r, c)
        x.value = f"={L(c)}{gt}+{L(c)}{gm}"
        x.number_format, x.alignment = NUM[c], opts.RGT
    grand = r
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

    # the tiles share the table's columns, so each one sits over the column it summarises
    TILES31 = [(h, f"=${L(3 + i)}${grand}", NUM[3 + i])
               for i, h in enumerate(H31[1:])]
    for i, (lab, v, f) in enumerate(TILES31):
        c = 3 + i
        h = ws.cell(tiles, c)
        h.value, h.font, h.fill, h.alignment, h.border = (lab, opts.HDRF,
                                                          opts.fl(opts.NAVY), opts.CEN,
                                                          opts.BOX)
        x = ws.cell(tiles + 1, c)
        x.value, x.font, x.number_format = v, opts.BIG, f
        x.alignment, x.border, x.fill = opts.CEN, opts.BOX, opts.fl(opts.GREY)
    ws.row_dimensions[tiles].height = max(32, 14 * max(
        opts.wrap_lines(lab, W31[i + 1])
        for i, (lab, _, _) in enumerate(TILES31)) + 6)
    ws.row_dimensions[tiles + 1].height = 30
    ws.freeze_panes = f"C{9}"
    return {"total": gt, "grand": grand, "gm": gm,
            "arch": s1, "direct": s2, "coe": s3, "overhead": s4}


H32 = ["How the cost is funded", "Archetype cost ($m)", "Actual cost ($m)",
       "Over/(under) archetype ($m)", "Cost after vacancy decisions ($m)", "Roles"]
W32 = [52, 15, 14, 17, 19, 9]
H32B = ["Overhead line", "Rate ($m)", "Times applied", "Allowance ($m)",
        "Roles in the portfolios", "Cost in the portfolios ($m)",
        "Not covered by the allowance ($m)", "Roles inside the COEs",
        "Cost inside the COEs ($m)", "Allowance drawn in the portfolios"]
# Both tables on 3.2 sit in the same columns, so they cannot each set their own widths -
# the second call used to overwrite the first, dropping column B from 58 to 26 and clipping
# every label in the block above. One profile, wide enough for both.
W32 = [42, 15, 14, 17, 19, 15, 19, 12, 15, 17]
BLOCKS = [("Squads priced by an archetype", "arch"),
          ("Directly funded programmes and platforms", "direct"),
          ("COEs and EGI", "coe"),
          ("Overhead roles in the portfolios", "overhead")]
# the label rows below a table carry no figures in column B's neighbours, so they can run
# long; a row with figures beside it clips at the column width and has to be short.


def build_32(wb, anchors, a31, wcol):
    ws = wb["3.2 Total Cost"]
    f2.wipe(ws)
    ws.column_dimensions["A"].width = 2
    ws.cell(2, 2).value = "Total Cost - archetype against actual"
    ws.cell(2, 2).font = opts.TITLE
    r = opts.bar(ws, 4, 2, len(H32), "How the organisation is funded")
    r = opts.head(ws, r, 2, H32, W32)
    st = r
    # the subtotal rows on 3.1 are not contiguous, because a section label sits between
    # them, so they come from build_31 rather than being counted back from the total
    for lab, key in BLOCKS:
        ws.cell(r, 2).value = lab
        ws.cell(r, 2).font = opts.BODY
        ws.cell(r, 2).alignment = opts.LFT
        for c in (3, 4, 5, 6):
            f2._m(ws, r, c, f"='3.1 Group Summary'!${L(c)}${a31[key]}")
        f2._m(ws, r, 7, f"='3.1 Group Summary'!$G${a31[key]}", opts.CT)
        r += 1
    opts.row(ws, r, 2, ["Cost of the organisation today"] + [None] * (len(H32) - 1),
             [None] * len(H32), bg=opts.MID, bold=True, top=True)
    ws.cell(r, 2).alignment = opts.LFT
    for c in range(3, 8):
        x = ws.cell(r, c)
        x.value = f"=SUM({L(c)}{st}:{L(c)}{r-1})"
        x.number_format = opts.CT if c == 7 else opts.M2
        x.alignment = opts.RGT
    tot = r
    r += 1
    for lab, f, cf in (
            ("Leadership - the 8 GMs, outside the 525-role ledger",
             "=N(Lists!$AG$12)", "=N(Lists!$AG$11)"),
            ("Of which people in seat today",
             f'=SUMIFS({REV}!$AA$2:$AA${LAST},{REV}!$AK$2:$AK${LAST},"Filled")/1000000',
             f'=COUNTIFS({REV}!$AK$2:$AK${LAST},"Filled")'),
            ("Of which vacancies not yet filled",
             f'=SUMIFS({REV}!$AA$2:$AA${LAST},{REV}!$AK$2:$AK${LAST},"Vacant")/1000000',
             f'=COUNTIFS({REV}!$AK$2:$AK${LAST},"Vacant")')):
        # $115.11m is not payroll: it prices 135 vacancies nobody has been hired into, and
        # it excludes the GM layer, which has no role in the ledger to price
        ws.cell(r, 2).value = lab
        ws.cell(r, 2).font = opts.BODY
        ws.cell(r, 2).alignment = opts.LFT
        f2._m(ws, r, 4, f)
        f2._m(ws, r, 7, cf, opts.CT)
        r += 1
    opts.row(ws, r, 2, ["Total including the GM layer"] + [None] * (len(H32) - 1),
             [None] * len(H32), bg=opts.MID, bold=True, top=True)
    ws.cell(r, 2).alignment = opts.LFT
    for c in (3, 4, 5, 6):
        f2._m(ws, r, c, f"='3.1 Group Summary'!${L(c)}${a31['grand']}")
    f2._m(ws, r, 7, f"='3.1 Group Summary'!$G${a31['grand']}", opts.CT)
    for c in range(3, 8):
        ws.cell(r, c).font = opts.BOLD
    r += 2

    r = opts.bar(ws, r, 2, len(H32B), "Overhead roles - line by line")
    r = opts.head(ws, r, 2, H32B, W32)
    st2 = r
    for i in range(2, 8):
        ws.cell(r, 2).value = f"=Lists!$AF${i}"
        ws.cell(r, 2).font = opts.BODY
        ws.cell(r, 2).alignment = opts.LFT
        f2._m(ws, r, 3, f"=Lists!$AG${i}", opts.M3)
        f2._m(ws, r, 4, f"=Lists!$AH${i}", opts.CT)
        f2._m(ws, r, 5, f"=Lists!$AJ${i}")
        # A role carrying an overhead title inside a COE has AT set to its COE squad, not
        # to the overhead line, so AR = AT selects the roles that sit in a portfolio. That
        # is the split: counting every role with an overhead title against the portfolio
        # allowance was what produced 62 roles and $22.9m against 43 and $11.7m.
        both = f"{REV}!$AR$2:$AR${LAST},$B{r}"
        pf_only = f"{both},{REV}!$AT$2:$AT${LAST},$B{r}"
        gm = '=IF($B{r}="Leadership - 8 GMs",{v},{e})'
        f2._m(ws, r, 6, gm.format(r=r, v="N(Lists!$AG$11)",
                                  e=f"COUNTIFS({pf_only})"), opts.CT)
        f2._m(ws, r, 7, gm.format(
            r=r, v="N(Lists!$AG$12)",
            e=f"SUMIFS({REV}!$AA$2:$AA${LAST},{pf_only})/1000000"))
        f2._m(ws, r, 8, f"=$G{r}-$E{r}")
        f2._m(ws, r, 9, gm.format(r=r, v="0", e=f"COUNTIFS({both})-$F{r}"), opts.CT)
        f2._m(ws, r, 10, gm.format(
            r=r, v="0",
            e=f"SUMIFS({REV}!$AA$2:$AA${LAST},{both})/1000000-$G{r}"))
        ws.cell(r, 11).value = f"=Lists!${wcol}${i}"
        ws.cell(r, 11).font = opts.BODY
        ws.cell(r, 11).alignment = opts.CEN
        r += 1
    opts.row(ws, r, 2, ["Every overhead line"] + [None] * (len(H32B) - 1),
             [None] * len(H32B), bg=opts.MID, bold=True, top=True)
    ws.cell(r, 2).alignment = opts.LFT
    for c, nf in ((5, opts.M2), (6, opts.CT), (7, opts.M2), (8, opts.M2),
                  (9, opts.CT), (10, opts.M2)):
        x = ws.cell(r, c)
        x.value = f"=SUM({L(c)}{st2}:{L(c)}{r-1})"
        x.number_format, x.alignment = nf, opts.RGT
    r += 1
    # the one line that reconciles this block to the overhead row on 3.1
    opts.row(ws, r, 2, ["Of which drawn in the portfolios"] +
             [None] * (len(H32B) - 1), [None] * len(H32B), bg=opts.GREY, bold=True)
    ws.cell(r, 2).alignment = opts.LFT
    f2._m(ws, r, 5, "=N(Lists!$AJ$9)")
    f2._m(ws, r, 6, f"='3.1 Group Summary'!$G${a31['overhead']}", opts.CT)
    f2._m(ws, r, 7, f"='3.1 Group Summary'!$D${a31['overhead']}")
    f2._m(ws, r, 8, f"=$G{r}-$E{r}")
    for c in (5, 6, 7, 8):
        ws.cell(r, c).font = opts.BOLD
    ws.freeze_panes = "C6"
    return {"total": tot}


H33 = ["Portfolio", "How it is funded", "Squad", "Archetype Type", "Size",
       "Archetype roles", "Roles", "Filled", "Vacant", "Roles after decisions",
       "Archetype cost ($m)", "Actual cost ($m)", "Variance to archetype ($m)",
       "Cost after vacancy decisions ($m)"]
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
    ws.freeze_panes = "E6"
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
        f"3.1 Group Summary: four funding blocks, group total row {a31['total']}",
        "3.2 Total Cost: how the cost is funded, then the overhead lines with the one "
        "line that reconciles them to 3.1",
        f"3.3 Squad Detail: every squad tagged by how it is funded, total {a33['group_total']}"]


if __name__ == "__main__":
    import sys
    for x in run(sys.argv[1], sys.argv[2]):
        print("  ", x)
