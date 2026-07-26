"""All fourteen working tabs, layout 2A, out of one function.

Every 2.x tab is written by build() and nothing else touches them, so they cannot drift
apart the way the old ones did (nine different column-width profiles across fourteen tabs,
two title styles, because three different scripts had written them over time).

Layout, top to bottom:

    title
    Squads priced by an archetype     bar + header + rows + subtotal
    Directly funded                   bar + header + rows + subtotal
    Overhead roles                    bar + header + rows + subtotal
    Total portfolio
    two control lines, both must read 0
    <Portfolio> FTE        bar
      header
      squad band carrying that squad's two totals
        one row per person, the lever the only yellow cell

All three blocks appear on all fourteen tabs. Where a portfolio has nothing in a block the
block carries one row reading "None" rather than being dropped, so the fourteen tabs are
structurally identical and the absence is stated rather than left to be inferred - the COEs
have no overhead by decision, and a reader should be able to see that on the tab.
"""
import collections
import copy
import json
import re

import openpyxl
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

import opts
from build_2xfix import DESIGN, squad_table_bounds

REVIEW = "REVIEW - Complete Role Mapping"
REV = f"'{REVIEW}'"
A3 = "'0.3 Squad Archetypes'"
LAST = 528

# ---- squad summary columns, B onwards ----
# "Roles after decisions" replaced "Vacancies remaining", which counted vacancies set to
# Hold and so read as though a cancelled vacancy were still outstanding. Roles after
# decisions is roles less anything put On hold, which is what D20 promised: pull Hold and
# the headcount moves, not just the cost.
S = dict(squad=2, type=3, size=4, aroles=5, roles=6, filled=7, vacant=8, hire=9,
         offshore=10, hold=11, rafter=12, acost=13, actual=14, var=15, after=16,
         newvar=17)
# The owner's own column names, taken off his markup of 2.8.
S_HDR = ["Squad", "Archetype Type", "Squad Size", "Archetype roles", "Total roles",
         "Filled", "Vacant", "To hire", "To offshore", "On hold",
         "Total roles after decisions", "Archetype cost ($m)", "Actual cost ($m)",
         "Variance to archetype ($m)", "Squad cost after decisions ($m)",
         "New variance ($m)"]
SECTION = {"arch": "Squads priced by an archetype",
           "direct": "Directly funded programmes and platforms",
           "oh": "Overhead roles"}
NONE_TEXT = {"arch": "No squad here is priced by an archetype",
             "direct": "No directly funded programme here",
             "oh": "The COEs carry no overhead"}
S_W = [30, 26, 11, 11, 9, 8, 8, 8, 11, 8, 14, 13, 12, 14, 17, 14]

# ---- FTE columns ----
P = dict(name=2, role=3, status=4, lever=5, cost=6, after=7)
P_HDR = ["Name", "Role", "Status", "Vacancy lever", "Role cost ($)",
         "Cost after decision ($)"]
P_W = [30, 44, 11, 15, 15, 19]

OH_ORDER = ["Head of Technology", "Business Partner", "Domain Architect",
            "Delivery Manager", "Technology Manager", "Leadership",
            "Leadership - squad not stated"]


def ledger(wv):
    R = wv[REVIEW]
    out = []
    for i in range(2, LAST + 1):
        if not str(R.cell(i, 2).value or "").strip():
            continue
        pf = R.cell(i, 36).value
        is_coe = str(pf or "").startswith("COE") or pf == "EGI"
        out.append({"row": i, "pf": pf, "grp": R.cell(i, 46).value,
                    "status": R.cell(i, 37).value,
                    "oh": (not is_coe) and R.cell(i, 44).value != "Squad"})
    return out


def groups(rows):
    seen = {}
    for r in rows:
        seen.setdefault(r["grp"], r["oh"])
    delivery = sorted(g for g, oh in seen.items() if not oh)
    overhead = [g for g in OH_ORDER if seen.get(g)]
    overhead += sorted(g for g, oh in seen.items() if oh and g not in OH_ORDER)
    return delivery, overhead


def has_archetype(wb, wv, design, squad):
    """True when the squad appears in the design tab's squad table AND that row carries a
    type and size the archetype library prices. Strategic programmes deliberately do not
    (register item 21), and a squad missing from the design cannot either.

    This partition matters: the old build put the archetype for SOME squads against the
    actual for ALL of them in one subtotal, so the group read 11.49m over design when the
    squads that actually have an archetype are 0.56m under it.
    """
    if design is None:
        return False
    lo, hi = squad_table_bounds(wb, design)
    ws = wv[design]
    keys = {str(wv["0.3 Squad Archetypes"].cell(r, 1).value or "").strip()
            for r in range(5, 24)}
    for r in range(lo, hi + 1):
        if str(ws.cell(r, 2).value or "").strip() != squad:
            continue
        ty = str(ws.cell(r, 3).value or "").strip()
        sz = str(ws.cell(r, 4).value or "").strip()
        return f"{ty}|{sz}" in keys
    return False


def wipe(ws):
    blank = copy.copy(openpyxl.cell.cell.Cell(ws)._style)
    for m in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(m))
    for row in ws.iter_rows():
        for c in row:
            c.value = None
            c._style = copy.copy(blank)
    for k in list(ws.column_dimensions):
        del ws.column_dimensions[k]
    for k in list(ws.row_dimensions):
        del ws.row_dimensions[k]
    ws.conditional_formatting._cf_rules.clear()
    ws.data_validations.dataValidation = []
    ws.freeze_panes = None
    ws.sheet_view.showGridLines = False


def build(wb, wv, tab, rows, bounds):
    ws = wb[tab]
    design = DESIGN.get(tab)
    pf = rows[0]["pf"]
    delivery, overhead = groups(rows)
    lo, hi = bounds[design] if design else (0, 0)
    # split the delivery squads by whether the archetype library actually prices them
    arch = [g for g in delivery if has_archetype(wb, wv, design, g)]
    direct = [g for g in delivery if g not in arch]
    wipe(ws)
    ws.column_dimensions["A"].width = 2

    # ---- plan every row position first ----
    # Every block gets at least one row. An empty block used to collapse to a subtotal
    # whose SUM range ran backwards from the header - SUM(M7:M6), which Excel reads as
    # SUM(M6:M7) and so includes the subtotal cell itself. That was a circular reference
    # on all four COE tabs.
    HDR = 6
    r = HDR + 1
    srow, empty, label, sub = {}, set(), {}, {}
    for kind, names in (("arch", arch), ("direct", direct), ("oh", overhead)):
        label[kind] = r
        r += 1
        if names:
            for g in names:
                srow[g] = r
                r += 1
        else:
            empty.add(kind)
            srow[f"-{kind}"] = r
            r += 1
        sub[kind] = r
        r += 1
    r_tot = r
    r += 2
    r_ctl = r
    r += 3
    r_ftebar = r
    r_ftehdr = r + 1
    r = r_ftehdr + 1
    band, people = {}, {}
    for g in delivery + overhead:
        band[g] = r
        r += 1
        for x in [y for y in rows if y["grp"] == g]:
            people.setdefault(g, []).append((r, x))
            r += 1

    # ---- header ----
    ws.cell(2, 2).value = f"{pf} - working copy"
    ws.cell(2, 2).font = opts.TITLE
    ws.cell(3, 2).value = "Portfolio"
    ws.cell(3, 2).font = opts.BOLD
    ws.cell(3, 3).value = pf
    ws.cell(3, 3).font = opts.BODY
    opts.bar(ws, 4, 2, len(S_HDR), f"{pf} - archetype cost against actual cost")
    # One table, one header row, the overhead lines inside it. They used to sit in a second
    # table below with a header of their own; the owner's instruction is that the overheads
    # belong up top with everything else.
    W = list(S_W)
    opts.head(ws, HDR, 2, list(S_HDR), W)
    ws.row_dimensions[5].height = 6

    def span(g):
        p = people.get(g, [])
        return (p[0][0], p[-1][0]) if p else (band[g], band[g])

    def squad(rw, g, kind):
        a, b = span(g)
        lev = f"${L(P['lever'])}${a}:${L(P['lever'])}${b}"
        st = f"${L(P['status'])}${a}:${L(P['status'])}${b}"
        ws.cell(rw, S["squad"]).value = g
        ws.cell(rw, S["squad"]).font = opts.BODY
        ws.cell(rw, S["squad"]).alignment = opts.LFT
        if kind == "oh" or design is None:
            # the bar above each block already says what the block is, so the type
            # column stays a dash rather than repeating a sentence on every row
            ws.cell(rw, S["type"]).value = '="-"'
            ws.cell(rw, S["type"]).alignment = opts.RGT
            for k in ("size", "aroles"):
                x = ws.cell(rw, S[k])
                x.value = '="-"'
                x.alignment = opts.RGT
            if kind == "oh":
                for k in ("acost", "var", "newvar"):
                    x = ws.cell(rw, S[k])
                    x.value = '="-"'
                    x.alignment = opts.RGT
            else:
                # a COE squad has no archetype and no per-squad design row. Its design is
                # what its own roles cost, which is how the owner asked the COEs to be
                # compared, so the variance is nil by construction and says so.
                _m(ws, rw, S["acost"], f'=${L(S["actual"])}{rw}')
                _m(ws, rw, S["var"],
                   f'=ROUND(${L(S["actual"])}{rw}-${L(S["acost"])}{rw},6)')
                _m(ws, rw, S["newvar"],
                   f'=ROUND(${L(S["after"])}{rw}-${L(S["acost"])}{rw},6)')
        else:
            key = f'${L(S["type"])}{rw}&"|"&${L(S["size"])}{rw}'
            m = f"MATCH(${L(S['squad'])}{rw},'{design}'!$B${lo}:$B${hi},0)"
            ws.cell(rw, S["type"]).value = (
                f"=IFERROR(INDEX('{design}'!$C${lo}:$C${hi},{m}),\"Not on the 1.x tab\")")
            # INDEX over an empty cell returns 0, not an error, so a strategic programme
            # with no size printed 0 in the Size column
            ws.cell(rw, S["size"]).value = (
                f"=IFERROR(IF(INDEX('{design}'!$D${lo}:$D${hi},{m})=\"\",\"-\","
                f"INDEX('{design}'!$D${lo}:$D${hi},{m})),\"-\")")
            if kind == "arch":
                _m(ws, rw, S["aroles"],
                   f"=IFERROR(INDEX({A3}!$F$5:$F$23,MATCH({key},{A3}!$A$5:$A$23,0)),"
                   f'"-")', opts.C1)
                _m(ws, rw, S["acost"],
                   f"=IFERROR(IF(INDEX('{design}'!$E${lo}:$E${hi},{m})=\"Offshore\","
                   f"INDEX({A3}!$H$5:$H$23,MATCH({key},{A3}!$A$5:$A$23,0)),"
                   f'INDEX({A3}!$G$5:$G$23,MATCH({key},{A3}!$A$5:$A$23,0))),"-")')
            else:
                # Directly funded. No archetype prices it - the owner's instruction is
                # that strategic programmes are directly funded and do not fit one - so
                # the design figure is the amount typed against the programme on its own
                # platform block, column H of the design tab. EGI is the exception: the
                # owner ruled EGI is actuals, and four of the six EGI inputs are typed
                # zero or left blank, so pricing off them would report the whole of EGI
                # as overspend against nothing.
                _m(ws, rw, S["aroles"], '="-"', opts.C1)
                if g.startswith("EGI"):
                    _m(ws, rw, S["acost"], f'=${L(S["actual"])}{rw}')
                else:
                    _m(ws, rw, S["acost"],
                       f"=IFERROR(INDEX('{design}'!$H${lo}:$H${hi},{m}),\"-\")")
            # rounded, or a squad priced exactly at its archetype shows (0.00) on a
            # residual of a few billionths
            _m(ws, rw, S["var"],
               f'=IFERROR(ROUND(${L(S["actual"])}{rw}-${L(S["acost"])}{rw},6),"-")')
            _m(ws, rw, S["newvar"],
               f'=IFERROR(ROUND(${L(S["after"])}{rw}-${L(S["acost"])}{rw},6),"-")')
        for k in ("type", "size"):
            ws.cell(rw, S[k]).font = opts.BODY
            ws.cell(rw, S[k]).alignment = opts.LFT
        # To hire only applies to a vacancy. To offshore and On hold can apply to a
        # filled role too, because offshoring a filled resource is a decision the tool
        # exists to price. Vacancies remaining therefore has to look at vacant rows
        # only: counting an offshored filled person against it drove the count negative.
        for k, f in (("roles", f"=COUNTA(${L(P['name'])}${a}:${L(P['name'])}${b})"),
                     ("filled", f'=COUNTIFS({st},"Filled")'),
                     ("vacant", f'=COUNTIFS({st},"Vacant")'),
                     ("hire", f'=COUNTIFS({st},"Vacant",{lev},"Hire")'),
                     ("offshore", f'=COUNTIFS({lev},"Offshore")'),
                     ("hold", f'=COUNTIFS({lev},"Hold")'),
                     ("rafter", f'=${L(S["roles"])}{rw}-${L(S["hold"])}{rw}')):
            _m(ws, rw, S[k], f, opts.CT)
        _m(ws, rw, S["actual"],
           f"=SUMIFS({REV}!$AA$2:$AA${LAST},{REV}!$AJ$2:$AJ${LAST},$C$3,"
           f"{REV}!$AT$2:$AT${LAST},${L(S['squad'])}{rw})/1000000")
        _m(ws, rw, S["after"],
           f"=SUM(${L(P['after'])}${a}:${L(P['after'])}${b})/1000000")

    def none_row(rw, why):
        """A block with nothing in it still gets a row, so the tab states the absence."""
        ws.cell(rw, S["squad"]).value = "None"
        ws.cell(rw, S["squad"]).font = opts.BODY
        ws.cell(rw, S["squad"]).alignment = opts.LFT
        ws.cell(rw, S["type"]).value = why
        ws.cell(rw, S["type"]).font = opts.BODY
        ws.cell(rw, S["type"]).alignment = opts.LFT
        for k in ("size", "aroles", "roles", "filled", "vacant", "hire", "offshore",
                  "hold", "rafter", "acost", "actual", "var", "after", "newvar"):
            x = ws.cell(rw, S[k])
            x.value = '="-"'
            x.font, x.alignment = opts.BODY, opts.RGT

    for g in arch:
        squad(srow[g], g, "arch")
    for g in direct:
        squad(srow[g], g, "direct")
    for g in overhead:
        squad(srow[g], g, "oh")

    def total(rw, label, r0, r1, bg, line=False):
        opts.row(ws, rw, 2, [label] + [None] * (len(S_HDR) - 1),
                 [None] * len(S_HDR), bg=bg, bold=True, top=line)
        ws.cell(rw, S["squad"]).alignment = opts.LFT
        for k in ("aroles", "roles", "filled", "vacant", "hire", "offshore", "hold",
                  "rafter", "acost", "actual", "var", "after", "newvar"):
            c = S[k]
            x = ws.cell(rw, c)
            # a variance on a total is actual less archetype. Summing the row variances
            # drops every row whose archetype column holds "-", which understated the
            # directly funded subtotal on six of the fourteen tabs.
            # and where the block has no archetype at all - the overhead lines - the
            # variance is a dash, not the whole cost measured against zero
            if k == "var":
                x.value = (f'=IF(N(${L(S["acost"])}{rw})=0,"-",'
                           f"ROUND(${L(S['actual'])}{rw}-${L(S['acost'])}{rw},6))")
            elif k == "newvar":
                x.value = (f'=IF(N(${L(S["acost"])}{rw})=0,"-",'
                           f"ROUND(${L(S['after'])}{rw}-${L(S['acost'])}{rw},6))")
            else:
                x.value = f"=SUM({L(c)}{r0}:{L(c)}{r1})"
            x.number_format = (opts.M2 if c >= S["acost"] else
                               (opts.C1 if k == "aroles" else opts.CT))
            x.alignment = opts.RGT

    for kind in ("arch", "direct", "oh"):
        opts.row(ws, label[kind], 2, [SECTION[kind]] + [None] * (len(S_HDR) - 1),
                 [None] * len(S_HDR), bg=opts.PALE, bold=True)
        ws.cell(label[kind], 2).alignment = opts.LFT
        if kind in empty:
            none_row(srow[f"-{kind}"], NONE_TEXT[kind])
        total(sub[kind], f"{SECTION[kind]} total", label[kind] + 1, sub[kind] - 1,
              opts.GREY)

    opts.row(ws, r_tot, 2, ["Total portfolio"] + [None] * (len(S_HDR) - 1),
             [None] * len(S_HDR), bg=opts.MID, bold=True, top=True)
    ws.cell(r_tot, S["squad"]).alignment = opts.LFT
    for k in ("aroles", "roles", "filled", "vacant", "hire", "offshore", "hold",
              "rafter", "acost", "actual", "var", "after", "newvar"):
        c = S[k]
        x = ws.cell(r_tot, c)
        # Across the whole portfolio the variance is the sum of the three block
        # variances, not actual less archetype: the archetype column only prices the first
        # block, so actual-less-archetype at portfolio level would charge the overhead cost
        # against a squad archetype. Each block variance is already on one basis.
        if k in ("var", "newvar"):
            x.value = "=" + "+".join(f"N({L(c)}{sub[kk]})"
                                     for kk in ("arch", "direct", "oh"))
        else:
            x.value = "=" + "+".join(f"N({L(c)}{sub[kk]})"
                                     for kk in ("arch", "direct", "oh"))
        x.number_format = (opts.M2 if c >= S["acost"] else
                           (opts.C1 if k == "aroles" else opts.CT))
        x.alignment = opts.RGT

    for i, (lab, col, f, nf) in enumerate([
            ("Control - roles against the ledger, must be 0", S["roles"],
             f"=${L(S['roles'])}${r_tot}-COUNTIFS({REV}!$AJ$2:$AJ${LAST},$C$3)",
             opts.CTL_C),
            ("Control - cost against the ledger ($m), must be 0", S["actual"],
             f"=ROUND(${L(S['actual'])}${r_tot}-SUMIFS({REV}!$AA$2:$AA${LAST},"
             f"{REV}!$AJ$2:$AJ${LAST},$C$3)/1000000,6)", opts.CTL_M)]):
        ws.cell(r_ctl + i, 2).value = lab
        ws.cell(r_ctl + i, 2).font = opts.BODY
        _m(ws, r_ctl + i, col, f, nf)

    # ---- FTE ----
    opts.bar(ws, r_ftebar, 2, len(P_HDR), f"{pf} FTE")
    for i, w in enumerate(P_W):
        W[i] = max(W[i], w)
    opts.head(ws, r_ftehdr, 2, list(P_HDR), W)
    dv = DataValidation(type="list", formula1='"Filled,Hire,Hold,Offshore"',
                        allow_blank=False, showDropDown=False)
    ws.add_data_validation(dv)
    R = wv[REVIEW]
    for g in delivery + overhead:
        a, b = span(g)
        bd = band[g]
        opts.row(ws, bd, 2, [g] + [None] * (len(P_HDR) - 1), [None] * len(P_HDR),
                 bg=opts.PALE, bold=True)
        ws.cell(bd, P["name"]).alignment = opts.LFT
        n = f'COUNTA(${L(P["name"])}${a}:${L(P["name"])}${b})'
        ws.cell(bd, P["role"]).value = f'={n}&IF({n}=1," role"," roles")'
        ws.cell(bd, P["role"]).alignment = opts.LFT
        for k in ("cost", "after"):
            x = ws.cell(bd, P[k])
            x.value = f"=SUM({L(P[k])}{a}:{L(P[k])}{b})"
            x.number_format, x.alignment = opts.M0, opts.RGT
        for rw, x in people.get(g, []):
            i = x["row"]
            for k, col in (("name", "B"), ("role", "C"), ("status", "AK")):
                cell = ws.cell(rw, P[k])
                cell.value = f"=INDEX({REV}!${col}:${col},{i})"
                cell.font, cell.alignment = opts.BODY, opts.LFT
            lv = ws.cell(rw, P["lever"])
            lv.value = "Filled" if x["status"] == "Filled" else "Hire"
            lv.fill, lv.border = opts.fl(opts.YEL), opts.BOX
            lv.font, lv.alignment = opts.BODY, opts.CEN
            dv.add(lv)
            _m(ws, rw, P["cost"], f"=INDEX({REV}!$AA:$AA,{i})", opts.M0)
            _m(ws, rw, P["after"],
               f"=${L(P['cost'])}{rw}*IFERROR(INDEX(Lists!$AD:$AD,"
               f"MATCH(${L(P['lever'])}{rw},Lists!$AC:$AC,0)),1)", opts.M0)

    # no frozen panes anywhere in this workbook - owner's instruction
    return {"tab": tab, "pf": pf, "total_row": r_tot, "delivery_row": sub["arch"],
            "direct_row": sub["direct"], "overhead_row": sub["oh"], "header_row": HDR,
            "first_squad": label["arch"] + 1, "last_squad": sub["arch"] - 1,
            "squads": arch, "direct": direct, "overhead": overhead,
            "srow": srow, "people": sum(len(v) for v in people.values()),
            "cols": S}


def _m(ws, r, c, formula, fmt=None):
    x = ws.cell(r, c)
    x.value = formula
    x.number_format = fmt or opts.M2
    x.alignment = opts.RGT
    x.font = opts.BODY
    return x


def run(src, dst):
    wb = openpyxl.load_workbook(src)
    wv = openpyxl.load_workbook(src, data_only=True)
    rows = ledger(wv)
    bypf = collections.defaultdict(list)
    for x in rows:
        bypf[x["pf"]].append(x)
    bounds = {d: squad_table_bounds(wb, d) for d in set(DESIGN.values())}
    tabs = [s for s in wb.sheetnames if re.match(r"^2\.\d+ ", s)]
    anchors, out = {}, []
    for tab in tabs:
        pf = wv[tab]["C3"].value
        if pf not in bypf:
            pf = next(p for p in bypf if p and tab.endswith(str(p)))
        a = build(wb, wv, tab, bypf[pf], bounds)
        anchors[tab] = a
        out.append(f"{tab}: {len(a['squads'])} archetyped, {len(a['direct'])} directly "
                   f"funded, {len(a['overhead'])} overhead, {a['people']} people")
    json.dump(anchors, open("anchors_final.json", "w"), indent=1)
    wb.save(dst)
    return out


if __name__ == "__main__":
    import sys
    for x in run(sys.argv[1], sys.argv[2]):
        print("  ", x)
