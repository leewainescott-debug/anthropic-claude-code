"""All fourteen working tabs, layout 2A, out of one function.

Every 2.x tab is written by build() and nothing else touches them, so they cannot drift
apart the way the old ones did (nine different column-width profiles across fourteen tabs,
two title styles, because three different scripts had written them over time).

Layout, top to bottom:

    title
    Squad summary          bar
      header
      one row per delivery squad
      Delivery squads      subtotal
    Overhead roles         bar + header + rows + subtotal   (absent on the COEs and EGI)
    Total portfolio
    two control lines, both must read 0
    <Portfolio> FTE        bar
      header
      squad band carrying that squad's two totals
        one row per person, the lever the only yellow cell
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
S = dict(squad=2, type=3, size=4, aroles=5, roles=6, filled=7, vacant=8, hire=9,
         offshore=10, hold=11, remaining=12, acost=13, actual=14, var=15, after=16,
         newvar=17)
S_HDR = ["Squad", "Archetype Type", "Size", "Archetype roles", "Roles", "Filled",
         "Vacant", "To hire", "To offshore", "On hold", "Vacancies remaining",
         "Archetype cost ($m)", "Actual cost ($m)", "Variance to archetype ($m)",
         "Cost after vacancy decisions ($m)", "New variance ($m)"]
S_W = [30, 26, 7, 11, 7, 7, 8, 8, 11, 8, 12, 13, 12, 14, 17, 12]

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
    wipe(ws)
    ws.column_dimensions["A"].width = 2

    # ---- plan every row position first ----
    HDR = 6
    r = HDR + 1
    srow = {}
    for g in delivery:
        srow[g] = r
        r += 1
    r_del = r
    r += 1
    r_ohbar = r_ohhdr = r_oh = None
    if overhead:
        r += 1
        r_ohbar = r
        r_ohhdr = r + 1
        r = r_ohhdr + 1
        for g in overhead:
            srow[g] = r
            r += 1
        r_oh = r
        r += 1
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
    opts.bar(ws, 4, 2, len(S_HDR), "Squad summary")
    opts.head(ws, HDR, 2, S_HDR, S_W)
    ws.row_dimensions[5].height = 6

    def span(g):
        p = people.get(g, [])
        return (p[0][0], p[-1][0]) if p else (band[g], band[g])

    def squad(rw, g, is_oh):
        a, b = span(g)
        lev = f"${L(P['lever'])}${a}:${L(P['lever'])}${b}"
        st = f"${L(P['status'])}${a}:${L(P['status'])}${b}"
        ws.cell(rw, S["squad"]).value = g
        ws.cell(rw, S["squad"]).font = opts.BODY
        ws.cell(rw, S["squad"]).alignment = opts.LFT
        dash = ('="-"',)
        if is_oh or design is None:
            ws.cell(rw, S["type"]).value = ("Overhead - see 3.2" if is_oh
                                            else "COE - measured on budget, see 3.2")
            for k in ("size", "aroles", "acost", "var", "newvar"):
                x = ws.cell(rw, S[k])
                x.value = dash[0]
                x.alignment = opts.RGT
        else:
            key = f'${L(S["type"])}{rw}&"|"&${L(S["size"])}{rw}'
            m = f"MATCH(${L(S['squad'])}{rw},'{design}'!$B${lo}:$B${hi},0)"
            ws.cell(rw, S["type"]).value = (
                f"=IFERROR(INDEX('{design}'!$C${lo}:$C${hi},{m}),\"Not in the design\")")
            ws.cell(rw, S["size"]).value = (
                f"=IFERROR(INDEX('{design}'!$D${lo}:$D${hi},{m}),\"-\")")
            opts.money = None  # guard against accidental use
            _m(ws, rw, S["aroles"],
               f"=IFERROR(INDEX({A3}!$F$5:$F$23,MATCH({key},{A3}!$A$5:$A$23,0)),\"-\")",
               opts.C1)
            _m(ws, rw, S["acost"],
               f"=IFERROR(IF(INDEX('{design}'!$E${lo}:$E${hi},{m})=\"Offshore\","
               f"INDEX({A3}!$H$5:$H$23,MATCH({key},{A3}!$A$5:$A$23,0)),"
               f'INDEX({A3}!$G$5:$G$23,MATCH({key},{A3}!$A$5:$A$23,0))),"-")')
            _m(ws, rw, S["var"],
               f'=IFERROR(${L(S["actual"])}{rw}-${L(S["acost"])}{rw},"-")')
            _m(ws, rw, S["newvar"],
               f'=IFERROR(${L(S["after"])}{rw}-${L(S["acost"])}{rw},"-")')
        for k in ("type", "size"):
            ws.cell(rw, S[k]).font = opts.BODY
            ws.cell(rw, S[k]).alignment = opts.LFT
        for k, f in (("roles", f"=COUNTA(${L(P['name'])}${a}:${L(P['name'])}${b})"),
                     ("filled", f'=COUNTIFS({st},"Filled")'),
                     ("vacant", f'=COUNTIFS({st},"Vacant")'),
                     ("hire", f'=COUNTIFS({lev},"Hire")'),
                     ("offshore", f'=COUNTIFS({lev},"Offshore")'),
                     ("hold", f'=COUNTIFS({lev},"Hold")'),
                     ("remaining", f'=COUNTIFS({st},"Vacant")'
                                   f'-COUNTIFS({lev},"Hire")'
                                   f'-COUNTIFS({lev},"Offshore")')):
            _m(ws, rw, S[k], f, opts.CT)
        _m(ws, rw, S["actual"],
           f"=SUMIFS({REV}!$AA$2:$AA${LAST},{REV}!$AJ$2:$AJ${LAST},$C$3,"
           f"{REV}!$AT$2:$AT${LAST},${L(S['squad'])}{rw})/1000000")
        _m(ws, rw, S["after"],
           f"=SUM(${L(P['after'])}${a}:${L(P['after'])}${b})/1000000")

    for g in delivery:
        squad(srow[g], g, False)
    for g in overhead:
        squad(srow[g], g, True)

    def total(rw, label, r0, r1, bg, line=False):
        opts.row(ws, rw, 2, [label] + [None] * (len(S_HDR) - 1),
                 [None] * len(S_HDR), bg=bg, bold=True, top=line)
        ws.cell(rw, S["squad"]).alignment = opts.LFT
        for k in ("aroles", "roles", "filled", "vacant", "hire", "offshore", "hold",
                  "remaining", "acost", "actual", "var", "after", "newvar"):
            c = S[k]
            x = ws.cell(rw, c)
            x.value = f"=SUM({L(c)}{r0}:{L(c)}{r1})"
            x.number_format = (opts.M2 if c >= S["acost"] else
                               (opts.C1 if k == "aroles" else opts.CT))
            x.alignment = opts.RGT

    total(r_del, "Delivery squads", HDR + 1, r_del - 1, opts.GREY)
    if overhead:
        opts.bar(ws, r_ohbar, 2, len(S_HDR), "Overhead roles")
        opts.head(ws, r_ohhdr, 2, S_HDR, S_W)
        total(r_oh, "Overhead roles total", r_ohhdr + 1, r_oh - 1, opts.GREY)

    opts.row(ws, r_tot, 2, ["Total portfolio"] + [None] * (len(S_HDR) - 1),
             [None] * len(S_HDR), bg=opts.MID, bold=True, top=True)
    ws.cell(r_tot, S["squad"]).alignment = opts.LFT
    for k in ("aroles", "roles", "filled", "vacant", "hire", "offshore", "hold",
              "remaining", "acost", "actual", "var", "after", "newvar"):
        c = S[k]
        x = ws.cell(r_tot, c)
        x.value = (f"=N({L(c)}{r_del})+N({L(c)}{r_oh})" if overhead
                   else f"=N({L(c)}{r_del})")
        x.number_format = (opts.M2 if c >= S["acost"] else
                           (opts.C1 if k == "aroles" else opts.CT))
        x.alignment = opts.RGT

    for i, (lab, col, f, nf) in enumerate([
            ("Control - roles against the ledger, must be 0", S["roles"],
             f"=${L(S['roles'])}${r_tot}-COUNTIFS({REV}!$AJ$2:$AJ${LAST},$C$3)",
             opts.CT),
            ("Control - cost against the ledger ($m), must be 0", S["actual"],
             f"=ROUND(${L(S['actual'])}${r_tot}-SUMIFS({REV}!$AA$2:$AA${LAST},"
             f"{REV}!$AJ$2:$AJ${LAST},$C$3)/1000000,6)", opts.M2)]):
        ws.cell(r_ctl + i, 2).value = lab
        ws.cell(r_ctl + i, 2).font = opts.BODY
        _m(ws, r_ctl + i, col, f, nf)

    # ---- FTE ----
    opts.bar(ws, r_ftebar, 2, len(P_HDR), f"{pf} FTE")
    opts.head(ws, r_ftehdr, 2, P_HDR, P_W)
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
        ws.cell(bd, P["role"]).value = (
            f'=COUNTA(${L(P["name"])}${a}:${L(P["name"])}${b})&" roles"')
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

    ws.freeze_panes = f"C{HDR + 1}"
    return {"tab": tab, "pf": pf, "total_row": r_tot, "delivery_row": r_del,
            "overhead_row": r_oh, "header_row": HDR, "first_squad": HDR + 1,
            "last_squad": r_del - 1, "squads": delivery, "overhead": overhead,
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
        out.append(f"{tab}: {len(a['squads'])} squads, {len(a['overhead'])} overhead, "
                   f"{a['people']} people")
    json.dump(anchors, open("anchors_final.json", "w"), indent=1)
    wb.save(dst)
    return out


if __name__ == "__main__":
    import sys
    for x in run(sys.argv[1], sys.argv[2]):
        print("  ", x)
