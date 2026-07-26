"""Rewire everything that read the old 3.x row positions, and rebuild Exec and 4.0.

Rebuilding 3.1, 3.2 and 3.3 moved every row on them, and 59 formulas elsewhere still
pointed at where those numbers used to be. One of them mattered a lot: Lists!AH2 counted
the portfolios from '3.1'!B6:B15 to price the overhead allowance, and with the rows moved
the allowance silently fell from 12.925 to 12.13. That is the join-by-row failure mode
exactly, so the replacements below key off labels wherever a label exists.

Exec Summary and 4.0 Data QA are rebuilt rather than patched, because both were written
against the budget comparison that has been dropped.
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


def find_row(ws, label, col=2, limit=140):
    for r in range(1, min(ws.max_row, limit) + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.strip().startswith(label):
            return r
    raise KeyError(f"{ws.title}: no row starting {label!r}")


def anchors(wb):
    s1, s2, s3 = wb["3.1 Group Summary"], wb["3.2 Total Cost"], wb["3.3 FTE View"]
    return {
        "g31": find_row(s1, "Group total"),
        "del31": find_row(s1, "Delivery squads in the portfolios"),
        "coe31": find_row(s1, "COEs and EGI"),
        "oh31": find_row(s1, "Overhead roles"),
        "first31": find_row(s1, "Portfolio") + 1,
        "g32": find_row(s2, "Cost of the organisation today"),
        "ohtot32": find_row(s2, "Overhead total"),
        "ohfirst32": find_row(s2, "Overhead line") + 1,
        "g33": find_row(s3, "Group total"),
        "first33": find_row(s3, "Portfolio") + 1,
    }


def repoint(wb, a):
    """Replace the stale references. Each entry is (old, new, what it means)."""
    g1, g3 = a["g31"], a["g33"]
    d1, o1 = a["del31"], a["oh31"]
    MAP = {
        # 3.2's old total row carried the group facts; they live on 3.1 now
        "'3.2 Total Cost'!$C$21": f"'3.1 Group Summary'!$G${g1}",      # roles
        "'3.2 Total Cost'!$D$21": f"'3.1 Group Summary'!$H${g1}",      # filled
        "'3.2 Total Cost'!$E$21": f"'3.1 Group Summary'!$I${g1}",      # vacant
        "'3.2 Total Cost'!$F$21": f"'3.1 Group Summary'!$D${g1}",      # actual cost
        "'3.2 Total Cost'!$H$21": f"'3.1 Group Summary'!$D${d1}",      # squad cost
        "'3.2 Total Cost'!$M$21": f"'3.1 Group Summary'!$D${o1}",      # overhead cost
        "'3.2 Total Cost'!$J$16": f"'3.1 Group Summary'!$C${d1}",      # archetype cost
        "'3.2 Total Cost'!$K$16": f"'3.1 Group Summary'!$E${d1}",      # variance
        "'3.2 Total Cost'!$I$16": f"'3.3 FTE View'!$F${g3}",           # archetype roles
        "'3.2 Total Cost'!$H$36": f"'3.2 Total Cost'!$H${a['ohtot32']}",
        "'3.1 Group Summary'!$D$20": f"'3.1 Group Summary'!$D${g1}",
        "'3.1 Group Summary'!$J$20": f"'3.1 Group Summary'!$G${g1}",
        "'3.3 FTE View'!$I$113": f"'3.3 FTE View'!$G${g3}",
        # the portfolio count that prices the overhead allowance
        "COUNTA('3.1 Group Summary'!$B$6:$B$15)":
            f"COUNTA('3.1 Group Summary'!$B${a['first31']}:$B${d1-1})",
        # 3.3's archetype-roles range, per portfolio
        "'3.3 FTE View'!$F$6:$F$94":
            f"'3.3 FTE View'!$F${a['first33']}:$F${g3-1}",
        "'3.3 FTE View'!$B$6:$B$94":
            f"'3.3 FTE View'!$B${a['first33']}:$B${g3-1}",
        "'3.3 FTE View'!$C$6:$C$94":
            f"'3.3 FTE View'!$C${a['first33']}:$C${g3-1}",
        '"<>Portfolio total"': '"<>*total"',
    }
    n = 0
    for ws in wb.worksheets:
        if ws.title in ("3.1 Group Summary", "3.2 Total Cost", "3.3 FTE View"):
            continue
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not (isinstance(v, str) and v.startswith("=")):
                    continue
                new = v
                for old, rep in MAP.items():
                    if old in new:
                        new = new.replace(old, rep)
                if new != v:
                    c.value = new
                    n += 1
    return n


# ------------------------------------------------------------------- Exec Summary
def build_exec(wb, a):
    ws = wb["Exec Summary"]
    f2.wipe(ws)
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 62
    ws.column_dimensions["C"].width = 18
    g1, d1, c1, o1, g3 = a["g31"], a["del31"], a["coe31"], a["oh31"], a["g33"]
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
        ("Roles in the ledger", f"='3.1 Group Summary'!$G${g1}", opts.CT),
        ("Filled", f"='3.1 Group Summary'!$H${g1}", opts.CT),
        ("Vacant", f"='3.1 Group Summary'!$I${g1}", opts.CT),
        ("Cost today ($m)", f"='3.1 Group Summary'!$D${g1}", opts.M2)])

    r = block(r, "Against the design", [
        ("Delivery squads - archetype cost ($m)",
         f"='3.1 Group Summary'!$C${d1}", opts.M2),
        ("Delivery squads - actual cost ($m)",
         f"='3.1 Group Summary'!$D${d1}", opts.M2),
        ("Delivery squads over/(under) the archetype ($m)",
         f"='3.1 Group Summary'!$E${d1}", opts.M2),
        ("Overhead roles over/(under) their allowance ($m)",
         f"='3.1 Group Summary'!$E${o1}", opts.M2),
        ("COEs and EGI over/(under) their 1.x design ($m)",
         f"='3.1 Group Summary'!$E${c1}", opts.M2),
        ("Total over/(under) design ($m)", f"='3.1 Group Summary'!$E${g1}", opts.M2)])

    r = block(r, "The vacancy decision", [
        ("Vacant roles", f"='3.1 Group Summary'!$I${g1}", opts.CT),
        ("Cost of hiring every vacancy ($m)",
         f"=SUMIFS({REV}!$AA$2:$AA${LAST},{REV}!$AK$2:$AK${LAST},\"Vacant\")/1000000",
         opts.M2),
        ("Cost after the decisions set today ($m)",
         f"='3.1 Group Summary'!$F${g1}", opts.M2),
        ("Impact of those decisions ($m)",
         f"='3.1 Group Summary'!$F${g1}-'3.1 Group Summary'!$D${g1}", opts.M2)])

    # portfolio drill-down
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
    names = [str(wb["3.1 Group Summary"].cell(x, 2).value)
             for x in range(a["first31"], g1)
             if wb["3.1 Group Summary"].cell(x, 2).value
             and not str(wb["3.1 Group Summary"].cell(x, 2).value).startswith(
                 ("Delivery squads", "COEs and EGI", "Overhead roles"))]
    dv = DataValidation(type="list", formula1='"' + ",".join(names) + '"',
                        allow_blank=False, showDropDown=False)
    ws.add_data_validation(dv)
    dv.add(pick)
    r += 1
    lo, hi = a["first31"], g1 - 1
    for lab, col, nf in (("Design cost ($m)", "C", opts.M2),
                         ("Actual cost ($m)", "D", opts.M2),
                         ("Over/(under) design ($m)", "E", opts.M2),
                         ("Cost after vacancy decisions ($m)", "F", opts.M2),
                         ("Roles", "G", opts.CT), ("Filled", "H", opts.CT),
                         ("Vacant", "I", opts.CT)):
        ws.cell(r, 2).value = lab
        ws.cell(r, 2).font = opts.BODY
        ws.cell(r, 2).alignment = opts.LFT
        ws.cell(r, 2).border = opts.BOX
        x = ws.cell(r, 3)
        x.value = (f"=IFERROR(INDEX('3.1 Group Summary'!${col}${lo}:${col}${hi},"
                   f"MATCH($C${sel},'3.1 Group Summary'!$B${lo}:$B${hi},0)),\"-\")")
        x.number_format, x.alignment = nf, opts.RGT
        x.font, x.border = opts.BODY, opts.BOX
        r += 1
    ws.freeze_panes = "C4"
    return r


# ---------------------------------------------------------------------- 4.0 Data QA
def build_qa(wb, a, a2):
    ws = wb["4.0 Data QA"]
    f2.wipe(ws)
    ws.column_dimensions["A"].width = 2
    ws.cell(2, 2).value = "Data QA - every difference must read zero"
    ws.cell(2, 2).font = opts.TITLE
    HDR = 4
    opts.head(ws, HDR, 2, ["Check", "Model", "Expected", "Difference"], [66, 16, 16, 14])
    g1, g3, g32 = a["g31"], a["g33"], a["g32"]
    checks = [
        ("Roles on 3.1 against the ledger", f"='3.1 Group Summary'!$G${g1}",
         f"=COUNTA({REV}!$B$2:$B${LAST})", opts.CT),
        ("Filled on 3.1 against the ledger", f"='3.1 Group Summary'!$H${g1}",
         f'=COUNTIFS({REV}!$AK$2:$AK${LAST},"Filled")', opts.CT),
        ("Vacant on 3.1 against the ledger", f"='3.1 Group Summary'!$I${g1}",
         f'=COUNTIFS({REV}!$AK$2:$AK${LAST},"Vacant")', opts.CT),
        ("Cost on 3.1 against the ledger ($m)", f"='3.1 Group Summary'!$D${g1}",
         f"=SUM({REV}!$AA$2:$AA${LAST})/1000000", opts.M2),
        ("Roles on 3.3 against 3.1", f"='3.3 FTE View'!$G${g3}",
         f"='3.1 Group Summary'!$G${g1}", opts.CT),
        ("Cost on 3.3 against 3.1 ($m)", f"='3.3 FTE View'!$K${g3}",
         f"='3.1 Group Summary'!$D${g1}", opts.M2),
        ("Cost on 3.2 against 3.1 ($m)", f"='3.2 Total Cost'!$D${g32}",
         f"='3.1 Group Summary'!$D${g1}", opts.M2),
        ("Design cost on 3.2 against 3.1 ($m)", f"='3.2 Total Cost'!$C${g32}",
         f"='3.1 Group Summary'!$C${g1}", opts.M2),
        ("Overhead allowance on 3.2 against Lists ($m)",
         f"='3.2 Total Cost'!$F${a['ohtot32']}", "=N(Lists!$AJ$8)", opts.M2),
    ]
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
                opts.CTL_C if nf == opts.CT else opts.CTL_M)
            x.alignment, x.font, x.border = opts.RGT, opts.BODY, opts.BOX
        r += 1
    opts.row(ws, r, 2, ["Checks failing", None, None, None], [None] * 4,
             bg=opts.MID, bold=True, top=True)
    ws.cell(r, 2).alignment = opts.LFT
    x = ws.cell(r, 5)
    x.value = f'=COUNTIF($E{HDR+1}:$E{r-1},"<>0")'
    x.number_format, x.alignment = opts.CT, opts.RGT
    ws.freeze_panes = "C5"
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
    build_exec(wb, a)
    k = build_qa(wb, a, a2)
    nc = cream(wb)
    wb.save(dst)
    return [f"{n} formulas repointed at the rebuilt 3.x rows",
            f"{nc} inputs recoloured to cream across the live model",
            "Exec Summary rebuilt on design against actual, with a portfolio drill-down",
            f"4.0 Data QA rebuilt: {k} checks, model / expected / difference",
            f"anchors: 3.1 total r{a['g31']}, 3.2 total r{a['g32']}, "
            f"3.3 total r{a['g33']}"]


if __name__ == "__main__":
    import sys
    for x in run(sys.argv[1], sys.argv[2]):
        print("  ", x)
