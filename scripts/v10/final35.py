"""3.4 COE detail and 3.5 Source Reconciliation, on the same layout as the rest of 3.x.

3.4 carried two definitions of the same number - column F was literally `=$K6`, headed
"People cost, gross ($m)" against K's "Gross people cost ($m)" - and priced the COEs
against a "budget to draw down" the owner ruled out: 12.00 against 27.77 actual is not a
comparison, it is two unrelated figures on one row.

What 3.4 is for now is the thing no other tab shows: where COE cost sits geographically,
and which COE roles carry an overhead title. It reads the same squad grouping as every
other tab rather than a second cut built off the department column.

3.5 reconciles the retired Squads tab to REVIEW. It stays as evidence; it was the only
visible tab still showing Excel gridlines, and its headers used a Greek delta.
"""
import json

import openpyxl
from openpyxl.utils import get_column_letter as L

import final2x as f2
import opts

REV = f2.REV
LAST = f2.LAST
S = f2.S
COE_ORDER = ["COE Cyber", "COE BP&T", "COE SA&D", "EGI"]

H34 = ["COE", "Squad", "Roles", "Filled", "Vacant", "Cost ($m)", "Cost - AU ($m)",
       "Cost - NZ ($m)", "Cost - elsewhere ($m)", "Roles carrying an overhead title",
       "Cost of those roles ($m)"]
W34 = [16, 34, 8, 8, 8, 13, 13, 13, 15, 15, 15]
F34 = [None, None, opts.CT, opts.CT, opts.CT, opts.M2, opts.M2, opts.M2, opts.M2,
       opts.CT, opts.M2]
SUMS = (4, 5, 6, 7, 8, 9, 10, 11, 12)


def build_34(wb, anchors):
    ws = wb["3.4 COE Summary"]
    f2.wipe(ws)
    ws.column_dimensions["A"].width = 2
    ws.cell(2, 2).value = "COE detail - where the cost sits"
    ws.cell(2, 2).font = opts.TITLE
    by = {a["pf"]: a for a in anchors.values()}
    r = opts.bar(ws, 4, 2, len(H34), "The COEs and EGI, squad by squad")
    r = opts.head(ws, r, 2, H34, W34)
    subs = []
    for pf in COE_ORDER:
        a = by[pf]
        st = r
        for sq in a["direct"] + a["squads"] + a["overhead"]:
            base = (f'{REV}!$AA$2:$AA${LAST},{REV}!$AJ$2:$AJ${LAST},$B{r},'
                    f"{REV}!$AT$2:$AT${LAST},$C{r}")
            cnt = (f'{REV}!$AJ$2:$AJ${LAST},$B{r},{REV}!$AT$2:$AT${LAST},$C{r}')
            for c, v in ((2, pf), (3, sq)):
                x = ws.cell(r, c)
                x.value, x.font, x.alignment = v, opts.BODY, opts.LFT
            f2._m(ws, r, 4, f"=COUNTIFS({cnt})", opts.CT)
            f2._m(ws, r, 5, f'=COUNTIFS({cnt},{REV}!$AK$2:$AK${LAST},"Filled")', opts.CT)
            f2._m(ws, r, 6, f'=COUNTIFS({cnt},{REV}!$AK$2:$AK${LAST},"Vacant")', opts.CT)
            f2._m(ws, r, 7, f"=SUMIFS({base})/1000000")
            f2._m(ws, r, 8,
                  f'=SUMIFS({base},{REV}!$M$2:$M${LAST},"Australia")/1000000')
            f2._m(ws, r, 9, f'=SUMIFS({base},{REV}!$AL$2:$AL${LAST},"NZ")/1000000')
            # everything that is neither Australia nor mapped to NZ. The owner's AU/NZ
            # column maps a Singapore role to AU, so this column is the only place the
            # cost outside both countries is visible.
            f2._m(ws, r, 10, f"=$G{r}-$H{r}-$I{r}")
            f2._m(ws, r, 11,
                  f'=COUNTIFS({cnt},{REV}!$AR$2:$AR${LAST},"<>Squad")', opts.CT)
            f2._m(ws, r, 12,
                  f'=SUMIFS({base},{REV}!$AR$2:$AR${LAST},"<>Squad")/1000000')
            r += 1
        opts.row(ws, r, 2, [f"{pf} total"] + [None] * (len(H34) - 1),
                 [None] * len(H34), bg=opts.GREY, bold=True)
        ws.cell(r, 2).alignment = opts.LFT
        for c in SUMS:
            x = ws.cell(r, c)
            x.value = f"=SUM({L(c)}{st}:{L(c)}{r-1})"
            x.number_format, x.alignment = F34[c - 2], opts.RGT
        subs.append(r)
        r += 1
    opts.row(ws, r, 2, ["COEs and EGI total"] + [None] * (len(H34) - 1),
             [None] * len(H34), bg=opts.MID, bold=True, top=True)
    ws.cell(r, 2).alignment = opts.LFT
    for c in SUMS:
        x = ws.cell(r, c)
        x.value = "=" + "+".join(f"{L(c)}{p}" for p in subs)
        x.number_format, x.alignment = F34[c - 2], opts.RGT
    gt = r
    r += 2
    for lab, col, f, nf in (
            ("Control - roles against the ledger, must be 0", 4,
             "=$D%d-(%s)" % (gt, "+".join(
                 f'COUNTIFS({REV}!$AJ$2:$AJ${LAST},"{p}")' for p in COE_ORDER)),
             opts.CTL_C),
            ("Control - cost against the ledger ($m), must be 0", 7,
             "=ROUND($G%d-(%s)/1000000,6)" % (gt, "+".join(
                 f'SUMIFS({REV}!$AA$2:$AA${LAST},{REV}!$AJ$2:$AJ${LAST},"{p}")'
                 for p in COE_ORDER)), opts.CTL_M),
            ("Control - AU plus NZ plus elsewhere against cost ($m), must be 0", 10,
             f"=ROUND($H{gt}+$I{gt}+$J{gt}-$G{gt},6)", opts.CTL_M)):
        ws.cell(r, 2).value = lab
        ws.cell(r, 2).font = opts.BODY
        ws.cell(r, 2).alignment = opts.LFT
        f2._m(ws, r, col, f, nf)
        r += 1
    ws.freeze_panes = "D6"
    return {"total": gt}


H35 = ["Portfolio", "Squads roles", "Squads filled", "Squads vacant", "REVIEW roles",
       "REVIEW filled", "REVIEW vacant", "Difference - roles", "Difference - filled",
       "Difference - vacant"]
W35 = [34, 12, 12, 12, 12, 12, 12, 13, 13, 13]


def build_35(wb):
    """Relabel and restyle. The population logic is left exactly as it is - it is the
    evidence that the retired Squads tab was replaced by REVIEW, and rewriting it would
    destroy the thing it exists to prove."""
    ws = wb["3.5 Source Reconciliation"]
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.cell(2, 2).value = "Source Reconciliation - the retired Squads tab against REVIEW"
    ws.cell(2, 2).font = opts.TITLE
    ws.cell(3, 2).value = None                       # a sentence in a data area
    opts.bar(ws, 4, 2, len(H35), "Where the two sources disagree, role by role")
    opts.head(ws, 5, 2, H35, W35)
    last = 20
    for r in range(6, last + 1):
        band = opts.GREY if r % 2 == 1 else None
        for c in range(2, 12):
            x = ws.cell(r, c)
            x.border = opts.BOX
            x.font = opts.BOLD if r == last else opts.BODY
            x.alignment = opts.LFT if c == 2 else opts.RGT
            if c > 2:
                x.number_format = opts.CT
            if r == last:
                x.fill = opts.fl(opts.MID)
            elif band:
                x.fill = opts.fl(band)
    ws.freeze_panes = "C6"
    return last


def run(src, dst):
    wb = openpyxl.load_workbook(src)
    anchors = json.load(open("anchors_final.json"))
    a34 = build_34(wb, anchors)
    n = build_35(wb)
    wb.save(dst)
    return [f"3.4 COE detail rebuilt: squad by squad, AU / NZ / elsewhere, "
            f"overhead titles, total row {a34['total']}",
            f"3.5 Source Reconciliation restyled to {n} rows, gridlines off, "
            "plain-English headers"]


if __name__ == "__main__":
    import sys
    for x in run(sys.argv[1], sys.argv[2]):
        print("  ", x)
