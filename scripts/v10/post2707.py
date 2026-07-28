"""The owner's 2707 adoptions that need the finished blocks in place.

Four jobs, each one of his edits generalised the way he applied it:

1. The Actuals column. On 1.7, 1.8 and 1.9 he added "Actuals" to the Portfolio Summary
   table, reading the tab's own actual-after-decisions total, with a "Variance to
   actuals" line under the block (summary Total minus Actuals). Applied to all ten
   portfolio tabs, wired by label to the block actuals.py just built - by label, and to
   the whole tab rather than to column B, because that block has since moved from the foot
   of the tab to the top and been redrawn from five lines to three, and the wiring had to
   follow it both times without being told where it went.
2. His live levers. The COE design tabs now carry his On/Off states (Hold and Offshore),
   and 2.7 carries two hand-set levers. The working tabs' lever cells are set to match,
   by person, so his decisions price through the whole cascade.
3. The 0.2 COE spend cells. His rev formulas read the old '3.4 COE Summary'; the intended
   figures are the COE groupings' net planned spend, which live on the COE tabs
   themselves. Each is repointed at the cell whose cached value his own file shows.
4. The 1.13 bar, extended over his new On/Off column - done here because the bar
   normaliser earlier in the chain repaints it to the old width.
"""
import json
import re

import openpyxl
from openpyxl.utils import get_column_letter as L

import opts

REVIEW = "REVIEW - Complete Role Mapping"
# the line of the 1.x actuals table this column quotes, and the head of the column its
# figure sits in. Both are the table's own words, which is the only part of that block that
# has stayed put through two redesigns and one move from the foot of the tab to the top.
ACTUALS_ROW = "Actual portfolio"
COST_HEAD = "Cost ($m)"


# ---------------------------------------------------------------- 1. the Actuals column
def actuals_column(wb, wv, out):
    a = json.load(open("anchors_final.json"))
    for one in [t for t in wb.sheetnames if re.match(r"^1\.(10|14|[1-9]) ", t)]:
        ws, wsv = wb[one], wv[one]
        # The actuals table's own actual line and the column its cost sits in, both found by
        # what they say rather than by where they are. The table used to be at the foot of
        # the tab with its labels in column B; it is now at the top beside the budget box
        # with them in K, and its five decomposition lines are now three - Actual portfolio,
        # Archetype portfolio, Variance - because the owner mocked it that way. Wiring by
        # label is what let it move and then change shape without this step being rewritten
        # twice, and is why the scan is over every column rather than over column B.
        foot = lab_col = act_col = None
        for r in range(1, ws.max_row + 1):
            for c in range(2, 21):
                if str(ws.cell(r, c).value or "").strip() == ACTUALS_ROW:
                    foot, lab_col = r, c
        if foot is None:
            out.append(f"{one}: no {ACTUALS_ROW!r} line on the actuals table - skipped")
            continue
        # the cost column is the one the table's own header calls Cost ($m), read upwards
        # from the actual line; the old rule - the rightmost cell on the row that reads a
        # working tab - is kept as the fallback if that head is ever reworded
        for k in range(foot - 1, max(foot - 8, 0), -1):
            hit = next((c for c in range(lab_col, 21)
                        if str(ws.cell(k, c).value or "").strip() == COST_HEAD), None)
            if hit:
                act_col = hit
                break
        if act_col is None:
            for c in range(lab_col + 1, 30):
                v = ws.cell(foot, c).value
                if isinstance(v, str) and v.startswith("=") and "'2." in v:
                    act_col = c
        # the Portfolio Summary table: bar, header row, Total Cost row
        bar = hdr = tot = None
        for r in range(1, 20):
            b = str(wsv.cell(r, 2).value or "").strip()
            if b == "Portfolio Summary":
                bar = r
            elif b == "Cost" and bar and r > bar:
                hdr = r
            elif b == "Total Cost" and hdr:
                tot = r
                break
        if not (hdr and tot and act_col):
            out.append(f"{one}: no Portfolio Summary table - skipped")
            continue
        if bar:
            ws.cell(bar, 7).fill = opts.fl(opts.BARC)
            ws.cell(bar, 7).font = opts.BARF
        h = ws.cell(hdr, 7)
        h.value = "Actuals"
        h.font, h.fill, h.alignment, h.border = (opts.HDRF, opts.fl(opts.NAVY),
                                                 opts.CEN, opts.BOX)
        x = ws.cell(tot, 7)
        x.value = f"=${L(act_col)}${foot}"
        x.number_format, x.alignment, x.font, x.border = (opts.M2, opts.RGT, opts.BOLD,
                                                          opts.BOX)
        # the variance line, on the first free row under the summary block
        placed = False
        for r in range(tot + 1, tot + 9):
            if all(wsv.cell(r, c).value is None and wb[one].cell(r, c).value is None
                   for c in (5, 6, 7)):
                ws.cell(r, 5).value = "Variance to actuals"
                ws.cell(r, 5).font = opts.BOLD
                ws.cell(r, 5).alignment = opts.LFT
                v = ws.cell(r, 6)
                v.value = f"=$F${tot}-$G${tot}"
                v.number_format, v.alignment, v.font = opts.M2, opts.RGT, opts.BOLD
                placed = True
                break
        out.append(f"{one}: Actuals column on the summary (G{tot} <- {L(act_col)}{foot}, "
                   f"the actuals table's {ACTUALS_ROW} cost)"
                   f"{', variance line placed' if placed else ''}")


# ---------------------------------------------------------------- 2. his live levers
# The ledger row behind an FTE line, recovered from the formula that names the person.
# Two shapes are live: the direct reference the 2.x tabs now use - 'REVIEW…'!$B$36 - and
# the older INDEX($B:$B,36) form, which is still what a workbook built before the join
# was made insert-safe will carry. Matching both is what stops this step going quietly
# blind and dropping ten of the owner's levers.
_NAME_REF = re.compile(r"'REVIEW - Complete Role Mapping'!\$B(?::\$B,|\$)(\d+)")


def _fte_rows(wb, tab):
    """REVIEW row -> FTE-block row, read out of the name formulas.

    The name cells are live INDEX formulas into the ledger, so the ledger row each one
    points at is in the formula text - the one join that survives every rename and every
    cache strip.
    """
    ws = wb[tab]
    rows = {}
    for r in range(1, ws.max_row + 1):
        f = ws.cell(r, 2).value
        if isinstance(f, str):
            m = _NAME_REF.search(f)
            if m and ws.cell(r, 5).value is not None:
                rows.setdefault(int(m.group(1)), []).append(r)
    return rows


def _design_levers(wb, wv, tab, hdr_col=8):
    """(REVIEW row, state) parsed from a COE design tab's cost-engine formulas."""
    ws = wb[tab]
    out = []
    pat = re.compile(r"\$AA\$?(\d+)")
    for r in range(1, ws.max_row + 1):
        f = ws.cell(r, 20).value                        # column T, the cost engine
        state = str(wv[tab].cell(r, hdr_col).value or "").strip()
        if isinstance(f, str) and REVIEW in f and state in ("Hold", "Offshore"):
            m = pat.search(f)
            if m:
                out.append((int(m.group(1)), state))
    return out


def sync_levers(wb, wv, out):
    R = wv[REVIEW]
    # design tab -> its working tab
    for dtab, wtab in (("1.11 BP&T", "2.12 COE BP&T"), ("1.12 SA&D", "2.13 COE SA&D"),
                       ("1.13 Cyber Roles", "2.11 COE Cyber")):
        pairs = _design_levers(wb, wv, dtab)
        fte = _fte_rows(wb, wtab)
        n = 0
        for rev_row, state in pairs:
            for r in fte.get(rev_row, []):
                if str(wb[wtab].cell(r, 5).value or "") != state:
                    wb[wtab].cell(r, 5).value = state
                    n += 1
                break
        out.append(f"{wtab}: {n} levers set from {dtab} ({len(pairs)} states on the tab)")
    # his two hand-set levers on 2.7, found by person in the ledger
    fte = _fte_rows(wb, "2.7 Infrastructure")
    targets = []
    # the portfolio column is a formula whose cache does not survive an openpyxl save,
    # so membership of 2.7 is proven by the FTE block itself: only 2.7's own ledger rows
    # are in its name-formula map
    for i in range(2, R.max_row + 1):
        nm = str(R.cell(i, 2).value or "").strip().lower()
        ro = str(R.cell(i, 3).value or "").strip().lower()
        if nm == "stevani kho" and ro == "delivery lead":
            targets.append((i, "Offshore", "Stevani Kho, Delivery Lead"))
        elif i == 431 and nm == "vacant" and ro == "quality assurance":
            # his 2707 set exactly one of Infrastructure's vacant QA roles - the name
            # formula on his own edited row points at ledger row 431
            targets.append((i, "Filled", "the vacant Quality Assurance role (r431)"))
    for rev_row, state, who in targets:
        rows = fte.get(rev_row, [])
        if rows:
            wb["2.7 Infrastructure"].cell(rows[0], 5).value = state
            out.append(f"2.7 Infrastructure: {who} -> {state}")
        else:
            out.append(f"2.7 Infrastructure: {who} (ledger row {rev_row}) not in the "
                       f"FTE block - NOT set")


# ---------------------------------------------------------------- 3. the 0.2 COE cells
# His rev formulas read the retired '3.4 COE Summary', whose F column was planned spend
# per COE grouping. Those figures live on the COE tabs' own grouping rows, so each cell
# is wired straight to its source and moves with the data - the mapping is his own, read
# off the old 3.4's labels: F6 Strategy & Architecture, F8 Transformation, F9 Business
# Partnering, F10 Data.
CFG_COE = {6: "='1.12 SA&D'!$G$6", 8: "='1.11 BP&T'!$F$7",
           9: "='1.11 BP&T'!$F$6", 10: "='1.12 SA&D'!$G$7"}


def repoint_02(wb, wv, rev_path, out):
    ws = wb["0.2 Data Config"]
    for r, f in CFG_COE.items():
        cur = ws.cell(r, 6).value
        if isinstance(cur, str) and "'3.4 COE " in cur:
            ws.cell(r, 6).value = f
            out.append(f"0.2!F{r} -> {f[1:]} (the grouping's planned spend, live)")
        else:
            out.append(f"0.2!F{r} holds {str(cur)[:40]!r} - not the 3.4 read, left alone")


# ---------------------------------------------------------------- 4. the 1.13 bar
def cyber_bar(wb, out):
    ws = wb["1.13 Cyber Roles"]
    for r in range(1, 30):
        if str(ws.cell(r, 2).value or "").strip() == "Roles" \
                and (ws.cell(r, 2).fill.patternType or ws.cell(r, 3).fill.patternType):
            for c in range(2, 9):
                ws.cell(r, c).fill = opts.fl(opts.BARC)
                ws.cell(r, c).font = opts.BARF
            out.append(f"1.13!B{r} bar extended over the On/Off column")
            return


def run(src, dst, rev_path="rev.xlsx"):
    wb = openpyxl.load_workbook(src)
    before = {t: {} for t in wb.sheetnames}
    for t in wb.sheetnames:
        for row in wb[t].iter_rows():
            for c in row:
                if c.value is not None:
                    before[t][c.coordinate] = c.value
    wv = openpyxl.load_workbook(src, data_only=True)
    out = []
    actuals_column(wb, wv, out)
    sync_levers(wb, wv, out)
    repoint_02(wb, wv, rev_path, out)
    cyber_bar(wb, out)
    # the manifest of every cell this step wrote, so the shipped-workbook diff can tell a
    # declared edit from an accident
    touched = []
    for t in wb.sheetnames:
        for row in wb[t].iter_rows():
            for c in row:
                if c.value != before[t].get(c.coordinate):
                    touched.append([t, c.coordinate])
    json.dump(touched, open("post2707_manifest.json", "w"))
    out.append(f"{len(touched)} cells declared in post2707_manifest.json")
    wb.save(dst)
    return out


if __name__ == "__main__":
    import sys
    for x in run(sys.argv[1], sys.argv[2]):
        print("  ", x)
