"""Logic fixes on the design tabs. No column is added or removed and no squad is repriced.

1.12 SA&D's roles list was three roles short of the ledger. Its own control cell said so -
"Check - roles listed vs counted (must be 0)" reading -3 - and nothing else in the file
noticed, because 4.0 tested roles and cost against the ledger and never read a control
cell on another tab. Two of the three are the roles DECISIONS D7 moved into COE SA&D; the
third was never written. The counts and the planned spend on that tab read the whole ledger
directly and were always right, but the AU / NZ split sums the roles list, so the split came
to 5.7816 against a planned spend of 6.5295 on the same row.

The rest of that gap is the offshore toggle. 1.11 and 1.12 carried a note saying that
setting a role to Offshore prices it at 40% and that "the totals above and every summary
follow". They do not: the toggle feeds the AU / NZ split only. Rather than rewire the
owner's tab, the discount is now stated on its own line, so the split adds back to planned
spend on the face of the page, and the note says where the decision that moves cost is
actually made.

Three platform blocks were empty shells: a navy platform bar, a full column-header row, no
squad row, a blank overhead line and a grey total reading zero. One of them, EGI Data on
1.3, DECISIONS D6a records as removed. They are collapsed to the single line that carries
the fact, so the tabs stop showing tables that are not finished.
"""
import re

import openpyxl
from openpyxl.styles import Border, PatternFill
from openpyxl.utils import get_column_letter as L

import opts

REVIEW = "REVIEW - Complete Role Mapping"
REV = f"'{REVIEW}'"
LAST = 528
NONE_FILL = PatternFill()
NO_BORDER = Border()

SHELLS = [("1.1 Ampol Retail", "Platform: Pricing & WFM"),
          ("1.3 Enterprise Data", "Platform: EGI Data"),
          ("1.5 P&C", "Platform: EGI P&C")]
SHELL_NOTE = {"Platform: Pricing & WFM": "Platform: Pricing & WFM - combined into Above "
                                         "Store, no squad of its own",
              "Platform: EGI Data": "Platform: EGI Data - removed, no roles",
              "Platform: EGI P&C": "Platform: EGI P&C - the one role is management and "
                                   "sits on the P&C management line"}
# tab, the portfolio its roles list covers, the list's first row
COE = [("1.11 BP&T", "COE BP&T"), ("1.12 SA&D", "COE SA&D"),
       ("1.13 Cyber Roles", "COE Cyber")]
ROWREF = re.compile(r"\$([A-Z]{1,2})\$(\d+)")


def block_rows(ws, label, limit=90):
    start = None
    for r in range(1, min(ws.max_row, limit) + 1):
        v = str(ws.cell(r, 2).value or "").strip()
        if start is None:
            if v.startswith(label):
                start = r
            continue
        if v.startswith(("Platform:", "Total ", "Reconciled")):
            return start, r - 1
        if v.endswith(" Total"):
            return start, r
    return (start, min(ws.max_row, limit)) if start else (None, None)


def collapse_shells(wb, wv):
    out = []
    for tab, label in SHELLS:
        ws, wsv = wb[tab], wv[tab]
        lo, hi = block_rows(ws, label)
        if lo is None:
            out.append(f"{tab}: no {label!r} block")
            continue
        # a real block has a squad row: a name that is not a header, an overhead label or
        # a total, carrying a cost in column H
        real = any(
            (n := str(wsv.cell(r, 2).value or "").strip())
            and not n.startswith(("Platform Overhead", "Squad"))
            and not n.endswith(" Total")
            and isinstance(wsv.cell(r, 8).value, (int, float))
            and wsv.cell(r, 8).value
            for r in range(lo + 1, hi + 1))
        if real:
            out.append(f"{tab}: {label!r} has a squad in it, left alone")
            continue
        for m in [m for m in list(ws.merged_cells.ranges)
                  if m.min_row >= lo and m.max_row <= hi]:
            ws.unmerge_cells(str(m))
        for r in range(lo, hi + 1):
            for c in range(2, 13):
                x = ws.cell(r, c)
                x.value = None
                x.fill, x.border, x.font = NONE_FILL, NO_BORDER, opts.BODY
            ws.row_dimensions[r].height = None
        x = ws.cell(lo, 2)
        x.value, x.font, x.alignment = SHELL_NOTE[label], opts.BOLD, opts.LFT
        out.append(f"{tab}: {label!r} collapsed to one line, rows {lo}-{hi}")
    return out


def list_bounds(ws):
    """The roles list: the header row carrying 'Name', and the last row its formulas
    reference. The SUMIFS windows above run to row 50, so the list can grow into it."""
    hdr = next((r for r in range(1, ws.max_row + 1)
                if str(ws.cell(r, 2).value or "").strip() == "Name"), None)
    if hdr is None:
        return None, None, None
    last = hdr
    for r in range(hdr + 1, ws.max_row + 1):
        if isinstance(ws.cell(r, 2).value, str) and ws.cell(r, 2).value.startswith("="):
            last = r
        elif str(ws.cell(r, 2).value or "").strip():
            break
    win = 50
    for row in ws.iter_rows(min_row=1, max_row=hdr):
        for c in row:
            if isinstance(c.value, str) and "$50" in c.value:
                win = 50
    return hdr, last, win


def extend_lists(wb, lg):
    """Write into each COE roles list any role the ledger has and the list does not."""
    out = []
    for tab, pf in COE:
        ws = wb[tab]
        hdr, last, win = list_bounds(ws)
        if hdr is None:
            out.append(f"{tab}: no roles list found")
            continue
        listed = set()
        for r in range(hdr + 1, last + 1):
            v = ws.cell(r, 2).value
            if isinstance(v, str):
                m = ROWREF.search(v)
                if m:
                    listed.add(int(m.group(2)))
        want = [i for i in range(2, LAST + 1)
                if str(lg.cell(i, 36).value or "") == pf]
        missing = [i for i in want if i not in listed]
        if not missing:
            out.append(f"{tab}: roles list complete, {len(listed)} roles")
            continue
        if last + len(missing) > win:
            out.append(f"{tab}: {len(missing)} roles will not fit before row {win}")
            continue
        cols = [c for c in range(2, 30) if ws.cell(last, c).value is not None]
        r = last + 1
        for i in missing:
            for c in cols:
                new, old = ws.cell(r, c), ws.cell(last, c)
                new._style = old._style
                new.number_format = old.number_format
                v = old.value
                if isinstance(v, str) and v.startswith("="):
                    # the reference carries the REVIEW row; the local row number moves too
                    v = ROWREF.sub(
                        lambda m: f"${m.group(1)}${i}" if int(m.group(2)) in listed
                        or int(m.group(2)) > 200 else m.group(0), v)
                    v = re.sub(r"(?<![$\d])" + str(last) + r"(?![\d])", str(r), v)
                new.value = v
            r += 1
        out.append(f"{tab}: {len(missing)} roles written in "
                   f"(REVIEW rows {', '.join(str(i) for i in missing)})")
    return out


def offshore_line(wb):
    """State the offshore discount, so the AU / NZ split adds back to planned spend.

    Only where the roles list carries an On/Off column. 1.13 has none, so a discount line
    there would be a permanent zero under a heading that implies a mechanic.
    """
    out = []
    for tab, _ in COE:
        ws = wb[tab]
        hdr, _, _ = list_bounds(ws)
        if hdr is None or not any(
                str(ws.cell(hdr, c).value or "").strip() in ("On/Off",
                                                             "Onshore / Offshore")
                for c in range(2, 14)):
            out.append(f"{tab}: no On/Off column, no discount line needed")
            continue
        tot = next((r for r in range(1, 30)
                    if str(ws.cell(r, 2).value or "").strip() == "Total"), None)
        if tot is None:
            out.append(f"{tab}: no Total row on the summary block")
            continue
        # the planned spend and the two split columns, found by their headers
        hdr = tot - 3
        while hdr > 1 and str(ws.cell(hdr, 2).value or "").strip() != "Grouping":
            hdr -= 1
        cols = {}
        for c in range(3, 14):
            v = str(ws.cell(hdr, c).value or "").strip()
            if v.startswith("Planned spend"):
                cols["spend"] = c
            elif v.startswith("Cost - AU"):
                cols["au"] = c
            elif v.startswith("Cost - NZ"):
                cols["nz"] = c
        if len(cols) < 3:
            out.append(f"{tab}: could not find the spend and split columns")
            continue
        r = tot + 1
        while str(ws.cell(r, 2).value or "").strip():
            r += 1
        x = ws.cell(r, 2)
        x.value = "Offshore discount in the AU / NZ split ($m)"
        x.font, x.alignment = opts.BODY, opts.LFT
        y = ws.cell(r, cols["spend"])
        y.value = (f"={L(cols['spend'])}{tot}-{L(cols['au'])}{tot}"
                   f"-{L(cols['nz'])}{tot}")
        y.font, y.number_format, y.alignment = opts.BODY, opts.M2, opts.RGT
        out.append(f"{tab}: offshore discount stated at {L(cols['spend'])}{r}")
    return out


NOTE_FIX = {
    "On/Off: set a role to Offshore and it is priced at 40% of the onshore cost. The "
    "totals above and every summary follow.":
        "On/Off discounts the AU / NZ split only, and the discount is shown on its own "
        "line. The decision that moves cost and every summary is made on the working tab.",
    "set a role to Offshore and it is priced at 40% of the onshore cost. The totals "
    "above and every summary follow.":
        "the AU / NZ split is discounted and the discount is shown on its own line. The "
        "decision that moves cost is made on the working tab.",
    # 38 characters in an 11-wide column: it rendered as "spend - budg" clipped at both
    # ends on all three COE tabs
    "Left to fund (spend - budget, + = over)": "Left to fund ($m, + = over budget)",
    "2.0 Group Summary": "3.1 Group Summary",
    "(1.3 / 4.3)": "(1.3 / 2.3)",
    "the yellow cell": "the cream cell",
    "yellow cell": "cream cell",
}


def fix_notes(wb):
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not isinstance(v, str) or v.startswith("="):
                    continue
                new = v
                for old, rep in NOTE_FIX.items():
                    if old in new:
                        new = new.replace(old, rep)
                if new != v:
                    c.value = new
                    n += 1
    return [f"{n} labels corrected - the offshore claim, two stale tab references and "
            f"the yellow-cell wording the recolour made wrong"]


def run(src, dst, ledger="w1r.xlsx"):
    wb = openpyxl.load_workbook(src)
    wv = openpyxl.load_workbook(src, data_only=True)
    lg = openpyxl.load_workbook(ledger, data_only=True)[REVIEW]
    out = (extend_lists(wb, lg) + offshore_line(wb) + collapse_shells(wb, wv)
           + fix_notes(wb))
    wb.save(dst)
    return out


if __name__ == "__main__":
    import sys
    for x in run(*sys.argv[1:]):
        print("  ", x)
