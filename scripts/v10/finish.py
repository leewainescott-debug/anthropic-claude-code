"""The last presentation pass, from an independent review of the rendered tabs.

Every item here was found by looking at the sheets, not the cells, and several of them had
survived earlier passes because a test was subtly wrong rather than absent:

  the [Red] number-format strip tested for "[Red]" and the file writes "[RED]", so 240 cells
  across 0.2, 1.11, 1.12, 1.13 and REVIEW still printed negatives in red

  0.1 Budget Table (Fin) and 0.4 Presentation Pack are raw pastes from Finance and the deck -
  a full red / amber / green traffic-light grid, 741 [RED] formats, Arial and Aptos, four
  extra blues, hidden columns. They are the source and they stay in the file, but they cannot
  ship visible

  0.2 Data Config used the reserved subtotal grey as a zebra stripe, in patches, and one data
  row was 76pt tall against 14.25 for the rest, which dragged the second table's header band
  to a different depth from the first

  fifteen sentences were sitting in data cells, one of them the only Cambria cell in the
  workbook, and one of them a 122-character section label I had written myself
"""
import re

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill
from openpyxl.utils import get_column_letter as L

import opts

NONE_FILL = PatternFill()
NO_BORDER = Border()
HIDE = ["0.1 Budget Table (Fin)", "0.4 Presentation Pack"]
REVIEW = "REVIEW - Complete Role Mapping"

# sentences sitting in a data cell. Cleared, not reworded: the fact each one carried is
# either on the face of the model now or belongs in the decisions log.
PROSE = {"1.5 P&C": ["B11"], "1.11 BP&T": ["B9", "B47", "B49"],
         "1.12 SA&D": ["B9", "B53", "B55"], "1.2 Customer": ["L15", "L18"],
         "1.8 Energy Solutions & B2B": ["E17", "E18", "B21"],
         "0.2 Data Config": ["H9", "H10", "H14"], "1.13 Cyber Roles": ["B73"]}
# labels that were too long to sit on a row carrying figures
SHORTEN = {"3.1 Cost Bridge": {
    "Directly funded programmes and platforms - no archetype prices them, so the "
    "comparison is the amount funded on the 1.x tab":
        "Directly funded programmes and platforms - funded on the 1.x tab",
    "Overhead roles in the portfolios - the allowance is the comparison, line by line "
    "on 3.2": "Overhead roles in the portfolios - the allowance is on 3.2",
    "COEs and EGI - priced off the planned spend on their own 1.x tabs":
        "COEs and EGI - planned spend on their own 1.x tabs",
    "Squads priced by an archetype - detail on 3.3":
        "Squads priced by an archetype - detail on 3.3"}}
TITLES = {"3.2 Overhead & Leadership": "Overhead & Leadership - the allowance against "
                                       "what it costs",
          "1.13 Cyber Roles": "Cyber, Risk & Service Operations - roles and funding"}
# one width profile for the three COE design tabs, wide enough for all three
COE_W = {"B": 46, "C": 38, "D": 24, "E": 11, "F": 11, "G": 15, "H": 15, "I": 15, "J": 15,
         "K": 15}
COE_TABS = ["1.11 BP&T", "1.12 SA&D", "1.13 Cyber Roles"]


def no_red(wb):
    """The strip that missed 240 cells because Excel writes [RED], not [Red]."""
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                nf = c.number_format or ""
                if "[red]" not in nf.lower():
                    continue
                c.number_format = (opts.M2 if ".00" in nf else
                                   (opts.M3 if ".000" in nf else
                                    (opts.PCT if "%" in nf else opts.CT)))
                n += 1
    return [f"[RED] stripped from {n} number formats, case-insensitively this time"]


def hide_sources(wb):
    out = []
    for t in HIDE:
        if t in wb.sheetnames and wb[t].sheet_state == "visible":
            wb[t].sheet_state = "hidden"
            out.append(f"{t}: hidden - it is a raw paste, not a built tab")
    return out


def drop_prose(wb):
    n = 0
    for tab, cells in PROSE.items():
        if tab not in wb.sheetnames:
            continue
        ws = wb[tab]
        for ref in cells:
            if ws[ref].value is not None:
                ws[ref].value = None
                ws[ref].font = opts.BODY
                n += 1
    for tab, m in SHORTEN.items():
        ws = wb[tab]
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.strip() in m:
                    c.value = m[c.value.strip()]
    return [f"{n} sentences cleared from data cells, and the long section labels on 3.1 "
            f"cut to fit their row"]


def fix_02(wb):
    """0.2 Data Config: one row height, no zebra, a real total row."""
    ws = wb["0.2 Data Config"]
    out = []
    for r in range(1, 40):
        h = ws.row_dimensions[r].height
        if h and h > 50:
            ws.row_dimensions[r].height = 14.25
            out.append(f"0.2 row {r} height {h} -> 14.25")
    # every header row on the tab, sized by the same wrap arithmetic the builders use,
    # so the two overhead bands come out the same depth as each other and nothing clips
    for r in range(1, 30):
        cells = [c for c in range(2, 14) if _fill(ws.cell(r, c)) == opts.NAVY
                 and isinstance(ws.cell(r, c).value, str) and ws.cell(r, c).value.strip()]
        if len(cells) < 3:
            continue
        lines = 1
        for c in cells:
            ws.cell(r, c).alignment = opts.CEN
            w = ws.column_dimensions[L(c)].width or 8.43
            lines = max(lines, opts.wrap_lines(ws.cell(r, c).value, w))
        ws.row_dimensions[r].height = max(32, 14 * lines + 6)
    # the subtotal grey is reserved. Strip it from every row that is not a total.
    n = 0
    for r in range(6, 30):
        lab = str(ws.cell(r, 2).value or "").strip()
        istotal = lab.startswith(("Total", "Budget", "Variance"))
        for c in range(2, 8):
            x = ws.cell(r, c)
            if _fill(x) != opts.GREY:
                continue
            if istotal:
                continue
            x.fill = NONE_FILL
            n += 1
    # and the total row gets the block-total grey across the table
    for r in range(6, 30):
        if str(ws.cell(r, 2).value or "").strip() == "Total":
            for c in range(2, 8):
                ws.cell(r, c).fill = opts.fl(opts.MID)
                ws.cell(r, c).font = opts.BOLD
            out.append(f"0.2 row {r} is the total row, shaded as one")
    out.append(f"0.2: subtotal grey removed from {n} data cells")
    return out


def _fill(c):
    try:
        return str(c.fill.start_color.rgb or "").upper() if c.fill and c.fill.patternType \
            else ""
    except Exception:
        return ""


def bars_and_totals(wb):
    """Bars that stop short of their table, and total rows shaded in two greys."""
    out = []
    # a bar in column H over an H..K table, merged on 1.2 and not on 1.1
    ws = wb["1.1 Ampol Retail"]
    for r in range(1, 30):
        v = str(ws.cell(r, 8).value or "").strip()
        if v.startswith("Other funding") and _fill(ws.cell(r, 8)) == opts.BARC:
            ws.cell(r, 8).value = "Other funding"          # it had a trailing space
            for c in range(9, 12):
                ws.cell(r, c).fill = opts.fl(opts.BARC)
                ws.cell(r, c).font = opts.BARF
            out.append(f"1.1!H{r} bar extended across H..K")
    # three empty outlined boxes hanging off the end of 1.13's Roles bar
    ws = wb["1.13 Cyber Roles"]
    for r in range(1, 30):
        if str(ws.cell(r, 2).value or "").strip() == "Roles":
            for c in range(8, 12):
                x = ws.cell(r, c)
                if x.value is None:
                    x.border, x.fill = NO_BORDER, NONE_FILL
            out.append(f"1.13 row {r}: empty bordered cells past the bar cleared")
    # every total row in one grey, across the cells that carry the table
    for tab in [t for t in wb.sheetnames if re.match(r"^1\.\d+ ", t)]:
        ws = wb[tab]
        for r in range(1, min(ws.max_row, 90) + 1):
            greys = [c for c in range(2, 13)
                     if _fill(ws.cell(r, c)) in (opts.GREY, opts.MID)]
            if not greys:
                continue
            has = any(ws.cell(r, c).value is not None for c in range(2, 13))
            if not has:                                     # an empty shaded row
                for c in greys:
                    ws.cell(r, c).fill = NONE_FILL
                    ws.cell(r, c).border = NO_BORDER
                continue
            if len({_fill(ws.cell(r, c)) for c in greys}) > 1:
                for c in range(min(greys), max(greys) + 1):
                    ws.cell(r, c).fill = opts.fl(opts.MID)
    out.append("total rows shaded in one grey, empty shaded rows cleared")
    return out


def empty_inputs(wb):
    """Cream on an empty cell reads as a box waiting for a number that nothing wants."""
    n = 0
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        for row in ws.iter_rows():
            for c in row:
                if c.value is None and _fill(c) == opts.YEL:
                    # keep it where the row it sits on carries a label; otherwise it is
                    # an orphan
                    if not any(isinstance(ws.cell(c.row, k).value, str)
                               and ws.cell(c.row, k).value.strip()
                               for k in range(2, c.column)):
                        c.fill, c.border = NONE_FILL, NO_BORDER
                        n += 1
    return [f"{n} orphan cream cells cleared"]


def review_font(wb):
    """Calibri everywhere a reader can see. The strays were the ledger's header row, two
    notes in its spare column, and four empty cells that had kept the theme default."""
    n = 0
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        for row in ws.iter_rows():
            for c in row:
                f = c.font
                if f and (f.name or opts.FN) != opts.FN:
                    c.font = Font(name=opts.FN, size=max(f.size or 11, 10), bold=f.bold,
                                  color=(f.color.rgb if f.color and isinstance(
                                      getattr(f.color, "rgb", None), str) else None))
                    n += 1
    return [f"{n} cells set to Calibri across every visible tab"]


def _unused(wb):
    ws = wb[REVIEW]
    n = 0
    for c in ws[1]:
        f = c.font
        if f and (f.name or "") != opts.FN:
            c.font = Font(name=opts.FN, size=max(f.size or 11, 10), bold=f.bold,
                          color=(f.color.rgb if f.color and
                                 isinstance(getattr(f.color, "rgb", None), str)
                                 else None))
            n += 1
    return [f"{n} header cells on the ledger set to Calibri"]


def coe_widths(wb):
    for t in COE_TABS:
        for k, v in COE_W.items():
            wb[t].column_dimensions[k].width = v
    return [f"one width profile across {', '.join(COE_TABS)}"]


def titles(wb):
    n = 0
    for tab, t in TITLES.items():
        if tab in wb.sheetnames:
            wb[tab].cell(2, 2).value = t
            wb[tab].cell(2, 2).font = opts.TITLE
            n += 1
    return [f"{n} titles matched to their tab name"]


def strays(wb):
    out = []
    # a naked ratio with no label, in an otherwise blank area
    for tab, ref in (("1.1 Ampol Retail", "E18"), ("1.9 Commercial Fuels", "E17"),
                     ("1.10 Z Retail", "E16")):
        ws = wb[tab]
        if isinstance(ws[ref].value, str) and ws[ref].value.startswith("="):
            ws[ref].value = None
            out.append(f"{tab}!{ref} unlabelled ratio cleared")
    return out


def review_cream(wb):
    """The input colour on the ledger marks cells nobody types into.

    `cream()` recolours every bright-yellow cell in the live model, and REVIEW had four -
    including a column header. The one place an input genuinely belongs on the ledger is the
    cost-override column, so that is the only place it is kept.
    """
    ws = wb[REVIEW]
    n = 0
    for row in ws.iter_rows():
        for c in row:
            if _fill(c) == opts.YEL and c.column != 47:
                c.fill = NONE_FILL
                n += 1
    return [f"{n} stray input-coloured cells cleared from the ledger"]


def align_3x(wb):
    """The four summary tabs are meant to read as one family; their label columns were
    34 / 26 / 22 / 16 wide, so the left edge of the table landed somewhere different on
    every one of them."""
    n = 0
    for t in [x for x in wb.sheetnames if re.match(r"^3\.\d ", x)]:
        wb[t].column_dimensions["B"].width = 34
        n += 1
    return [f"column B set to 34 on all {n} summary tabs, so the left edge lines up"]


def qa_bar(wb):
    """4.0 was the only built tab with a header row and no bar over it."""
    ws = wb["4.0 Data QA"]
    if _fill(ws.cell(3, 2)) != opts.BARC:
        opts.bar(ws, 3, 2, 4, "Every difference must read zero")
    return ["4.0 given a section bar, like every other built tab"]


def run(src, dst):
    wb = openpyxl.load_workbook(src)
    out = (no_red(wb) + hide_sources(wb) + drop_prose(wb) + fix_02(wb)
           + bars_and_totals(wb) + empty_inputs(wb) + review_font(wb) + coe_widths(wb)
           + titles(wb) + strays(wb) + review_cream(wb) + align_3x(wb) + qa_bar(wb))
    wb.save(dst)
    return out


if __name__ == "__main__":
    import sys
    for x in run(sys.argv[1], sys.argv[2]):
        print("  ", x)
