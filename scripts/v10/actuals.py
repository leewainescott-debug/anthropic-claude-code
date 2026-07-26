"""Actual cost against archetype cost, on the 1.x design tabs. Two designs.

The design tabs price squads from the archetype library and never say what those squads
actually cost, so the comparison could only be read on 2.x or 3.x. This puts it on the design
tab itself, driven by formula off the working tab's cost-after-decisions column, so it moves
the moment a lever is pulled.

Design A - inline. Two columns are appended to every squad table and every platform total
row: the actual cost after decisions, and the variance to the archetype. You read the
comparison on the row you are already reading. Nothing moves; K and L were empty on every
squad table.

Design B - one comparison table. Every existing table is left exactly as it is, and one
"Archetype against actual" table is added at the foot of each tab: every squad on the tab,
grouped by platform, with platform subtotals, the portfolio total, and the overhead roles
stated separately so the figure ties back to the working tab.

Both put the same three facts on the portfolio block: what the archetype prices, what those
squads actually cost after decisions, and the difference. Neither touches a single existing
formula.
"""
import copy
import json
import re

import openpyxl
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter as L

import opts
from build_2xfix import DESIGN

NOTE_AL = Alignment(horizontal="left", vertical="top", wrap_text=True)

# 1.x tab -> its working tab
PAIR = {v: k for k, v in DESIGN.items()}
# Every column on a 1.x tab already belongs to a table. The squad tables run B:J, the budget
# block above them runs I:N, and setting a width for one shrinks a column of the other - C
# went 26 to 16 and truncated "Configuration / Integration", K went 22 to 20 and squeezed the
# budget block's own last two columns. Nothing here sets a width. The three figures go under
# H, K and L, which is where the archetype cost already sits and where the squad tables have
# nothing, so every figure of a kind stays in one column all the way down the tab.
ARCH, ACT, VAR = 8, 11, 12                          # H, K, L
LAST = 12                                           # the band runs B:L
A_HDR = {ACT: "Actual cost after decisions ($m)", VAR: "Variance to archetype ($m)"}


def blocks(ws, wsv, limit=95):
    """Every platform block on a design tab: its squad rows and its total row.

    A block is a "Platform: x" bar, a header row whose column B reads "Squad", the squad
    rows, an optional platform overhead line, and a total row. A block with no squad row -
    a platform folded into another, or one that was removed - is skipped.
    """
    out = []
    r = 1
    while r <= min(ws.max_row, limit):
        b = str(wsv.cell(r, 2).value or "").strip()
        if not b.startswith("Platform:"):
            r += 1
            continue
        name = b[9:].strip()
        hdr = r + 1
        if str(wsv.cell(hdr, 2).value or "").strip() != "Squad":
            r += 1
            continue
        squads, total = [], None
        k = hdr + 1
        while k <= min(ws.max_row, limit):
            v = str(wsv.cell(k, 2).value or "").strip()
            if v.endswith(" Total"):
                total = k
                break
            if v.startswith("Platform:"):
                break
            # a squad row carries a squad type. Testing for a cost instead skipped the
            # EGI P&C row on 1.5, whose funded input the owner has not set yet, and pushed
            # its 0.24 into the residual line as though the squad had no row at all.
            if v and v != "Platform Overhead" and \
                    str(wsv.cell(k, 3).value or "").strip():
                squads.append(k)
            k += 1
        if squads and total:
            # a block is comparable only if an archetype prices every squad in it. EGI P&C
            # on 1.5 has no size set, so its block has no archetype side at all and cannot
            # be added to one that has.
            comparable = all(isinstance(wsv.cell(s, 8).value, (int, float)) for s in squads)
            out.append({"name": name, "hdr": hdr, "squads": squads, "total": total,
                        "comparable": comparable})
        r = k + 1
    return out


def lookup(tab, lo, hi, row, after):
    """The squad's cost after decisions on its working tab, found by name."""
    return (f"=IFERROR(INDEX('{tab}'!${after}${lo}:${after}${hi},"
            f"MATCH($B{row},'{tab}'!$B${lo}:$B${hi},0)),\"-\")")


def diff(a, b):
    """b less a, stated only when both sides are figures.

    EGI P&C on 1.5 carries no archetype - the owner has not set its size - so K-H read the
    whole 0.24 as an overspend against nothing. A variance needs two figures on the same
    basis or it is not a variance, and the honest answer is a dash.
    """
    return f'=IF(AND(ISNUMBER({a}),ISNUMBER({b})),ROUND({b}-{a},6),"-")'


def priced(cells):
    """TRUE when every squad in a block carries an archetype cost."""
    return f"COUNT({','.join(cells)})={len(cells)}"


def _span(ws, r, spec, font, fill, al, height=None):
    """One row of a table, merged where a label needs more than its own column.

    Widths are never touched, so a label that will not fit in one column is given two or
    five. The bands run B:L on every table here, which is what the squad tables above
    already span once design A has added its two columns.
    """
    for c0, c1, text in spec:
        if c1 > c0:
            ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c1)
        x = ws.cell(r, c0)
        x.value = text
        for c in range(c0, c1 + 1):
            y = ws.cell(r, c)
            y.font, y.border = font, opts.BOX
            if fill:
                y.fill = opts.fl(fill)
        x.alignment = al
    if height:
        ws.row_dimensions[r].height = height
    return r + 1


def _head(ws, r, spec):
    """A header row that widens nothing and grows its own height instead."""
    lines = max(opts.wrap_lines(t, sum(ws.column_dimensions[L(c)].width or 8.43
                                       for c in range(c0, c1 + 1)))
                for c0, c1, t in spec if t)
    return _span(ws, r, spec, opts.HDRF, opts.NAVY, opts.CEN, max(32, 14 * lines + 6))


def free_pair(ws, wsv, blks, start=11, limit=40):
    """The first two adjacent columns that are empty on every row design A writes to.

    The squad tables are not all the same width. 1.4, 1.5 and 1.6 carry five more columns
    the owner added - Nbr Archetype Roles, Published Roles, Review Outcome, Vacant Now,
    FY27 - and writing at a fixed K and L destroyed two of them on 1.6 without a single
    check noticing: they hold typed numbers, so nothing recalculated and nothing broke.
    """
    rows, note = set(), set()
    for b in blks:
        rows.update([b["hdr"] - 1, b["hdr"], b["total"]] + b["squads"])
        edge = max((c for c in range(2, limit) if wsv.cell(b["hdr"], c).value is not None),
                   default=2)
        # a free-floating note past the table's own last column will be moved along the row,
        # so it must not push the columns further out. Counting it as occupied left K empty
        # and the figures one column adrift from the table on three tabs.
        note.update((r, c) for r in range(b["hdr"] + 1, b["total"] + 1)
                    for c in range(edge + 1, limit)
                    if isinstance(wsv.cell(r, c).value, str))
    for c in range(start, limit):
        if all((wsv.cell(r, k).value is None or (r, k) in note)
               for r in rows for k in (c, c + 1)):
            return c, c + 1
    raise SystemExit("no free column pair on this tab")


def clipped(ws, wsv, r, c):
    """A note to the left of column c that would stop being readable if c were filled.

    The owner's notes sit in a single cell and run across the empty columns beside them -
    "People in this program today cost 0.24m. Set the agreed cost in the cream cell." on
    1.6 needs four columns to show. Putting a figure in the fourth truncates it to "People
    in this program", which loses the instruction without touching a cell.
    """
    for k in range(c - 1, 1, -1):
        v = wsv.cell(r, k).value
        if v is None:
            continue
        if not isinstance(v, str):
            return None
        room = sum(ws.column_dimensions[L(x)].width or 8.43 for x in range(k, c))
        return k if opts.wrap_lines(v, room) > 1 else None
    return None


def shift_notes(ws, wsv, blks, act, var):
    """Move a note the new columns would truncate to the far side of them.

    Nothing is lost and nothing is reworded - the note stays on its own row, and it is the
    same move a reader would make by hand after widening a table under an annotation. The
    alternative, placing the columns beyond the widest note, left eleven blank columns
    between the table and its own figures on three tabs.

    Only a squad row is considered, and only a cell past the right-hand edge of that
    block's own header. Walking left from the new column on any row instead picked up
    "Funded outside TDD ($m)" - a header, not a note - and moved the table's own last
    column out of the table on six tabs.
    """
    out = []
    for b in blks:
        edge = max((c for c in range(2, act) if wsv.cell(b["hdr"], c).value is not None),
                   default=2)
        for r in range(b["hdr"] + 1, b["total"] + 1):
            # a note the columns will sit on top of has to move whether it would have been
            # clipped or not. Every row of the block, not only the squad rows: 1.9 carries
            # one on its overhead line and another on its total.
            k = next((c for c in range(edge + 1, var + 1)
                      if isinstance(wsv.cell(r, c).value, str)), None)
            if k is None:
                k = clipped(ws, wsv, r, act)
            if k is None or k <= edge:
                continue
            to = var + 1
            while wsv.cell(r, to).value is not None or ws.cell(r, to).value is not None:
                to += 1
            src, dst = ws.cell(r, k), ws.cell(r, to)
            dst.value, dst.font, dst.alignment = src.value, copy.copy(src.font), opts.LFT
            src.value = None
            out.append(f"note on row {r} moved {L(k)} -> {L(to)}, past the new columns")
    return out


def empty_col(ws, wsv, c, limit=95):
    """A column with nothing anywhere on the tab, so its width is nobody else's."""
    return all(wsv.cell(r, c).value is None
               for r in range(1, min(ws.max_row, limit) + 1))


def design_a(ws, wsv, blks, tab, lo, hi, after):
    act, var = free_pair(ws, wsv, blks)
    moved = shift_notes(ws, wsv, blks, act, var)
    for c in (act, var):
        # a width may only be set on a column no other table on this tab is using
        if empty_col(ws, wsv, c) and (ws.column_dimensions[L(c)].width or 8.43) < 20:
            ws.column_dimensions[L(c)].width = 22
    for b in blks:
        # the table just got two columns wider, so its bar has to follow it. The platform
        # bars are merged B:J, and a merge followed by painted cells reads as one bar.
        for c in (act, var):
            x = ws.cell(b["hdr"] - 1, c)
            x.fill, x.font = opts.fl(opts.BARC), opts.BARF
        _head(ws, b["hdr"], [(act, act, "Actual cost after decisions ($m)"),
                             (var, var, "Variance to archetype ($m)")])
        for r in b["squads"]:
            _m(ws, r, act, lookup(tab, lo, hi, r, after))
            _m(ws, r, var, diff(f"${L(ARCH)}{r}", f"${L(act)}{r}"))
        t = b["total"]
        # named cells, not a range: a block whose overhead line sits between its squads
        # would otherwise be summed with the overhead in it
        _m(ws, t, act, "=SUM(" + ",".join(f"${L(act)}{r}" for r in b["squads"]) + ")",
           bold=True)
        # the block total compares only when every squad in it is priced by an archetype,
        # or the archetype side is short a squad the actual side is carrying
        _m(ws, t, var, f'=IF({priced([f"${L(ARCH)}{r}" for r in b["squads"]])},'
                       f'{diff(f"${L(ARCH)}{t}", f"${L(act)}{t}")[1:]},"-")', bold=True)
        for c in (act, var):
            ws.cell(t, c).fill = copy.copy(ws.cell(t, ARCH).fill)
    return act, var, moved


# B:C the platform, D:G the squad, then the squad tables' own three cost columns and the
# two new ones, each under the column it already lives in further up the tab
B_SPEC = [(2, 3, "Platform"), (4, 7, "Squad"), (8, 8, "Total Squad Cost ($m)"),
          (9, 9, "TDD Cost ($m)"), (10, 10, "Funded outside TDD ($m)"),
          (11, 11, "Actual cost after decisions ($m)"),
          (12, 12, "Variance to archetype ($m)")]


def design_b(ws, blks, tab, lo, hi, after, anchor):
    """One table at the foot of the tab. Returns the row its portfolio block starts on."""
    r = (ws.max_row or 1) + 3
    r = opts.bar(ws, r, 2, LAST - 1, "Archetype against actual - every squad on this tab")
    r = _head(ws, r, B_SPEC)
    subs, unpriced = [], []
    for b in blks:
        st = r
        for src in b["squads"]:
            _span(ws, r, [(2, 3, b["name"]), (4, 7, f"=$B{src}")], opts.BODY, None, opts.LFT)
            # a blank archetype cell reads back as 0 through "=$H", which would print
            # $0.00m as though the archetype priced this squad at nothing
            for c in (ARCH, 9, 10):
                _m(ws, r, c, f'=IF({L(c)}{src}="","-",{L(c)}{src})')
            _m(ws, r, ACT, lookup(tab, lo, hi, src, after).replace(f"$B{src}", f"$D{r}"))
            _m(ws, r, VAR, diff(f"${L(ARCH)}{r}", f"${L(ACT)}{r}"))
            r += 1
        _span(ws, r, [(2, 7, f"{b['name']} total")] + [(c, c, None) for c in range(8, 13)],
              opts.BOLD, opts.GREY, opts.LFT)
        for c in (ARCH, 9, 10):
            cells = [f"${L(c)}{x}" for x in range(st, r)]
            _m(ws, r, c, f'=IF({priced(cells)},SUM({",".join(cells)}),"-")', bold=True)
        _m(ws, r, ACT, f"=SUM(${L(ACT)}{st}:${L(ACT)}{r-1})", bold=True)
        _m(ws, r, VAR, diff(f"${L(ARCH)}{r}", f"${L(ACT)}{r}"), bold=True)
        (subs if b["comparable"] else unpriced).append((f"${L(ARCH)}{r}", f"${L(ACT)}{r}"))
        r += 1
    return _foot(ws, r, subs, unpriced, tab, anchor, (ARCH, ACT, VAR))


def _line(ws, r, text, cols, arch, act, var, band=False):
    ca, cb, cv = cols
    _span(ws, r, [(2, ca - 1, text)] + [(c, c, None) for c in range(ca, cv + 1)],
          opts.BOLD if band else opts.BODY, opts.MID if band else None, opts.LFT)
    if band:
        for c in range(2, cv + 1):
            ws.cell(r, c).border = opts.TOPR
    for c, f in ((ca, arch), (cb, act), (cv, var)):
        _m(ws, r, c, f, bold=band)
    return r + 1


def _foot(ws, r, subs, unpriced, tab, anchor, cols):
    """The portfolio lines every design ends on.

    Squads an archetype prices, then - only where there are any - squads it does not, then
    overhead, then whatever the working tab carries and this tab has no row for, then the
    total. Every line the archetype does not price states a dash rather than a figure, so
    nothing on this block adds two different bases together.
    """
    A, oh = L(anchor["cols"]["after"]), anchor["overhead_row"]
    ca, cb, _ = cols
    K = L(cb)
    add = []
    r = _line(ws, r, "Squads priced by an archetype", cols,
              "=" + "+".join(f"N({a})" for a, _ in subs),
              "=" + "+".join(f"N({b})" for _, b in subs),
              diff(f"${L(ca)}{r}", f"${K}{r}"), band=True)
    add.append(r - 1)
    if unpriced:
        r = _line(ws, r, "Squads with no archetype to price them", cols, '="-"',
                  "=" + "+".join(f"N({b})" for _, b in unpriced), '="-"')
        add.append(r - 1)
    r = _line(ws, r, "Overhead roles in this portfolio - no archetype prices them", cols,
              '="-"', f"=N('{tab}'!${A}${oh})" if oh else '="-"', '="-"')
    add.append(r - 1)
    # every squad on the working tab that has no row on this design tab. Without this the
    # control could not read zero, because the working tab carries groups the design does
    # not - a one-person EGI programme, a Leadership group - and saying so is the point.
    r = _line(ws, r, "On the working tab with no row on this tab", cols, '="-"',
              f"=ROUND(N('{tab}'!${A}${anchor['total_row']})-"
              + "-".join(f"${K}{x}" for x in add) + ",6)", '="-"')
    add.append(r - 1)
    r = _line(ws, r, "Total actual cost after decisions - ties to the working tab", cols,
              '="-"', f"='{tab}'!${A}${anchor['total_row']}", '="-"', band=True)
    tot = r - 1
    _span(ws, r, [(2, ca - 1,
                   "Control - every line above against the working tab, must be 0")],
          opts.BODY, None, opts.LFT)
    _m(ws, r, cb, "=ROUND(" + "+".join(f"${K}{x}" for x in add) + f"-${K}{tot},6)",
       fmt=opts.CTL_M)
    return add[0]


def _m(ws, r, c, f, fmt=None, bold=False):
    x = ws.cell(r, c)
    x.value = f
    x.number_format = fmt or opts.M2
    x.alignment, x.font = opts.RGT, (opts.BOLD if bold else opts.BODY)
    x.border = opts.BOX
    return x


def portfolio_block(ws, blks, tab, anchor, cols):
    """Design A states the portfolio comparison at the foot of the tab too, so both designs
    answer the same question in the same words.

    Its three figures sit in the same columns the squad tables above use, so a reader can
    run an eye straight down the tab.
    """
    ca, cb, cv = cols
    r = (ws.max_row or 1) + 3
    r = opts.bar(ws, r, 2, cv - 1, "Archetype against actual - this portfolio")
    r = _head(ws, r, [(2, ca - 1, "Line"), (ca, ca, "Archetype cost ($m)")]
              + [(c, c, None) for c in range(ca + 1, cb)]
              + [(cb, cb, "Actual cost after decisions ($m)"),
                 (cv, cv, "Variance to archetype ($m)")])
    ref = [(f"${L(ca)}{b['total']}", f"${L(cb)}{b['total']}") for b in blks]
    _foot(ws, r, [x for x, b in zip(ref, blks) if b["comparable"]],
          [x for x, b in zip(ref, blks) if not b["comparable"]], tab, anchor, cols)


COE_NOTE = ("These are centres of excellence, funded by allocation rather than priced by an "
            "archetype, so there is no archetype cost to compare against. The actual cost "
            "after decisions for these groups is on 3.4 COE Detail and on the working tab.")


def coe_note(wb, wv):
    """The three COE tabs say why they carry no comparison.

    They group by department - Business Partnering is Commercial plus TDD Business Partner -
    while the working tab groups by squad, so the split cannot be taken from either tab
    without inventing a mapping. Saying so is better than a silent gap on three of thirteen
    tabs.
    """
    out = []
    for one in ("1.11 BP&T", "1.12 SA&D", "1.13 Cyber Roles"):
        if one not in wb.sheetnames:
            continue
        ws, wsv = wb[one], wv[one]
        tot = next((r for r in range(1, min(ws.max_row, 30) + 1)
                    if str(wsv.cell(r, 2).value or "").strip() == "Total"), None)
        if tot is None:
            out.append(f"{one}: no Total row on the summary table")
            continue
        r = tot + 1
        if str(wsv.cell(r, 2).value or "").strip():          # never write over a live row
            out.append(f"{one}: row {r} is not free")
            continue
        ws.cell(r, 2).value = COE_NOTE
        ws.cell(r, 2).font = opts.BODY
        ws.cell(r, 2).alignment = NOTE_AL
        ws.row_dimensions[r].height = 14 * opts.wrap_lines(COE_NOTE, 60) + 4
        out.append(f"{one}: note on row {r} - no archetype prices a COE")
    return out


def run(src, dst, variant="A", anchors="anchors_final.json"):
    wb = openpyxl.load_workbook(src)
    wv = openpyxl.load_workbook(src, data_only=True)
    a = json.load(open(anchors))
    # the working tabs were renamed after the anchors were written, so map by portfolio
    live = {str(wv[t]["C3"].value): t for t in wb.sheetnames if t.startswith("2.")}
    out = []
    for one in [t for t in wb.sheetnames if re.match(r"^1\.\d+ ", t) and t in PAIR]:
        anchor = next(x for x in a.values() if PAIR[one] == x["tab"])
        tab = live[anchor["pf"]]
        lo, hi = anchor["first_squad"], anchor["total_row"]
        after = L(anchor["cols"]["after"])
        ws, wsv = wb[one], wv[one]
        blks = blocks(ws, wsv)
        if not blks:
            out.append(f"{one}: no platform block with a squad in it")
            continue
        if variant == "A":
            act, var, moved = design_a(ws, wsv, blks, tab, lo, hi, after)
            portfolio_block(ws, blks, tab, anchor, (ARCH, act, var))
            where = f"in {L(act)} and {L(var)}"
            out += [f"  {one}: {m}" for m in moved]
        else:
            design_b(ws, blks, tab, lo, hi, after, anchor)
            where = "in one table at the foot"
        out.append(f"{one}: {len(blks)} platform blocks, "
                   f"{sum(len(b['squads']) for b in blks)} squads, actual against "
                   f"archetype from {tab}, {where}")
    out += coe_note(wb, wv)
    wb.save(dst)
    return out


if __name__ == "__main__":
    import sys
    for x in run(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else "A"):
        print("  ", x)
