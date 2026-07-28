"""What the COE audit found that is still mine to fix.

This file used to rewrite `1.11!C15` and `1.12!C15` from the owner's `=C14` to
`=C14+C13`, on the finding that the portfolio-funded line above was feeding nothing. That
was right two generations ago, when planned spend on both tabs was gross. It is wrong now.
In rev he restructured both tabs so planned spend is net of C13 - his own note at
`1.11!B9` says so in as many words - and the rewrite compared a net spend to a gross
budget, which put both tabs out against 0.2 Data Config by exactly C13: 2.2 on 1.11, 1.4
on 1.12. His `=C14` is correct and it survives.

What remains: the banned Category columns survived as three defined names - BPTCat,
SADCat, CYBCat - pointing at Lists!E2:G4. No formula used them and no COE tab has a
Category column, so the money was never affected, but the columns the owner ruled out
were still in the file and would reappear in any dropdown built off a name.

--- the "formulas that make no sense" round -------------------------------------------

THE LEVER PRICE LIVES IN ONE PLACE.  `Lists!AC2:AD5` is the model's lever table - Filled
1, Hire 1, Hold 0, Offshore 0.4 - and it is what the 2.x working tabs price their levers
off. The three COE design tabs did not read it: all 102 of their T-column cost engines
(24 on 1.11, 26 on 1.12, 52 on 1.13) carried
`IF($H21="Hold",0,IF($H21="Offshore",0.4,1))` inline, so the 0.4 the owner can retype on
Lists moved the working tabs and left the design tabs exactly where they were. They
look it up now, in the house idiom the 2.x tabs already use:

    IFERROR(INDEX(Lists!$AD$2:$AD$5,MATCH($H21,Lists!$AC$2:$AC$5,0)),1)

The IFERROR default of 1 is not decoration - it is what preserves the behaviour exactly.
The inline IF charged full price for anything that was not "Hold" or "Offshore", which on
these tabs means "Onshore" and a blank cell, and neither is in the lever table; both fall
through to 1 as they did before. "Filled" and "Hire" are in the table and are 1 there, so
they are unmoved either way. 1.11's Hold branch, added a round earlier in fix1x, survives
as the Hold row of the table.

THE PLATFORM TOTAL ROWS.  Column H of a platform total stops at the last squad row while
column I runs one further and takes in the Platform Overhead line. It reads like a typo
and it is not: an overhead is not a squad and has no archetype price, so it exists in the
TDD Cost column only, and the H cell on every overhead row is empty. Widening H would
double-count against the portfolio summary, which already reads the overhead out of I.
Checked on all 29 total rows across all fourteen 1.x tabs - every one of them is the same
way round. They carry a comment now so the next reader does not "fix" them. This runs
here rather than in repair_design because fix1x collapses the empty platform blocks in
between, and a comment on a collapsed row would outlive its cell.
"""
import re

import openpyxl
from openpyxl.styles import Border, PatternFill

import opts

BANNED_NAMES = ("BPTCat", "SADCat", "CYBCat")


def drop_names(wb):
    out = []
    for n in BANNED_NAMES:
        if n in wb.defined_names:
            del wb.defined_names[n]
            out.append(f"defined name {n} removed")
    l = wb["Lists"]
    n = 0
    for r in range(1, 6):
        for c in (5, 6, 7):
            if l.cell(r, c).value is not None:
                l.cell(r, c).value = None
                l.cell(r, c).fill = PatternFill()
                l.cell(r, c).border = Border()
                n += 1
    out.append(f"Lists!E1:G5 cleared, {n} cells - the banned Category columns")
    return out


# ---------------------------------------------------------------- the design tabs
def fix_validation(wb):
    """SupportPct stopped at 90% while six live cells hold 100%.

    Lists!D11 is 1, sitting just outside the named range, so `1.1!G27`, `1.2!G33`, `G39`,
    `G40`, `1.4!G21` and `1.6!G27` all breached their own dropdown.
    """
    out = []
    dn = wb.defined_names.get("SupportPct")
    if dn is not None and dn.value.endswith("$D$10"):
        dn.value = dn.value.replace("$D$10", "$D$11")
        out.append("SupportPct extended to include 100%")
    return out


# tab -> {cell: (what it should say or compute, why)}
SIGNS = {
    # 1.4's variance ran budget less spend and was labelled "(Over)/ Under", the opposite
    # convention to every other design tab, and the funding block below it reads that cell
    # as though it were spend less budget
    "1.4 TDD Group Functions": {"H8": "TDD over/(under) budget ($m)", "I8": "=I7-I6"},
    # 1.7's budget box sits one row higher than the others, so I8+I9 added the NZ variance
    # to a total that already contained it. It read correctly only because NZ is zero.
    "1.7 Infrastructure": {"C17": "=$I$7+$I$8"},
    # 1.5 carried the same variance twice under the same label, and the funding block read
    # the second copy
    "1.5 P&C": {"H10": None, "I10": None, "C16": "=$I$8"},
}
HEADERS = {"1.4 TDD Group Functions": {"F20": "AU / NZ", "F29": "AU / NZ"},
           "1.5 P&C": {"F24": "AU / NZ"}}

def fix_signs(wb):
    out = []
    for tab, cells in SIGNS.items():
        ws = wb[tab]
        for cell, val in cells.items():
            was = ws[cell].value
            ws[cell] = val
            out.append(f"{tab}!{cell} {was!r} -> {val!r}")
    for tab, cells in HEADERS.items():
        ws = wb[tab]
        for cell, val in cells.items():
            if ws[cell].value != val:
                out.append(f"{tab}!{cell} {ws[cell].value!r} -> {val!r}")
                ws[cell] = val
    return out


def drop_orphan_overheads(wb):
    """A "Platform Overhead" label with nothing beside it reads as an unfinished line.

    The strategic-programme platforms have no squads to support, so they carry no platform
    overhead. The label was left behind; the blank cell it points at is summed into the
    portfolio overhead line either way, so removing the label changes no figure.
    """
    out = []
    for tab in [t for t in wb.sheetnames if t.startswith("1.") and " " in t]:
        ws = wb[tab]
        for r in range(1, min(ws.max_row, 90) + 1):
            if str(ws.cell(r, 2).value or "").strip() == "Platform Overhead" \
                    and ws.cell(r, 9).value is None:
                ws.cell(r, 2).value = None
                out.append(f"{tab}!B{r} orphan 'Platform Overhead' label removed")
    return out


LISTS_LEVER = ("IFERROR(INDEX(Lists!$AD$2:$AD$5,"
               "MATCH({h},Lists!$AC$2:$AC$5,0)),1)")
# the inline engine as all three tabs write it, with the lever cell captured
INLINE = re.compile(r'IF\((\$[A-Z]{1,2}\d+)="Hold",0,IF\(\1="Offshore",0\.4,1\)\)')
COE_TABS = ("1.11 BP&T", "1.12 SA&D", "1.13 Cyber Roles")


def lever_from_lists(wb):
    """The T-column engines look the lever factor up instead of carrying it."""
    out = []
    if "Lists" not in wb.sheetnames:
        return ["Lists is not in the workbook - lever factors left inline"]
    l = wb["Lists"]
    table = {str(l.cell(r, 29).value): l.cell(r, 30).value for r in range(2, 6)}
    if table.get("Hold") != 0 or table.get("Offshore") != 0.4:
        return [f"Lists!AC2:AD5 reads {table} - not the factors the engines carry, "
                f"left inline"]
    for tab in COE_TABS:
        if tab not in wb.sheetnames:
            out.append(f"{tab}: not in the workbook")
            continue
        ws = wb[tab]
        n = 0
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            for c in row:
                v = c.value
                if not (isinstance(v, str) and v.startswith("=")):
                    continue
                new, k = INLINE.subn(
                    lambda m: LISTS_LEVER.format(h=m.group(1)), v)
                if k:
                    c.value = new
                    n += k
        out.append(f"{tab}: {n} lever factors now read Lists!AC2:AD5"
                   if n else f"{tab}: no inline lever factor left to repoint")
    out.append("lever factors (Hold 0, Offshore 0.4, everything else 1) now priced from "
               "Lists!AC2:AD5 on the design tabs as well as the working tabs")
    return out


TOTAL_NOTE = ("Total Squad Cost stops at the last squad row on purpose. A platform "
              "overhead is not a squad and has no archetype price, so it exists in the "
              "TDD Cost column only and the cell beside this one on the overhead row is "
              "empty by design. Widening this SUM to take the overhead row in would "
              "double-count it against the portfolio summary above, which already reads "
              "the overhead out of column I. Checked on all fourteen 1.x tabs.")


def note_total_asymmetry(wb, out):
    """The H-vs-I total rows get a comment so nobody 'fixes' them."""
    from openpyxl.comments import Comment
    n = 0
    for tab in [t for t in wb.sheetnames if re.match(r"^1\.\d+ ", t)]:
        ws = wb[tab]
        for r in range(1, min(ws.max_row, 95) + 1):
            if not str(ws.cell(r, 2).value or "").strip().endswith(" Total"):
                continue
            h, i = ws.cell(r, 8), ws.cell(r, 9)
            if not (isinstance(h.value, str) and isinstance(i.value, str)):
                continue
            # the overhead row is the one this total's I range reaches and its H does not
            if h.value == i.value.replace("I", "H"):
                continue                       # symmetric block, no overhead row
            h.comment = Comment(TOTAL_NOTE, "TDD cost model")
            n += 1
    out.append(f"{n} platform total rows noted: H stops at the squads, I takes in the "
               f"Platform Overhead row, and that is deliberate")



def run(src, dst):
    wb = openpyxl.load_workbook(src)
    out = (drop_names(wb) + fix_validation(wb) + fix_signs(wb)
           + drop_orphan_overheads(wb) + lever_from_lists(wb))
    note_total_asymmetry(wb, out)
    wb.save(dst)
    return out


if __name__ == "__main__":
    import sys
    for x in run(sys.argv[1], sys.argv[2]):
        print("  ", x)
