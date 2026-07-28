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
"""
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


def run(src, dst):
    wb = openpyxl.load_workbook(src)
    out = (drop_names(wb) + fix_validation(wb) + fix_signs(wb)
           + drop_orphan_overheads(wb))
    wb.save(dst)
    return out


if __name__ == "__main__":
    import sys
    for x in run(sys.argv[1], sys.argv[2]):
        print("  ", x)
