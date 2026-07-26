"""Two things the COE audit found that are mine to fix.

1. "Total Business Partnering budget ($m)" excluded the portfolio-funded amount sitting on
   the line directly above it. `1.11!C15` was `=C14`, so `C13` - the 2.20 of Business
   Partner cost funded out of portfolio overheads - fed nothing at all, and "Left to fund"
   read 2.400562 against a real 0.200562. Same shape on 1.12, where the 1.40 of Domain
   Architect funding was unused and left-to-fund read 2.529494 against 1.129494.

   Register item 44 asks for portfolio funding plus both allocations. Nothing outside the
   two tabs reads these cells, so the change is contained: the COE design figure 3.1 reads
   is planned spend, not budget.

2. The banned Category columns survived as three defined names - BPTCat, SADCat, CYBCat -
   pointing at Lists!E2:G4. No formula used them and no COE tab has a Category column, so
   the money was never affected, but the columns the owner ruled out were still in the file
   and would reappear in any dropdown built off a name.
"""
import openpyxl
from openpyxl.styles import Border, PatternFill

import opts

# tab -> the label whose total has to include the portfolio-funded line above it
TOTALS = {"1.11 BP&T": ("Business Partner funding from portfolio overheads",
                        "Total Business Partnering budget"),
          "1.12 SA&D": ("Domain Architect funding from portfolio overheads",
                        "Total Strategy & Architecture budget")}
BANNED_NAMES = ("BPTCat", "SADCat", "CYBCat")


def row_of(ws, label, col=2, limit=60):
    for r in range(1, min(ws.max_row, limit) + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str) and v.strip().startswith(label):
            return r
    return None


def fix_totals(wb):
    out = []
    for tab, (funded, total) in TOTALS.items():
        ws = wb[tab]
        rf, rt = row_of(ws, funded), row_of(ws, total)
        if rf is None or rt is None:
            out.append(f"{tab}: could not find {funded!r} / {total!r}")
            continue
        cur = ws.cell(rt, 3).value
        alloc = rt - 1                      # the COE's own allocation, directly above
        ws.cell(rt, 3).value = f"=C{alloc}+C{rf}"
        out.append(f"{tab}!C{rt} {cur!r} -> =C{alloc}+C{rf}, so the {ws.cell(rf, 2).value}"
                   f" on row {rf} is counted")
    return out


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


def run(src, dst):
    wb = openpyxl.load_workbook(src)
    out = fix_totals(wb) + drop_names(wb)
    wb.save(dst)
    return out


if __name__ == "__main__":
    import sys
    for x in run(sys.argv[1], sys.argv[2]):
        print("  ", x)
