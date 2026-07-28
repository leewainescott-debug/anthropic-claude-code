"""Re-apply the design-tab fixes the review workbook's branch predates.

rev.xlsx forked before two documented rounds of work, so its 1.x tabs still carry:

  the pre-D4 squad names - the ledger is the source of truth and the design tabs were
  renamed to follow it ("Z Energy Apps" is the ledger's "Z App and Web", and so on);
  the D6a zero-people squads - "We cannot have squads with 0 people in them" (owner) -
  which charge archetype cost against nobody. The new Customer dataset still has no
  people in any of them, so the rule still holds.

Renaming a row keeps every figure the owner typed on it, which is the point: his squad
cost for "Z Energy Martech" belongs to the squad the ledger calls "Z Loyalty & Martech",
and the rename carries it there. The one exception is Digital Support NZ, which is
removed outright; its typed 0.217 goes with it and is logged.

Also here: the one-cell typo his 1.2 edit introduced (platform overhead summing the
Total Squad Cost column instead of the TDD Cost column).
"""
import openpyxl

RENAME = {
    "1.1 Ampol Retail": {"Network / QSR": "Network & QSR"},
    "1.2 Customer": {"Z Energy Apps": "Z App and Web",
                     "Z Energy Martech": "Z Loyalty & Martech",
                     "AU CRM & Martech": "Ampol Loyalty & Martech",
                     "Customer AI": "Customer, AI"},
    "1.4 TDD Group Functions": {"Network & Infrastructure": "Cloud, Network & Infra Ops",
                                "DevOps & Engineering": "DevOps & QE",
                                "Integration": "Integration & Process Automation"},
    "1.5 P&C": {"P&C - RTA": "P&C RTA"},
    "1.6 Finance": {"AU Finance": "SAP ERP"},
    "1.7 Infrastructure": {"Manufacturing & Group Projects": "Manufacturing Group Projects"},
}
# D6a: design rows with nobody in them, confirmed still empty in the new dataset
REMOVE = {
    "1.2 Customer": ["Digital Support NZ"],
    "1.3 Enterprise Data": ["EGI Data", "Enterprise Data Delivery"],
}
FIX = {
    # his C7 points platform overhead at column H (Total Squad Cost); the overhead lives
    # in I on this tab's platform-overhead rows, as it did before and as C7 did in the
    # ancestor
    "1.2 Customer": [("C7", None, None)],
    # his 1.5 edit removed the NZ column, so 0.2's P&C spend read wraps the dead D9 the
    # house way
    "0.2 Data Config": [("F18", "=('1.5 P&C'!$C$9+'1.5 P&C'!$D$9)",
                         "='1.5 P&C'!$C$9+N('1.5 P&C'!$D$9)")],
}


# his typed inputs on tabs the chain does not repaint - declared cream so the
# stale-literal police reads them as inputs, which they are
CREAM = [("1.2 Customer", "I54"), ("1.8 Energy Solutions & B2B", "E12"),
         ("1.8 Energy Solutions & B2B", "I14"), ("1.8 Energy Solutions & B2B", "I15")]


def wrap_empty_budget_reads(wb, out):
    """His zeroing rewires point budget cells at empty 0.1 cells; N() is the house style
    for a read that may legitimately be blank, and it is what keeps the adversarial pass
    able to tell a deliberate empty read from a broken one."""
    import re
    pat = re.compile(r"^='0\.1 Budget Table \(Fin\)'!([A-Z]{1,2}\d+)$")
    fin = wb["0.1 Budget Table (Fin)"]
    n = 0
    for t in [s_ for s_ in wb.sheetnames if re.match(r"^1\.\d+ ", s_)]:
        ws = wb[t]
        for row in ws.iter_rows(min_row=1, max_row=30, min_col=8, max_col=11):
            for c in row:
                if isinstance(c.value, str):
                    m = pat.match(c.value)
                    if m and fin[m.group(1)].value is None:
                        c.value = f"=N('0.1 Budget Table (Fin)'!{m.group(1)})"
                        n += 1
    out.append(f"{n} budget reads of empty 0.1 cells wrapped in N()")


def run(src, dst):
    import opts as _o
    wb = openpyxl.load_workbook(src)
    out = []
    wrap_empty_budget_reads(wb, out)
    for tab, cell in CREAM:
        wb[tab][cell].fill = _o.fl(_o.YEL)
    out.append(f"{len(CREAM)} of his typed inputs declared cream")
    # his new 0.2 headers and 1.x note headers run past their columns; wrap them the way
    # every other header in the file wraps
    from openpyxl.styles import Alignment
    WRAP = [("0.2 Data Config", x) for x in ("N5", "N13", "N21", "O21", "L21", "M21")]
    for t in wb.sheetnames:
        import re as _re
        if _re.match(r"^1\.\d+ ", t):
            for r_ in range(10, 16):
                for c_ in (10, 11):
                    v = wb[t].cell(r_, c_).value
                    if isinstance(v, str) and len(v) > 24:
                        WRAP.append((t, wb[t].cell(r_, c_).coordinate))
    n_w = 0
    for t, cell in WRAP:
        x = wb[t][cell]
        if isinstance(x.value, str):
            x.alignment = Alignment(horizontal=x.alignment.horizontal or "left",
                                    vertical="center", wrap_text=True)
            n_w += 1
    out.append(f"{n_w} long headers set to wrap")
    # his new On/Off column widened the 1.13 roles table to H; the bar above follows it
    ws13 = wb["1.13 Cyber Roles"]
    for r_ in range(1, 30):
        v = str(ws13.cell(r_, 2).value or "")
        if v.strip() == "Roles":
            for c_ in range(2, 9):
                ws13.cell(r_, c_).fill = _o.fl(_o.BARC)
                ws13.cell(r_, c_).font = _o.BARF
            out.append(f"1.13!B{r_} bar extended across the On/Off column")
            break
    for tab, pairs in RENAME.items():
        ws = wb[tab]
        for r in range(1, min(ws.max_row, 95) + 1):
            v = ws.cell(r, 2).value
            if isinstance(v, str) and v.strip() in pairs:
                new = pairs[v.strip()]
                ws.cell(r, 2).value = new
                out.append(f"{tab}!B{r} {v.strip()!r} -> {new!r} (ledger name)")
    for tab, names in REMOVE.items():
        ws = wb[tab]
        for r in range(1, min(ws.max_row, 95) + 1):
            v = str(ws.cell(r, 2).value or "").strip()
            if v in names:
                dropped = [f"{openpyxl.utils.get_column_letter(c)}={ws.cell(r, c).value!r}"
                           for c in range(2, 15) if ws.cell(r, c).value is not None]
                for c in range(2, 15):
                    ws.cell(r, c).value = None
                out.append(f"{tab}!B{r} {v!r} removed - zero people in the ledger; "
                           f"carried {'; '.join(dropped[:4])}")
    for tab, fixes in FIX.items():
        ws = wb[tab]
        for cell, expect, repl in fixes:
            cur = ws[cell].value
            if expect is None:
                # his 1.2 C7 wraps the platform-overhead sum in a NZ/AU condition; only
                # the column letters inside are wrong (H, the squad-cost column, where
                # the overhead lives in I)
                if isinstance(cur, str) and "SUM(H34,H42,H49)" in cur:
                    ws[cell] = cur.replace("SUM(H34,H42,H49)", "SUM(I34,I42,I49)")
                    out.append(f"{tab}!{cell}: H34/H42/H49 -> I columns inside his IF")
                continue
            if cur == expect:
                ws[cell] = repl
                out.append(f"{tab}!{cell} -> {repl}")
            else:
                out.append(f"{tab}!{cell} holds {str(cur)[:50]!r} - left alone")
    wb.save(dst)
    return out


if __name__ == "__main__":
    import sys
    for x in run(sys.argv[1], sys.argv[2]):
        print("  ", x)
