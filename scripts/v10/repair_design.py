"""Re-apply the design-tab fixes the review workbook's branch predates.

rev.xlsx forked before two documented rounds of work, so its 1.x tabs still carry:

  the pre-D4 squad names - the ledger is the source of truth and the design tabs were
  renamed to follow it ("Z Energy Apps" is the ledger's "Z App and Web", and so on);
  the D6a zero-people squads - "We cannot have squads with 0 people in them" (owner) -
  which charge archetype cost against nobody. The new Customer dataset still has no
  people in any of them, so the rule still holds.

Renaming a row keeps every figure the owner typed on it, which is the point: his squad
cost for "Z Energy Martech" belongs to the squad the ledger calls "Z Loyalty & Martech",
and the rename carries it there. A rename now follows the name onto the two rows that
frame the block - "Platform: {name}" above it and "{name} Total" below - because matching
column B exactly left 1.1 with a platform bar and a total row still calling the squad
"Network / QSR" over a squad row renamed "Network & QSR".

The one exception is Digital Support NZ, which is removed outright; its typed 0.217 goes
with it and is logged. Removing it used to blank the row out to column N, which took his
note at 1.2!K41 with it. A removed row loses its figures, not his writing: anything he
typed in K..N moves along the row to M and stays.

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
# 1.2's platform overhead, once the H/I typo above is out of the way, is the identical
# formula in the AU column and the NZ column: both branches count the whole 0.495, so the
# three Customer platforms are charged twice and F7 reads 0.99.
#
# Customer is the one dual-country portfolio on this tab set - 0.2 gives it an AU line
# (row 13) and a NZ line (row 14) - and row 6 directly above says what that means here:
# portfolio overhead is the same IF over the same two 0.2 cells with *0.5 on each side, so
# the tab splits its overhead half to each country rather than choosing one. Row 7 is the
# same overhead on a per-platform basis and follows the same split. Its nine single-country
# siblings use the exclusive form instead - all of it to one column, nothing to the other -
# which is why none of them doubles.
CUST_PLAT = ("=IF(('0.2 Data Config'!$D$13+'0.2 Data Config'!$D$14)>"
             "('0.2 Data Config'!$C$13+'0.2 Data Config'!$C$14),SUM(I34,I42,I49),0)")
FIX = {
    # his C7 points platform overhead at column H (Total Squad Cost); the overhead lives
    # in I on this tab's platform-overhead rows, as it did before and as C7 did in the
    # ancestor
    "1.2 Customer": [("C7", None, None),
                     # keyed on the byte-identical text the line above produces, so it can
                     # only fire on the doubled shape and never on a corrected one
                     ("C7", CUST_PLAT, CUST_PLAT + "*0.5"),
                     ("D7", CUST_PLAT, CUST_PLAT + "*0.5")],
    # his 1.5 edit removed the NZ column, so 0.2's P&C spend read wraps the dead D9 the
    # house way
    "0.2 Data Config": [("F18", "=('1.5 P&C'!$C$9+'1.5 P&C'!$D$9)",
                         "='1.5 P&C'!$C$9+N('1.5 P&C'!$D$9)"),
                        # every other row of this table reads its 1.x tab absolutely; two
                        # read relatively, which is a copy away from pointing at the wrong
                        # row
                        ("F13", "='1.2 Customer'!C9", "='1.2 Customer'!$C$9"),
                        ("F14", "='1.2 Customer'!D9", "='1.2 Customer'!$D$9"),
                        ("F15", "='1.9 Commercial Fuels'!C9",
                         "='1.9 Commercial Fuels'!$C$9"),
                        # his row 23 is labelled "TDD Cyber incl. COE ..." - once the
                        # 1.14 TDD Cyber tab exists its spend belongs on this row too,
                        # or the row understates the thing its own label names
                        ("F23", "='1.13 Cyber Roles'!$F$11",
                         "='1.13 Cyber Roles'!$F$11+N('1.14 TDD Cyber'!$F$9)")],
    # 1.10's Data NZ platform draws no overhead - the owner's own note on the squad row,
    # 1.10!K39 "No Overhead required", says so, and his C7 already prices two platforms.
    # Only D7 still drags the dead I40 reference from before he removed the charge, so the
    # two branches of the same switch priced different sets. D7 is aligned to his C7; his
    # note, his empty I40 and his B40 label are not touched.
    "1.10 Z Retail": [("D7", "=IF(('0.2 Data Config'!$D$12)>('0.2 Data Config'!$C$12),"
                             "SUM(I27,I34,I40),0)",
                       "=IF(('0.2 Data Config'!$D$12)>('0.2 Data Config'!$C$12),"
                       "SUM(I27,I34),0)")],
}

# rows of 0.2's budget table with no Spend cell and no Variance cell, where all their
# siblings carry both. Written as the sibling pair - a hardcoded 0 spend the way row 7
# does it, and E less F beside it - so the table reads the same way all the way down and
# G26 totals a complete column. Row 24 is his restored EG row, which gets the same pair.
CFG_GAPS = [20, 24, 25]


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


def close_config_gaps(wb, out):
    """0.2's Spend and Variance cells, on the rows carrying neither."""
    ws = wb["0.2 Data Config"]
    for r in CFG_GAPS:
        f_, g_ = ws.cell(r, 6), ws.cell(r, 7)
        if f_.value is not None or g_.value is not None:
            out.append(f"0.2!F{r}/G{r} hold {f_.value!r} / {g_.value!r} - left alone")
            continue
        f_.value = 0
        g_.value = f"=E{r}-F{r}"
        out.append(f"0.2!{ws.cell(r, 2).value!r}: F{r} = 0 and G{r} = E{r}-F{r}, the pair "
                   f"every other row of this table carries")


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
        # a squad's name appears three times in its block: on the platform bar above it,
        # on the squad row, and on the total row below. All three follow the ledger.
        wide = {}
        for old, new in pairs.items():
            wide[old] = new
            wide[f"Platform: {old}"] = f"Platform: {new}"
            wide[f"{old} Total"] = f"{new} Total"
        for r in range(1, min(ws.max_row, 95) + 1):
            v = ws.cell(r, 2).value
            if isinstance(v, str) and v.strip() in wide:
                new = wide[v.strip()]
                # whatever padding he typed around the name is his and stays
                lead = v[: len(v) - len(v.lstrip())]
                trail = v[len(v.rstrip()):]
                ws.cell(r, 2).value = lead + new + trail
                out.append(f"{tab}!B{r} {v.strip()!r} -> {new!r} (ledger name)")
    for tab, names in REMOVE.items():
        ws = wb[tab]
        for r in range(1, min(ws.max_row, 95) + 1):
            v = str(ws.cell(r, 2).value or "").strip()
            if v in names:
                dropped = [f"{openpyxl.utils.get_column_letter(c)}={ws.cell(r, c).value!r}"
                           for c in range(2, 11) if ws.cell(r, c).value is not None]
                for c in range(2, 11):              # B..J, the figures and the levers
                    ws.cell(r, c).value = None
                # K..N is the note margin. His writing survives the row it was sitting on:
                # it moves to M, where the design tabs carry a note, and the cell it came
                # from is cleared.
                for c in range(11, 15):
                    x = ws.cell(r, c)
                    if not (isinstance(x.value, str) and x.value.strip()
                            and not x.value.startswith("=")):
                        continue
                    if c in (13, 14):               # already in the note margin
                        continue
                    free = next((k for k in (13, 14) if ws.cell(r, k).value is None), None)
                    if free is None:
                        out.append(f"{tab}!{openpyxl.utils.get_column_letter(c)}{r} "
                                   f"{x.value.strip()!r} kept where it is - M and N are "
                                   f"both taken")
                        continue
                    ws.cell(r, free).value = x.value
                    ws.cell(r, free)._style = x._style
                    out.append(f"{tab}!{openpyxl.utils.get_column_letter(c)}{r} "
                               f"{x.value.strip()!r} moved to "
                               f"{openpyxl.utils.get_column_letter(free)}{r} - his note "
                               f"outlives the row it sat on")
                    x.value = None
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
    close_config_gaps(wb, out)
    wb.save(dst)
    return out


if __name__ == "__main__":
    import sys
    for x in run(sys.argv[1], sys.argv[2]):
        print("  ", x)
