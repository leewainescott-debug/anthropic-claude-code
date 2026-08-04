#!/usr/bin/env python3
"""u4 - spec section D (and E, folded into it): the funded outside architecture.

  python3 u4_outside.py <in.xlsx> <out.xlsx>

Every lever modelling tab learns two things it never carried: how much of a
squad somebody else pays for, and what is left for TDD. The variance to the
archetype then compares the archetype against the TDD funded cost, not against
the whole cost, and a lever can no longer zero a funding disclosure.

Also carries integrator ruling 1: the third role the owner typed REMOVE over
(2.4 B38, ledger row 476 in his file) leaves the ledger the same way A4's two
did, so every count re-derives off 528 roles.

Idempotent: handed its own output it copies it through untouched.
"""
import sys, os, shutil
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from _xl import (REVIEW, Log, load, save, shift_rows, shift_cols, swap_col,
                 copy_style, ledger, extend_review)

# Lists gains one table: Squad | Funded by | Basis  (D1)
# AU / AV / AW, rows 2 to 10.  The two "EGI" squad rows the spec names
# separately - 2.3 Enterprise Data's and 2.14's - are the same squad name and
# the same basis, so they key on one row and both tabs resolve through it.
FUND_TABLE = [
    ("EGI Retail",   "EGI",                    "actual"),
    ("EGI TDD",      "EGI",                    "actual"),
    ("EGI Customer", "EGI",                    "actual"),
    ("EGI",          "EGI",                    "actual"),
    ("EGI P&C",      "EGI",                    "actual"),
    ("EGI Finance",  "EGI",                    "actual"),
    ("CTRM",         "CTRM programme",         3.8),
    ("AmPOS",        "AmPOS programme",        1.404),
    ("Cyber Uplift", "cyber uplift programme", 1.2998),
]
FIRST, LAST = 2, 1 + len(FUND_TABLE)

OUTSIDE = "Funded outside TDD ($m)"
TDDFUND = "TDD-funded cost ($m)"

# =$O if the squad is funded outside at actual, else the typed basis, else 0
FUND_F = ('=IFERROR(IF(INDEX(Lists!$AW${a}:$AW${z},MATCH($B{r},Lists!$AU${a}:$AU${z},0))'
          '="actual",$O{r},INDEX(Lists!$AW${a}:$AW${z},'
          'MATCH($B{r},Lists!$AU${a}:$AU${z},0))),0)')

# D4 - the five 1.x funding disclosures, moved off the after levers column
DISCLOSURES = [("1.1 Ampol Retail", "H66", "2.1 Ampol Retail", 18),
               ("1.2 Customer", "H54", "2.2 Customer", 17),
               ("1.4 TDD Group Functions", "H30", "2.4 TDD Group Functions", 14),
               ("1.5 P&C", "H32", "2.5 P&C", 12),
               ("1.6 Finance", "H33", "2.6 Finance", 12)]

# integrator ruling 1 - his third REMOVE, as it stands after A4's two deletions
R1_TAB, R1_ROW, R1_LEDGER, R1_NAME = "2.4 TDD Group Functions", 38, 475, \
    "Ring fenced selection"


def tabs2x(wb):
    return [ws.title for ws in wb.worksheets if ws.title.startswith("2.")]


def total_row(ws):
    for r in range(5, 60):
        if ws.cell(r, 2).value == "Total portfolio":
            return r
    raise SystemExit("STOP: no 'Total portfolio' row on %s" % ws.title)


def main(src, dst):
    log = Log("u4_outside")
    wb = load(src)

    if wb["Lists"]["AU1"].value is not None:
        print("input already carries the funded outside architecture - copying through")
        shutil.copy(src, dst)
        log.tail()
        print("wrote", dst)
        return

    L = wb["Lists"]
    rv = wb[REVIEW]

    # ------------------------------------------------- integrator ruling 1
    log.head("ruling 1  the third REMOVE role leaves the ledger properly")
    g4 = wb[R1_TAB]
    want = "='" + REVIEW + "'!$B$%d" % R1_LEDGER
    if g4.cell(R1_ROW, 2).value != want or \
            str(rv.cell(R1_LEDGER, 2).value).strip() != R1_NAME:
        print("STOP: %s!B%d is %r and REVIEW B%d is %r - not the REMOVE role"
              % (R1_TAB, R1_ROW, g4.cell(R1_ROW, 2).value, R1_LEDGER,
                 rv.cell(R1_LEDGER, 2).value))
        raise SystemExit(2)
    ttl = rv.cell(R1_LEDGER, 3).value
    shift_rows(wb, R1_TAB, R1_ROW, -1)
    log("D0", "%s row %d" % (R1_TAB, R1_ROW),
        "helper row for REVIEW %d removed with the role" % R1_LEDGER)
    shift_rows(wb, REVIEW, R1_LEDGER, -1)
    log("D0", "REVIEW row %d" % R1_LEDGER,
        "role deleted from the ledger (%r / %r)" % (R1_NAME, ttl))
    rv = wb[REVIEW]
    log("D0", "REVIEW", "roles now %d, every count re-derives" % len(ledger(rv)))

    # ---------------------------------------------------------------- D1
    log.head("D1  Lists gains the funded outside table")
    for col, head in ((47, "Squad"), (48, "Funded by"), (49, "Basis")):
        copy_style(L.cell(1, 42), L.cell(1, col))
        L.cell(1, col).value = head
    for i, (sq, by, basis) in enumerate(FUND_TABLE):
        r = FIRST + i
        for col, val in ((47, sq), (48, by), (49, basis)):
            copy_style(L.cell(2, 42), L.cell(r, col))
            L.cell(r, col).value = val
        L.cell(r, 49).number_format = "0.0000" if not isinstance(basis, str) \
            else "General"
    log("D1", "Lists!AU1:AW%d" % LAST,
        "Squad | Funded by | Basis, %d rows (EGI at actual, CTRM 3.8, "
        "AmPOS 1.404, Cyber Uplift 1.2998)" % len(FUND_TABLE))
    log.note("D1", "the spec names the 2.3 and the 2.14 EGI rows separately; "
                   "both squads are called EGI and both are at actual, so they "
                   "key on the one row")

    # ---------------------------------------------------------------- D2
    log.head("D2  every lever modelling tab gains the two columns after O")
    for title in tabs2x(wb):
        ws = wb[title]
        tot = total_row(ws)
        keep = {r: ws.cell(r, 15).value for r in range(7, tot + 1)}
        var = {r: ws.cell(r, 16).value for r in range(7, tot + 1)}
        shift_cols(wb, title, 16, 2)
        ws = wb[title]
        for col, head in ((16, OUTSIDE), (17, TDDFUND)):
            copy_style(ws.cell(6, 15), ws.cell(6, col))
            ws.cell(6, col).value = head
        n_data = n_tot = 0
        for r in range(7, tot + 1):
            o = keep[r]
            if o is None:
                continue
            copy_style(ws.cell(r, 15), ws.cell(r, 16))
            copy_style(ws.cell(r, 15), ws.cell(r, 17))
            if o == '=""':
                ws.cell(r, 16).value = '=""'
                ws.cell(r, 17).value = '=""'
            elif o.startswith("=SUM("):
                ws.cell(r, 16).value = swap_col(o, title, "O", "P")
                ws.cell(r, 17).value = swap_col(o, title, "O", "Q")
                n_tot += 1
            else:
                ws.cell(r, 16).value = FUND_F.format(r=r, a=FIRST, z=LAST)
                ws.cell(r, 17).value = "=$O{r}-$P{r}".format(r=r)
                n_data += 1
            # the variance now measures the archetype against TDD funded cost
            v = var[r]
            if isinstance(v, str) and ("$O%d" % r) in v:
                ws.cell(r, 18).value = v.replace("$O%d" % r, "$Q%d" % r)
        log("D2", title,
            "columns P and Q added (%d squad rows, %d total rows); variance "
            "in R now reads TDD-funded against archetype" % (n_data, n_tot))

    # ---------------------------------------------------------------- D3
    log.head("D3  3.1 gains the same two columns and a TDD-funded total row")
    t31 = "3.1 Archetype to Actuals"
    s31 = wb[t31]
    act = {r: s31.cell(r, 5).value for r in range(6, 22)}
    shift_cols(wb, t31, 6, 2)
    s31 = wb[t31]
    for col, head in ((6, OUTSIDE), (7, TDDFUND)):
        copy_style(s31.cell(4, 5), s31.cell(4, col))
        s31.cell(4, col).value = head
    for r in range(6, 22):
        e = act[r]
        if not isinstance(e, str):
            continue
        copy_style(s31.cell(r, 5), s31.cell(r, 6))
        copy_style(s31.cell(r, 5), s31.cell(r, 7))
        if r == 21:                       # the TDD total line, his N() chain
            s31.cell(r, 6).value = swap_col(e, t31, "E", "F")
            s31.cell(r, 7).value = swap_col(e, t31, "E", "G")
        else:
            s31.cell(r, 6).value = e.replace("$O$", "$P$")
            s31.cell(r, 7).value = e.replace("$O$", "$Q$")
    shift_rows(wb, t31, 22, 1)
    s31 = wb[t31]
    for c in range(2, 14):
        copy_style(s31.cell(21, c), s31.cell(22, c))
    s31["B22"].value = "TDD-funded total"
    s31["G22"].value = "=SUM($G$6:$G$20)"
    log("D3", "%s!F4:G22" % t31,
        "funded outside and TDD-funded columns, and the TDD-funded total row "
        "directly under the TDD total")
    # the GM layer sits outside the role mapping: none of it is funded outside
    s31["F23"].value = "=0"
    s31["G23"].value = "=$E$23"
    s31["F24"].value = "=N(F21)+N(F23)"
    s31["G24"].value = "=N(G21)+N(G23)"
    for c in (6, 7):
        copy_style(s31.cell(23, 5), s31.cell(23, c))
        copy_style(s31.cell(24, 5), s31.cell(24, c))
    log("D3", "%s!F23:G24" % t31,
        "the GM layer carries 0 funded outside, so the total including it ties")

    # ---------------------------------------------------------------- D4
    log.head("D4  the funding disclosures read the actual basis, never a lever")
    for sheet, coord, tab, row in DISCLOSURES:
        w = wb[sheet]
        new = "='%s'!$O$%d" % (tab, row)
        log("D4", "%s!%s" % (sheet, coord),
            "%r -> %s (actual, not after levers)" % (w[coord].value, new))
        w[coord].value = new
    a21 = wb["2.1 Ampol Retail"]
    log("D4", "2.1 Ampol Retail!N19",
        "%r -> =$O19 (the directly funded variance stops going blank)"
        % a21["N19"].value)
    a21["N19"].value = "=$O19"
    q = wb["4.0 Data QA"]
    d39 = "='1.1 Ampol Retail'!$F$9+N('2.1 Ampol Retail'!$N$19)"
    if q["D39"].value != d39:
        q["D39"].value = d39
        log("D4", "4.0 Data QA!D39",
            "identity restated: 2.1 carries a directly funded line 1.1 has no "
            "block for (the same shape A1 gave 2.3 and B4 gave 2.2)")

    # ------------------------------------------------------------------ E
    log.note("E", "2.14 EGI reads the EGI row of the table, so its funded "
                  "outside is its whole cost and its TDD-funded cost is 0")

    # ---------------------------------------------------------------- G3
    log.head("G3  the REVIEW ranges the deletion shortened go back to row 700")
    n = extend_review(wb, 700)
    log("G3", "workbook", "%d formulas read REVIEW $2:$700 again" % n)

    save(wb, dst)
    log.tail()
    print("wrote", dst)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
