#!/usr/bin/env python3
"""u1 - spec section A: reconcile the owner's 30/07 edits.

  python3 u1_reconcile.py <in.xlsx> <out.xlsx>

The owner's file is the base. Every edit below is a disposition he ruled on.
Idempotent: it fingerprints the input first. Handed its own output it copies it
through untouched, so re-running the pipeline can never delete a role twice.
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from _xl import (REVIEW, LEVERS, Log, load, save, shift_rows, repoint_sheet,
                 extend_review, white, row_style, copy_style, ledger,
                 rewrite_refs, read_block, write_block, set_dv, wire_summary)

OLD31 = "3.1 Archetype to Acutals"
NEW31 = "3.1 Archetype to Actuals"

# the 8 Enterprise Data roles he tagged Squad="EGI" (A1)
EGI8 = [279, 298, 308, 310, 311, 318, 320, 329]

# ------------------------------------------------------------------- script

def main(src, dst):
    log = Log("u1_reconcile")
    wb = load(src)
    rv = wb[REVIEW]

    # section A moves rows, so it must only ever run against his 30/07 base
    if NEW31 in wb.sheetnames and len(ledger(rv)) == 529:
        print("input is already reconciled - copying through untouched")
        import shutil
        shutil.copy(src, dst)
        log.tail()
        print("wrote", dst)
        return
    fp = [("3.1 tab name", OLD31 in wb.sheetnames),
          ("531 roles", len(ledger(rv)) == 531),
          ("REVIEW B420 = Remove", str(rv["B420"].value).strip() == "Remove"),
          ("2.6 B41 = Remove", str(wb["2.6 Finance"]["B41"].value).strip() == "Remove")]
    if not all(ok for _, ok in fp):
        print("STOP: input is not the owner's 30/07 base: %r"
              % [n for n, ok in fp if not ok])
        raise SystemExit(2)

    # ---------------------------------------------------------------- A5
    log.head("A5  Significant Items budget is 4.5")
    b = wb["0.1 Budget Table (Fin)"]
    if b["O13"].value != 4.5:
        log("A5", "0.1 Budget Table (Fin)!O13", "%r -> 4.5" % b["O13"].value)
        b["O13"].value = 4.5
    c12 = wb["1.2 Customer"]
    want = "='0.1 Budget Table (Fin)'!O13"
    if c12["I18"].value != want:
        log("A5", "1.2 Customer!I18", "%r -> %s" % (c12["I18"].value, want))
        c12["I18"].value = want

    # ---------------------------------------------------------------- A13
    log.head("A13  ledger spelling and the five live #VALUE! note cells")
    if rv["C132"].value == "Engineer - Mobile (ios)":
        log("A13", "REVIEW!C132", "(ios) -> (iOS)")
        rv["C132"].value = "Engineer - Mobile (iOS)"
    NOTE = ('="People in this program today cost "&TEXT(SUMIFS(\'' + REVIEW +
            "'!$AA$2:$AA$700,'" + REVIEW + "'!$AT$2:$AT$700,$B{row},'" + REVIEW +
            "'!$AJ$2:$AJ$700,'{tab}'!$C$3)/1000000,\"0.00\")&"
            '"m. Set the agreed cost in the cream cell."')
    for sheet, coord, row, tab in [
            ("1.1 Ampol Retail", "N48", 48, "2.1 Ampol Retail"),
            ("1.1 Ampol Retail", "N66", 66, "2.1 Ampol Retail"),
            ("1.4 TDD Group Functions", "M30", 30, "2.4 TDD Group Functions"),
            ("1.5 P&C", "N32", 32, "2.5 P&C"),
            ("1.6 Finance", "N33", 33, "2.6 Finance")]:
        w = wb[sheet]
        new = NOTE.format(row=row, tab=tab)
        if w[coord].value != new:
            log("A13", "%s!%s" % (sheet, coord),
                "note SUMIFS repaired (AA over AT=squad, AJ=portfolio)")
            w[coord].value = new

    # ---------------------------------------------------------------- A11
    log.head("A11  his wording sweep finished")
    words = [
        ("3.3 Squad Actuals to Archetype", "K5", "Total roles after decisions",
         "Total roles after levers"),
        ("3.3 Squad Actuals to Archetype", "O5", "Cost after decisions ($m)",
         "Cost after levers ($m)"),
        ("Exec Summary", "B31", "Roles after the decisions set today",
         "Roles after the levers set today"),
        ("Exec Summary", "B32", "Cost after the decisions set today ($m)",
         "Cost after the levers set today ($m)"),
        ("Exec Summary", "B33", "Impact of those decisions ($m)",
         "Impact of those levers ($m)"),
        ("Exec Summary", "B40", "Roles after decisions", "Roles after levers"),
        ("Exec Summary", "B45", "Cost after vacancy decisions ($m)",
         "Cost after vacancy levers ($m)"),
        ("4.0 Data QA", "B31", "Roles after decisions against roles less anything on hold",
         "Roles after levers against roles less anything on hold"),
        ("2.5 P&C", "P6", "Variance to levers ($m)", "Variance to archetype ($m)"),
        ("2.5 P&C", "Q6", "Squad cost after decisions ($m)",
         "Squad cost after levers ($m)"),
    ]
    for sheet, coord, old, new in words:
        w = wb[sheet]
        if w[coord].value == old:
            log("A11", "%s!%s" % (sheet, coord), "%r -> %r" % (old, new))
            w[coord].value = new
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v == "Actual cost after decisions ($m)":
                    cell.value = "Actual cost after levers ($m)"
                    n += 1
    log("A11", "1.x K headers", "%d 'Actual cost after decisions ($m)' -> levers" % n)
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and not v.startswith("=") and "the working tabs" in v:
                    cell.value = v.replace("the working tabs", "the lever modelling tabs")
                    n += 1
    log("A11", "workbook", "%d 'the working tabs' -> 'the lever modelling tabs'" % n)
    stale = [
        ("1.3 Enterprise Data", "H15", "3.4 COE Detail", "3.4 COE Breakdown"),
        ("1.4 TDD Group Functions", "H14", "3.4 COE Detail", "3.4 COE Breakdown"),
        ("1.11 BP&T", "B10", "3.1 Cost Bridge", NEW31),
        ("1.12 SA&D", "B10", "3.1 Cost Bridge", NEW31),
        ("1.12 SA&D", "B53", "3.3 Squad Detail", "3.3 Squad Actuals to Archetype"),
        ("1.13 Cyber Roles", "B9", "3.4 COE Detail", "3.4 COE Breakdown"),
        ("1.13 Cyber Roles", "B9", "3.1 Cost Bridge", NEW31),
        ("3.4 COE Breakdown", "B2", "COE detail", "COE breakdown"),
    ]
    for sheet, coord, old, new in stale:
        w = wb[sheet]
        v = w[coord].value
        if isinstance(v, str) and old in v:
            w[coord].value = v.replace(old, new)
            log("A11", "%s!%s" % (sheet, coord), "%r -> %r" % (old, new))

    # ---------------------------------------------------------------- A10
    log.head("A10  his tab name spelt right, and the label count re-derived")
    if OLD31 in wb.sheetnames:
        hits = repoint_sheet(wb, OLD31, NEW31)
        wb[OLD31].title = NEW31
        log("A10", "3.1 tab", "%r -> %r (%d refs repointed)" % (OLD31, NEW31, hits))
    t31 = wb[NEW31]
    lbl = ('="TDD total ( "&TEXT(COUNTA(\'' + REVIEW +
           "'!$B$2:$B$700),\"0\")&\" roles in consolidated role mapping)\"")
    if t31["B21"].value != lbl:
        log("A10", "%s!B21" % NEW31,
            "%r -> live count, 'consolidated role mapping'" % t31["B21"].value)
        t31["B21"].value = lbl

    # ---------------------------------------------------------------- A6
    log.head("A6  3.2 controls restored, white font")
    o = wb["3.2 Overhead & Leadership"]
    if o["B14"].value is None:
        o["B14"].value = ("Control - Times applied set above, against the count the "
                          "model carries (the ten portfolios and the platforms the "
                          "1.x tabs draw), must be 0")
        log("A6", "3.2!B14", "control label restored")
    if o["I14"].value is None:
        o["I14"].value = "=ROUND($I11-SUMPRODUCT(Lists!$AG$2:$AG$7,Lists!$AH$2:$AH$7),6)"
        log("A6", "3.2!I14", "control formula restored")
    if o["B18"].value is None:
        o["B18"].value = ("Control - the portfolios plus the COEs and EGI against "
                          "the role mapping, must be 0")
        log("A6", "3.2!B18", "control label restored")
    if o["B14"].font.color is None or o["B14"].font.color.rgb != "FFFFFFFF":
        white(o, "B14", "I14", "B18", "G18")
        log("A6", "3.2!B14 I14 B18 G18", "white font FFFFFFFF (invisible, functional)")

    # ---------------------------------------------------------------- A9
    log.head("A9  1.3 row 30 completed as a net new squad")
    e13 = wb["1.3 Enterprise Data"]
    if e13["H30"].value is None:
        bump = lambda sh, c1, r1, c2, r2: (
            (c1, r1 + 1, c2, None if r2 is None else r2 + 1)
            if sh == "1.3 Enterprise Data" and r1 == 29 else (c1, r1, c2, r2))
        for col in "HIJKL":
            src = e13["%s29" % col].value
            e13["%s30" % col].value = rewrite_refs(src, "1.3 Enterprise Data", bump)
            copy_style(e13["%s29" % col], e13["%s30" % col])
        log("A9", "1.3!H30:L30",
            "archetype machinery for Enterprise Data Delivery (EDI|M|Offshore, support 0.15)")
    if e13["K32"].value == "=SUM($K27,$K28,$K29)":
        e13["K32"].value = "=SUM($K27,$K28,$K29,$K30)"
        log("A9", "1.3!K32", "row 30 into the platform total")
    if "COUNT($H27,$H28,$H29)=3" in str(e13["L32"].value):
        e13["L32"].value = str(e13["L32"].value).replace(
            "COUNT($H27,$H28,$H29)=3", "COUNT($H27,$H28,$H29,$H30)=4")
        log("A9", "1.3!L32", "row 30 into the variance guard")

    # ---------------------------------------------------------------- A2
    log.head("A2  the Hold hacks come off")
    g4 = wb["2.4 TDD Group Functions"]
    for r in range(52, 56):
        if g4.cell(r, 5).value == "Hold":
            g4.cell(r, 5).value = "Filled"
            log("A2", "2.4 TDD Group Functions!E%d" % r, "Hold -> Filled")
    if g4["F51"].value == 0:
        g4["F51"].value = "=SUM(F52:F55)"
        log("A2", "2.4 TDD Group Functions!F51", "0 -> =SUM(F52:F55)")
    p5 = wb["2.5 P&C"]
    if p5["E28"].value == "Hold":
        p5["E28"].value = "Filled"
        log("A2", "2.5 P&C!E28", "Hold -> Filled (EGI P&C row)")

    # ---------------------------------------------------------------- A4a
    log.head("A4  named vacancies are filled; the mis-wire and the two removals")
    if rv["B418"].value == "Nico Lender":
        rv["B418"].value = "Vacant"
        log("A4", "REVIEW!B418", "'Nico Lender' -> 'Vacant' (413 is the one he fills)")
    fixes = [("2.4 TDD Group Functions", "B71", 517, "Hardik Trivedi"),
             ("2.5 P&C", "B39", 465, None),
             ("2.5 P&C", "B41", 503, "Vikram Chhahira"),
             ("2.5 P&C", "B50", 510, "Dave O'Keefe"),
             ("2.4 TDD Group Functions", "B38", 476, None)]
    for sheet, coord, rr, name in fixes:
        w = wb[sheet]
        typed = w[coord].value
        f = "='" + REVIEW + "'!$B$%d" % rr
        if isinstance(typed, str) and not typed.startswith("="):
            w[coord].value = f
            log("A4", "%s!%s" % (sheet, coord),
                "%r -> the proper cell $B$%d" % (typed, rr))
        if name and rv.cell(rr, 2).value != name:
            log("A4", "REVIEW!B%d" % rr,
                "%r -> %r (his name, in the ledger)" % (rv.cell(rr, 2).value, name))
            rv.cell(rr, 2).value = name
    f6 = wb["2.6 Finance"]
    if f6["B43"].value == "='" + REVIEW + "'!$B$500":
        log("A4", "2.6 Finance!B43", "%r -> $B$501 (mis-wire)" % f6["B43"].value)
        f6["B43"].value = "='" + REVIEW + "'!$B$501"

    # ---------------------------------------------------------------- A1
    log.head("A1  Enterprise Data gets its directly funded block")
    e23 = wb["2.3 Enterprise Data"]
    if e23["B13"].value != "EGI":
        old = read_block(e23, 28, 64)
        shift_rows(wb, "2.3 Enterprise Data", 12, 2)
        e23 = wb["2.3 Enterprise Data"]
        row_style(e23, 16, 12, 2, 18)
        e23["B12"].value = "Directly funded programs and platforms"
        row_style(e23, 15, 13, 2, 18)
        e23["B13"].value = "EGI"
        T = "'1.3 Enterprise Data'"
        e23["C13"].value = ('=IFERROR(IF(INDEX(%s!$C$25:$C$38,MATCH($B13,%s!$B$25:$B$38,0))'
                            '="","",INDEX(%s!$C$25:$C$38,MATCH($B13,%s!$B$25:$B$38,0))),'
                            '"Not on the 1.x tab")' % (T, T, T, T))
        e23["D13"].value = ('=IFERROR(IF(INDEX(%s!$D$25:$D$38,MATCH($B13,%s!$B$25:$B$38,0))'
                            '="","",INDEX(%s!$D$25:$D$38,MATCH($B13,%s!$B$25:$B$38,0))),"")'
                            % (T, T, T, T))
        e23["E13"].value = '=""'
        e23["G13"].value = ("=ROUND(SUMIFS('%s'!$O$2:$O$534,'%s'!$AJ$2:$AJ$534,$C$3,"
                            "'%s'!$AT$2:$AT$534,$B13),2)" % (REVIEW, REVIEW, REVIEW))
        e23["M13"].value = "=$F13-$L13"
        e23["N13"].value = "=$O13"
        e23["O13"].value = ("=SUMIFS('%s'!$AA$2:$AA$534,'%s'!$AJ$2:$AJ$534,$C$3,"
                            "'%s'!$AT$2:$AT$534,$B13)/1000000" % (REVIEW, REVIEW, REVIEW))
        e23["P13"].value = '=IFERROR(ROUND($O13-$N13,6),"")'
        e23["R13"].value = '=IFERROR(ROUND($Q13-$N13,6),"")'
        log("A1", "2.3 Enterprise Data!B12:R13",
            "'Directly funded programs and platforms' + EGI row (N = $O)")

        lev = {rr: lv for k, nm, rr, lv in old if k == "role"}
        for rr in EGI8:
            fill = str(rv.cell(rr, 2).value or "").lower().find("vacant") < 0
            lev[rr] = "Filled" if fill else "Hire"
            log("A1", "2.3 Enterprise Data lever (REVIEW %d)" % rr,
                "Hold -> %s (ruling 2, the Hold comes off)" % lev[rr])
        groups, cur = [], None
        for kind, nm, rr, lv in old:
            if kind == "group":
                cur = (nm, [])
                groups.append(cur)
            elif rr not in EGI8:
                cur[1].append((rr, lev[rr]))
        groups.append(("EGI", [(rr, lev[rr]) for rr in sorted(EGI8)]))
        order = ["Data Platforms", "Data Science", "EGI", "Leadership",
                 "Reporting & Analytics", "Head of Technology", "Technology Manager"]
        groups.sort(key=lambda g: order.index(g[0]))
        ranges, endrow = write_block(e23, 30, groups, (30, 31), 70)
        summary = {"Data Platforms": 8, "Data Science": 9, "Reporting & Analytics": 10,
                   "EGI": 13, "Leadership": 15, "Head of Technology": 17,
                   "Technology Manager": 19}
        for nm, row in summary.items():
            a, z = ranges[nm]
            wire_summary(e23, row, a, z)
        e23["F13"].value = '=COUNTIF($B${a}:$B${b},"?*")'.format(
            a=ranges["EGI"][0], b=ranges["EGI"][1])
        e23["Q13"].value = "=SUM($G${a}:$G${b})/1000000".format(
            a=ranges["EGI"][0], b=ranges["EGI"][1])
        log("A1", "2.3 Enterprise Data!B30:G%d" % endrow,
            "FTE block rebuilt, 7 groups, his lever states carried")
        picks = "E8:E10,E13:E13,E15:E15,E17:E19,E21"
        for col in "EFGHIJKLMOQ":
            e23["%s23" % col].value = "=SUM(%s)" % picks.replace("E", col)
        e23["N23"].value = ('=IF(COUNT(%s)=0,"",SUM(%s))'
                            % (picks.replace("E", "N"), picks.replace("E", "N")))
        e23["P23"].value = '=IF(ISNUMBER($N23),ROUND($O23-$N23,6),"")'
        e23["R23"].value = '=IF(ISNUMBER($N23),ROUND($Q23-$N23,6),"")'
        log("A1", "2.3 Enterprise Data!E23:R23", "EGI row into every portfolio total")
        rows = [r for a, z in ranges.values() for r in range(a, z + 1)]
        set_dv(e23, sorted(rows))

    # 3.3 carries every squad on every lever modelling tab, so it carries this one
    r33 = wb["3.3 Squad Actuals to Archetype"]
    egi33 = "='2.3 Enterprise Data'!$B$13"
    if not any(r33.cell(r, 4).value == egi33 for r in range(30, 50)):
        shift_rows(wb, "3.3 Squad Actuals to Archetype", 37, 1)
        r33 = wb["3.3 Squad Actuals to Archetype"]
        row_style(r33, 36, 37, 2, 15)
        r33.cell(37, 2).value = "Enterprise Data"
        r33.cell(37, 3).value = "Directly funded"
        for col, src in zip(range(4, 16),
                            ["B", "C", "D", "E", "F", "H", "I", "M", "N", "O", "P", "Q"]):
            r33.cell(37, col).value = "='2.3 Enterprise Data'!$%s$13" % src
        log("A1", "3.3 Squad Actuals to Archetype row 37",
            "Enterprise Data / Directly funded / EGI line added")
    q = wb["4.0 Data QA"]
    if "'2.3 Enterprise Data'!$N$13" not in str(q["D17"].value):
        q["D17"].value = str(q["D17"].value) + "+N('2.3 Enterprise Data'!$N$13)"
        log("A1", "4.0 Data QA!D17", "directly funded list gains 2.3's EGI line")
    c41 = "='2.3 Enterprise Data'!$N$23+N('1.3 Enterprise Data'!$H$30)"
    d41 = "='1.3 Enterprise Data'!$F$9+N('2.3 Enterprise Data'!$N$13)"
    if q["C41"].value != c41:
        q["C41"].value, q["D41"].value = c41, d41
        log("A1", "4.0 Data QA!C41 D41",
            "identity restated: 1.3 carries a net new squad 2.3 has no row for, "
            "2.3 carries a directly funded line 1.3 has no block for")

    # ---------------------------------------------------------------- A12
    log.head("A12  his grid notes move to the ledger commentry")
    notes = [("2.2 Customer", "H54", "B54"), ("2.2 Customer", "H78", "B78"),
             ("2.6 Finance", "H45", "B45"), ("2.13 COE SA&D", "H29", "B29")]
    import re as _re
    for tab, ncoord, bcoord in notes:
        w = wb[tab]
        txt = w[ncoord].value
        if not isinstance(txt, str) or not txt.strip():
            continue
        ref = str(w[bcoord].value or "")
        m = _re.search(r"\$(\d+)$", ref)
        if not m:
            log("A12", "%s!%s" % (tab, ncoord), "STOP: no ledger row behind %s" % bcoord)
            raise SystemExit(2)
        r = int(m.group(1))
        cur = rv.cell(r, 50).value
        new = txt if not cur else ("%s %s" % (cur, txt))
        rv.cell(r, 50).value = new
        w[ncoord].value = None
        log("A12", "%s!%s -> REVIEW!AX%d" % (tab, ncoord, r), "%r (his words)" % txt)
    s13 = wb["2.13 COE SA&D"]
    for co in ("O14", "P14", "O15", "P15"):
        if s13[co].value is not None:
            log("A12", "2.13 COE SA&D!%s" % co, "%r deleted (stray)" % s13[co].value)
            s13[co].value = None

    # ---------------------------------------------------------------- A4b
    log.head("A4  the two 'Remove' roles leave the ledger")
    for tab, coord, rr in [("2.6 Finance", "B41", 498),
                           ("2.9 Commercial Fuels", "B59", 420)]:
        w = wb[tab]
        row = int(coord[1:])
        if str(w[coord].value).strip().lower() == "remove" or \
                ("$%d" % rr) in str(w.cell(row, 3).value or ""):
            shift_rows(wb, tab, row, -1)
            log("A4", "%s row %d" % (tab, row),
                "helper row for REVIEW %d removed with the role" % rr)
    for rr in sorted([498, 420], reverse=True):
        nm = rv.cell(rr, 2).value
        ttl = rv.cell(rr, 3).value
        shift_rows(wb, REVIEW, rr, -1)
        log("A4", "REVIEW row %d" % rr,
            "role deleted from the ledger (%r / %r)" % (nm, ttl))
    rv = wb[REVIEW]
    log("A4", "REVIEW", "roles now %d" % len(ledger(rv)))

    # ------------------------------------------------------------ A7 / A8 / A3
    log.head("A7 / A8 / A3  what stays exactly as he left it")
    hidden = ["Exec Summary", "4.0 Data QA", "Lists", "DATA >>",
              "0.1 Budget Table (Fin)", "0.4 Presentation Pack"]
    bad = [h for h in hidden if wb[h].sheet_state != "hidden"]
    if bad:
        log("A7", "hidden tabs", "STOP: %r are not hidden" % bad)
        raise SystemExit(2)
    log.note("A7", "his six hidden tabs stay hidden; no model health strip added")
    sad = wb["1.12 SA&D"]
    if sad["H37"].value != "Hold" or sad["I37"].value is not None:
        log("A8", "1.12 SA&D!H37", "STOP: his Hold or his deleted note has moved")
        raise SystemExit(2)
    log.note("A8", "1.12 SA&D H37 Hold kept; his deleted note stays deleted")
    bad = []
    for ws in wb.worksheets:
        if not ws.title.startswith("2."):
            continue
        for r in range(1, ws.max_row + 1):
            d = ws.cell(r, 4).value
            if not (isinstance(d, str) and REVIEW in d and "$AK$" in d):
                continue                      # not a role row in the FTE block
            v = ws.cell(r, 5).value
            if v not in LEVERS:
                bad.append("%s!E%d=%r" % (ws.title, r, v))
    if bad:
        log("A3", "lever cells", "STOP: not one of the four values: %r" % bad[:6])
        raise SystemExit(2)
    log.note("A3", "every other lever cell keeps his value")

    # ---------------------------------------------------------------- G3
    log.head("G3  every REVIEW range this build touches runs to row 700")
    n = extend_review(wb, 700)
    log("G3", "workbook", "%d formulas now read REVIEW $2:$700" % n)

    save(wb, dst)
    log.tail()
    print("wrote", dst)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
