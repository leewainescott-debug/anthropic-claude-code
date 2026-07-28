"""Three agreed role moves, made through the visible override table on Lists.

The owner's raw columns on REVIEW are never edited. The grouping columns read an override
table first and fall back to the raw data, so every reassignment is one visible row a
reader can check, and the ledger still totals $115,113,262.27.

  r136 Jasper Na       squad Energy -> Ampol Web
        Associate Engineer - BE reporting to Jin Zhong, Lead Engineer - BE, who sits in
        Ampol Web; the only other person under Jin Zhong is also Ampol Web. "Energy" was
        a one-person squad that appears on no design tab.

Two further moves were proposed and withdrawn on the owner's instruction: Viren Khatri
stays in Ampol Retail, which is his home, and Vikram Chhahira stays in the EGI P&C squad,
which is a squad in its own right. Neither is in the table.

The table was three rows wide and hardcoded at $AN$2:$AN$4. It is now ten, and it carries
an overhead-line override as well as portfolio and squad, because moving a role onto an
overhead line has to change REVIEW!AR or the role lands in a delivery block under an
overhead name.
"""
import re

import openpyxl
from openpyxl.utils import get_column_letter as L

import opts

REVIEW = "REVIEW - Complete Role Mapping"
LAST = None                             # measured from the ledger in run()
N = 11                                  # override window runs Lists!$AN$2:$AN$11

# The three moves the owner instructed (D7). They lived as typed rows in the Lists
# override table; the review workbook's branch predates that table, so they are
# reinstated here or the roles silently snap back to their raw homes.
NEW = [(283, "COE SA&D", "Group Data", None),
       (313, "COE SA&D", "Group Data", None),
       (528, None, "SAP ERP", None)]

# Folds to take back out of Lists!W:X. The table is meant to fix typing - AmPos to AmPOS,
# Manuacturing to Manufacturing - and three rows in it were doing something else:
#
#   Customer, AI      -> Customer AI              a rename, not a typo. The owner's column
#                                                 K has the comma and so does his own list.
#   Z Energy Martech  -> Z Loyalty & Martech      two real squads merged into one
#   AU CRM & Martech  -> Ampol Loyalty & Martech  two real squads merged into one
#
# Merging them hid two of Customer's twelve squads completely: the owner's list has
# AU CRM & Martech at 2.0 FTE and Z Energy Martech at 2.0, and neither appeared anywhere in
# the workbook. Undoing the merge also puts Ampol Loyalty & Martech back to 6.8 and
# Z Loyalty & Martech back to 12.6, which is what his column K says.
UNFOLD = ["Customer, AI", "Z Energy Martech", "AU CRM & Martech"]

# The move of Jasper Na off the one-person Energy squad was mine, not the owner's, and his
# own reconciliation lists Energy at 1.0 FTE. It comes out. The three moves he did instruct -
# r283 and r313 to the Data COE, r528 to AU Finance - stay.
DROP_OVERRIDE = [136]

# A filled role priced at zero. REVIEW row 491, Nidhi Aggarwal, Snr Engineer - Boomi, NZ,
# FTE 1.0, status Filled. Column T carries her local base of 150,000 but U is empty, and the
# cost formula prices off U, so she counted in the 525 headcount and contributed nothing to
# the total. She is the only role in the ledger with T populated and U empty, and no check in
# the workbook could see it: all fifty-six reconcile TO the ledger, and a zero inside the
# ledger reconciles perfectly.
#
# The figure is derived from her own cohort, not invented. Every one of the thirteen other NZ
# roles in TDD Group Functions prices the same way: U = T x 0.92, then STI 0.15, pensions
# 0.05, CPI 0.03, no payroll component. 150,000 x 0.92 x 1.23 = 169,740. It goes in the
# cost-override column the formula already honours, so her raw columns stay untouched.
COST = {491: (169740.0,
              "Priced from her local base of 150,000 at the 0.92 rate her cohort uses, "
              "then STI 0.15, pensions 0.05, CPI 0.03 - the pattern all thirteen other NZ "
              "roles in this portfolio use. Column U was blank, so the formula read zero.")}
# build scaffolding and stray input colour left in the ledger
CLEAR_NOTES = [(191, 29), (491, 29)]
LABELS = {"AS1": "Squad name on the 1.x tab"}

PORTFOLIOS = ["Ampol Retail", "Customer", "Enterprise Data", "TDD Group Functions",
              "P&C", "Finance", "Infrastructure", "Energy Solutions & B2B",
              "Commercial Fuels", "Z Retail"]

# the keyword chain that classifies a position title into an overhead line. Unchanged;
# it just moves inside the override wrapper.
KEYWORD = ('IF(ISNUMBER(SEARCH("head of ",$C{r})),"Head of Technology",'
           'IF(ISNUMBER(SEARCH("TDD BP",$C{r})),"Business Partner",'
           'IF(OR(ISNUMBER(SEARCH("domain architect",$C{r})),'
           'ISNUMBER(SEARCH("enterprise architect",$C{r}))),"Domain Architect",'
           'IF(ISNUMBER(SEARCH("delivery man",$C{r})),"Delivery Manager",'
           'IF(OR(ISNUMBER(SEARCH("technology manager",$C{r})),'
           'ISNUMBER(SEARCH("technology manger",$C{r})),'
           'ISNUMBER(SEARCH("tech manager",$C{r}))),"Technology Manager","Squad"))))')


# Where a design tab and the ledger disagree on a squad name, the design tab is renamed -
# the ledger is the source of truth. This has to happen before the working tabs are built:
# they decide which section a squad belongs in by looking its name up on the design tab, and
# renaming afterwards left "Customer, AI" filed under "no archetype" with a 0.80 archetype
# sitting beside it.
SQUAD_RENAME = {"1.2 Customer": {"Customer AI": "Customer, AI"}}


def rename_squads(wb):
    out = []
    for tab, pairs in SQUAD_RENAME.items():
        ws = wb[tab]
        for r in range(1, min(ws.max_row, 95) + 1):
            v = ws.cell(r, 2).value
            if isinstance(v, str) and v.strip() in pairs:
                new = pairs[v.strip()]
                ws.cell(r, 2).value = new
                out.append(f"{tab}!B{r} {v.strip()!r} -> {new!r}, to match the ledger")
    return out


def ovr(col):
    """The override lookup for one column of the table."""
    return (f'IFERROR(INDEX(Lists!${col}$2:${col}${N},'
            f'MATCH(ROW(),Lists!$AN$2:$AN${N},0)),"")')


def formulas(r):
    o = ovr("AO")
    aj = (f'=IF(TRIM($B{r})="","",IF({o}<>"",{o},'
          f'IFERROR(INDEX(Lists!$U:$U,MATCH(TRIM($I{r}),Lists!$T:$T,0)),TRIM($I{r}))))')
    q = ovr("AQ")
    ar = f'=IF(TRIM($B{r})="","",IF({q}<>"",{q},{KEYWORD.format(r=r)}))'
    p = ovr("AP")
    at = (f'=IF(TRIM($B{r})="","",IF({p}<>"",{p},'
          f'IF(OR(LEFT($AJ{r},3)="COE",$AJ{r}="EGI"),$AP{r},'
          f'IF($AR{r}<>"Squad",$AR{r},$AP{r}))))')
    return {36: aj, 44: ar, 46: at}


def run(src, dst):
    global LAST
    wb = openpyxl.load_workbook(src)
    LAST = opts.ledger_last(wb)
    l = wb["Lists"]
    out = []

    # The columns the new table needs must be empty, or already hold this table from a
    # previous run - the build has to be re-runnable against its own output, or the second
    # run stops on the header the first one wrote.
    OURS = {43: "Overhead line override", 45: "Portfolios (10)"}
    for c, mine in OURS.items():
        if l.cell(1, c).value in (None, mine):
            continue
        raise SystemExit(f"Lists!{L(c)}1 holds {l.cell(1, c).value!r}, not "
                         f"{mine!r} - pick another column")

    # the table already carries the earlier agreed moves (r283 and r313 to the Data COE,
    # r528 to AU Finance). Append, never overwrite - writing over row 2 silently pulled
    # two roles back out of COE SA&D and moved a third out of SAP ERP.
    # ---- take the three merges back out of the fold table ----
    gone = []
    for r in range(2, 20):
        if str(l.cell(r, 23).value or "").strip() in UNFOLD:
            gone.append(f"{l.cell(r, 23).value} -> {l.cell(r, 24).value}")
            l.cell(r, 23).value = None
            l.cell(r, 24).value = None
    if gone:
        out.append("fold removed: " + "; ".join(gone))

    keep = []
    for r in range(2, 40):
        row = l.cell(r, 40).value
        if isinstance(row, (int, float)) and int(row) not in DROP_OVERRIDE:
            keep.append((int(row), l.cell(r, 41).value, l.cell(r, 42).value, None))
        elif isinstance(row, (int, float)):
            out.append(f"override removed: REVIEW row {int(row)} -> "
                       f"{l.cell(r, 42).value!r}")
    have = {k[0] for k in keep}
    moves = keep + [m for m in NEW if m[0] not in have]
    if len(moves) + 1 > N:
        raise SystemExit(f"{len(moves)} moves will not fit in {N - 1} slots")
    out.append("kept " + ", ".join(f"r{k[0]}" for k in keep))

    l.cell(1, 43).value = "Overhead line override"
    for c in (40, 41, 42, 43):
        x = l.cell(1, c)
        x.font, x.fill, x.alignment = opts.HDRF, opts.fl(opts.NAVY), opts.CEN
        l.column_dimensions[L(c)].width = 24
    for i, (row, pf, sq, oh) in enumerate(moves):
        r = 2 + i
        for c, v in ((40, row), (41, pf), (42, sq), (43, oh)):
            x = l.cell(r, c)
            x.value = v
            x.font, x.border = opts.BODY, opts.BOX
            x.fill = opts.fl(opts.YEL)
            x.alignment = opts.RGT if c == 40 else opts.LFT
    for r in range(2 + len(moves), N + 1):              # spare, declared and empty
        for c in (40, 41, 42, 43):
            x = l.cell(r, c)
            x.value = None
            x.font, x.border, x.fill = opts.BODY, opts.BOX, opts.fl(opts.YEL)
    l.cell(N + 2, 40).value = ("Agreed moves. REVIEW's own columns are untouched; these "
                               "override the grouping only.")
    l.cell(N + 2, 40).font = opts.BODY
    out.append(f"Lists override table: {len(moves)} moves, {N - 1} slots")

    # a clean list of the ten portfolios, so the allowance stops counting a block of
    # cells on 3.1 whose row numbers move every time the tab is rebuilt
    l.cell(1, 45).value = "Portfolios (10)"
    l.cell(1, 45).font = opts.HDRF
    l.cell(1, 45).fill, l.cell(1, 45).alignment = opts.fl(opts.NAVY), opts.CEN
    l.column_dimensions["AS"].width = 26
    for i, p in enumerate(PORTFOLIOS):
        x = l.cell(2 + i, 45)
        x.value, x.font, x.border = p, opts.BODY, opts.BOX
        x.alignment = opts.LFT
    out.append("Lists!AS2:AS11: the ten portfolios, named")

    # ---- repoint every reader of the old three-row window ----
    # every column of the table, not only the key column. Widening AN alone left
    # INDEX(Lists!$AO$2:$AO$4, MATCH(ROW(), Lists!$AN$2:$AN$11, 0)): a match at position
    # four indexes past the end of a three-row range, so an override in one of the new
    # slots silently returned nothing.
    old = re.compile(r"\$(AN|AO|AP|AQ)\$2:\$(AN|AO|AP|AQ)\$4\b")
    hits = 0
    for s in wb.sheetnames:
        for row in wb[s].iter_rows():
            for c in row:
                if isinstance(c.value, str) and old.search(c.value):
                    c.value = old.sub(lambda m: f"${m.group(1)}$2:${m.group(2)}${N}",
                                      c.value)
                    hits += 1
    out.append(f"widened override window in {hits} formulas")

    # ---- rebuild the three grouping columns ----
    R = wb[REVIEW]
    n = 0
    # every row, populated or not. The formula guards itself with IF(TRIM($B)="",""), and
    # skipping the empty rows left two of them carrying the old mismatched ranges, so a
    # role typed into row 191 would not have honoured an override.
    for r in range(2, LAST + 1):
        for col, f in formulas(r).items():
            R.cell(r, col).value = f
        n += 1
    out.append(f"REVIEW AJ / AR / AT rebuilt on {n} rows")

    # ---- a filled role priced at zero, and the scaffolding around it ----
    for row, (cost, why) in COST.items():
        R.cell(row, 47).value = cost                     # AU, the cost override
        R.cell(row, 47).number_format = '#,##0'
        R.cell(row, 48).value = why                      # AV, its provenance
        R.cell(row, 48).font = opts.BODY
        out.append(f"REVIEW row {row}: cost override {cost:,.0f} - {why[:48]}...")
    for row, col in CLEAR_NOTES:
        R.cell(row, col).value = None
    for ref, val in LABELS.items():
        R[ref].value = val
    out.append("REVIEW: build notes removed, AS1 relabelled")
    out += rename_squads(wb)

    # ---- every fixed window over the ledger follows the measured extent ----
    # The design tabs the owner edits carry SUMIFS/COUNTIFS over REVIEW with the end row
    # typed in - $AJ$2:$AJ$528, $AA$2:$AA$530 and friends. A ledger that grows past the
    # typed end silently drops the new rows out of every one of those formulas, which is
    # the worst failure mode this workbook has: wrong and green. Any range anchored at
    # row 2 whose typed end sits in 500..999 is a ledger window and follows LAST.
    win = re.compile(r"(\$[A-Z]{1,2})\$2:(\$[A-Z]{1,2})\$(5\d\d|[6-9]\d\d)\b")
    wided = 0
    for s_ in wb.sheetnames:
        for row in wb[s_].iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("=") and win.search(c.value):
                    new = win.sub(lambda m: f"{m.group(1)}$2:{m.group(2)}${LAST}", c.value)
                    if new != c.value:
                        c.value = new
                        wided += 1
    out.append(f"ledger windows widened to row {LAST} in {wided} formulas")

    # ---- the allowance stops counting rows on a summary tab ----
    fixed = 0
    pat = re.compile(r"COUNTA\('3\.1 Group Summary'!\$B\$\d+:\$B\$\d+\)")
    for s in wb.sheetnames:
        for row in wb[s].iter_rows():
            for c in row:
                if isinstance(c.value, str) and pat.search(c.value):
                    c.value = pat.sub("COUNTA(Lists!$AS$2:$AS$11)", c.value)
                    fixed += 1
    out.append(f"portfolio count keyed off the named list in {fixed} formulas")
    wb.save(dst)
    return out


if __name__ == "__main__":
    import sys
    for x in run(sys.argv[1], sys.argv[2]):
        print("  ", x)
