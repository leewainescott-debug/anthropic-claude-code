"""1.14 TDD Cyber - the design tab for the one-platform, one-squad cyber portfolio.

The owner's instruction: "build a tdd cyber 1.x tab ... exactly the same [as] the 1.x tabs
for portfolios, however, it just has one platform and one squad. platform is called tdd
cyber, squad is cyber uplift."

The tab that carried this name until now was a dead copy of 1.9 Commercial Fuels - its own
title said TDD Cyber while every figure on it, and every cell reference under it, read
Commercial Fuels' budget row and Commercial Fuels' three platforms. It is rebuilt here
from scratch rather than patched, because nothing on it was right except the name.

Two templates, both read at runtime rather than transcribed, so the new tab cannot drift
away from the family the next time a sibling restyles it:

  1.3 Enterprise Data   the block layout - Portfolio Summary, the budget box, Other
                        funding, Total to fund - and every style, fill, border and number
                        format those blocks carry. 1.3 is the closest sibling in shape:
                        one funded platform, the reduced Other funding table, no budget
                        reconciliation block.
  1.9 Commercial Fuels  the squad row itself - the archetype formula, the TDD/Funded split
                        beside it, the cream input cells and their four dropdowns. Its own
                        squad row is row 26 and so is this one, so the formula copies over
                        with no renumbering at all.

Two decisions the shape states rather than hides:

  the portfolio overhead is NOT drawn. TDD Cyber is not one of the ten funded portfolios
  on 0.2 Data Config, so it draws no Head of Technology / Business Partner / Domain
  Architect / Leadership allowance. Row 6 is zero on every column and the reason is
  written at B11, in words, under the block it explains.
  the platform overhead IS drawn - one platform, one 0.2!N16 - and it is the whole of the
  tab's cost until the owner sizes the squad.

What this tab deliberately does NOT carry: the actuals footer and the K/L columns
(actuals.py builds those), and the Actuals figure at G9 with the variance line under it
(post2707.py writes those). The header cell at G5 and the bar over it are here so those
two steps have a column to write into.

Shipped, with the squad's cream cells still empty: F7 and F9 read 0.165, H26 reads
"check size" - the family's own convention for a squad whose type and size are not set
yet - and no cell on the tab is an error.

The empty 2.15 TDD Cyber shell is created here too, so the pair exists from this point in
the chain onward; final2x.py builds its contents.
"""
import copy

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.datavalidation import DataValidation

TAB = "1.14 TDD Cyber"
WORK = "2.15 TDD Cyber"
T3 = "1.3 Enterprise Data"          # block layout and styles
T9 = "1.9 Commercial Fuels"         # the squad row: formulas, cream cells, dropdowns
CFG = "'0.2 Data Config'"
CFGROW = 23                         # 0.2's "TDD Cyber incl. COE - Cyber, Risk & Service Ops"

# The shipped column profile, the same on all thirteen design tabs. K and L are the two
# columns actuals.py writes into and are sized for it here so the tab does not resize
# under the reader later.
WIDTHS = {"A": 2.0, "B": 40.0, "C": 27.0, "D": 9.5, "E": 17.8, "F": 10.0, "G": 11.0,
          "H": 43.5, "I": 14.0, "J": 45.0, "K": 24.8, "L": 25.5}

SQUAD = "Cyber Uplift"
PLATFORM = "TDD Cyber"

WHY_NO_PF_OVERHEAD = (
    "TDD Cyber does not draw a portfolio overhead - it is not one of the ten funded "
    "portfolios on 0.2 Data Config. The platform overhead is drawn. No roles in the "
    "role mapping carry TDD Cyber yet - today's cyber roles are on 1.13, and the "
    "working tab fills itself as roles move here.")
WHY_BUDGET = "This is the same TDD Cyber budget that 1.13 Cyber Roles draws on."
HOW_TO_PRICE = ("Set the Squad Type, Size, On/Off and Support % and this squad prices "
                "itself")

# the four family dropdowns plus the country list, and the cell each one lands on
DVS = [("SquadTypes", "C26"), ("SquadSizes", "D26"), ("OnOff", "E26"),
       ("SupportPct", "G26"), ('"AU,NZ"', "F26")]


def wipe(ws):
    """Everything off the tab. The old 1.14 was a Commercial Fuels copy down to its
    merges and its column widths; anything left behind would be a figure from another
    portfolio sitting on this one."""
    blank = copy.copy(openpyxl.cell.cell.Cell(ws)._style)
    for m in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(m))
    for row in ws.iter_rows():
        for c in row:
            c.value = None
            c._style = copy.copy(blank)
    for k in list(ws.column_dimensions):
        del ws.column_dimensions[k]
    for k in list(ws.row_dimensions):
        del ws.row_dimensions[k]
    ws.conditional_formatting._cf_rules.clear()
    ws.data_validations.dataValidation = []
    ws.freeze_panes = None
    ws.sheet_view.showGridLines = False


def build(wb, out):
    ws, t3, t9 = wb[TAB], wb[T3], wb[T9]
    wipe(ws)
    for col, w in WIDTHS.items():
        ws.column_dimensions[col].width = w
    ws.sheet_properties.tabColor = copy.copy(t3.sheet_properties.tabColor)

    def put(coord, src, src_coord, value=None):
        """One cell: the template's style, and a value where this tab has one."""
        ws[coord]._style = copy.copy(src[src_coord]._style)
        if value is not None:
            ws[coord] = value

    def like(coord, src, src_coord):
        """The template's style and the template's own text - a label the family owns."""
        put(coord, src, src_coord, src[src_coord].value)

    def note(coord, text):
        c = ws[coord]
        c.value = text
        c.font = Font(name="Calibri", size=10)
        c.alignment = Alignment(horizontal="left", vertical="center")

    def bar(coord, value=None):
        put(coord, t3, "B4", value)      # 1.3!B4 is the one true bar: FF002F6C, white bold

    # ---- title -------------------------------------------------------------------
    put("B2", t3, "B2", PLATFORM)
    ws.row_dimensions[2].height = t3.row_dimensions[2].height

    # ---- Portfolio Summary ---------------------------------------------------------
    # The bar runs the width of the table under it, which on this tab is B to G: the
    # Actuals column is part of the summary here, so the strip covers F and G too rather
    # than stopping at the merge.
    like("B4", t3, "B4")
    for c in "CDEFG":
        bar(f"{c}4")
    bar("H4", t3["H4"].value)
    bar("I4")
    ws.row_dimensions[4].height = 19

    for c in "BCDEF":
        like(f"{c}5", t3, f"{c}5")
    put("G5", t3, "F5", "Actuals")       # post2707 rewrites the value; the column is here
    ws.row_dimensions[5].height = t3.row_dimensions[5].height

    # row 6 - no portfolio overhead. Zero in each column rather than a blank, so the row
    # is a stated nil and F6 stays a live sum of three cells like its ten siblings.
    like("B6", t3, "B6")
    for c, f in (("C", "=0"), ("D", "=0"), ("E", "=0"), ("F", "=C6+D6+E6")):
        put(f"{c}6", t3, f"{c}6", f)

    # row 7 - one platform, so one I-cell in the SUM. The AU/NZ branch is the family's,
    # written exactly the way its ten siblings write it: if the portfolio's NZ budget on
    # 0.2 is the bigger of the two, the cost lands in the NZ column, else in AU.
    like("B7", t3, "B7")
    put("C7", t3, "C7",
        f'=IF({CFG}!$D${CFGROW}>{CFG}!$C${CFGROW},0,SUM(I27))')
    put("D7", t3, "D7",
        f'=IF({CFG}!$D${CFGROW}>{CFG}!$C${CFGROW},SUM(I27),0)')
    put("E7", t3, "E7", "=0")
    put("F7", t3, "F7", "=C7+D7+E7")

    # row 8 - one squad, so a one-cell SUMIF range on each side
    like("B8", t3, "B8")
    put("C8", t3, "C8", '=SUMIF(F26:F26,"AU",I26:I26)')
    put("D8", t3, "D8", '=SUMIF(F26:F26,"NZ",I26:I26)')
    put("E8", t3, "E8", "=SUM(J26)")
    put("F8", t3, "F8", "=C8+D8+E8")

    like("B9", t3, "B9")
    for c in "CDEF":
        put(f"{c}9", t3, f"{c}9", f"=SUM({c}6:{c}8)")
    # G9 and the E10/F10 variance line are post2707's - left empty, not left wrong

    # ---- Budget vs TDD Cost --------------------------------------------------------
    like("H5", t3, "H5")
    put("I5", t3, "I5", f"={CFG}!$E${CFGROW}")
    like("H6", t3, "H6")
    put("I6", t3, "I6", f"={CFG}!$C${CFGROW}")
    like("H7", t3, "H7")
    put("I7", t3, "I7", f"={CFG}!$D${CFGROW}")
    like("H8", t3, "H8")
    put("I8", t3, "I8", "=C9-I6")
    like("H9", t3, "H9")
    put("I9", t3, "I9", "=D9-I7")
    like("H10", t3, "H10")
    put("I10", t3, "I10", "=I8+I9")

    note("B11", WHY_NO_PF_OVERHEAD)
    note("H11", WHY_BUDGET)

    # ---- Other funding -------------------------------------------------------------
    # 1.3's reduced form: no Significant Items line, because there is no EGI here. The
    # funded line is the cyber uplift bucket the owner set aside, and it is a typed input.
    bar("H13", "Other funding")
    bar("I13")
    bar("J13")
    ws.row_dimensions[13].height = 19
    for c in "HIJ":
        like(f"{c}14", t3, f"{c}13")
    for c in "HIJ":
        put(f"{c}15", t3, f"{c}14")
    ws["H15"] = t3["H14"].value
    ws["J15"] = "=E9"
    for c in "HIJ":
        put(f"{c}16", t3, f"{c}15")
    ws["H16"] = "Funded from the cyber uplift bucket ($m)"
    ws["J16"] = 0
    for c in "HIJ":
        put(f"{c}17", t3, f"{c}17")
    ws["H17"] = t3["H17"].value
    ws["J17"] = "=J15-J16"

    # ---- Total to fund -------------------------------------------------------------
    like("B16", t3, "B15")
    put("C16", t3, "C15")
    put("B17", t3, "B16")
    put("C17", t3, "C16")
    like("B18", t3, "B17")
    put("C18", t3, "C17", "=I8+I9")
    like("B19", t3, "B18")
    put("C19", t3, "C18", "=J17")
    like("B20", t3, "B19")
    put("C20", t3, "C19", "=C18+C19")

    # ---- Platform: TDD Cyber -------------------------------------------------------
    for c in "BCDEFGHIJ":
        bar(f"{c}24")
    ws["B24"] = f"Platform: {PLATFORM}"
    ws.row_dimensions[24].height = 19

    for c in "BCDEFGHIJ":
        like(f"{c}25", t9, f"{c}25")     # K25/L25 stay empty - actuals.py heads them

    # The squad row. 1.9's own squad row is row 26, so its archetype formula - the
    # Onshore / Hybrid / Offshore lookup into 0.3 - copies across verbatim, with the cream
    # cells left empty for the owner. Until he sets them H26 reads "check size", which is
    # what every unsized squad in this file reads.
    for c in "BCDEFGHIJ":
        put(f"{c}26", t9, f"{c}26")
    ws["B26"] = SQUAD
    ws["F26"] = "AU"
    for c in ("H", "I", "J"):
        ws[f"{c}26"] = t9[f"{c}26"].value
    note("M26", HOW_TO_PRICE)

    for c in "BCDEFGHIJ":
        put(f"{c}27", t9, f"{c}28")
    ws["B27"] = t9["B28"].value           # "Platform Overhead"
    ws["I27"] = t9["I28"].value           # ='0.2 Data Config'!$N$16

    for c in "BCDEFGHIJ":
        put(f"{c}28", t9, f"{c}29")
    # 1.9's own total row leaves F unfilled, a hole in the grey band that the polish step
    # closes later. The band is written closed here rather than shipped broken.
    ws["F28"]._style = copy.copy(t9["E29"]._style)
    ws["B28"] = f"{PLATFORM} Total"
    ws["H28"] = "=SUM(H26:H26)"
    ws["I28"] = "=SUM(I26:I27)"
    ws["J28"] = "=SUM(J26:J26)"

    # ---- merges, last, so no cell is written through a MergedCell -------------------
    for rng in ("B4:E4", "H4:I4", "H13:J13", "B24:J24"):
        ws.merge_cells(rng)

    # ---- the four dropdowns, copied off 1.9 rather than re-declared -----------------
    src = {dv.formula1: dv for dv in t9.data_validations.dataValidation}
    for f1, cell in DVS:
        s = src.get(f1)
        if s is None:
            out.append(f"{TAB}: 1.9 carries no {f1} validation - {cell} left without one")
            continue
        d = DataValidation(type=s.type, operator=s.operator, formula1=s.formula1,
                           allow_blank=s.allow_blank, showDropDown=s.showDropDown,
                           showErrorMessage=s.showErrorMessage,
                           showInputMessage=s.showInputMessage,
                           errorTitle=s.errorTitle, error=s.error, errorStyle=s.errorStyle,
                           promptTitle=s.promptTitle, prompt=s.prompt)
        ws.add_data_validation(d)
        d.add(cell)
    out.append(f"{TAB} rebuilt: one platform, one squad ({SQUAD}), platform overhead "
               f"drawn, portfolio overhead not - layout off {T3}, squad row off {T9}")


def pair(wb, out):
    """The working copy's shell. final2x.py writes its contents; it only needs the sheet
    to exist, and it has to exist before the chain recalculates."""
    if WORK in wb.sheetnames:
        out.append(f"{WORK} already present - left for final2x")
        return
    prev = "2.14 EGI"
    idx = wb.sheetnames.index(prev) + 1 if prev in wb.sheetnames else None
    ws = wb.create_sheet(WORK, idx)
    ws.sheet_view.showGridLines = False
    if prev in wb.sheetnames:
        ws.sheet_properties.tabColor = copy.copy(wb[prev].sheet_properties.tabColor)
    out.append(f"{WORK} created empty after {prev} - final2x builds it")


def run(src, dst):
    wb = openpyxl.load_workbook(src)
    out = []
    if TAB not in wb.sheetnames:
        wb.create_sheet(TAB, wb.sheetnames.index(T9) + 1 if T9 in wb.sheetnames else None)
        out.append(f"{TAB} did not exist - created")
    build(wb, out)
    pair(wb, out)
    wb.save(dst)
    return out


if __name__ == "__main__":
    import sys

    for x in run(sys.argv[1], sys.argv[2]):
        print("  ", x)
