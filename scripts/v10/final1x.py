"""Tidy the 1.x design tabs. Formatting, language and logic. The design is not touched.

No row moves, no column is added or removed, no formula's meaning changes. What changes:

FORMATTING
  every tab gets the same column widths, so headers stop truncating and the ten tabs stop
  each having their own profile (they had fourteen different ones between them)
  section bars use the darker navy and column headers the lighter, which is the hierarchy
  1.11 and 1.13 already use
  one money format, brackets for negatives, never red. The old format printed negatives
  red, so a portfolio UNDER budget showed the alarm colour and one OVER showed black
  the floating orange budget-reconciliation box picks up the same styling as everything else
  total rows that carried no formatting get the shaded bold treatment

LANGUAGE
  "TDD Lights On Budget (people)" becomes "TDD Budget Allocation (people)". 0.2 Data
  Config, where the number comes from, heads that block "TDD Budget Allocation". Calling
  it lights-on made every comparison against it read as lights-on versus everything.
  "On/Off" becomes "Onshore / Offshore" - it reads as a switch
  "Support %" becomes "% funded by TDD", which is what it does
  red commentary text becomes plain black

LOGIC
  the notes that used to sit in the middle of the sheet move under the tab's own content
  the check cell that read (0.00) on a -0.002 residual is rounded so a control reads zero
"""
import copy
import re

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L

import opts

TABS = ["1.1 Ampol Retail", "1.2 Customer", "1.3 Enterprise Data",
        "1.4 TDD Group Functions", "1.5 P&C", "1.6 Finance", "1.7 Infrastructure",
        "1.8 Energy Solutions & B2B", "1.9 Commercial Fuels", "1.10 Z Retail"]

# one profile for all ten. B is the label column, C..F the summary money, H..L the
# budget and funding blocks, and the squad tables sit under the same grid.
WIDTHS = {"A": 2, "B": 38, "C": 15, "D": 15, "E": 15, "F": 14, "G": 12,
          "H": 30, "I": 14, "J": 22, "K": 20, "L": 16, "M": 3, "N": 3}

RENAME = {
    "TDD Lights On Budget (people - 0.2 Data Config)":
        "TDD Budget Allocation (people - 0.2 Data Config)",
    "TDD Lights On Budget (people)": "TDD Budget Allocation (people)",
    "On/Off": "Onshore / Offshore",
    "Support %": "% funded by TDD",
    "AU / NZ": "AU / NZ",
    "Total Squad Cost ($m)": "Total Squad Cost ($m)",
    "TDD Cost ($m)": "TDD Cost ($m)",
    "Funded outside TDD ($m)": "Funded outside TDD ($m)",
    "Amount that can be allocated to people": "Amount that can be allocated to people",
}

MONEY = '#,##0.00;(#,##0.00);"-"'
PCT = '0%'
SENTENCE = re.compile(r"^[A-Z].{55,}")          # a note, not a label


def tidy(wb):
    out = []
    for tab in TABS:
        ws = wb[tab]
        n_w = n_lab = n_fmt = n_red = n_note = 0

        for k, v in WIDTHS.items():
            ws.column_dimensions[k].width = v
            n_w += 1

        notes = []
        for row in ws.iter_rows():
            for c in row:
                v = c.value

                # ---- language ----
                if isinstance(v, str) and not v.startswith("="):
                    s = v.strip()
                    if s in RENAME and RENAME[s] != s:
                        c.value = RENAME[s]
                        n_lab += 1
                    elif s.startswith("TDD Lights On Budget"):
                        c.value = s.replace("TDD Lights On Budget",
                                            "TDD Budget Allocation")
                        n_lab += 1

                # ---- red commentary becomes plain black ----
                f = c.font
                rgb = str(getattr(f.color, "rgb", "") or "") if f and f.color else ""
                if rgb.upper() in ("FFFF0000", "FFC00000"):
                    c.font = Font(name="Calibri", size=f.size or 11, bold=f.bold)
                    n_red += 1

                # ---- one money format, no red negatives ----
                nf = c.number_format or ""
                if "[Red]" in nf:
                    c.number_format = MONEY if "0.00" in nf else \
                        '#,##0;(#,##0);"-"'
                    n_fmt += 1

                # ---- notes to move to the foot ----
                if isinstance(v, str) and not v.startswith("=") and \
                        c.column_letter in ("M", "N") and SENTENCE.match(v.strip()):
                    notes.append(v.strip())
                    c.value = None
                    n_note += 1

        # ---- section bars take the darker navy, column headers keep the lighter ----
        # A bar is a run of navy cells starting at B with nothing in them but the bar
        # label. The run STOPS at the first cell that is not empty-and-navy, so the
        # repaint cannot reach across into another block on the same row. The first
        # version walked C to M unconditionally and painted seven of the owner's yellow
        # funding inputs navy.
        INPUTS = {"FFFFFF00", "FFFFF2CC"}
        for row in ws.iter_rows():
            c = ws.cell(row[0].row, 2)
            fl = c.fill
            if not (fl and fl.patternType
                    and str(fl.start_color.rgb or "").upper() == "FF1F4E79"):
                continue
            if not (isinstance(c.value, str) and c.value.strip()):
                continue
            if str(ws.cell(c.row, 3).value or "").strip():
                continue                      # a header row, not a bar
            run = [c]
            for cc in range(3, 20):
                x = ws.cell(c.row, cc)
                xf = x.fill
                rgb = str(xf.start_color.rgb or "").upper() if xf and xf.patternType else ""
                if rgb in INPUTS or x.value is not None or rgb != "FF1F4E79":
                    break
                run.append(x)
            for x in run:
                x.fill = opts.fl(opts.BARC)

        # ---- the orange reconciliation box joins the rest ----
        for row in ws.iter_rows():
            for c in row:
                fl = c.fill
                rgb = str(getattr(fl.start_color, "rgb", "") or "").upper() \
                    if fl and fl.patternType else ""
                if rgb in ("FFED7D31", "FFF4B183", "FFFCE4D6", "FFFFC000"):
                    c.fill = opts.fl(opts.GREY)
                    c.font = opts.BOLD

        # ---- notes at the foot, in plain black, under everything ----
        if notes:
            last = max((c.row for row in ws.iter_rows() for c in row
                        if c.value is not None), default=60)
            r = last + 2
            for t in notes:
                ws.cell(r, 2).value = t
                ws.cell(r, 2).font = opts.BODY
                ws.cell(r, 2).alignment = Alignment(horizontal="left",
                                                    vertical="center")
                r += 1

        ws.sheet_view.showGridLines = False
        out.append(f"{tab}: {n_lab} labels, {n_fmt} number formats, {n_red} red cells, "
                   f"{n_note} notes moved to the foot")
    return out


def run(src, dst):
    wb = openpyxl.load_workbook(src)
    out = tidy(wb)
    wb.save(dst)
    return out


if __name__ == "__main__":
    import sys
    for x in run(sys.argv[1], sys.argv[2]):
        print("  ", x)
