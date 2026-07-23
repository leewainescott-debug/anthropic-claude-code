#!/usr/bin/env python3
"""Restore the correct old squad/platform mapping (base = clean_reroute, whose
working tabs and 3.3 already use the curated platform->squad structure with
Leadership in overhead) and fix the COE netting so 1.11/1.12, 3.1 and 3.4 agree:
the portfolio-funded Business Partner / Domain Architect amount is netted OFF the
COE spend, not ADDED to the COE budget. The group 'Less:' line is retired because
the netting now happens at each COE line (same total, no double count)."""
import openpyxl
from openpyxl.styles import Font
from openpyxl.worksheet.views import Selection

SRC="clean_reroute.xlsx"; OUT="clean_v5.xlsx"
wb=openpyxl.load_workbook(SRC)
IT=Font(italic=True, size=10)

# ---- 1.11 BP&T: budget = COE allocation only; planned spend = gross net of the
#      Business Partner FTEs funded inside portfolio overheads (C13) ----
b=wb["1.11 BP&T"]
b["C15"].value="=C14"                                   # Total BP budget = COE allocation only
b["F6"].value=('=SUMIFS($T$21:$T$44,$D$21:$D$44,"TDD Business Partner")'
               '+SUMIFS($T$21:$T$44,$D$21:$D$44,"Commercial")-C13')   # net spend
b["B9"].value=("Planned spend is net of the Business Partner FTEs funded inside portfolio "
               "overheads (row 13); the COE draws down its own allocation only.")
b["B9"].font=IT

# ---- 1.12 SA&D: same pattern with the Domain Architect amount (C13) ----
s=wb["1.12 SA&D"]
s["C15"].value="=C14"                                   # Total S&A budget = COE allocation only
s["G6"].value=('=SUM($T$22:$T$50)-SUMIFS($T$22:$T$50,$D$22:$D$50,"Group Data")-C13')  # net spend
s["B9"].value=("Planned spend is net of the Domain Architect FTEs funded inside portfolio "
               "overheads (row 13); the COE draws down its own allocation only.")
s["B9"].font=IT

# ---- 3.2 Total Cost is the total-cost view: keep the COE planned on a GROSS basis
#      and net the portfolio double count once at the group 'Less:' line (row 23,
#      unchanged). Because 1.11!F6 / 1.12!G6 are now net, add the netted amount back
#      so 3.2 stays gross and its grand total is unchanged. The funding view (net)
#      lives on 1.11/1.12/3.1/3.4. ----
tc=wb["3.2 Total Cost"]
tc["C16"].value="='1.11 BP&T'!$F$6+'1.11 BP&T'!$C$13"   # BP gross = net + netted-back
tc["C18"].value="='1.12 SA&D'!$G$6+'1.12 SA&D'!$C$13"   # S&A gross = net + netted-back

# ---- remove ALL frozen panes cleanly (owner instruction) ----
for w in wb.worksheets:
    w.freeze_panes=None; w.sheet_view.pane=None
    w.sheet_view.selection=[Selection(activeCell="A1",sqref="A1")]

wb.save(OUT)
print("saved",OUT)
print("1.11 F6 =",b["F6"].value)
print("1.11 C15 =",b["C15"].value)
print("1.12 G6 =",s["G6"].value)
print("3.2 C23 =",tc["C23"].value)
# confirm the correct mapping survived (Customer has no Leadership squad)
cust=[wb["2.2 Customer"][f"B{r}"].value for r in range(6,13)]
print("2.2 Customer squads:", [c for c in cust if c])