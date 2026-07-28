set -e
# The consolidation chain: the owner's review workbook becomes the base, his datasets load,
# and the standard chain rebuilds everything downstream. docs/ORCHESTRATION_2707.md is the
# plan this implements.
python3 assemble_base.py rev.xlsx b0.xlsx
python3 - <<'EOF'
import merge_review
wb, out, bad = merge_review.run("b0.xlsx", "b1.xlsx")
wb.save("b1.xlsx")
for x in out: print("  ", x)
# the gate: with the overrides in, the ledger's Customer block must equal the dataset's
# own total to the cent
import openpyxl, merge_review as mr
wb2 = openpyxl.load_workbook("b1.xlsx")
R = wb2[mr.REVIEW]
total = 0.0
for r in range(mr.CUST_LO, mr.CUST_HI + 1):
    raw = {c: R.cell(r, c).value for c in range(1, 27)}
    au = R.cell(r, 47).value
    total += au if isinstance(au, (int, float)) else mr.d8_cost(raw)
import sys
if abs(total - 16522075.33) > 0.05:
    sys.exit(f"Customer block prices at {total:,.2f}, dataset says 16,522,075.33 - stop")
print(f"   Customer block prices at {total:,.2f} == dataset total, gate passed")
EOF
python3 repair_design.py b1.xlsx b2.xlsx
python3 cyber14.py b2.xlsx b2c.xlsx
# after cyber14, so 1.14's copied squad row is swept with the other thirteen tabs: the
# Hybrid branch on every 1.x squad prices 2 roles onshore and the rest offshore, off a
# settable input on 0.3 (K7/K8 - the only two cells the chain writes on the owner's tab)
python3 hybrid.py b2c.xlsx b2h.xlsx
python3 ensure_lists.py b2h.xlsx base2.xlsx
cp base2.xlsx w0.xlsx
python3 overrides.py w0.xlsx w1.xlsx
python3 -c "import wbio,shutil; shutil.copy(wbio.recalc('w1.xlsx'),'w1r.xlsx')"
python3 final2x.py w1r.xlsx w2.xlsx
python3 final3x.py w2.xlsx w3.xlsx
python3 final4x.py w3.xlsx w4.xlsx
python3 final35.py w4.xlsx w5.xlsx
python3 fix1x.py w5.xlsx w6.xlsx w1r.xlsx
python3 fixcoe.py w6.xlsx w6a.xlsx
python3 purge.py w6a.xlsx w6b.xlsx
python3 polish.py w6b.xlsx w7a.xlsx w1r.xlsx
python3 design2707.py w7a.xlsx w7d.xlsx
python3 finish.py w7d.xlsx w7.xlsx
python3 -c "
import wbio
rc,st=wbio.build('w7.xlsx','cand.xlsx'); print('  injected',st)
e,b=wbio.audit('cand.xlsx'); print('  errors',len(e),'blank',len(b))
for x in e[:10]: print('    ',x)
"
