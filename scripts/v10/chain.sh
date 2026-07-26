set -e
# base_ship.xlsx is the previous commit's workbook: git show HEAD~1:TDD_Cost_Calc.xlsx
cp base_ship.xlsx w0.xlsx
python3 overrides.py w0.xlsx w1.xlsx
python3 -c "import wbio,shutil; shutil.copy(wbio.recalc('w1.xlsx'),'w1r.xlsx')"
python3 final2x.py w1r.xlsx w2.xlsx
python3 final3x.py w2.xlsx w3.xlsx
python3 final4x.py w3.xlsx w4.xlsx
python3 final35.py w4.xlsx w5.xlsx
python3 fix1x.py w5.xlsx w6.xlsx w1r.xlsx
python3 fixcoe.py w6.xlsx w6a.xlsx
python3 purge.py w6a.xlsx w6b.xlsx
python3 polish.py w6b.xlsx w7.xlsx w1r.xlsx
python3 -c "
import wbio
rc,st=wbio.build('w7.xlsx','cand.xlsx'); print('  injected',st)
e,b=wbio.audit('cand.xlsx'); print('  errors',len(e),'blank',len(b))
for x in e[:10]: print('   ',x)
"
