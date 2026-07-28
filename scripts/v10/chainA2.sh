set -e
# design-A actuals on the 1.x tabs, then the owner's 2707 adoptions that need the built
# blocks in place: the Actuals column on the Portfolio Summary tables, his live levers,
# the 0.2 COE spend repoints and the 1.13 bar
python3 actuals.py cand.xlsx pA.xlsx A
python3 post2707.py pA.xlsx pB.xlsx
python3 -c "
import wbio
rc,st=wbio.build('pB.xlsx','cand_A.xlsx'); print('  A injected',st)
e,b=wbio.audit('cand_A.xlsx'); print('  A errors',len(e),'blank',len(b))
for x in e[:10]: print('    ',x)
"
