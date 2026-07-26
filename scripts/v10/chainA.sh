set -e
# Option A only - the owner picked it over B. Built from the shipped file, so the two differ
# only in where the comparison sits.
python3 actuals.py cand.xlsx wAA.xlsx A
python3 -c "
import wbio
rc,st=wbio.build('wAA.xlsx','cand_A.xlsx'); print('  A injected',st)
e,b=wbio.audit('cand_A.xlsx'); print('  A errors',len(e),'blank',len(b))
for x in e[:8]: print('    ',x)
"
