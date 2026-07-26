set -e
# both workbooks are built from the same shipped file, so they differ only in the design
for V in A B; do
  python3 actuals.py cand.xlsx wA$V.xlsx $V
  python3 -c "
import wbio
rc,st=wbio.build('wA$V.xlsx','cand_$V.xlsx'); print('  $V injected',st)
e,b=wbio.audit('cand_$V.xlsx'); print('  $V errors',len(e),'blank',len(b))
for x in e[:8]: print('    ',x)
"
done
