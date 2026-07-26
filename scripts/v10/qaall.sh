#!/bin/bash
# Five passes, in order of how much they can see. Run against a built workbook.
F=${1:-cand.xlsx}
echo "===== 1. formula errors and cached values (wbio)"
python3 -c "
import wbio; e,b=wbio.audit('$F')
print(f'   formula errors {len(e)}')
print(f'   formula cells with no cached value {len(b)} (REVIEW rows 191-192 are empty by design)')"
echo "===== 2. the workbook's own live checks (4.0 Data QA)"
python3 -c "
import openpyxl; ws=openpyxl.load_workbook('$F',data_only=True)['4.0 Data QA']
n=0
for r in range(5,ws.max_row+1):
    if ws.cell(r,2).value=='Checks failing':
        print(f'   {r-5} checks, {ws.cell(r,5).value} failing'); break"
echo "===== 3. adversarial (qa.py)"
python3 qa.py "$F" 2>&1 | tail -2 | sed 's/^/   /'
echo "===== 4. layout, coverage, banned words, bars, headers (verify.py)"
python3 verify.py "$F" 2>&1 | tail -2 | sed 's/^/   /'
echo "===== 5. every reader-visible figure rebuilt from the ledger (recompute.py)"
python3 recompute.py "$F" 2>&1 | tail -1 | sed 's/^/   /'
echo "===== lever, recalculated end to end"
python3 -c "
import verify
for x in verify.lever('$F'): print('  ',x)" 2>&1 | tail -10
