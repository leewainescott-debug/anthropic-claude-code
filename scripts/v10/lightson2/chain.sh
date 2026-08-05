set -e
cd /tmp/claude-0/-home-user-anthropic-claude-code/e550b440-3996-5abb-87e5-bafafe598f82/scratchpad
V=/home/user/anthropic-claude-code/scripts/v10
export PYTHONPATH=$V
SNAP=vstages/tw3/old_model.xlsx

rm -f buildw/w1.xlsx buildw/w1r.xlsx buildw/w2.xlsx buildw/w4.xlsx

echo "=== W1 START ==="
python3 vstages/w1_map.py /home/user/anthropic-claude-code/TDD_Cost_Calc.xlsx buildw/w1.xlsx
echo "=== W1 APPLIED ==="

python3 -c "import sys;sys.path.insert(0,'$V');import wbio,shutil; shutil.copy(wbio.recalc('buildw/w1.xlsx'),'buildw/w1r.xlsx'); print('w1 recalced')"
echo "=== W1 RECALCED ==="

W2_RECALC=1 python3 vstages/w2_lightson2.py buildw/w1r.xlsx buildw/w2.xlsx
echo "=== W2 APPLIED ==="

python3 vstages/w4_guard.py buildw/w2.xlsx buildw/w4.xlsx
echo "=== W4 APPLIED ==="

python3 vstages/w_gate.py buildw/w4.xlsx --prev $SNAP
echo "=== GATE 1 GREEN ==="
python3 vstages/w_gate.py buildw/w4.xlsx --prev $SNAP
echo "=== GATE 2 GREEN ==="

python3 vstages/w3_verify.py buildw/w4.xlsx
echo "=== W3 VERIFY DONE ==="
echo "=== CHAIN COMPLETE ==="
