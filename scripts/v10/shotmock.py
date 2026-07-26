import os,glob,re
import shots
OUT=os.path.join(shots.SP,"shots")
os.makedirs(OUT,exist_ok=True)
for f in sorted(glob.glob(os.path.join(shots.SP,"mocks","*.xlsx"))):
    k=os.path.basename(f)[:-5]
    p=shots.to_png(f,"opt_"+k)
    print(("  ok   " if p else "  FAIL ")+k)
