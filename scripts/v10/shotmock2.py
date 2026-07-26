import os,glob,re,openpyxl
import shots, wbio
for f in sorted(glob.glob(os.path.join(shots.SP,"mocks2","*.xlsx"))):
    k=os.path.basename(f)[:-5]
    # openpyxl writes formulas with no cached value, so a mock must be recalculated
    # before its sheets are extracted or every total renders blank
    rc=wbio.recalc(f); os.replace(rc,f)
    wb=openpyxl.load_workbook(f,read_only=True); names=wb.sheetnames; wb.close()
    for n in names:
        tmp=os.path.join(shots.SP,"_m2.xlsx")
        shots.one_sheet(f,n,tmp)
        stem=f"o2_{k}_"+re.sub(r'[^A-Za-z0-9]+','_',n).strip('_')
        p=shots.to_png(tmp,stem)
        print(("  ok   " if p else "  FAIL ")+k+" / "+n)
    if os.path.exists(os.path.join(shots.SP,"_m2.xlsx")): os.remove(os.path.join(shots.SP,"_m2.xlsx"))
