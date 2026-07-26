import os,glob,re,openpyxl,shots,wbio
for f in sorted(glob.glob(os.path.join(shots.SP,"opt","*.xlsx"))):
    k=os.path.basename(f)[:-5]
    rc=wbio.recalc(f); os.replace(rc,f)
    wb=openpyxl.load_workbook(f,read_only=True); names=wb.sheetnames; wb.close()
    for n in names:
        tmp=os.path.join(shots.SP,"_o.xlsx"); shots.one_sheet(f,n,tmp)
        stem=f"OPT_{k}_"+re.sub(r'[^A-Za-z0-9]+','_',n).strip('_')
        print(("  ok   " if shots.to_png(tmp,stem) else "  FAIL ")+k+" / "+n)
if os.path.exists(os.path.join(shots.SP,"_o.xlsx")): os.remove(os.path.join(shots.SP,"_o.xlsx"))
