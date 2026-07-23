#!/usr/bin/env python3
import formulas, re, logging
logging.getLogger().setLevel(logging.ERROR)
PATH="clean_v3.xlsx"
xl=formulas.ExcelModel().loads(PATH).finish(); sol=xl.calculate()
val={}
for k,v in sol.items():
    m=re.match(r"^'?\[[^\]]*\]([^!]*?)'?!([A-Z]+\d+)$", k)
    if not m: continue
    x=v.value
    try: x=x[0,0]
    except: pass
    if not isinstance(x,(int,float,str,bool)) and x is not None: x=str(x)
    val[(m.group(1).strip().upper(), m.group(2))]=x
def gv(t,a): return val.get((t.upper(),a))
def isnum(x): return isinstance(x,(int,float)) and not isinstance(x,bool)
ERR=("#REF!","#DIV/0!","#VALUE!","#N/A","#NAME?","#NUM!","#NULL!","#CYCLE!")
errs=[(s,c,v) for (s,c),v in val.items() if isinstance(v,str) and v.strip() in ERR]
from collections import Counter
print("=== ERROR CELLS:", len(errs))
for s,n in Counter(s for s,_,_ in errs).most_common(): print("   ",s,n)
for s,c,v in errs[:30]: print("     ",s,"!",c,"=",v)
WT={6:"2.1 Ampol Retail",7:"2.2 Customer",8:"2.3 Enterprise Data",9:"2.4 TDD Group Functions",
10:"2.5 P&C",11:"2.6 Finance",12:"2.7 Infrastructure",13:"2.8 Energy Solutions & B2B",
14:"2.9 Commercial Fuels",15:"2.10 Z Retail",20:"2.11 TDD Cyber"}
TR={"2.1 Ampol Retail":17,"2.2 Customer":16,"2.3 Enterprise Data":10,"2.4 TDD Group Functions":13,
"2.5 P&C":10,"2.6 Finance":10,"2.7 Infrastructure":11,"2.8 Energy Solutions & B2B":10,
"2.9 Commercial Fuels":11,"2.10 Z Retail":11,"2.11 TDD Cyber":7}
def fm(x): return f"{x:.3f}" if isnum(x) else str(x)
print("\n=== FOOTING (working tab default=hold vs 3.2) ===")
print(f"{'tab':26s} {'Itot':>8s} {'3.2K':>8s} {'Htot':>8s} {'3.2L':>8s} {'Dfill/3.2I':>12s} {'Evac/3.2M':>11s} {'3.2F':>8s}")
bad=0
for r,t in WT.items():
    tr=TR[t]
    I=gv(t,f"I{tr}"); H=gv(t,f"H{tr}"); D=gv(t,f"D{tr}"); E=gv(t,f"E{tr}")
    k=gv("3.2 Total Cost",f"K{r}"); l=gv("3.2 Total Cost",f"L{r}")
    Ic=gv("3.2 Total Cost",f"I{r}"); Mc=gv("3.2 Total Cost",f"M{r}"); Fc=gv("3.2 Total Cost",f"F{r}")
    ok=(isnum(I) and isnum(k) and abs(I-k)<1e-4 and isnum(H) and isnum(l) and abs(H-l)<1e-4
        and isnum(D) and isnum(Ic) and abs(D-Ic)<1e-9 and isnum(E) and isnum(Mc) and abs(E-Mc)<1e-9
        and isnum(Fc) and isnum(I) and abs(Fc-I)<1e-6)
    if not ok: bad+=1
    print(f"{t:26s} {fm(I):>8s} {fm(k):>8s} {fm(H):>8s} {fm(l):>8s} {fm(D):>4s}/{fm(Ic):<6s} {fm(E):>4s}/{fm(Mc):<5s} {fm(Fc):>8s}{'  <--BAD' if not ok else ''}")
print("footing mismatches:", bad)
c52=gv("Exec Summary","C52"); hsum=sum(gv(t,f"H{TR[t]}") for t in TR if isnum(gv(t,f"H{TR[t]}")))
print(f"\nExec C52 = {fm(c52)}  sum(H-tot) = {hsum:.4f}  ok={isnum(c52) and abs(c52-hsum)<1e-4}")
print("3.2 TOTAL row:", {c:fm(gv('3.2 Total Cost',f'{c}24')) for c in 'CDEFKLIMN'})
