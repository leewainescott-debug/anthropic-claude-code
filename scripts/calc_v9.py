#!/usr/bin/env python3
"""Recalculate v9 with the formulas engine; report errors + key numbers."""
import formulas, logging, re, sys
logging.getLogger().setLevel(logging.ERROR)
SRC = "TDD_Cost_Calc_v9.xlsx"
xl = formulas.ExcelModel().loads(SRC).finish()
sol = xl.calculate()
vals = {}
for k, v in sol.items():
    m = re.match(r"^'?\[[^\]]*\]([^!]*?)'?!([A-Z]+\d+)$", k)
    if not m: continue
    val = v.value
    try: val = val[0, 0]
    except Exception: pass
    vals[(m.group(1).strip().upper(), m.group(2))] = val
errs = [(k, v) for k, v in vals.items() if isinstance(v, str) and v.startswith("#")
        and k[0] not in ("SQUADS", "ADDED DATA", "SHEET2")]
print("ERR COUNT (model tabs):", len(errs))
for k, v in errs[:40]: print("  ", k, v)
def g(sheet, cell): return vals.get((sheet.upper(), cell))
print("--- key numbers ---")
print("2.3: ports", g("2.3 BP&T","C10"), "| BP draw", g("2.3 BP&T","C13"), "| BP budget", g("2.3 BP&T","C15"),
      "| Tr budget", g("2.3 BP&T","C16"), "| roles", g("2.3 BP&T","C8"), "| spend", g("2.3 BP&T","F8"),
      "| left", g("2.3 BP&T","H8"), "| check", g("2.3 BP&T","C45"))
print("2.4: ports", g("2.4 SA&D","C10"), "| DA draw", g("2.4 SA&D","C13"), "| SA budget", g("2.4 SA&D","C15"),
      "| Data budget", g("2.4 SA&D","C16"), "| roles", g("2.4 SA&D","C8"), "| spend", g("2.4 SA&D","G8"),
      "| paused", g("2.4 SA&D","C18"), "| left", g("2.4 SA&D","I8"), "| check", g("2.4 SA&D","C41"))
print("2.5: roles", g("2.5 Cyber Roles","C8"), "| filled", g("2.5 Cyber Roles","D8"), "| vac", g("2.5 Cyber Roles","E8"),
      "| spend", g("2.5 Cyber Roles","F8"), "| left", g("2.5 Cyber Roles","H8"), "| check", g("2.5 Cyber Roles","C71"))
print("2.2: E8", g("2.2 COE","E8"), "| E11", g("2.2 COE","E11"), "| D13", g("2.2 COE","D13"), "| F13", g("2.2 COE","F13"))
print("1.11: E9", g("1.11 TDD Cyber","E9"), "| C8", g("1.11 TDD Cyber","C8"), "| D8", g("1.11 TDD Cyber","D8"))
print("2.1: arch tot", g("2.1 Total Cost","C24"), "| filled", g("2.1 Total Cost","E24"), "| vac", g("2.1 Total Cost","G24"),
      "| act tot", g("2.1 Total Cost","I24"), "| over", g("2.1 Total Cost","K24"),
      "| FTE arch", g("2.1 Total Cost","D24"), "| FTE tot", g("2.1 Total Cost","J24"),
      "| dedup", g("2.1 Total Cost","C23"), "| restate", g("2.1 Total Cost","I26"), "| egi memo", g("2.1 Total Cost","I27"))
print("2.0: J26 net", g("2.0 Group Summary","J26"), "| C32 chk", g("2.0 Group Summary","C32"))
print("exec: C26 model", g("Exec Summary","C26"), "| C30 actual", g("Exec Summary","C30"),
      "| C50", g("Exec Summary","C50"), "| C52 lever", g("Exec Summary","C52"), "| C57", g("Exec Summary","C57"),
      "| C58 cyber", g("Exec Summary","C58"), "| C59 coe", g("Exec Summary","C59"))
print("4.1: G15 hires", g("4.1 Ampol Retail","G15"), "| H15 vac-after", g("4.1 Ampol Retail","H15"),
      "| F15 vac", g("4.1 Ampol Retail","F15"), "| C16", g("4.1 Ampol Retail","C16"), "| C17", g("4.1 Ampol Retail","C17"),
      "| B2", g("4.1 Ampol Retail","B2"))
print("3.0: EGI n", g("3.0 FTE View","C185"), "| EGI $", g("3.0 FTE View","F185"), "| xcheck", g("3.0 FTE View","C164"))
