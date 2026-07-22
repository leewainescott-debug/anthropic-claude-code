#!/usr/bin/env python3
"""Build the validation evidence pack: every QA assertion, what it checks, and the
key numbers - generated only from a green run."""
import re, sys

SCR = "/tmp/claude-0/-home-user-anthropic-claude-code/6161aafe-2dad-5bc5-ad55-d8a92ce554cc/scratchpad/"
log = open(SCR + sys.argv[1]).read() if len(sys.argv) > 1 else open(SCR + "qa_v9_run3.log").read()
if "FAILS: 0" not in log or "EXIT=0" not in log:
    print("QA NOT GREEN - refusing to build evidence pack"); sys.exit(1)

src = open(SCR + "qa_v9.py").read()
labels = re.findall(r'chk\("([^"]+)"', src)
extra = re.findall(r'fails\.append\(f?"([^"{]+)', src)

out = ["# TDD Cost Calculator - validation evidence pack", "",
       "QA harness result: FAILS: 0 (exit 0). Three passes:",
       "  PASS 1 - independent Python recompute straight from raw data / Added data (no Excel formulas)",
       "  PASS 2 - full formula-engine evaluation of every cell: zero formula errors",
       "  PASS 3 - the assertion battery below, plus a second engine run with the offshore",
       "           toggle flipped to prove Offshore = exactly 40% of onshore cost", ""]
key = [l for l in log.splitlines() if l.startswith(("PASS1:", "KEY:"))]
out += ["## Key numbers from this run"] + ["  " + k for k in key] + [""]
out.append(f"## The {len(labels)} named assertions (every one passed)")
for i, l in enumerate(labels, 1):
    out.append(f"  {i:3d}. {l}")
out.append("")
out.append(f"## Structural guards (fail the build if violated): {len(extra)}")
seen = set()
for e in extra:
    e = e.strip().rstrip(':').rstrip()
    if e and e not in seen:
        seen.add(e)
        out.append(f"  - {e}")
open(SCR + "VALIDATION_EVIDENCE.md", "w").write("\n".join(out) + "\n")
print("wrote VALIDATION_EVIDENCE.md:", len(labels), "assertions,", len(seen), "guards")
