# Four reviewers on the layout options

A financial modeller, a spreadsheet designer, a CTO / instruction-compliance auditor and an
HR / workforce specialist reviewed the mockups. Every claim below that could be checked
against the live workbook has been checked, and the check is recorded.

**Verdict: the options are not ready.** Not because the layouts are wrong, but because of
how they were built.

---

## The root cause

**I hand-typed the numbers into the mockups instead of reading them out of the workbook.**
Almost every serious finding traces to that one decision.

| What the mockup says | What the file says | Because |
|---|---|---|
| 3.2 Ampol Retail actual cost **14.26** | ledger **14.0052** | I typed 6.08 + 8.18, which is `1.1!F9`, the **design** total. A design number under a header saying "Actual". |
| 3.1 Enterprise Data budget **AU 3.50 / NZ 0.00** | `0.2!r22` **AU 2.5 / NZ 1.0** | typed. A million dollars moved from the NZ budget to the AU budget. |
| 3.1 group budget **50.50** | `0.2` r27 **53.80**, with **3.30 unallocated** | typed. The only free money in the model is invisible. |
| 1.x squad table totals **11.28 / 6.66** | portfolio summary **12.80 / 8.18** | I listed five platforms; the file has seven. `EGI Retail` (1.52) and `Pricing & WFM` were dropped, and the summary that includes them was kept. |
| AmPOS platform overhead **0.165** | `1.1!I49` is **empty** | invented. That blank is open question O3, deliberately unanswered. |
| 3D tiles: Other **34.85**, Total **115.11** | its own table three rows below: **34.88** and **115.14** | typed constants beside a calculated table. |
| 3.1 `Left to fund` | identical to `Other` on every row | the code writes the same variable twice. |

A mockup that disagrees with the workbook is worse than no mockup: it invites a decision
on a layout whose numbers would change the moment it was built for real.

**The fix is method, not layout.** Mockups get generated from `TDD_Cost_Calc.xlsx` itself,
never from typed constants. Then they cannot disagree with it.

---

## Errors in my own formatting standard

### The typeface was wrong in all twelve mockups. VERIFIED, FIXED.
openpyxl's default theme font is **Cambria**, a serif. Your tabs are **Calibri**. I never
named a font, so every mockup rendered in the wrong typeface. That is the first thing the
eye registers and the biggest single reason they looked like a different document.

### The two navies were collapsed into one. VERIFIED, FIXED.
`1.11 BP&T` uses two: the section bar is the **darker** `002F6C`, the column header row the
**lighter** `1F4E79`. That contrast is the main hierarchy device on the tabs you call clean.
My mockups used one navy for both, so bar and header merged into a slab.

(The reviewer had the two the wrong way round; the file settles it.)

### Notes were 9pt grey. FIXED.
Breaks instruction 73 (nothing under 10pt) and 74 (no grey note styling) simultaneously.
A formatting standard that breaks the formatting rules is worse than none.

### Sentences are still written into cells.
"The squad band carries that squad's totals." "Click the minus in the margin." Diagnosis
point 3 said stop doing this. I did it again.

### Headers still truncate.
`Portfolio overhead (see 0.2 Data Co`, `TDD lights-on budg`, `AU over/(under) bu`,
`to archetype ($m)` with "Variance" cut off, and a `###` cell on 1D. That is the exact
defect the diagnosis was written to fix.

---

## Arithmetic that does not hold

1. **3.2's total row does not subtract.** Archetype 64.20, actual 115.14, variance 23.40.
   115.14 − 64.20 = 50.94. The variance column totals the ten rows that have an archetype;
   the cost columns total all fourteen. **The live tab already handles this properly** — it
   compares *squad cost* (delivery only, 75.69) to archetype (64.20) and gets 11.49. My
   mockup compared *actual including overhead* to archetype, which is the apples-to-apples
   rule broken. The file was right and I broke it.
2. **Platform totals do not cross-foot.** 1A Store Operations: 5.70 total, 3.48 TDD, 2.39
   outside. 3.48 + 2.39 = 5.87. The platform overhead is added to one column and not the other.
3. **`Left to fund` on 1.x is a typed −0.002 plug.** There is 3.40 of people budget in the
   Other funding table with no cost against it, and the model prints (0.00). That is the
   "total to fund" figure instruction 21 asked to have fixed, still plugged.
4. **3.2's overhead total adds up a rate column and a units column** — 0.9625 and 100, both
   meaningless — and leaves the `Over/(under) allowance` cell blank. That blank is **9.98**,
   the largest single variance in the model, and instruction 99 asks for it by name.
5. **Overhead is counted twice on 3.2.** The portfolio rows already contain the 62 overhead
   people; the block underneath restates 22.91. Read top to bottom you add them.
6. **`Left to fund` means (0.00) on 1.x and 8.18 on 3.1** for the same portfolio, same pack.

---

## Instructions the options break

| # | Instruction | What the mockups do |
|---|---|---|
| 35 | The nine Group Summary columns you pasted | six of nine dropped: Archetype Support Cost, Cost of FTE non TDD funded, Amount identified as rechargeable, Total still left to fund, Total Cost |
| 34 | Budget impact of decisions: TDD vs business vs left to fund | absent from all four 2.x options |
| 91, 29, 27 | Filled, Planning to hire, Vacancies remaining, Cost after **vacancy** decisions | Filled missing; "Hire/offshore" merges two different decisions; "Roles after" is not vacancies remaining; the word "vacancy" dropped from the header again |
| 9 | The ladder to 53.8 and the 43.5 allocation | absent |
| 40, 13 | The FTE view restored as its own tab | replaced by "Squad Detail" |
| 38, 49, 98 | 3.4 COE Summary, gross and net, dollars and FTE | not in the pack |
| 43 | Exec Summary | not in the pack |
| 41 | Squad Detail reacts to the working tabs | 3.3 has no cost-after-decisions column, so the levers leave it frozen |
| 18, 5 | Platform overhead per platform, platform blocks and totals stay exactly as built | 1B/1C/1D delete the per-platform overhead line |
| 8, 61 | Strategic programme cost as a yellow input, with the people-cost note beside it | not yellow in any option; the notes deleted |
| 72 | Inputs are yellow, only inputs are yellow | three options yellow Size/On-Off/AU-NZ, one yellows Support % instead. Same table, two definitions |
| 19, 20 | The AU/NZ toggle, "crucial" | demonstrated on Ampol Retail, where every row is AU and both NZ columns are zero |
| 15, 80 | Never redesign a tab; your wording is never overwritten | "Amount that can be allocated to people" became "Allocated to people", turning a ceiling into a commitment |
| O1 | Cyber split — *do not build until answered* | 3.1 ships a COE Cyber budget of 2.50/1.00 |
| D34 | "Total to fund" is meaningless and gone | reinstated on 1A, 1B, 1C |
| D18, 101 | The lever sits on every role row, filled included | 2D removes filled roles from the decision table |

---

## What a CTO cannot decide from any option

1. **If I hold these 135 vacancies, what do I save?** Impact is zero on every row of every
   mockup, because vacancies default to Hire. The tab demonstrates the lever doing nothing.
2. **Does that saving help AU or NZ?** No AU/NZ on any 2.x tab.
3. **Does it reduce TDD lights-on, or a recharge the business was paying anyway?** No TDD
   vs Other column on 2.x. That is the difference between saving TDD money and saving nobody's.
4. **What is my real headroom?** 3.1 stops at the 50.5 allocation. The budget is 53.8.
5. **Which funding line absorbs the overspend?** The CapEx / OpEx / Significant Items walk
   exists per portfolio on 1.x and appears on no 3.x tab in any option. Worth noting: the
   drawdown already shows CapEx drawn 10.410 against a 4.9 pool — a 5.5m hole that surfaces
   on no summary.
6. **Which decisions are worth the most, across all portfolios?** There is no group-level
   decision register. A CTO would open fourteen tabs and merge by hand.

---

## Workforce findings

1. **The lever can terminate a named person and prices it at zero.** Hold removes the head
   and the cost. On 2A and 2B the dropdown sits on named individuals beside their salaries.
   No notice, no exit cost, no distinction from simply not filling a vacancy. Named salaries
   on a tab circulating to eight GMs is its own exposure.
2. **Nothing says which vacancy matters** — no days open, no critical flag, no vacant %.
   Data AU is 6 vacant of 8 and reads the same as a squad with 1 of 8.
3. **Headcount variance is never calculated though both numbers are printed.** Ampol Retail
   is roughly 68% over design headcount at 1.7% over design cost. The dollar variance says
   on plan; the headcount says the opposite.
4. **30% allocated managers are shown as whole heads.** Three Technology Managers at 0.77
   is 257k each — a full cost, not 30%.
5. **Overhead has no allowance on the 2.x tab**, so on the tab where decisions are made it
   appears immovable and measured against nothing.

---

## What survived

- Brackets rather than red for negatives — correct, and fixes the backwards alarm colour.
- The orphan styled rows and the green/salmon conditional blocks are gone.
- **1A** is the only 1.x option that keeps platform blocks, platform overhead lines and
  platform totals together with the budget, funding position and Finance reconciliation.
- **3C**'s bridge is the most readable page in the set, but it must sit *under* the 3.2
  table rather than replace it.
- **2A**'s content model — squad summary, overhead, portfolio total, then the people — is
  the right one. It fails on execution: the FTE table is built on a different grid from the
  squad table above it, which is the "L-shaped hole" and most likely what "not quite there"
  means.
