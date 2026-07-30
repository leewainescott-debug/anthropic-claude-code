# Retired instructions and superseded decisions

Everything in this file is **dead**. It is kept only so nobody re-derives a rule from an
old document and applies it again. Nothing here is to be acted on.

If a rule appears both here and in `PLAN.md` or `DECISIONS.md`, the live file wins. If a
rule appears only here, it is not a rule.

---

## A. Instructions the owner has since overruled

| Retired rule | Where it came from | What replaced it, and when |
|---|---|---|
| Squads with no people in them are removed from the design tabs | D6a, from "we cannot have squads with 0 people in them" | Still live for the four already removed - Pricing & WFM, Digital Support NZ, Enterprise Data Delivery, EGI Data - which he confirmed stay out. **It is not a licence to remove anything further.** No block comes off a design tab again without him saying so. |
| Frozen panes on every table over 25 rows | Original register item 70 | Item 94, then his direct instruction: **no frozen panes anywhere**. |
| "Lights on budget" and "Total to fund" are banned wordings | An early banned-word ruling | His own newest workbook uses both, consistently. **His labels win.** (D83) |
| The 1.x actuals block sits at the foot of the tab | The first design-A build | **Up top, beside the budget box**, one clean fixed-shape table. (D102) |
| 3.2 states allowance vs organisation in my column order | My redesign | **His layout, his headings, his order** - he rewrote the tab himself. (D108) |
| The 525-role ledger | Pre-consolidation | **531 roles, $115,589,735.11.** Every "525" in an old doc is stale. |
| Tab named "3.5 Source Reconciliation", "Squads", "Added data", "Sheet2", "FY26 Budget (superseded)", "squad mapping (superseded)" | Original workbook | **Deleted.** REVIEW is the only source. (D51) |

## B. Changes I made that he did not ask for, and has now reversed

These were mine, not his. They are recorded so the reasoning is not repeated.

| What I did | Why I thought it was right | Why it was wrong | Status |
|---|---|---|---|
| Added a **Home country** column to 0.2 | The 1.x tabs decided AU vs NZ by comparing which budget figure was bigger, which is unguessable | He did not ask for it, it is his config tab, and the rule it replaced was working | **Removed** (D113) |
| "Fixed" **1.2!C7/D7** as a double count | Both branches of the same IF counted the whole platform overhead | It is his tab and his published figure | **Reverted** (D112) |
| Wrote **0.00** into 0.2's blank Legal, EG and EGI spend rows | So the column read consistently top to bottom | A typed zero is a statement he did not make | **Reverted to blank** (D113) |
| Renamed **"Budget to draw down ($m)"** to "Budget available ($m)" on 1.11, 1.12, 1.13, and "Alloc %" to "Allocation %" on 0.2!M13 | It went in on the same pass that replaced my word "ledger" with his "role mapping" | "Ledger" was mine to fix. These are **his own labels**, and D83 already settled that his labels win | **Restored** (D113) |

Two regression checks went with them. `regress2707.py` used to assert that 0.2 carried a
Home country column and that **no** tab decided its country by comparing budget cells. Both
encoded the instruction above, and both are retired: the file now checks that the column is
gone and that all eleven 1.x tabs use his comparison, identically.

Six 3.2 checks were re-derived rather than re-baselined: they compute their expected figure
from the model's own platform count and tie 3.2 to 3.1, so they test the thing that matters
whatever the count is.

**Retired with wave K (D112 corrected): the claim that 1.2's double-counted platform
overhead was "his shape".** It was not - both his workbooks price it once (0.495), and the
double count was the build's own corruption of his C7. The three checks that pinned 0.99 /
16.0575 / "C7 == D7" as his are gone; the pins now hold his true 27/07 complement shape.
Anyone reading an old log or the first D112 and concluding 0.99 is his: it never was.

## B2. Retired with wave M (D120)

| Retired | What replaced it |
|---|---|
| **3.1's step-by-step walk** - forty rows: squads priced by an archetype, directly funded programmes by name, overhead against the allowance, the COEs, the groups nothing prices, and two subtotals that could only carry a dash | 3.1 is one row per group, his approved layout. Every step of the walk is still on the tab it came from - 3.3 by squad, 3.2 by overhead line, each working tab by block. |
| **3.1's "Retail - Ampol and Z together" subtotal row** | The two rows sit adjacent with no subtotal. The row is removed, not hidden - it was a display row inside a column of group rows and it was in no total. |
| **The dash convention on 3.1's total and grand rows** (a dash in the archetype and variance columns) | The COEs and EGI price archetype = actual, so every row carries a figure and the total column adds up. |
| **1.13's 0.5 Cyber CapEx input**, its bucket total, its "Left to fund" row and the "Planned spend less CapEx" cell beside it | Two funding lines: the COE allocation off 0.2, then the total to draw down. The seventh summary column is a Variance. |
| **0.2's combined "TDD Cyber incl. COE" line** and the F23 read that summed 1.13 and 1.14 | Row 7 is the COE, row 23 is TDD Cyber, each reading its own tab. The allocated total is unchanged at 50.5. |
| **1.2!I54's typed 2.21** for EGI Customer | His review workbook's own formula. The EGI squads are funded at the actual cost of their roles (D120 item 9), which settles the question D118 left open. |
| **Digital Support NZ's removal under D6a** | The squad is back. D6a still holds for a squad with no people AND no line of his behind it; this one has his line behind it. |
| **The accounting-dash section of every money and count number format** | Two sections, so a zero renders 0.00 or 0. No cell in the workbook shows a dash. |
| **Exec's five walk-subtotal lines** (squads priced by an archetype x3, directly funded, COEs and EGI actual, overhead in the COEs, everything comparable, groups with no archetype and no funded figure) | Five lines off 3.1's group total plus two "of which" lines read straight off the working tabs. |

Four gates moved with them. Two were **re-derived, not re-baselined**: 1.2!F9's typed
15.5625 is now a tie between the tab's own three columns and 2.2's archetype total, and the
Exec hold count is a recount of every vacant role held off the working tabs. The 40-squad
hybrid count is derived rather than typed. Two were **retired**: 1.13's CapEx-follows-its-
input pin, and "1.14 Cyber Uplift awaits its inputs".

## C. Superseded process rules

- **Seven QA passes** as a fixed list - superseded by `audit.py` plus the five scripted
  passes and `regress2707.py`. The number is not the point; the coverage is.
- **Multi-agent verification waves as the default** - superseded. Arithmetic is verified
  by script, in seconds, on every build. Agents are for judgement, not for counting.
  A 127-agent workflow to check a spreadsheet was the wrong tool and cost a quarter of a
  week's usage.
- **"QA five times"** - superseded by the same. The gate runs every build, not five times
  at the end.

## D. Files that are history, not instruction

These describe states the model has left behind. Read them for context only.

`DEFECT_REGISTER.md`, `DESIGN_DIAGNOSIS.md`, `REDESIGN_OPTIONS.md`, `REVIEW_FINDINGS.md`,
`VALIDATION_EVIDENCE.md`, `compliance_checklist.md`, `requirements_brief.md`,
`vacancy_audit.md`, and sections 6-8a of `PLAN.md`.

`INSTRUCTION_REGISTER.md` holds the original 101-item brief. It still stands **except**
where this file or `DECISIONS.md` supersedes it.
