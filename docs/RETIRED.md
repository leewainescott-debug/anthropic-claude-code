# Retired instructions and superseded decisions

Everything in this file is **dead**. It is kept only so nobody re-derives a rule from an
old document and applies it again. Nothing here is to be acted on.

If a rule appears both here and in `PLAN.md` or `DECISIONS.md`, the live file wins. If a
rule appears only here, it is not a rule.

---

## A. Instructions the owner has since overruled

| Retired rule | Where it came from | What replaced it, and when |
|---|---|---|
| Squads with no people in them are removed from the design tabs | D6a, from "we cannot have squads with 0 people in them" | Still live for the three already removed - Pricing & WFM, Digital Support NZ, Enterprise Data Delivery, EGI Data - which he confirmed stay out. **It is not a licence to remove anything further.** No block comes off a design tab again without him saying so. |
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

Six 3.2 checks were re-derived rather than re-baselined. They pinned a snapshot - 22
platforms, 5.005, 31.2 priced for, a 38.8-role gap - which the 1.2 revert moved. They now
compute the expected figure from the model's own platform count and tie 3.2 to 3.1, so they
test the thing that matters and keep telling the truth whichever way he rules on 1.2.

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
