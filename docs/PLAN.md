# TDD Cost Calculator - plan and instruction log

Live document. Updated every time work lands or an instruction changes.
Companion file: `docs/DECISIONS.md` (what was decided and why).
The original 101-item brief is `docs/INSTRUCTION_REGISTER.md` and still stands except
where superseded below.

Last updated: after the cost-after-decisions fix. QA is at zero findings on all eight
checks and the workbook is ready to review.

---

## 1. What this workbook is for

Price the TDD organisation two ways and compare them:

- **Design** - what the squads *should* cost, from the archetype library on `0.3`.
- **Actual** - what the named people in `REVIEW - Complete Role Mapping` *do* cost.

Then let a GM pull a lever on each vacancy (Hire / Hold / Offshore) and see the cost move
through to the group.

## 2. The one rule everything else follows

**`REVIEW - Complete Role Mapping` is the only source of truth.** 525 roles,
$115,113,262.27, 390 filled, 135 vacant.

Nothing overrides it. Where the design tabs disagreed with REVIEW on a squad name, the
design tab was renamed - never the reverse. Where a role needs re-assigning, that sits in
a visible override table on `Lists`, not by editing Lee's raw columns.

## 3. Tab structure and flow

```
0.x  inputs        0.1 Budget Table (Fin)   0.2 Data Config   0.3 Squad Archetypes
1.x  design        1.1 - 1.10 portfolios    1.11 BP&T  1.12 SA&D  1.13/1.14 Cyber
2.x  working       2.1 - 2.14, one per portfolio, this is where decisions get made
3.x  summaries     3.1 Group  3.2 Total Cost  3.3 FTE  3.4 COE  3.5 Reconciliation
4.0  Data QA       live checks
Exec Summary       the story
```

Flow is one direction: 1 into 2 into 3. Numbering is the flow.

## 4. Instructions given, in order

### The original brief
See `docs/INSTRUCTION_REGISTER.md` - 101 numbered items covering the model, structure,
the 1.x and 2.x tabs, the summaries, the COEs, data rules, formatting, process and a
banned list. Still live except where superseded below.

### This engagement

| # | Instruction | Status |
|---|---|---|
| 1 | The 2.x levers do not move the numbers. Make offshoring a filled resource actually change cost. | Done |
| 2 | Work out where overhead gets baked in and put the judgement calls back to me. | Done |
| 3 | Confirmed the ten ledger/design squad name pairs, including Ampol Loyalty & Martech and SAP ERP. | Done |
| 4 | "Leadership overhead" is the 8 GMs priced at $5.1m. Nothing else. | Done |
| 5 | Tech Managers and Delivery Managers are priced at 30% allocated to one platform. There is no 1:1 rule. | Done |
| 6 | All six business partners make sense. | Done |
| 7 | Every squad in 1.x, 2.x, 3.x and REVIEW must ladder up. No discrepancies. Flag everything. | Done |
| 8 | Every dollar and every person accounted for. 525 people, every dollar, on a 2.x tab. | Done |
| 9 | Archetype cost vs actual squad cost must be apples to apples. No leftover dollars. | Done |
| 10 | Headers must make sense. "Total to fund" as a block header makes no sense. | Done |
| 11 | Tables and formatting must be consistent across 2.x and across 3.x. | Done |
| 12 | QA five times. Use multiple agents. Look at the whole thing, not one tab at a time. | Done |
| 13 | Errors means formatting, formulas, logic, design, structure and sense - not just #REF. | Done |
| 14 | Filled roles SHOW a cost. (Supersedes register items 30, 88, 101.) | Done |
| 15 | The COEs have no overhead. | Done |
| 16 | Fix the two budget variances, 3.4, the offshore lever, COE gross vs net, the sign conventions, the hardcoded cost cells. | Done |
| 17 | Loop: see the problem, fix, review, QA, re-review. Do not come back until it is resolved. | Standing |
| 18 | REVIEW is the only source of truth. The design tab names are wrong and must follow REVIEW. | Done |
| 19 | Customer's unassigned roles are just Leadership. Keep as is. | Done |
| 20 | r283 and r313 to the Data COE. r364 to an Infrastructure squad. r528 to AU Finance. | Done |
| 20a | George Moun (r364) is Leadership in Infrastructure. | Done, already grouped that way |
| 21 | Strategic programmes are directly funded and do not fit an archetype. | Done |
| 22 | Leave the two squad folds merged. | Done |
| 23 | 0.1 column G is empty by design. Leave it. | Closed |
| 24 | An archetype with no XS should return nothing or NA. | Already correct |
| 25 | Do not bring me things I did not ask for. | Standing |
| 26 | Plain English, in the workbook and in chat. | Standing |
| 27 | Keep a plan file and a decisions file, updated as we go. | This file and `DECISIONS.md` |

## 5. Where the model stands

| Check | Result |
|---|---|
| Ledger | 525 roles, $115,113,262.27 |
| On a 2.x tab | 525, none twice, none missing |
| 3.1 / 3.2 / 3.3 / Exec | all tie to the ledger |
| 3.4 COE + EGI | 113 roles, $27.771m |
| Every control row | 0 |
| 4.0 live checks | 18, all passing |
| Formula errors after full recalculation | 0 |
| Cost cells that are formulas | 527 of 527 |
| Adversarial QA, all eight checks | 0 findings |
| Sheets | 51, no gridline or tab-colour change |

Portfolio split after the agreed assignments:

| Portfolio | Roles | Cost |
|---|---|---|
| Customer | 83 | $17,126,867 |
| Ampol Retail | 70 | $14,005,227 |
| COE Cyber | 46 | $9,897,858 |
| TDD Group Functions | 46 | $10,031,229 |
| Commercial Fuels | 42 | $10,197,670 |
| Z Retail | 39 | $7,337,294 |
| Infrastructure | 36 | $7,748,364 |
| Energy Solutions & B2B | 33 | $6,447,816 |
| Enterprise Data | 28 | $6,418,404 |
| COE SA&D | 26 | $6,529,494 |
| COE BP&T | 24 | $6,400,562 |
| P&C | 18 | $4,184,441 |
| EGI | 17 | $4,943,419 |
| Finance | 17 | $3,844,616 |
| **Total** | **525** | **$115,113,262** |

## 6. Squad coverage - all 525 accounted for

| | Roles |
|---|---|
| In a delivery squad that exists on its design tab | 470 |
| On an overhead line (no design squad by definition) | 53 |
| In a squad with no row on its design tab | 2 |
| **Total** | **525** |

Every design squad now has people in it. The three that did not are gone.

## 7. Open items

| Item | Detail | Needs |
|---|---|---|
| Viren Khatri (r104) | Data Migration Lead. Department EGI, team EGI TDD, but he sits in the Ampol Retail portfolio. EGI TDD is a strategic programme on 1.4, not an Ampol Retail squad. | Lee to name the squad |
| Jasper Na (r136) | Associate Engineer - BE. Department Ampol Digital, team "Energy". Customer portfolio. There is no "Energy" squad on 1.2; Ampol Digital is a platform there. | Lee to name the squad |
| Two input colours | Bright yellow (49 cells) and cream (102 cells) both mean "typed input", side by side on the same tabs. One should win. | Lee to pick |
| Bridge tab | One page walking archetype cost to actual cost, line by line. Explained; not yet built. | Lee to say yes or no |

## 8. How to work on this

1. Read this file and `docs/DECISIONS.md` first.
2. Definitions before formulas. Most rework on this job came from building against a
   definition that turned out to be wrong.
3. Join tabs by name, never by row number. Row positions have moved three times and each
   time something downstream read a wrong-but-valid number with no error.
4. Check `docs/INSTRUCTION_REGISTER.md` on every build.
5. Never openpyxl-save a populated workbook and ship it - it strips cached values and the
   file opens blank. Build with `scripts/v10/wbio.py`, which recalculates and injects
   values back, and asserts completeness.
6. Run `scripts/v10/qa.py` after every build. An error count only catches #REF!; the
   adversarial checks catch silent zeros, dangling references and label-versus-value
   contradictions, which is where the real defects were.
