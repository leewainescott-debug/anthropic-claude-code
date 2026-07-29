# TDD Cost Calculator - plan and instruction log

Live document. **Read `RETIRED.md` first**: anything in it is dead and must not be acted
on, however reasonable it sounds. `DECISIONS.md` records what was decided and why.
The original brief is `INSTRUCTION_REGISTER.md`, live except where superseded.

**Where it stands.** The role mapping holds 531 roles, 529.30 FTE, $115,589,735.11
(386 filled, 145 vacant). After the decisions currently set - 8 on hold, 3 offshored,
1 vacancy set to fill - 523 roles and $113,777,703.16.

**The rule that outranks every other rule in this file: his content is his.** A figure he
typed, a formula he wrote, a label he chose and a block he built are not the build's to
correct. Where the model disagrees with something of his, the model says so beside the
cell; it does not change it. Four things were changed without asking and have been
reversed - they are listed in `RETIRED.md` section B. That section is the standing
reminder, not a closed incident.

The one exception, and it is narrow: a **name** he typed that does not match the role
mapping is normalised to the role mapping's spelling, because a label that does not match
does not join and the roles behind it vanish off the tab. Wording is never touched.
`whatsgone.py` proves the distinction on every build - it lists every value he typed that
is not in the file, so a wording change cannot hide among the name normalisations.

**Verification is scripted, not delegated.** `audit.py` recomputes every figure from the
role mapping, proves all 531 people appear exactly once on the tab it assigns, checks every
control, and applies the model-review tests - column consistency, constants buried in
formulas, unmarked inputs, rounding inside chains, controls incapable of failing,
unbounded references. It runs on every build in seconds. `regress2707.py` asserts every
defect ever found stays dead. `qa.py`, `verify.py`, `recompute.py` and `qa1x.py` are the
independent recomputations. `whatsgone.py` answers "what of his is not in the file".
Agents are for judgement only.

A check that a decision of his makes fail is **retired or re-derived, never re-baselined
to the new number**. Re-baselining turns a gate into a rubber stamp: the six 3.2 checks the
1.2 revert moved now compute their expected figure from the model's own platform count, so
they test that every tab counts the same platforms rather than that 3.2 still says 5.005.

## 1. What this workbook is for

Price the TDD organisation two ways and compare them:

- **Design** - what the squads *should* cost, from the archetype library on `0.3`.
- **Actual** - what the named people in `REVIEW - Complete Role Mapping` *do* cost.

Then let a GM pull a lever on each vacancy (Hire / Hold / Offshore) and see the cost move
through to the group.

## 2. The one rule everything else follows

**`REVIEW - Complete Role Mapping` is the only source of truth.** 531 roles,
$115,589,735.11, 386 filled, 145 vacant.

Nothing overrides it. Where the design tabs disagreed with REVIEW on a squad name, the
design tab was renamed - never the reverse. Where a role needs re-assigning, that sits in
a visible override table on `Lists`, not by editing Lee's raw columns.

## 3. Tab structure and flow

```
0.x  inputs        0.1 Budget Table (Fin)  0.2 Data Config  0.3 Squad Archetypes
                   0.4 Presentation Pack   REVIEW - Complete Role Mapping
1.x  design        1.1 - 1.10 portfolios   1.11 BP&T  1.12 SA&D  1.13 Cyber Roles
2.x  working       2.1 - 2.14, one per portfolio, this is where decisions get made
                   one table per tab: archetyped squads, directly funded, overhead
3.x  summaries     3.1 Cost Bridge  3.2 Overhead & Leadership
                   3.3 Squad Detail  3.4 COE Detail
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
| 28 | Leave the 1.x tabs as they are. Tidy the formatting, language and logic only. | Done |
| 29 | Forget budget. Compare full actual cost to archetype. | Done |
| 30 | Archetype cost must include the COEs so the comparison is comparable. | Done |
| 31 | The COEs and EGI are actuals. | Done |
| 32 | Inputs are cream, not bright yellow. | Done |
| 33 | Layout: 2A on the working tabs, 3D on the summaries. | Done |
| 34 | The archetype prices some squads and not others. Never put the archetype for some against the actual for all. | Done |
| 35 | The directly funded programmes all sit somewhere. Say where, and price them against what they are funded. | Done |
| 36 | Viren Khatri stays in Ampol Retail. EGI P&C is a squad; leave it. | Done |
| 37 | Do not call it design. It is the archetype. | Done |
| 38 | A tab's name must match its own title. | Done |
| 39 | Delete Sheet2. REVIEW is the only source we care about. | Done - and every other retired source with it |
| 40 | No frozen panes anywhere in this workbook. | Done |
| 41 | 3.x never followed the instruction. 3.1 is layout 3D - a cost bridge. | Done |
| 42 | Name the directly funded squads explicitly, e.g. CTRM. | Done - all ten on 3.1 by name |
| 43 | 3.2 Total Cost gave nothing. Delete it, replace it. | Done - it is Overhead & Leadership |
| 44 | The word "seat" is never used. | Done |
| 45 | The 2.x overheads belong in the table up top, not in a second table below. | Done - one table, one header |
| 46 | Do not give me something that is not right. | Standing |

## 5. How the cost is compared

The organisation is funded four ways, so the comparison is built four ways and each block
says which one it is. Netting them into one line is what produced a group reading +11.49
over an archetype that prices two thirds of the roles.

`3.1 Cost Bridge` is layout 3D. It starts at the archetype cost of the squads an archetype
prices and walks, one named line at a time, to what TDD actually costs. Every directly
funded programme is on the page by name.

Every variance on the page is like for like. Where there is nothing to compare against, the
column says so rather than measuring a cost against a zero.

| Step | Archetype / funded | Actual | Variance | Roles |
|---|---|---|---|---|
| Squads priced by an archetype, one line per portfolio | 64.20 | 63.98 | (0.22) | 316 |
| AmPOS | 1.40 | 2.12 | 0.71 | 10 |
| EGI Retail | 1.52 | 1.22 | (0.30) | 6 |
| EGI TDD (Ampol Retail) | - | 0.30 | - | 1 |
| EGI Customer | - | 2.21 | - | 10 |
| EGI TDD (TDD Group Functions) | - | 1.02 | - | 4 |
| EGI P&C | - | 0.24 | - | 1 |
| EGI Finance | - | 0.24 | - | 1 |
| CTRM | 3.80 | 3.22 | (0.58) | 14 |
| **Directly funded, where the funded figure is set** | **6.72** | **6.56** | **(0.17)** | **30** |
| **Directly funded, where no funded figure is set yet** | **-** | **4.01** | **-** | **17** |
| COE Cyber, COE BP&T, COE SA&D, EGI | - | 27.77 | - | 113 |
| **COEs and EGI** | **-** | **27.77** | **-** | **113** |
| Overhead roles in the portfolios, against the allowance | 4.84 | 11.65 | 6.81 | 43 |
| **Everything with a figure to compare** | **75.76** | **82.19** | **6.43** | **389** |
| Leadership (Customer) | - | 1.05 | - | 5 |
| Leadership (Infrastructure) | - | 0.26 | - | 1 |
| **Groups with no archetype and no funded figure** | **-** | **1.30** | **-** | **6** |
| **Cost of the 525 roles in the ledger** | **75.76** | **115.28** | **39.52** | **525** |
| Leadership - the 8 GMs, outside the ledger | 3.00 | 5.10 | 2.10 | 8 |
| **Total cost of TDD including the GM layer** | **78.76** | **120.38** | **41.62** | **533** |

Only three steps carry an archetype figure that prices the whole of their own actual, and
those three are what the comparable subtotal adds. The totals below it carry the comparison
too, on the owner's instruction: archetype against actual with the difference beside it. The
38.03 on the ledger row is everything the archetype does not reach - the COEs, the programmes
with no funded figure set, Leadership - plus the overspend on what it does. Every one of
those is a named line above with a dash in the archetype column, so the figure cannot be read
as anything else.

Six of the eight directly funded programmes have no funded figure set against them on their
1.x tab, so there is nothing to compare them to and the column says so. Set one and that
programme moves to the line above by itself - the split is a formula, not a list.

Of the $115.11m, $86.58m is people in seat and $28.53m is 135 vacancies nobody has been
hired into yet. Both are on 3.2 and on Exec, so the headline is not read as payroll.

The overhead allowance is 12.925 across six lines. Only three of them draw in the
portfolios: the Business Partner and Domain Architect allowances are drawn inside the COEs,
where all thirteen of those people sit, and the 8 GMs sit above the 525-role ledger. 3.2
states all six line by line, splits the roles between the portfolios and the COEs, and
carries the one line that reconciles the block to 3.1.

## 6. Where the model stands

| Check | Result |
|---|---|
| Ledger | 525 roles, $115,283,002.27 |
| On a 2.x tab | 525, none twice, none missing |
| 3.1 / 3.2 / 3.3 / 3.4 / Exec | all tie to the ledger |
| Every control row | 0 |
| 4.0 live checks | 56, all zero |
| Recomputed from the ledger | every reader-visible figure, 0 disagreements |
| Formula errors after full recalculation | 0 |
| Adversarial QA (`qa.py`), eight checks | 0 findings |
| Layout QA (`verify.py`), five checks | 0 findings |
| Lever, recalculated end to end | cost and headcount both move, cost today does not |
| Sheets | 44, ordered to match the numbering, one hidden (Lists) |
| Frozen panes | none |
| 2.x structural profiles | 1 across all fourteen tabs |

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

## 7. Squad coverage - all 525 accounted for

| | Roles |
|---|---|
| In a delivery squad that exists on its design tab | 470 |
| On an overhead line (no design squad by definition) | 53 |
| In a squad with no row on its design tab | 2 |
| **Total** | **525** |

Every design squad now has people in it. The three that did not are gone.

## 8. Open items

| Item | Detail | Needs |
|---|---|---|
| Overhead allowance basis | The allowance prices half a Head of Technology per portfolio (0.1375 of 0.275) and 30% of a Delivery or Technology Manager per platform, against 43 whole heads costing $11.65m. The variance is labelled "not covered by the allowance" rather than overspend, but the basis itself is a judgement only Lee can confirm. | Lee to confirm |
| Platform count on Lists | Lists prices the per-platform overhead over 30 platforms; the ten design tabs carry 21, a difference of 1.485 in the allowance. 3.1 and the 2.x tabs now both use the design tabs. | Lee to confirm 21 or 30 |
| Cyber vacancies | `1.13!E7` carries your comment "there are a total of 8 vacant roles" against a cell reading 4. REVIEW has 4. If 8 is right, four vacancies are missing from REVIEW and the fix belongs there. | Lee to settle |
| 1.5 EGI P&C funded amount | `1.5!H32` is a blank cream input, so that platform prices at zero while the squad's one role costs $0.24m. Until it is set, both option workbooks state that squad's variance as a dash and hold its $0.24m on a separate "squads with no archetype to price them" line rather than adding it to a comparison it is not part of. | Lee to set |
| ~~Which option ships~~ | Settled: Option A. `TDD_Cost_Calc.xlsx` is Option A; Option B is not built. | done |
| Frozen panes | Register item 70 asks for them on every long table, item 94 records that they were all removed. Applied consistently on every table over 25 rows. One line to reverse. | Lee to pick |
| Bridge tab | One page walking archetype cost to actual cost, line by line. Explained; not yet built. | Lee to say yes or no |

## 8a. Actual against archetype on the 1.x tabs - two options

The owner picked Option A. `scripts/v10/chainA.sh` builds it and it ships as
`TDD_Cost_Calc.xlsx`. Option B is kept in `actuals.py` but is not built.

Two columns are appended to every squad table and every platform total on 1.1 - 1.10:
*Actual cost after decisions ($m)* and *Variance to archetype ($m)*. They go in K and L on
nine tabs and in P and Q on 1.6, where the owner's own five columns run out to O.

Each tab then ends with a portfolio block: squads an archetype prices, squads it does not,
overhead roles, anything on the working tab with no row on this tab, the total, and a control
that must read 0. All ten controls read 0.

Every figure is `=INDEX(...MATCH(...))` into the working tab's cost-after-decisions column,
matched on the squad name, so pulling a lever moves the 1.x tab. Proved by test rather than
asserted: a $176,565 filled role set to Hold moves that squad's figure by exactly that
amount and moves no other squad.

1.11, 1.12 and 1.13 carry no comparison and say why. They are COEs, funded by allocation, and
no archetype prices them; their groupings are by department while the working tab groups by
squad, so the split cannot be taken from either tab without inventing a mapping.

## 9. The seven QA passes

`scripts/v10/qaall.sh` runs the five standing passes against a built workbook, and
`scripts/v10/qa1x.py` adds two more for the 1.x comparison. They are deliberately different
in what they can see, because four of them passed while two variance bugs were live - both
were internally consistent additions of the wrong things - and six of them passed while two
of the owner's own columns had been overwritten on 1.6.

| Pass | What it can see |
|---|---|
| `wbio.audit` | formula errors, and formula cells shipped with no cached value |
| `4.0 Data QA` | 56 live checks inside the workbook, every difference must read zero |
| `qa.py` | adversarial: silent zeros, dangling references, bad SUM ranges, stale literals, family drift, cross-tab disagreement |
| `verify.py` | layout: truncated headers measured against their own column width, bars against the table under them, role coverage, banned words |
| `recompute.py` | every figure a reader sees, rebuilt from REVIEW in Python. Knows nothing about the workbook's formulas |
| lever test | a vacancy set to Hold, the workbook recalculated, and cost and headcount checked all the way through |
| `qa1x.py` figures | every new 1.x figure rebuilt from the ledger, columns found by reading header text so the checker cannot inherit the writer's placement, plus its own lever test |
| `qa1x.py` untouched | every cell of the shipped workbook, formula and value, against the variant. Nothing may be lost; a note that moves along its own row is reported as a move, not folded into "unchanged" |

## 10. How to work on this

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
