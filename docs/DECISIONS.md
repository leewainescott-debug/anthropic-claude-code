# TDD Cost Calculator - decision log

Live document. Every decision that shapes the workbook, why it was made, and who made it.
Companion file: `docs/PLAN.md` (what we are doing and what is outstanding).

"Lee" means you decided it. "Build" means I decided it and it is reversible - each one is
flagged so you can overturn it without hunting for where it lives.

Last updated: after the squad-assignment and design-rename build.

---

## A. Source of truth and data handling

### D1. `REVIEW - Complete Role Mapping` is the only ledger. Lee.
525 roles, $115,113,262.27, 390 filled, 135 vacant.

Every count and every dollar on Exec, 1.x, 2.x, 3.1, 3.2, 3.3, 3.4 and 4.0 reads REVIEW.
`Squads`, `Added data` and `Sheet2` are retired from the model. The one exception is
`3.5 Source Reconciliation`, which exists precisely to compare REVIEW against the old
`Squads` sheet.

**Why:** two population snapshots were in play and they disagreed. Everything downstream
inherited the disagreement. One ledger removes the argument.

### D2. Your raw columns are never edited. Build, standing rule.
Where a role needs re-assigning, that sits in an override table on `Lists` (AN:AP) that
names the REVIEW row, the portfolio and the squad. The derivation columns read the
override first and your raw column second.

**Why:** if I edit column K, your source data and my model become the same thing and you
lose the ability to check my work. An override table is visible, listable and reversible.

### D3. Squad name typos are folded by a table, not by hand. Build.
`Lists!W:X` holds ten fold pairs, for example `Manuacturing Group Projects` to
`Manufacturing Group Projects` and `Data - AU` to `Data AU`.

**Why:** same reason as D2. A fold you can read is a fold you can argue with.

### D4. Where a design tab and REVIEW disagreed on a squad name, the design tab was renamed. Lee.
Ten squads renamed, including `Z Energy Martech` to `Z Loyalty & Martech`,
`AU CRM & Martech` to `Ampol Loyalty & Martech` and `Z Energy Apps` to `Z App and Web`.

**Why:** you said it plainly - REVIEW is the only source of truth, so nothing translates
REVIEW into something else.

### D5. The `Lists!Z:AA` name map is retired. Build, reversing my own earlier build.
It translated REVIEW names *into* design-tab names, which made the design tab the
authority. That is backwards. The cells now carry a one-line note saying it is retired.

### D6. The two squad folds stay merged. Lee.
You looked at them and said leave them.

### D6a. A design squad with nobody in it is removed. Lee.
"We cannot have squads with 0 people in them." Three went: `Digital Support NZ` on 1.2,
`EGI Data` and `Enterprise Data Delivery` on 1.3.

They were charging archetype cost against zero actual people, which is the one thing an
archetype-versus-actual comparison must not do.

### D7. Agreed role assignments. Lee.
| REVIEW row | Person | Goes to |
|---|---|---|
| 283 | | COE SA&D, Group Data |
| 313 | | COE SA&D, Group Data |
| 364 | George Moun | Infrastructure, Leadership |
| 528 | Jens Tom | Finance, SAP ERP |
| Customer's first five | | stay as Leadership, no squad |

Two roles are still open, in PLAN.md section 7: Viren Khatri (r104) and Jasper Na (r136).

---

## B. How cost is calculated

### D8. One formula for every cost cell. Build, at your instruction not to hardcode.
```
salary role   = base x (1 + STI + payroll + pension + CPI) + medical
day-rate role = day rate x days worked x (1 + CPI)
```
All 527 cells use that pair. 348 of them used to be typed numbers and the rest used three
different formula shapes.

**Why:** you asked why it was hardcoded and said we need consistency. The formula pair
reproduces 523 of the 524 stored values to the cent, so nothing was invented - the numbers
were already the components, just typed instead of calculated.

### D9. Days worked per year is an input, not a number inside formulas. Build.
`Lists!AG15`, currently 222, yellow. It used to be a 222 typed into 42 separate formulas.

### D10. Exactly one agreed cost override, and it is visible. Build, flagged to you.
Tim Corin (row 172) stores $275,810.25, which is a banded rate 26 other roles also carry.
His own components give $321,135. That is a commercial decision, not a formula, so it sits
in a yellow cell in column AU with a note beside it rather than being buried as a literal.

**To overturn:** clear `REVIEW!AU172` and he prices from his components like everyone else.

### D11. Cost is stated gross everywhere. Build, to end a contradiction you flagged.
The same $3.6m used to appear gross on 3.1 and 3.2 and net on 3.4, 0.2, 1.11 and 1.12,
with notes on 1.11 and 1.12 claiming net while their formulas were gross.

Presentation is gross on every tab. The portfolio-funded portion stays as its own column
(`3.4` column L) so the net figure is still visible without a second definition of the
same cost.

---

## C. Overhead and leadership

### D12. Overhead is six named lines, not a percentage. Lee's numbers, built out.
| Line | Rate ($m) | Units | Basis | Allowance ($m) |
|---|---|---|---|---|
| Head of Technology | 0.1375 | 10 | portfolios | 1.375 |
| Business Partner | 0.22 | 10 | portfolios | 2.200 |
| Domain Architect | 0.14 | 10 | portfolios | 1.400 |
| Delivery Manager | 0.084 | 30 | platforms | 2.520 |
| Technology Manager | 0.081 | 30 | platforms | 2.430 |
| Leadership - 8 GMs | 0.300 | 10 | portfolios | 3.000 |
| **Total** | | | | **12.925** |

Lives on `Lists!AF:AJ`. 62 of the 525 roles carry an overhead line; the other 463 sit in a
delivery squad.

### D13. "Leadership overhead" is the 8 GMs at $5.1m. Nothing else. Lee.
It excludes every other leadership role. The GMs are the only overhead line with no role
in REVIEW, so their number is a yellow input on `Lists` (AG11 count, AG12 cost) and sits
*above* the 525-role ledger rather than inside it.

**Why it matters:** if the GMs were counted inside the ledger the headcount would read 533
and every reconciliation would break. Sitting above it keeps 525 intact and still shows
the $5.1m.

### D14. There is no 1:1 allocation anywhere. Lee.
Technology Managers and Delivery Managers are priced at 30% allocated to one platform, as a
broad principle. That is why 30 units appear against a 10-portfolio model.

### D15. The COEs carry no overhead at all. Lee, twice.
`2.11 COE BP&T`, `2.12 COE SA&D` and `2.13 COE Cyber` have no overhead block. The rule is
in the grouping formula, so it cannot drift back in:

```
IF(OR(LEFT(portfolio,3)="COE", portfolio="EGI"), squad, ...)
```

### D16. Overhead logic was built last. Lee.
You asked for it in that order so the rest was locked first. It is now in.

---

## D. The 2.x working tabs

### D17. Delivery squads and overhead are two separate blocks on every 2.x tab. Build.
**Why:** archetype cost prices squads. If the 62 overhead people sit inside the squad rows
then archetype and actual are not comparing the same people, and you said there should not
be any leftover dollars. Split, the comparison is apples to apples and each block subtotals
on its own.

### D18. The lever sits on every role row, filled or vacant. Lee.
`Filled`, `Hire`, `Hold`, `Offshore`. Factors are a table on `Lists!AC:AD`:

| Lever | Cost factor |
|---|---|
| Filled | 1.0 |
| Hire | 1.0 |
| Hold | 0.0 |
| Offshore | 0.4 |

**Why a table:** you can change offshore from 0.4 to 0.35 in one cell and the whole
workbook moves. No formula edits.

**Proof it works:** Aaron Lu, filled, $235,334. Set him to Offshore and his squad drops to
$1.4500m and the portfolio to $15.3335m. Set him to Hold and it is $1.3559m and $15.2394m.
"Full cost of model" correctly stays still, because that line is the as-is.

### D19. Filled roles show a cost. Lee, explicitly reinstated.
This supersedes register items 30, 88 and 101, which banned costs against filled people.
You said: "I explicitly asked for the filled cost to show a cost."

### D20. Offshore keeps the headcount, Hold removes it. Build, flagged.
An offshored role is still a person, just a cheaper one. A held role is a decision not to
fill, so it leaves the FTE count.

### D20a. Cost after decisions is column H and nothing else. Build, fixing a regression.
Column H on the FTE rows is role cost times lever factor, one row per person. Once D19 put
a lever on filled roles too, H already contained everybody. The squad formula was still
adding the filled total on top of it, so every filled person was counted twice.

Ampol Web: eight people, all filled, no vacancies. Actual $1.591m, reported after
decisions $3.182m, change $1.591m - on a tab where nobody had touched a dropdown.

Now: after decisions = SUMIFS over column H by squad. With no lever pulled, after equals
actual on all fourteen tabs, and the sum equals the ledger.

### D21. Vacancies default to Hire. Build, flagged, reversible.
**Why:** a default of Hold would show the organisation costing less than it does simply
because nobody has touched a dropdown yet. Hire states the full cost of the plan as written
and every lever pull is then a visible saving against it.

**To overturn:** the default is set in one place per tab and can be flipped to Hold.

### D22. Grouping is done by a derived column, not by squad name directly. Build.
`REVIEW!AT` returns the canonical squad for a delivery role, the overhead line name for an
overhead role, and the squad for anyone in a COE or EGI. Every 2.x tab groups on AT.

**Why:** it puts the whole classification in one auditable formula. When the rule changed
for COEs, one column changed and thirteen tabs followed.

### D23. Derived helper columns on REVIEW, all formula-driven. Build.
AJ portfolio, AK status, AP canonical squad, AQ leadership flag, AR overhead line,
AS design squad name, AT grouping column, AU agreed cost override.

---

## E. Summaries and reconciliation

### D24. One sign convention. Build, at your instruction.
**Variance = actual minus budget. Positive means over.** Everywhere.

Four conventions used to be in play, two of them side by side in `3.1` row 6. Every
variance column is relabelled with the convention in the header so it cannot be misread.

### D25. One budget basis on `0.2` column F. Build.
It reads actual ledger cost for every row. It used to mix three bases: design cost for the
ten portfolios, net actual for two COEs, gross actual for Cyber, and nothing at all for EGI.

Result: the workbook stated two budget variances $54.3m apart. Both `0.2!G26` and
`3.1!E20` now read the same +$64.613m.

### D26. A portfolio that draws on two budget lines is charged against the first one only. Build.
Customer, Cyber, BP&T and SA&D each have two budget lines. The second shows nil with a note
saying "charged on the first line above".

**Why:** otherwise the column double counts the portfolio and stops tying to the ledger.

### D27. `3.4` funding drawdown is found by label, not by row number. Build.
The budget block sits at a different row on almost every 1.x tab. The formula walks column H
looking for the label.

It used to report $0.5m drawn because two of its three formulas pointed at blank cells and
the third was a typed zero. Now: OpEx $3.150m against a $3.0m pool, Significant Items
$11.924m against $20.2m, CapEx $10.410m against $4.9m.

### D28. The offshore choice on a 1.x design tab reaches 2.x. Build.
`0.3 Squad Archetypes` prices both - column G onshore, column H offshore at 40%. The 2.x
archetype cost used to read column G unconditionally, so six squads *designed* offshore
were *priced* onshore and the dropdown moved nothing. It now follows the design.

### D29. Join tabs by name, never by row. Build, standing rule.
Row positions have moved three times on this job. Each time, something downstream read a
wrong-but-valid number with no error to warn anyone. Three real defects came from this
(`Lists!K2:K12`, `3.4!I11`, the 3.3 group-total row).

---

## F. Formatting and presentation

### D30. Yellow means input, and only true inputs are yellow. Lee, standing.
QA treats a yellow cell as a declared input and does not flag it as a stale literal, so the
convention is load-bearing, not decorative.

### D31. 3.x tables follow 2.x. Lee.
Same header style, same money format, same decimal places, same column widths where the
columns mean the same thing.

### D32. No conditional formatting on 1.x. Lee.
Variance is plain over/(under). No red/green judgement colouring.

### D33. Plain English, your vocabulary. Lee, standing.
Archetype, TDD cost, lights on, left to fund, vacancy lever, vacancy decisions.
Banned permanently: "call"/"calls", "roster", "seats", an invented "Category" column,
em dashes and en dashes, possessive AI phrasing, "GM working copy".

### D34. Headers have to make sense on their face. Lee.
"Total to fund" as a block header was meaningless and is gone.

---

## G. Decisions reversed

### D35. Filled costs: hidden, then shown. Lee reversed me.
I removed them citing register items 30, 88 and 101. You reinstated them. Register items
30, 88 and 101 are superseded on this point and nothing else.

### D36. Spans and layers analysis: withdrawn. Lee.
"never asked for spans & layers bro."

### D37. Phasing: withdrawn. Lee.
"phasing? huh?"

### D38. Contingent-workforce analysis: withdrawn. Lee.
"188 roles contingent does not seem right. this roughly language and 'look' contingent
means you have not done any proper analysis." Correct - it was a guess dressed as a finding.
If it is wanted, it gets done properly from the data or not at all.

### D39. Overhead as a percentage: never built. Build, self-corrected.
An allocation percentage would have been quicker and would have hidden the fact that
overhead is six specific groups of named people. D12 is the version that survives scrutiny.

---

## H. How the workbook gets built

### D40. Never save a populated workbook with openpyxl and ship it. Build, hard-won.
openpyxl strips cached values on save, so the file opens blank in Excel until something
recalculates. Every build goes through `scripts/v10/wbio.py`, which copies, recalculates in
LibreOffice, harvests the values, injects them back into the XML and then **asserts** that
every formula cell the engine valued ended up carrying that value.

Two bugs in that harness caused silent data loss before they were found:
- the self-closing `<c/>` branch of the cell regex had to come first, otherwise an empty
  cell swallowed the next one and its formula shipped with no cached value;
- the sheet-to-part map assumed the order of attributes in the rels file, which differs
  between Excel and openpyxl, and silently injected zero cells.

### D41. QA is adversarial, not an error count. Build.
`scripts/v10/qa.py` assumes the workbook is wrong and tries to prove it. Seven checks:
formula errors, dangling references, silent zeros, bad SUM ranges, stale literals,
1.x/2.x family inconsistency, cross-tab fact disagreement.

**Why:** an error count only catches `#REF!`. Every serious defect on this job returned a
perfectly valid number. The Exec Summary reported 22 in a cell labelled "$m" - it was a role
count. `Lists!K2:K12` returned another portfolio's filled count. A budget label read
"(see 1.13)" where `0.2!B7` said "(see 1.13 Cyber Roles)", so COE Cyber's budget silently
vanished. None of those are errors. All of them are wrong.

### D42. Column AU, not column AB, holds the cost override. Build, after a near miss.
I first used AB. AB is "MyHR ee no" and holds 27 employee numbers. The formula read them as
costs and dropped $1.87m out of the model.

---

## I. Still open

These are in `docs/PLAN.md` section 6 and are not decided yet:

1. r364 George Moun's squad.
2. Whether to delete the three design squads with nobody in them.
3. Whether to build the bridge tab that walks $72.8m of archetype to $115.1m of actual.
4. Whether to restructure 1.4 so the funding block sits at the same row on every 1.x tab.
