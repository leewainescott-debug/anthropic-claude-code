# TDD Cost Calculator - decision log

Live document. Every decision that shapes the workbook, why it was made, and who made it.
Companion file: `docs/PLAN.md` (what we are doing and what is outstanding).

"Lee" means you decided it. "Build" means I decided it and it is reversible - each one is
flagged so you can overturn it without hunting for where it lives.

Last updated: after the squad-assignment and design-rename build.

---

## A. Source of truth and data handling

### D1. `REVIEW - Complete Role Mapping` is the only ledger. Lee.
525 roles, $115,283,002.27, 390 filled, 135 vacant.

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

### D10a. The COEs and EGI are actuals. Lee.
They have no squad archetype. Their design cost is the planned spend on their own 1.x tab,
which is built from the real roles: 1.11 gives 6.400562, 1.12 gives 6.529494, 1.13 gives
9.897858, each equal to its actual. EGI is actuals too. So all four read a nil variance,
and that is the honest statement rather than a gap invented to fill a column.

Two earlier attempts at this line were wrong and are recorded so they are not repeated:

- **3.4's "budget to draw down" as the design.** 12.00 against 27.77 actual, presented as
  the COEs being 15.77m over design. That is the budget comparison Lee had just ruled out,
  wearing a different word. There is no COE archetype for it to be a variance against.
- **EGI priced off the strategic-programme inputs.** 1.52 against 4.94, reported as a
  3.42m gap. Those inputs are not a design for EGI.

### D10b. One input colour: cream. Lee.
`FFFFF2CC`. The file had bright yellow and cream both meaning "typed input", side by side
on the same tabs. Every input in the live model is now cream. Bright yellow remains only on
the retired raw-data tabs - `Squads`, `Added data`, `Sheet2` - which are not part of the
model and carry Lee's own marks.

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

### D23a. No budget on the summaries. Lee.
"Forget budget because you're confused on the split." "No point using the budget."

3.1, 3.2 and 3.3 compare **design cost to actual cost**, full cost both sides. No budget
column anywhere on them.

**Why it had to go.** I had put the 0.2 allocation against full actual cost - 50.5 against
115.11 - which can only ever read massively over, because 36.01m of the design is
explicitly funded outside TDD and the actual has no TDD/outside split at all. Lee's
words: "How is that relevant. They will always be over budget?"

The file also holds three different figures that all get called the budget, which is why
this went round twice. For Ampol Retail: 5.50 (the 0.2 allocation), 11.58 (the amount that
can be allocated to people, summed across every funding line on 1.1) and 33.80 (the full
budget including non-people). Against 14.01 of actual cost the variance is +8.51, +2.42 or
-19.79 depending which you pick. Dropping budget from the summaries removes the ambiguity
rather than resolving it; the funding tables on the 1.x tabs still carry all three.

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


### D43. The archetype prices some squads. Never put it against the actual for all of them.
`2.x` splits its squads into three blocks and `3.1` into four. The archetype block compares
like for like; the directly funded block compares actual to the amount funded on the 1.x
tab; the COE block compares actual to the planned spend on its own 1.x tab; overhead
compares actual to the allowance.

**Why:** the delivery subtotal used to hold the archetype cost of the seven squads that have
one against the actual cost of all ten. The group read **+11.488 over**. The squads the
archetype actually prices are **0.389 under**. The difference was the directly funded
programmes being charged against a zero, and it published on Exec.

**What it is not:** the eleven directly funded lines are not homeless. Every one sits on a
platform block on a 1.x tab and always rolled into its portfolio total. Eight carry a funded
figure typed against them; four of those figures are typed zero or left blank, all of them
EGI, which is why "EGI is actuals" is the only ruling that works. AmPOS (funded 1.404,
actual 2.115) and CTRM (funded 3.800, actual 3.222) are the only two lines in that block
with a real comparison.

### D44. Only three of the six overhead lines draw in the portfolios.
`Lists` now carries a Yes/No column and `Lists!AJ9` totals the three that do: Head of
Technology, Delivery Manager, Technology Manager, $6.325m. `3.1`'s overhead row compares
that to the $11.65m of portfolio overhead cost. `3.2` states all six line by line and splits
each line's roles between the portfolios and the COEs.

**Why:** `3.2` gave two different answers for the same allowance seven rows apart - row 8
said 43 roles / $11.654m / **-1.271 under**, row 19 said 62 roles / $22.921m / **+9.996
over**. Row 19 counted the 19 roles carrying an overhead title inside the COEs, whose cost
is already in the COE block, plus the $5.1m of GMs, who are not in the ledger at all. The
netting also hid the shape: the Business Partner and Domain Architect allowances, $3.6m, are
not drawn in the portfolios at all, so setting them against portfolio cost only made an
overspend read as headroom.

**The GMs:** $5.1m, eight people, no role in REVIEW. Leaving them out of the headline
understated TDD by $5.1m on the one page a GM reads first, so 3.1, 3.2 and Exec now all
carry the ledger total, the GM line beneath it, and a grand total of **$120.213m / 533
roles**. The ledger row is untouched at $115,283,002.27 and every control and check still
ties to it. 4.0 tests the grand total against the ledger plus the GM input, and tests that
the five variance lines on Exec add to the total shown - they did not, because the total
included the GM layer and nothing listed it.

### D45. "Not covered by the allowance", not "over budget". Register item 99.
The overhead rates are allocated shares - half a Head of Technology per portfolio, 30% of a
manager per platform - and the actual is whole heads. Calling the difference an overspend
asserts an equivalence that does not exist. The column says what it is, the rate column
carries three decimals so rate x times applied equals the allowance on the face of the page,
and the basis sits on 0.2 Data Config.

### D46. Roles after decisions, not vacancies remaining.
The old column counted vacancies set to Hold and was headed "Vacancies remaining", which
reads as though a cancelled vacancy were still outstanding. It is now roles less anything on
hold, on 2.x, 3.1, 3.3 and Exec. Pulling Hold on a $202,853 vacancy moves cost by -0.2029
and headcount by -1, on every one of them, and leaves cost today untouched. Verified by
recalculating the workbook, not by reading the chain. D20 promised this and nothing
implemented it.

### D47. Three moves proposed, one kept.
`Lists!AN:AQ` is the visible override table, now ten slots wide and carrying an overhead-line
override as well as portfolio and squad.

- **Jasper Na (r136), Energy -> Ampol Web.** Kept. He reports to Jin Zhong, Lead Engineer -
  BE, who sits in Ampol Web, as does the only other person under Jin Zhong. "Energy" was a
  one-person squad on no design tab.
- **Viren Khatri (r104) -> TDD Group Functions.** Withdrawn on the owner's instruction:
  Ampol Retail is his home.
- **Vikram Chhahira (r448) -> the P&C management line.** Withdrawn: EGI P&C is a squad in
  its own right.

**A near miss worth recording:** the first version of this wrote the new moves over rows 2
to 4 of the table, which already held the three agreed moves. That silently pulled two roles
back out of COE SA&D and moved a third out of SAP ERP - the ledger still totalled
$115,283,002.27 and no control failed. The table is appended to, never overwritten. Widening
the key range from `$AN$2:$AN$4` to `$AN$2:$AN$11` and leaving `$AO$2:$AO$4` alone was the
same class of bug: a match at position four indexed past the end of a three-row range, so an
override in a new slot returned nothing.

### D48. 1.12's roles list was three roles short and said so out loud.
Its own control cell read **-3**. REVIEW rows 283, 301 and 313, $747,896, were never written
in. The counts and planned spend on that tab read the whole ledger and were always right,
but the AU / NZ split sums the list, so the split came to 5.7816 against a planned spend of
6.5295 on the same row. 4.0 now reads all three COE control cells, and checks each COE's
planned spend against the ledger.

The rest of that gap is the offshore toggle, which feeds the split and nothing else, while
the tab claimed "the totals above and every summary follow". The discount is now a line of
its own so the split adds back to planned spend, and the note says the decision that moves
cost is made on the working tab.

### D49. Two tables cannot each set the width of a column they share.
Fixed three times in this build, in three places, always the same way: the second header
call overwrote the first. On 3.2 it dropped column B from 58 to 26 and clipped every label
above. On 2.x the third block reset a column the second had widened, so "New variance to
funding ($m)" needed four lines in the height computed for three and its first line rendered
off the top. On the COE tabs a summary block counting roles in column C shares it with a
roles list holding position titles.

`opts.head` now sizes a header row from its own text: greedy word wrap at 85% of the column
width, widen the column to 24 first, then make the row tall enough for what is left. The COE
and input tabs are measured from their cached content, because every position title on those
lists is a formula and a length test on the formula text measures nothing.

### D50. A merged bar is not a one-cell bar.
The design review reported navy bars "1 cell wide over 9-column tables" on 1.11, 1.12 and
3.4. They were merged ranges; openpyxl reports the fill on the anchor cell only. `verify.py`
checks the merge first. Fifteen bars were genuinely wrong - one column short on the ten
portfolio summary bars and one column long on the COE roles bars - and `polish.py` now sizes
every bar from the header row beneath it and clears anything painted past it.

### D51. Do not call it design. It is the archetype.
Owner's instruction. Every column, header, tile, check name and block label that said
"design cost", "over/(under) design" or "variance to design" now says archetype. The 1.x
tabs are still where the archetype is priced; the word just does not appear as a
measurement.

### D62. A filled role priced at zero, and why nothing in the workbook could see it.
REVIEW row 491, Nidhi Aggarwal, Snr Engineer - Boomi, NZ, FTE 1.0, **status Filled, cost
$0**. Column T carries her local base of 150,000; column U, which the cost formula prices
off, is empty. She counted in the 525 headcount and contributed nothing to the total. She is
the only role in the ledger with T populated and U blank.

**Every one of the fifty-six live checks passed.** They all reconcile *to* the ledger, and a
zero inside the ledger reconciles perfectly. So does every control row, every cross-tab
agreement, and the recomputation pass, because that also builds from REVIEW. Only a reviewer
reading the raw columns found it.

The figure is derived from her own cohort, not invented: all thirteen other NZ roles in TDD
Group Functions price U = T x 0.92, then STI 0.15, pensions 0.05, CPI 0.03, no payroll
component. 150,000 x 0.92 x 1.23 = **169,740**. It sits in the cost-override column the
formula already honours, with its provenance in the cell beside it, so her raw columns stay
untouched. The ledger is now **$115,283,002.27**.

### D63. A total's variance must be measured on one basis, and where it cannot be, it says so.
Three rows were asserting a variance across two bases, and all three added up, which is why
four QA passes let them through.

- **The 2.x "Total portfolio" row.** Its archetype column prices two of the four sections and
  its actual covers all four. On 2.1 it read **(0.06)** beside two cells whose own difference
  was **+1.20**, and the sign was wrong on four tabs. It now carries a dash: the comparison
  belongs on the section subtotals, where both sides are on one basis.
- **The directly funded step.** It held eight programmes with a funded figure and two
  Leadership groups with none, so the subtotal charged **1.30 of leadership cost against
  nothing** and read 1.44 where the real comparison is 0.13. The groups with no figure are
  their own step now, on 2.x and on 3.1.
- **The ledger total and the grand total on 3.1.** Same arithmetic, same fix: a dash, with a
  new "Everything with a figure to compare" subtotal - **108.74 against 113.98, 5.24 over** -
  carrying the number that can be stated.

### D64. 0.1 and 0.4 are hidden, not restyled.
They are raw pastes: a red / amber / green traffic-light grid with the letters R, A and G in
the cells, 741 red number formats, Arial and Aptos, four blues that are not in the palette,
hidden columns, and a note reading "input retail SI number manually". Restyling evidence is
not tidying it, and it cannot ship visible. 0.2 Data Config is the built interface to both.

### D65. The [Red] strip was case-sensitive and Excel writes [RED].
971 number formats across 0.2, 1.11, 1.12, 1.13, REVIEW, 0.1 and 0.4 still printed negatives
in red after a pass that reported stripping zero. The test now lowercases.

### D56. 3.1 is a cost bridge. That is what 3D meant and the last build did not deliver it.
The owner picked layout 3D off the options paper. 3D is a cost bridge: start at the
archetype cost, walk to the actual cost, name every step, and push the per-portfolio detail
down to 3.3. What shipped instead was a four-block per-portfolio table, which is closer to
3A. That is why he said the 3.x tabs never followed instruction.

3.1 is now the bridge, and every directly funded programme is on it **by name** - AmPOS,
EGI Retail, EGI Customer, EGI TDD twice (Ampol Retail and TDD Group Functions, because
Viren Khatri stays in Retail), EGI P&C, EGI Finance, CTRM, and the two Leadership groups -
because "Directly funded, 10.44" tells a reader nothing about what the 10.44 is.

### D57. 3.2 restated 3.1 and was deleted. It is Overhead & Leadership now.
Its first block was 3.1's four subtotals copied into a second table; a reader got nothing
from it that 3.1 had not already given them. The one thing it carried that nothing else does
is the overhead allowance, line by line, so that is the whole tab: the six lines with rate,
times applied, allowance, roles and cost in the portfolios, roles and cost inside the COEs,
what is not covered, and a second block stating what the allowance is built from and where
each input is set.

### D58. Two more variance bugs, both of them a total measured on the wrong basis.
Neither was caught by any of the four QA passes, because every check tested totals against
the ledger and both of these were internally consistent additions of the wrong things.

- **A block subtotal summed its row variances.** A row with no archetype carries "-" in that
  column and drops out of a SUM, so the directly funded subtotal read **0.13 against a real
  1.44** and the bridge read 5.07 against 6.38. A subtotal variance is now actual less
  archetype, computed on the subtotal's own two figures.
- **A subtotal with no archetype at all reported the whole cost as a variance.** The overhead
  block's archetype column sums to zero, so actual-less-archetype made the entire $1.12m of
  overhead cost read as a variance. It reads "-" now.
- **And the portfolio total is the sum of the three block variances, not actual less
  archetype** - because the archetype column only prices the first block, so measuring the
  whole portfolio against it would charge overhead cost against a squad archetype.

### D59. No frozen panes. Owner's instruction, and it settles register 70 against 94.
Register item 70 asked for a frozen header on every long table; item 94 recorded that the
owner had them all removed. They were on. They are off everywhere, and `polish.py` strips
any that survive an earlier build rather than leaving it to chance.

### D60. "Seat" is never used.
Owner's instruction. "Of which people in seat today" is "Of which filled roles".

### D61. The 2.x overheads belong in the table up top.
They sat in a second table below the squads with a header row of their own, and that second
header carried column names that were wrong for it - "Archetype cost" over a block no
archetype prices. There is now one table, one header row, and three labelled sections inside
it: squads priced by an archetype, directly funded programmes, overhead roles. The column
names are the owner's own, off his markup of 2.8: Squad Size, Total roles, Total roles after
decisions, Squad cost after decisions.

All fourteen 2.x tabs measure as a single structural profile - identical column widths,
identical header text, identical section order.

### D62. Actual against archetype on the 1.x tabs, in two options.
Asked for a row beside the squad detail showing actual cost against archetype, driven by
formula off cost-after-decisions, on the squad tables and on the platform and portfolio
tables, as two workbooks with a different design each.

Both are built from the shipped file by `scripts/v10/chainAB.sh` and differ only in
placement. Option A appends two columns to every squad table and platform total. Option B
leaves every existing table exactly as it is and adds one table at the foot of each tab.
Neither changes a single existing formula, and neither changes a single column width.

### D63. A variance needs two figures on the same basis, at every level.
EGI P&C on 1.5 has no archetype - its size is a blank input the owner has not set - so
`actual - archetype` read its whole $0.24m as an overspend against nothing. The same shape
that produced D58 twice already, one level further down.

Three rules now, applied to every row of both options:
- a squad row states a variance only when both sides are figures, otherwise a dash;
- a platform total states one only when an archetype prices every squad in the block, so a
  block that is short a squad on the archetype side cannot report the gap as overspend;
- the portfolio block splits "squads priced by an archetype" from "squads with no archetype
  to price them" rather than adding them, which is exactly what 3.1 does.

On nine of ten tabs the second line does not appear at all, because there is nothing on it.

### D64. Nothing on the shipped tabs may be lost, and the checker has to prove it.
Option A first appended its two columns at a fixed K and L. The squad tables are not all the
same width: 1.4, 1.5 and 1.6 carry five more columns the owner added - Nbr Archetype Roles,
Published Roles, Review Outcome, Vacant Now, FY27 - and on 1.6 the first two of those were
silently replaced. Every one of the six QA passes stayed green. They hold typed numbers, so
nothing recalculated, no total moved and no check could see it.

The columns are now placed per tab, in the first adjacent pair that is empty on every row
the writer touches - K and L on nine tabs, P and Q on 1.6, after the owner's own columns.
A seventh pass compares every cell of the shipped workbook against the variant, formula and
cached value, and nothing may be lost.

### D65. Widening a table pushes the notes beside it along, and says so.
The owner's free-floating notes sit to the right of the squad tables - "Fully funded by
CTRM", "People in this program today cost 0.24m. Set the agreed cost in the cream cell." -
and run across the empty columns beside them. Two new columns in that space truncate them
to the first few words without touching a cell.

Placing the columns beyond the widest note instead left eleven blank columns between a table
and its own figures on three tabs, which is worse. So eight notes move along their own row to
the far side of the new columns, verbatim, and each move is listed in the build log and
reported by the QA pass as a move rather than folded into "unchanged". The first attempt at
this carried six table headers out of their own tables, because it walked left from the new
column and took whatever it found; it now only moves a cell that is past that block's own
last header column.

### D66. A checker must not inherit the writer's assumptions.
The first version of the 1.x QA pass found squad rows with the same block-walking code the
writer used, so a table the writer never saw would have been invisible to both. It now finds
its columns by reading the header text - which is how it survived the columns moving from K
to P on 1.6 - and finds its rows by matching against the group names on the working tab.

Three of its own bugs were found and fixed this way: it read the tab title on 1.5 as a squad
row, because the portfolio and one of its squads are both called P&C; it skipped every row of
option B, because those rows carry a B:C merge and the title filter keyed on merges; and it
compared a cached note string against a row of formulas and reported a moved note as lost.

### D67. The stale-literal check is scoped to a block, not a column.
`qa.py` flagged the owner's own Published Roles figures on 1.6 the moment option B put nine
formulas in the same column, forty rows below. A stale literal is a typed number sitting
among formulas in the same table; a table fifteen rows further down is a different table.
The check now splits a column into blocks on a run of empty rows. Confirmed still live by
planting a literal inside option B's own formula column and watching it fire.

### D68. The COE tabs say why they carry no comparison.
1.11, 1.12 and 1.13 are the only 1.x tabs with no actual-against-archetype figure. They are
COEs, funded by allocation, and no archetype prices them. Their summaries group by department
- Business Partnering is Commercial plus TDD Business Partner - while the working tabs group
by squad, so the split cannot be taken from either tab without inventing a mapping, and
inventing one would put a figure in front of the owner that nothing supports.

Each carries one line under its summary saying so and pointing at 3.4 COE Detail and the
working tab. A silent gap on three of thirteen tabs reads as an oversight.

### D53. A total that ignores the line above it.
`1.11!C15` "Total Business Partnering budget ($m)" was `=C14`, so `C13` - the $2.20m of
Business Partner cost funded out of portfolio overheads, sitting on the line directly above
it - fed nothing in the workbook. "Left to fund" read **2.400562** against a real
**0.200562**. Same shape on 1.12: the $1.40m of Domain Architect funding was unused and
left-to-fund read 2.529494 against 1.129494. Register item 44 asks for portfolio funding
plus both allocations. Nothing outside the two tabs reads those cells, so the fix is
contained - the COE figure 3.1 reads is planned spend, not budget.

### D54. The banned Category columns survived as defined names.
`BPTCat`, `SADCat` and `CYBCat` still pointed at `Lists!E2:G4`. No formula used them and no
COE tab has a Category column, so no number was ever affected - but the columns the owner
ruled out were still in the file and would have reappeared in any dropdown built off a
name. Names deleted, cells cleared.

### D55. The retired sources are deleted, not hidden. Owner's instruction.
`Sheet2` held an older cyber list - 52 roles, 10 vacant - against REVIEW's 46 and 4. That
disagreement got reported to the owner as a six-role hole in the ledger. It was not a hole.
It was a stale tab that should not have been in the file, and his answer was immediate: we
only care about REVIEW.

Deleted: `Sheet2`, `Squads`, `Added data`, `FY26 Budget (superseded)`,
`squad mapping (superseded)`, and `3.5 Source Reconciliation`. No formula read any of them
except 3.5, whose entire job - thirty-three formulas - was reconciling the retired Squads tab
to REVIEW. Once Squads is gone, 3.5 is comparing the source of truth against a source that
does not exist. The note on 1.13 claiming "roles and costs come straight from Sheet2" went
with them; it had not been true since that list was repointed at REVIEW.

44 sheets, one hidden (`Lists`). A hidden stale tab is still a stale tab: it ships, it can be
unhidden, and it will be read as a second opinion. `purge.py` refuses to delete a sheet any
formula still reads, so the check is in the build rather than in my head.

### D52c. An empty-looking block is not always an empty block.
1.5's EGI P&C block was collapsed as a shell because its funded input is blank. The squad is
real and carries a role in the ledger; the blank is an input the owner has not set. Restored
exactly as it was. Only 1.1's Pricing & WFM fold and 1.3's removed EGI Data are collapsed,
and both carry the fact on the line that remains.

### D52a. A tab's name must match its own title.
`3.3 FTE View` was headed "Squad Detail", `3.4 COE Summary` was headed "COE detail", and
`2.11 TDD Cyber` carried "COE Cyber" in its title and its portfolio cell while the ledger
calls it COE Cyber everywhere. Renamed to `3.3 Squad Detail`, `3.4 COE Detail` and
`2.11 COE Cyber`, with 105 references and labels repointed. The references are rewritten
before the sheet is renamed, because openpyxl does not follow a rename into formula text:
the tab would carry one name and every formula the other, which is a `#REF!` on open.

### D52b. The build has to run against its own output.
`overrides.py` asserted that the Lists columns it writes to are empty, so the second run
stopped on the header the first one wrote. Worse, `fix1x.py`'s empty-shell collapse read a
wider empty span each time it ran and ate three more rows on 1.5 on the second pass. The
assertion now accepts its own header, and the shipped file is built single-pass from the
previous commit's workbook rather than from the file the last run produced -
`scripts/v10/chain.sh` is that build, in order, and `base_ship.xlsx` is its input.

### D52. 1.14 TDD Cyber is gone.
A copy of 1.9 that reported $1.2925m for Cyber against the $9.898m on 1.13, with three black
bars from a theme fill, twenty styled rows holding nothing, and a 276-character build
changelog in its title bar. No formula anywhere referenced it. Removed rather than hidden,
because a hidden tab with wrong numbers still ships.

Also gone: 3.4's duplicate column (F was literally `=$K6`, headed "People cost, gross ($m)"
against K's "Gross people cost ($m)"), its budget comparison, the "EG" row on 0.2 that was a
broken duplicate of EGI, a stray `z` on 1.1, an unlabelled `1` on 1.9, three empty platform
shells with orphan headers and zero totals, 21 en dashes, and 239 `[Red]` number formats that
put the alarm colour on portfolios under their archetype.

---

## I. Still open

These are in `docs/PLAN.md` section 6 and are not decided yet:

1. The overhead allowance basis: half a Head of Technology per portfolio against whole
   heads. Labelled honestly, but only Lee can confirm the intent.
2. The 8 GMs at $5.1m, stated above the ledger rather than added into it.
3. Frozen panes: register item 70 asks for them, item 94 records their removal. Applied
   consistently on every table over 25 rows; one line to reverse.
4. Whether to build the bridge tab that walks archetype cost to actual cost line by line.
5. Whether to restructure 1.4 so the funding block sits at the same row on every 1.x tab.
