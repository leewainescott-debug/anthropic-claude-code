# TDD Cost Calculator - decision log

Live document. Every decision that shapes the workbook, why it was made, and who made it.
Companion file: `docs/PLAN.md` (what we are doing and what is outstanding).

"Lee" means you decided it. "Build" means I decided it and it is reversible - each one is
flagged so you can overturn it without hunting for where it lives.


## D121 - The model update on his 30/07 base (all his rulings, quoted)

His green light with rulings: "yes egi funded" (the eight Enterprise Data roles are
EGI-funded; they carry their own directly funded line on 2.3 and the 1,811,858.90
returns to every total); Holds used as exclusion hacks come off; "keep live" (his
8.98m of lever plays are the live scenario); "if we've named vacancies then they're
now named and not vacant" (all named vacancies are filled; the three rows carrying
"Remove" as a name are deleted properly - ledger 531 to 528); the Significant Items
budget "4.5" (0.1 updated, 1.2 relinked); Programme Management "customer only o/h"
(overhead line, allowance = its own cost, zero variance contribution); Ed Tacey
"follow data" (Head of Technology line); contractor end dates parked; Digital
Support NZ at the 0.32 archetype "but call out the 2.17 on the tab somewhere" (his
note now reads "CPI actuals; pull through (0.217)"); the four Move-to-Z people wait;
"best not increase cost" (the 13/07 EGI portfolio moves are not adopted - the
funded-outside split does the work); the 13/07 book is out of scope beyond its role
mapping tab; COE consolidation approved (one tab per COE, the 1.x trio removed);
no sheet protection ("i might've overtyped ur stupid formulas"); inputs are strict
dropdowns ("it should just be dropdowns"); controls stay functional but invisible
("controls and shit to be visible looks ai"); his language everywhere, and the
update builds ON his 30/07 file, never a regeneration of ours.

Verified state: 528 roles, actual 114.968474, funded outside TDD 18.384289 (EGI
11.880489, CTRM 3.8, AmPOS 1.404, Cyber Uplift 1.2998), TDD-funded 96.584185,
after levers 106.800980 across 511 roles, 0.2 spend 57.057227 vs 50.5. 87 pipeline
checks green twice plus an independent re-derivation of every headline from the
ledger. Pipeline: scripts/v10/update/u1-u7.

Last updated: after D120 - wave M, the cyber uplift restructure and the Customer corrections.

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

### D76. Customer, reconciled to the owner's own figures.
He sent his squad-and-FTE list for Customer and asked for it to be reflected on 1.x, 2.2 and
3.x. It is column K of REVIEW summed on column O, exactly, and four things in the build were
standing between his data and the workbook:

- **Two real squads were merged away.** `Lists!W:X` folded `Z Energy Martech` into
  `Z Loyalty & Martech` and `AU CRM & Martech` into `Ampol Loyalty & Martech`. Both are 2.0
  FTE squads in his own list and neither appeared anywhere in the file. Undoing it also puts
  Ampol Loyalty & Martech back to 6.8 and Z Loyalty & Martech to 12.6. Supersedes D6.
- **A squad was renamed.** The same table rewrote `Customer, AI` as `Customer AI`. His column
  K has the comma; the fold went the wrong way round, so the fold is out and the design tab
  is renamed to match the ledger instead - which is D4, applied properly.
- **A role I moved.** `Energy` is a one-person squad and I had moved Jasper Na off it onto
  Ampol Web. That was my call, not his, and his list has Energy at 1.0. The override is out.
  The three moves he did instruct stay.
- **FTE was not in the workbook at all.** He reconciles in FTE; the tabs counted heads. Seven
  roles in the ledger are part-time, so Customer is 83 people and 82.4 FTE, and his figures
  are the second one. FTE now sits beside Total roles on every working tab rather than
  replacing it - "525 roles" is used throughout the file and in every control.

Eleven of his twelve squads now match to the decimal and the portfolio total ties at 82.4
FTE. The twelfth is the open question in D77.

### D77. The eight overhead roles keep their own lines. Lee, asked and answered.
His list has `Leadership` at 12.0 FTE and `Z Loyalty & Martech` at 12.6. The workbook reads
5.0 and 11.6, because eight roles whose titles make them overhead are lifted out of those
squads onto their own lines so they can be measured against the allowance:

| Overhead line | FTE | Comes out of |
|---|---|---|
| Head of Technology | 2.0 | Leadership |
| Delivery Manager | 2.0 | Leadership 1.0, Z Loyalty & Martech 1.0 |
| Technology Manager | 4.0 | Leadership |

So Leadership 12.0 = 5.0 on its own line plus 7.0 across the three; Z Loyalty & Martech
12.6 = 11.6 plus the 1.0 Delivery Manager. The portfolio total is 82.4 FTE either way.

This was the one place his list and his previous instruction pointed different ways, so it
was put to him rather than guessed. He chose the overhead lines. The named people behind each
line are in the FTE block at the foot of the same tab, so the 12.0 can be traced back by
hand.

### D78. The 2707 consolidation round - what moved and why. Lee's instructions, executed in agent waves.
The owner sent three files: his review workbook (branched from an old generation, carrying
his 0.2/1.x redesign), the 2707 version (my latest ship plus his hand edits), and a new
Customer dataset. Fable planned and orchestrated; Opus agents analysed and verified;
docs/ORCHESTRATION_2707.md is the plan of record. The headline ledger moved
525 roles / $115,283,002 to **531 roles / 529.3 FTE / $115,589,735**:
+1,018,131 six new cyber vacancies (rows 529-534), -106,606 the SA&D re-level on row 321,
-604,792 the Customer restatement, +169,740 the standing Nidhi Aggarwal override.

### D79. The review workbook is the base; what its branch predates is re-applied, not lost.
rev.xlsx forked before the fold table, the lever factors, the allowance table, the
override table, the canonical-squad scaffold, the D4 design renames and the D6a
zero-people-squad removals existed. Everything on that list is recreated or re-applied by
the chain (ensure_lists.py, repair_design.py, merge_review.py), so his edits land on the
current model rather than dragging it back a generation. His three instructed role moves
(r283, r313, r528) are reinstated in the override table the same way.

### D80. The new Customer dataset replaces the block wholesale. Lee's data.
Same 83 people, FTE unchanged at 82.4, cost restated -604,792: the AU and NZ on-cost
parameter blocks were cross-contaminated in the old block (an Australian priced at 0.92 FX
with NZ pensions and medical) and the new file puts them back on the right people. Six
roles move out of Leadership into squads. Nine rows state a cost that disagrees with their
own components; each is priced at the stated cost through the AU agreed-cost override with
the discrepancy noted beside it in AV. Rob Jadrjevic ships at $0 with the dataset's own
note ("Cant find this record") parked visibly in the Commentry column - his file, his
figure, flagged rather than invented. Tim Corin's old banded-rate override (D10) is
superseded by the restatement.

### D81. His levers flow through the model, set once, priced everywhere. Lee's ask (T5).
His On/Off design: one lever column per COE tab (Onshore/Offshore/Hold; Hold added on
1.12 and 1.13, cost factor 0, role stays in the counts) and a squad-level Hybrid state
priced at the onshore/offshore midpoint. The working tabs' lever cells are synchronised
from his design-tab states by ledger row - five cyber Holds, three SA&D Holds plus three
Offshores, two BP&T Offshores, Stevani Kho Offshore, the vacant QA role at ledger row 431
set to Filled - so the file ships with his decisions live: after-decisions
$113.03m against $115.59m today. The old "after equals today with no lever pulled" check
became "the bridge's decision impact equals the sum of the working tabs' impacts".

### D82. 2.x archetype cost reads the design tab's H, not the library directly.
He typed bottom-up squad costs over the archetype lookups on 1.2. On nine tabs H still
computes the library rate, so nothing moves there; on 1.2 the model now shows his figures.
One source of truth - the tab he edits - and 2.2's archetype total ties to 1.2's Total
Cost to the cent. His EGI Customer funding sits in the TDD Cost column (I54=2.21), so the
directly-funded comparison reads H first and I second.

### D83. His labels supersede two old banned-word rulings.
He renamed every budget bar "TDD Lights On Budget" and every funding block "Total to
fund", consistently, across the ten portfolio tabs. The phrase police now allows both;
"seat" and the design-language bans stand.

### D84. His 1.x footer shape: "Additional costs", and the control moves off the tab.
He deleted the "On the working tab with no row on this tab" residual and the on-tab
control row, and renamed the residual "Additional costs" where he kept it. Adopted
everywhere: the residual line appears only where it is a real figure, and the zero-check
now runs in Python on every build (qa1x) instead of sitting on the tab.

### D85. 0.2's Spend column is live, wired to where the figures actually live.
His rev formulas read the retired 3.4; the four COE cells now read the grouping
planned-spend cells on 1.11/1.12 directly, and the empty-cell budget reads he pointed at
0.1 are wrapped in N() - the house style for a read that may legitimately be blank. The
1.7 row offset his copy carries stays: every consumer joins by label, and the consistency
check now compares block order and completeness rather than absolute row numbers.

### D86. The verification wave caught the chain inventing three Holds. Root cause fixed.
The instruction audit of the first 2707 candidate found 1.12 carrying six Holds where Lee
set exactly three (rev H31/H43/H44). The three extras were on the rows the chain itself
appends - the roles moved into the Data COE - because `fix1x.extend_lists` copies every
populated column from the template row, and on 1.12 the template row is one of Lee's Holds,
so his lever and his "Hold" annotation were duplicated onto three roles he never touched,
zeroing $747,896.05 including two filled employees. Appended rows now arrive Onshore with
no annotation, and only the owner sets levers. Every control read zero throughout, which is
the lesson: a workbook internally consistent with a wrong decision proves nothing about
the decision.

### D87. 1.11/1.12 budget basis follows rev: net against net.
`fixcoe.fix_totals` was still rewriting his C15 to add the portfolio-funded line - correct
two generations ago when planned spend was gross, wrong since his restructure made spend
net of that funding (his own B9 note says so). The rewrite made the tabs disagree with 0.2
by exactly the funded amount. Removed; his =C14 stands, and the line I used to label
"Offshore discount in the AU / NZ split" is relabelled to what it actually is - the
Business Partner / Domain Architect funding met by portfolio overheads - with the value
beside its label. His On/Off notes come back too: in this generation the lever does price
through the totals, so the note rewrite that said otherwise described a model that no
longer exists.

### D88. "Actuals" on the 1.x summary means after decisions, because Lee wired it that way.
The audit flagged the 1.x Actuals column reading the after-decisions footer rather than
pre-decision actual cost (visible only on 1.7, the one tab with a lever pulled: 7.5817 vs
7.7484). His own 2707 formula is `G9 = K52` - the after-decisions total, on the tab where
his own Stevani Kho offshore decision was live. His wiring, his meaning of the word;
generalised unchanged to all ten tabs and recorded here rather than "fixed".

### D89. finish.py's literal sweep now carries a rev whitelist. Owner text is unclearable.
One step was wiping ten blocks of his content - open questions on 1.8, the basis notes on
1.11/1.12, his On/Off and source notes, the 0.2 reconciliation remnants. Any literal cell
whose sheet, coordinate and value match rev.xlsx is now exempt from every clearing pass in
that step. The general rule replaces cell-by-cell rescue because the next round of his
edits would just hit the same sweep.

### D90. The 0.2 "EG" row is his again. polish.py stops deleting it.
An old ruling removed "EG" as a broken duplicate of EGI. His rev re-adds it with zeros and
a reconciliation note beside it. Newest file wins, same principle as D83; the stray-row
clear is gone and his row 24 - including "Reallocated 7m across Ampol & Z Retail" - ships.

### D91. Sheet visibility matches his 2707 exactly.
rev arrived with Exec Summary hidden - an old-generation state - and it leaked through to
the candidate, hiding the one tab built for the CTO. base_2707 shows his intent: hidden is
exactly 0.1 Budget Table (Fin), 0.4 Presentation Pack and Lists; everything else visible.
polish.py now enforces that set both ways.

### D92. His review-time working tables are content, not clutter.
1.4, 1.5 and 1.6 carry his scratch tables - Nbr Archetype Roles / Published Roles / Review
Outcome / Vacant Now / FY27 - holding decisions like "Remove Lead Engineer Software and
Test". They stay, formatted so nothing truncates. The 1.6 copy sat inside the squad table's
span and punched a hole through the header bar, so that one moves right of the table
verbatim; K/L then carry the Actual/Variance pair like every other 1.x tab, retiring the
P/Q exception.

### D120. Wave M: the cyber uplift is a portfolio, the COE is what is left, and Customer gets his own figures back. Lee.

Eleven rulings in one round. His words are quoted; every figure below is derived from the
role mapping, not typed.

**1. "looks nothing like the other coe tabs and doesn't tell me the budget vs spend"** -
about 1.13 Cyber Roles. He was right twice over. Its nine summary columns matched 1.11 and
1.12 except that the seventh was "Left to fund" - a funding gap - where its siblings carry
the variance of budget against spend, and its budget line was buried in a four-row bucket
table with a CapEx input in the middle of it. 1.13 now reads Grouping / Roles / Filled /
Vacant / Planned spend / Budget to draw down / **Variance** / Cost AU / Cost NZ, with the
budget on the total row where the allocation sits and the variance beside it: 2.00 against
7.02, so (5.02). The funding block is two lines in 1.11's own list style - the COE
allocation off 0.2 Data Config, then the total to draw down - and the 0.5 CapEx input, its
bucket total and the Left to fund row are gone with every reference to them.

**2. Nine roles leave the COE for a TDD Cyber portfolio of their own.** Five staff a Cyber
Uplift squad (Catherine Gire, Dan Balsamo, Kevin Sheerin, Tony Keeler, Vivienne Vasak) and
four staff Identity (Cameron Watman, Iwan Wibisono, Joe Nahma, Raymond Cheung). They move
through the Lists override table, keyed on "Name | Position Title" as every move does - his
raw Portfolio column still reads "COE - Cyber, Risk & Operations" and it is not the build's
to change. The window widened from ten slots to twenty; eleven are used.

**3. Five COE roles are part-charged to the uplift programme, and stay in the COE.** Chris
Lyons 50%, James Byrne 50%, Rahul Sahni 40%, Darshan Suvama 25%, Vanessa Castro 0%. 1.13
carries a cream **Uplift %** column beside the On/Off lever and its cost engine prices
cost x lever factor x (1 - uplift %); the slice itself is stated in its own column, and
1.14's funding block totals that column. The COE's planned spend and 2.11's cost after
decisions are both net of the slice, role for role, so it is counted once.

The maths ties to his own figure to the cent: the Cyber Uplift squad's five roles cost
$1,299,996.55 and the five part-charges come to $494,819.45. **$1,794,816.00** - exactly
the figure he gave for the cyber uplift people.

**4. The COE's offshoring levers.** Jack Jenkins, Jas Mann, Ritika Salaria, the two vacant
Leads, Vacant (AKL) Tech Support Technician, and the two vacant Operations Analysts, which
move off Hold onto Offshore - a role being offshored is not a role being cancelled. Vacant
(SYD) stays on Hold; he is recruiting it. After the moves the COE is 43 roles - Cyber & Risk
19 (16 filled, 3 vacant), Service Operations 24 (17, 7) - 2 to hire, 8 offshored, 3 held,
40 after decisions, 8.78 actual and 7.02 planned.

**5. 0.2's combined cyber line splits in two, which his own action note asked for**
("Separate out and incl. cyber uplift funding bucket"). Row 7 is the COE at 1.5 AU / 0.5 NZ
reading 1.13's planned spend; row 23 is **TDD Cyber** at 1.0 AU / 0.5 NZ reading 1.14's own
TDD cost. The two allocations still add to the 3.5 he had on one row, so the allocated total
is 50.5 as it was. His offshoring note moves onto the COE row with the COE.

**6. 1.14 TDD Cyber takes the standard 1.x shape with two squads.** Cyber Uplift is a typed
cream 1.2998 at support % 0 - no archetype in his library prices a cyber uplift programme,
and support % 0 means the whole of it shows as funded outside TDD, which is what "fully
funded from the programme" means on that table. Identity prices off the archetype library
(Operations, S) at 0.8 with his **80% support toggle**, so 0.64 is TDD and 0.16 is charged to
the programmes. The programme's funding block states the 2.8 he set aside and what it is
paying for - the squad, the 1.13 part-charges, Identity's share - leaving 0.845 for
non-people. There is no Left to fund line: the programme funds all of it.

**7. "the coes and egi need to show archetype = actual"** on 3.1. There is no independent
plan for those four groups, which the tab used to state as a dash - and a dash is not a
number, so the total column could never add up and the tab's boldest row carried no
archetype figure at all. Archetype = actual says the same thing in a way the column can
carry: no variance, and every row in the total.

**8. 3.1 is his approved layout, and the walk is deleted.** Title "Archetype cost to actual
cost", band "TDD", one row per group with Ampol Retail and Z Retail adjacent and **no
subtotal row between them** (removed, not hidden - it was a display row inside a column of
group rows and it was in no total), the cyber COE named "Cyber, Risk & Service Operations",
a Total roles after decisions column on every line, then the TDD total, the GM layer and the
grand total. The forty-row step-by-step walk below it is gone: every step of it is still on
the tab it came from - 3.3 by squad, 3.2 by overhead line, each working tab by block - and
the Exec now quotes the group total instead of five walk subtotals.

**9. The EGI squads are funded by EGI at the actual cost of their roles, in every
portfolio.** Support % 0, Total Squad Cost reading the working tab's own after-decisions
total, so no part of them is a TDD cost, none of them lands in Squad Support Costs, and the
whole of each shows in the portfolio's Other column and so in its total cost. This settles
the open question left by D118: **his 27/07 typed 2.21 in 1.2's TDD Cost column comes out**
and the cell is his review workbook's own formula again, with EGI Customer's actual (2.099)
flowing to Significant Items EGI.

**10. Digital Support NZ comes back.** It is a squad on his own review workbook, row 41,
with "Net New" in the gutter beside it, and taking it off (D6a) removed 0.32 of NZ platform
cost. With it back and the 2.21 out, Customer reads his review workbook's own figures:
**Ampol Customer 2.439 against 2.5, Z Customer 5.314 against 4.0.** His note off that row
moves to the note margin, because K is now a live column.

**11. All six WiPro roles in Customer are offshored, and none of them reprices.** Every one
costs 73,260, which IS the vendor's offshore rate; multiplying it by 0.4 would invent a 60%
saving on a rate that already has it. So a role whose Type on the role mapping contains
WIPRO prices at the full rate whatever its lever says, and the guard sits in the cost-after
formula on every role rather than on those six, so a seventh vendor role is covered the day
it arrives. The lever moves the status and the vacancy split and not the price: Customer goes
to 8 to hire and 6 to offshore, the Exec to 126 / 12 / 6 with 5 filled roles offshored, and
Customer's cost does not move.

**Tab renamed.** 2.11 is "2.11 Cyber Risk & Service Ops" - 29 characters, because Excel
stops at 31 - and its own title, its 3.1 line and 3.4's total row say
"Cyber, Risk & Service Operations" in full. What is deliberately NOT renamed: the portfolio
cell at 2.11!C3, column B of 3.4, and REVIEW's grouping columns all still read "COE Cyber",
because every count and every sum in the file joins the ledger on that string. Renaming a
join key would take 43 roles off four tabs with every control still reading zero, because
both halves of each control read the same renamed key.

**No dashes.** His ruling, and it is absolute: no cell in this workbook shows a dash. Every
`="-"` literal is blank, every `"-"` fallback inside an IF or IFERROR returns `""`, and the
third section of every money and count format - the section that decides what a zero renders
as - is gone, so a zero reads 0.00 or 0. Swept at source in the writers and again on the
finished file, which is what makes it checkable in one place.

**Two gates were re-derived rather than re-baselined, and two were retired.** 1.2's F9 was
pinned at his 15.5625; two of his own rulings move it, so it is now pinned to equal its own
three columns and 2.2's archetype total - both sides move together or the check fails. The
Exec hold count was pinned at 8; it is now a recount of every vacant role held, off the
working tabs. The "all 40 squad formulas" count is derived rather than typed. Retired: the
1.13 CapEx pin (the input is gone) and the "1.14 awaits its inputs" pin (both squads are
priced).

### D108. 3.2 is his layout, and the sentence column is his to write.
He redrew the tab himself and his order is better: rate, times applied, roles priced for,
actual roles, roles not applied, archetype cost, actual cost, variance - roles read
together, then costs, each pair beside what it is measured against. Built to his mock,
his headings, no section bar. "Where they sit" was a five-deep nested IF that computed
the same COUNTIFS six times to build one sentence; it is now plain text, seeded from the
ledger on every build, and he can reword it. A new column states the allocation in words
("50% across 10 portfolios") so the rate does not have to be reverse-engineered. Times
applied is a cream cell he sets, seeded from the count the model carries, and the total
row dashes it because ten portfolios plus twenty-two platforms is not eighty-four of
anything.

### D109. The ledger join stops depending on row numbers. The worst trap in the file.
2,124 cells on the working tabs found each person through INDEX over a whole column at a
hardcoded row. Insert one row in REVIEW and every name, title, status and cost below it
reattaches to the wrong person - silently, with every control still reading zero. The COE
tabs already used the direct reference, which tracks an insert; all fourteen working tabs
now do. The three agreed moves were keyed on rows 283/313/528 and are now keyed on
"Name | Position Title", with the build refusing to run if a key is not unique - name
alone would not do, because 143 rows are called "Vacant".

### D110. Formulas that could not be followed, or were not doing what their label said.
An independent Opus audit read all 15,694 formulas. Fixed: the Exec line promising a COE
over/(under) that printed a dash - there is no independent plan to compare against, so it
states actual cost and the label stops promising a variance; the Exec vacancy counts that
scanned whole columns holding squad sizes, right only because the words never collided;
the "allowed for elsewhere" row whose platform half cancelled exactly but unreadably; the
geography rule that decided AU or NZ by asking which budget number was bigger, replaced by
a Home country column he sets; 102 COE cost engines that hardcoded the offshore factor
instead of reading the one table it lives in; a magic 0.5 duplicating an input cell; a
typed 50.5 duplicating a computed one; ~4,200 inert cells including a lookup into two
empty columns repeated 531 times and four ledger columns nothing read; and a "variance"
comparing two different populations that nothing read.

Kept deliberately, with comments so the next reader does not "simplify" them: the core
cost formula, which carries three genuinely different pricing bases; the 2.x column
width, which is the cost of the model doing two comparisons at once; and one
SUMPRODUCT gate on 3.1 that looks redundant beside its twin one section up but is
load-bearing - the twin was simplified, that one was not.

Two audit findings were wrong and the builders said so rather than implementing them: the
"omitted" funding line follows a deliberate family convention, and the row-26 cancellation
was algebraically exact rather than luck. Both were reworked for readability instead.

### D111. recompute.py had stopped checking 3.2 and said so in a line that read like a note.
When 3.2 was rebuilt to his layout, the recomputation's column lookups kept the retired
headings, returned nothing, and skipped the entire tab behind the message "3.2 is missing
one of its columns". Every figure on that tab had been unverified since. Rewired, and the
gate now fails if that message ever appears again.

### D112. 1.2!C7 restored to his true shape - and the first version of this entry was wrong.
This entry as first written said his C7 and D7 were byte-identical and that 0.99 of platform
overhead was "his shape". **That was false.** The wave-K fidelity review checked both of his
workbooks and found:

- **His review workbook:** C7 sums column H (empty on the overhead rows), D7 sums column I.
  One branch carries a figure. F7 = 0.495.
- **His 27/07 workbook:** C7 is the exact complement of D7 - the `0` and the `SUM` swap
  branches, so exactly one fires whichever country is home. F7 = 0.495.
- **Neither of his books ever counted the overhead twice.** The double count was introduced
  by `repair_design.py`, which "corrected" his C7's H columns to I and made C7 identical to
  D7. My later `*0.5` "fix" halved my own defect; his "revert 1.2 plz" then removed the
  halving, restoring the defect, and this entry blamed the result on him.

C7 now ships in his 27/07 complement shape - his newest statement of the cell - and the
figures are his again: F7 0.495, F9 15.5625, 22 platforms on Lists!AH5, the 5.005 overhead
allowance on 3.1/3.2. `regress2707.py` pins the complement shape, F7, F9 and the exact
platform-count equality, so neither the corruption nor the false "his shape" story can
return without the gate going red.

Kept in this corrected form rather than deleted, because the lesson is the point: a claim
about "his shape" is checked against his files, not against the build's memory of them.

### D113. The Home country column is off 0.2, and the 1.x tabs decide it his way again. Lee.
"why have u put home country dat config???? pointless information i never asked for."

It was mine. The 1.x tabs decide whether a portfolio's cost lands in the AU or the NZ column
by comparing its two budget cells on 0.2 and taking the bigger, which is unguessable from the
formula; I added a column that said it outright. He did not ask for it, it is his config tab,
and the rule it replaced worked. Gone, and all eleven 1.x tabs are back on his comparison,
including 1.14 TDD Cyber, which was written directly against the new column and would have
been left pointing at an empty cell.

Two more of mine reversed in the same pass:
- **0.2's blank Legal, EG and EGI spend rows.** I wrote 0.00 into them so the column read
  consistently. A typed zero is a statement he did not make. Blank, as he has them.
- **"Budget to draw down" renamed to "Budget available"** on 1.11, 1.12 and 1.13.
  "Ledger" is my word and swapping it for "role mapping" was right. "Budget to draw down"
  is *his*, it is in his own review workbook, and D83 settled that his labels win. Restored.
  So is `0.2!M13`, which said "Alloc %" until a tidy-up made it "Allocation %".

### D114. What of his is not in the build, answered mechanically rather than from memory.
`whatsgone.py` takes every value he typed in his own workbooks - literals only, not formula
results, which are meant to move - and asks whether that content still appears anywhere on
the same tab of the build. It runs on every build at the end of `chainA2.sh`; the current
counts are always the ones in `chainA2_run.log`, and this entry deliberately quotes none
(the first version quoted four and the fidelity review found all four stale within a day).

What the list contains, honestly stated - the fidelity review corrected this entry's first
claim that it was "not wording, it is naming":
- **Names**, the bulk: his hand-typed squad and platform labels normalised to his own role
  mapping's spelling so the join finds the people ("Network / QSR" to "Network & QSR",
  "AU Finance" to "SAP ERP", "Network & Infrastructure" to "Cloud, Network & Infra Ops",
  and the rest). REVIEW is the source of truth (D1); a label that does not match it does
  not join.
- **Typo fixes** where a word would otherwise mislead ("acorss", "Siginificant").
- **Cross-references updated to renamed tabs** inside two of his notes (1.12!B53's
  "3.3 FTE View" is now "3.3 Squad Detail", and its "4.3" a "2.3").
- **Two headers**: 1.4!H8 "(Over)/ Under budget" became the family's "TDD over/(under)
  budget ($m)" with its sign flipped to match (meaning preserved), and 1.5!C5 "TDD  ($m)"
  gained its missing country word.

Sentence-level rewording of his prose is not on the list and is not done. Every item above
is visible in the whatsgone output on every build, so a wording change cannot hide among
the name normalisations. Any of them is reversible on his word.

One naming case surfaced here was resolved by his word the next day: his role mapping
carries BOTH "AU CRM & Martech" (2 roles) and "Ampol Loyalty & Martech" (7), and both
"Z Energy Martech" and "Z Loyalty & Martech" - renaming his two 1.2 rows to the larger
squads had left the two small ones unpriced. He ruled they price; the rows are back
(D117).

### D115. Wave K: three reviewers over the shipped file, and what their findings changed.
Three Opus reviewers read the shipped workbook cold - a GM story lens, a CFO model lens and
a fidelity lens - after his instruction to QA everything to a standard fit for the GMs.
Mechanical arithmetic was already proven by the scripts; these were judgement findings.
What changed, beyond D112 and D116:

- **3.1's bold bottom line no longer strikes a "variance" it cannot mean.** The ledger row
  put the 395-role archetype base under the 531-role actual and called the 40.72 difference
  "Variance to archetype"; the Exec says 10.42 for the comparable set. Both reviewers took
  40.72 as the answer. The ledger and grand rows now carry a dash in the archetype and
  variance columns, like 3.3's Group total; the comparable variance stays on its own named
  rows, which is what the Exec quotes.
- **"Actual portfolio" on the 1.x tabs now reads the actual.** The row (and the summary
  "Actuals" column beside the budget box) read the *after-decisions* column, so the one tab
  with a lever set - 1.7 - showed 7.58 under the word "Actual" while the Exec's drill-down
  said 7.75. Any lever pulled in the room would have silently rewritten the "Actuals"
  column. Rewired to the working tabs' Actual cost column; after-decisions figures live
  under labels that say so. The "Variance to actuals" line under the summary also flipped
  to actual-less-archetype, so it no longer shows the same number as the table beside it
  with the opposite sign.
- **The Exec now answers the budget question, and names all three slices of the overhead
  gap.** "Overhead roles - not covered by the allowance" gained "in the portfolios" (it is
  the portfolio slice of 3.2's larger total), a sibling line states the COE and EGI slice,
  and one line reads 0.2's own bottom line - over/(under) the allocated budget - which is
  the question a finance partner arrives with and the Exec never answered.
- **One offshore rate.** Lists!AD5 (the lever's 40%) and 0.3!K5 (his archetype Offshore
  rate) were two independently typed 0.4s; retyping his cell would have moved half the
  model. AD5 now reads his cell. And the ten portfolio working tabs now say in words what
  the lever does - the sentence previously existed only on the three COE tabs.
- **Cream means typed input, everywhere.** A handful of formula cells wore the input
  colour (1.5's "half of each budget line" rule among them); the paint is stripped from
  formulas on every build, and the gate now fails if a formula is ever cream again.
- **His review-notes columns on 1.4, 1.5 and 1.6 are named as a snapshot.** His typed
  role and vacancy counts there disagree with the live model on five of eight squads. His
  numbers stand untouched; a line above them now says they are a hand-typed snapshot, not
  live figures - the rule for content of his the model disagrees with.
- **Smaller repairs:** 1.13's note pointed at 3.4 for an after-decisions figure 3.4 does
  not carry (now points at 3.1 and the working tabs); 3.2's rate note was half a sentence;
  3.2's Times-applied control now says "must be 0" and a note explains how a reader gets
  from the platform headings they can count to the count the allowance prices; 0.2!N14:N16
  show three decimals so the platform rate foots on the page; 4.0's title no longer repeats
  as its own subtitle and three check labels lost their shorthand; 1.14 says on its face
  that no roles carry TDD Cyber yet; Lists!AF13 said "525-role ledger"; built prose now
  spells "programs" the way his own squad type does; 1.5's empty NZ column no longer gets
  a sibling-copied header (he removed that column - dressing it as live was mine, now
  gated).

Findings *rejected*, because the content is his: 1.8's typed 7.2 against budget lines
totalling 8.9 and its two open questions; the 0.2-versus-1.x sign conventions (both his,
both labelled); the three COE tabs' three different summary schemas; his workings notes in
the M columns (including "7 FTE in AmPOS" where the model carries 10); his "Hybrid" notes
beside squads set Onshore; the four spellings of Significant Items on 1.2; 0.2's zeroed
COE Cyber row; the dated "Position on 23/07" snapshot; hidden 0.1/0.4. All surfaced in the
hand-back for his ruling, none changed.

### D116. 0.3!C25 states his hybrid rule the right way round. His newer word over his older.
His note read "Hybrid = 2 roles offshore, rest offshore" - self-contradictory, and backwards
against the rule he set in as many words ("assume 2 roles would be onshore and the rest of
the roles offshore", D104). It is the only written statement of the rule in the workbook, on
the tab that defines the archetypes, so as typed it taught every reader the rule backwards.
Same precedence as D83: his newer instruction outranks his older text. Corrected to
"Hybrid = 2 roles onshore, rest offshore"; hybrid.py documents the exemption; the gate pins
the corrected sentence. One word from him reverses it.

### D119. Deepali and the vacant Service Transition Lead sit in Enterprise Data. Lee, twice.
"we literally said deepali and service transition lead DO NOT sit in COE- data, they sit
in Ent Data."

He had said it, explicitly - "Deepali and the service transition lead sit in the Ent Data
world **(not COE)** in the leadership group" - and D118 item 3 got it backwards: his
message's later sentence ("deepali is in enterprise data data coe ... also coe") read as
if it overrode the first, and the build followed the wrong one. When his own two sentences
disagree, the rule is now to ask, not choose; the explicit parenthetical was the ruling.

Both roles are home: no overrides, their own raw rows (Enterprise Data / Leadership)
place them, they price on 2.3 in Enterprise Data's leadership group, and the Data COE
returns to his 1.72 against its 2.0 allocation. The two D7 table entries that moved them
(283, 313) are superseded and removed; D7's remaining rows (Jens Tom) stand. The 0.2 note
he asked for flips to state the boundary the right way round: they are NOT in the COE -
Data line, they sit in Enterprise Data's leadership group. Rihan's D118 ruling
(Reporting & Analytics) is unchanged.

### D118. The two "missing squads" never existed; his org rulings land. Lee, in his words.
**Item 3 of this entry was implemented backwards and corrected the next day - see D119.**
Four rulings in one message, reversing D117's premise and reshaping 3.1:

**1. "z energy martech is z loyalty & martech. same frickin squad, not a new squad."**
His dataset's pivot lists old and new names with separate FTE lines, which is what read as
four squads; his person-level sheet puts every one of those people in one squad. The two
fold rows go back (only the "Customer, AI" rename stays unfolded), D117's two added 1.2
rows and their insertion machinery are gone, and folded, the squads carry exactly his own
pivot totals: Ampol Loyalty & Martech 9 roles / 8.8 FTE, Z Loyalty & Martech 17 / 16.6.
1.2!F9 returns to his 15.5625.

**2. "Rihan is in the reporting and analytics squad."** His raw row says Portfolio "COE -
Strategy, Architecture, Data" / Platform "Group Data", so Rihan had been counting in the
Data COE. A grouping override (raw columns untouched, as ever) moves him to Enterprise
Data, Reporting & Analytics. The Data COE drops to 2.22.

**3. "deepali is in enterprise data data coe (need to add a note in config tab to explain
this), same with the vacant service transition lead, also coe."** The two D7 moves stand -
they count in the Data COE - and 0.2 now carries the note he asked for, in its own Notes
column beside the COE - Data row: they sit in Enterprise Data's leadership group and are
counted in the Data COE.

**4. "in 3.1, energy and leadership need to be included in the total portfolio cost ...
coes need to be grouped with portfolios ... we need to show the total up there early.
ampol retail and z retail need to be grouped together."** 3.1 now opens with a
whole-organisation block: one row per group including the COEs and EGI, the actual and
after-decisions figures reading each working tab's own Total portfolio row - so
leadership, energy-type squads and directly funded programs are all in. The archetype
side is the tab's whole design build (same scope, so the group-level variance is apples
to apples); the COEs carry a dash. Ampol Retail and Z Retail sit together under a Retail
subtotal, and the TDD total row sits at the top of the tab. The step-by-step walk stays
below it, unchanged, as the reconciliation; a 4.0 check ties the block's total to the
walk's ledger row so the two can never drift.

Still open from this exchange: EGI Customer's 2.21 (his 27/07 typed figure in 1.2's TDD
Cost column; his review book's formula there returns 0). Until he rules, it stays -
Ampol Customer reads 4.65 against 2.5, and 2.44 of that question is this one cell.

### D117. Superseded by D118 - kept for the mechanics. Was: the two archetype rows on 1.2.
"we need to fix 5... they should [have archetype rows] if they're squads." His role mapping
carries four Martech/CRM squads in Customer; his 1.2 carried rows for two. The name
normalisation (D114) had matched his two typed rows to the two big squads, leaving
AU CRM & Martech (two WIPRO vacancies, $0.147m) and Z Energy Martech ($0.288m) with people
on the working tab, a named line in 3.1's "nothing prices these" bucket, and no archetype
row pricing them.

Two rows added by `repair_design.add_missing_squads`, wired identically to their siblings
(archetype lookup, TDD/funded-outside split, the working-tab actuals join, the hybrid
sweep). Z Energy Martech sits in row 41 - the row his removed Digital Support NZ squad
vacated, already inside every range on the tab. AU CRM & Martech is a true insertion in the
Group Customer block; the seven formulas that cross the insertion point are rewritten by
exact text match, the pass proves nothing elsewhere references the shifted rows before it
touches anything, and it refuses to run at all if any expected text is absent.

**The four pricing inputs on each row are seeds, not his numbers - they are cream and his
to retype.** Both seed as Configuration / Integration XS, the smallest archetype in his
library (1 role, $0.4m): AU CRM & Martech at support 0.2 like its block siblings, Z Energy
Martech at support 1 like its. The first shipped seed priced Z Energy Martech as Product S
- a 5.5-role, $1.3m archetype on a two-person squad - which moved Z Customer's 0.2 spend
from 4.99 to 6.29 overnight and he rejected it on sight ("z customer so far from
budget?"). The lesson is recorded: a seed on his tab is sized to overstate least, not to
match the squad's type family.

The gate pins 42 hybrid-swept squad formulas (was 40) and the new 1.2 totals; the two
squads leave 3.1's unpriced bucket and join Customer's archetype line.

### D104. Hybrid prices two roles onshore and the rest offshore, and the two is his to set.
His rule, replacing the 50/50 assumption: per-FTE cost is the archetype's squad cost over
its # of roles; two FTE price onshore, the remainder at the offshore rate. The "2" is a
cream input on 0.3 beside his offshore rate, so the rule re-prices the whole workbook
from one cell. MIN(k, n) makes the one-person archetype price fully onshore, and
fractional squad sizes fall out naturally. All 40 squad formulas across the eleven design
tabs carry the rule; nothing moves until a squad is set to Hybrid. 0.3 stays his tab -
the parity check now exempts exactly the two declared input cells and nothing else.

### D105. The 1.x comparison table is his mock, verbatim.
"Actuals vs archetype": Actual portfolio, Archetype portfolio, Variance - roles and cost,
live off the working tab's own totals, identical on all eleven tabs. Roles carry one
decimal because the archetype side is FTE and a rounded column stops adding on the page.
The Actual row reads the after-decisions total, which is what his own original Actuals
wiring read (D88); on 1.7 - the one tab with a lever pulled - that is 7.58 against a
pre-decision 7.75, and both facts are one column apart on 2.7.

### D106. 3.2 tells his story: applied to the portfolios against the organisation.
Per overhead line: the roles and dollars the allowance applies to the portfolios, the
roles and cost the organisation actually carries - all of them, everywhere - and the two
gaps between those numbers, with one plain-English cell saying where the people sit. The
Yes/No column and the extra split columns are gone. The counted-once row and its control
stay, and the partition now counts future TDD Cyber roles by name so a role can never
fall out of both halves.

### D107. The Exec vacancy block is a true partition, proven by the triple check.
Three Opus checkers verified the people layer three independent ways: a raw-column
recompute (531/531 on tab, squad, overhead line and cost), a full formula re-evaluation
(15,682 of 15,682 reproduce their cached values) and a person-by-person roster diff
(0 missing, 0 duplicated, 0 misplaced). The one logic defect found: the Exec vacancy
lines summed lever counts blind to status and added to 145 only because two errors
cancelled. The block is now five status-qualified lines - hire 134, offshore 2, hold 8,
fill-as-is 1 - that genuinely partition the 145, with the filled role he offshored
(Stevani Kho) stated on its own line. Data facts surfaced for his ruling, not changed:
Rob Jadrjevic (r190) ships at $0 with his file's own "can't find this record" note; six
ring-fenced placeholder rows count as Filled; seven part-timers may carry full-time
bases; the AU columns mean "not NZ" (six WIPRO rows, one USA, fourteen Singapore inside
them); his 1.12 grouping labels differ from the ledger's for five roles; and his S:W
review tables hold nine stale review-time figures.

### D100. 1.14 TDD Cyber is a real portfolio tab, integrated end to end.
His ask: a 1.x tab exactly like the portfolio family, one platform (TDD Cyber), one
squad (Cyber Uplift). Built as 1.14 after the COE tabs, with a 2.15 working tab seeded
from the design tab that fills itself the moment a ledger role carries the portfolio,
3.1/3.3 lines, and 4.0 ties. It sits OUTSIDE the ten-portfolio list on Lists - the EGI
precedent - because joining it would silently grant $0.7975m of Business Partner /
Domain Architect / GM allowance for people who do not exist. The platform overhead IS
priced (his "exactly the same" rule): the group platform count is 22 and the overhead
allowance moves 4.84 to 5.005 on 3.1/3.2 together, controls at zero throughout. The
portfolio overhead is NOT drawn, and the tab says why on its face. His 0.2 row 23 -
"TDD Cyber incl. COE" - now reads both 1.13 and 1.14, which is what its own label says.
The squad ships with cream inputs empty: it reads "check size" the family way, sits in
3.1's archetype section with dashes, and prices itself into every total the moment he
sets Squad Type and Size. The archetype subtotal on 3.1 covers priced members only,
by the same ISNUMBER mechanism the directly-funded split already used, so an unpriced
line cannot dash out a whole section's total.

### D101. 0.3 Squad Archetypes is his source tab. The chain is locked out.
He asked, in as many words, why the chain had changed his cost library. It is now a
source tab like 0.1 and 0.4: no title, no width profile, no gutter, no print fit, no
restyle of any kind, and regress2707 proves it cell-for-cell against rev.xlsx including
widths and heights. The verify sweeps skip it for the same reason.

### D102. The 1.x actuals table sits up top, and it is a table.
His ask, verbatim: "table at bottom of 1.x tabs needs to be up top. it's also not a
clean table." It now sits beside the budget box at K4:N10 (one row lower on 1.7),
mirroring that box exactly: navy bar and header, label / Roles / Cost columns that
each decompose to their own total, dash for zero, the same fixed five rows on every
tab, box borders, zero row insertions anywhere. The old bottom block is gone. The
Portfolio Summary's Actuals cell wires to the relocated total by label.

### D103. 3.2 answers his three questions on its face.
Business Partner and Domain Architect rows state their COE roles in words ("No - all 6
of these roles sit in the COEs" / "all 7"), the opaque Yes/No column is a sentence
derived from the same Lists cell that drives the maths, and one band states the whole
model - 412 portfolio roles + 119 in the COEs and EGI = 531, each counted once - with
a live control beside it that is the sum of three independent ledger counts.

### D96. The ledger's live filter is gone; one criteria-free filter spans the whole ledger.
REVIEW shipped with an active AutoFilter - two criteria hiding 519 of 531 rows - whose
range stopped at row 528 and column AH while the ledger runs to 534 and AX. One sort
through it would have reordered half the columns for most of the rows and left the ten
cost overrides and the newest roles where they were, with every control still zero.
Every filter and sortState in the file is swept; REVIEW keeps a single criteria-free
filter over A1:AX534 so a reader can still filter without ever sorting a subset.

### D97. 3.4 uses the owner's own country basis: AU is everything that is not NZ.
3.4 defined AU as literally "Australia" while his COE design tabs treat the one
Singapore role as AU, so the two published different AU figures for the same group. 3.4
now uses his basis, the "elsewhere" plug column - which made its own control a
tautology - is gone, and the control is a real AU + NZ = total check.

### D98. Customer splits its platform overhead half per country; Z Retail prices two platforms.
His 1.2 C7/D7 were byte-identical, so both countries counted the full 0.495 and the tab
overstated by $495k, flowing to 0.2, 3.1 and the Exec overhead line. They now follow his
own row 6 - half to AU, half to NZ. On 1.10 the QA agent read the unpriced overhead cell
as a missing row; his note on the squad row - "No Overhead required" - says otherwise,
so the dead reference came out of D7 and no overhead was added. His note outranks the
diagnosis.

### D99. His review tables live right of the working columns, declared and movable.
The role-review scratch tables on 1.4/1.5/1.6 move verbatim to columns S+ when the
Actual/Variance pair lands beside the squad table, so his working notes never sit inside
a header bar again. design2707 declares every cell it touches in a manifest, as post2707
does, so the shipped-workbook diff distinguishes a declared edit from an accident.

### D94. The platform-overhead allowance counts the platforms the model actually prices.
3.1 said the portfolios' overhead allowance was 5.335; 3.2 said 6.325 for the same roles.
The gap was exactly six platforms' worth of allowance (0.165 x 6): Lists carried a typed
"30 platforms" from an old generation while the ten design tabs carry 24. The count is now
measured off the design tabs at build time, so 3.2's allowance table ties to 3.1 to the
cent and a new plain-English line shows the slice allowed for people outside the
portfolios. Nothing on the actual-cost side moved.

### D95. One subtotal rule, one label pattern, controls beside their labels - all fourteen 2.x tabs.
A section gets a subtotal only when it has two or more rows and shares the tab with
another live section, which retires the 2.11-2.13 rows that duplicated the grand total,
and every subtotal is named for its section ("Squads total", not a bare "Total"). The
grand total sums the sections' data rows directly, so dropping a subtotal can never drop
rows from it. The two control values sit in column C beside their labels. The 2.x family
check in qa.py turned out to be dead - it compared fourteen identical tuples - and is now
live, catching exactly these defects on the old candidate and zero on the rebuild.

### D93. Ledger typos are fixed in the ledger, not papered over downstream.
"Project Manger", "Portfolio Mnager", "michelle Siegman", "EnterpriseProcess Analyst",
"DeveloperSAP ECC", "australia" and friends flow from the source datasets through INDEX
into every tab a GM reads. Corrected once in merge_review at load; nothing joins on those
strings (verified by grep before the change). Same treatment for two note typos
("Siginificant", "acorss") and the blank AR1/AT1 headers over live ledger columns.

### D74. The archetype prices overhead, and the working tabs were ignoring it.
Owner: "taking energy solutions & b2b as an example, we have 7.9m as the total cost for the
archetype, however, the actual total cost for the archetype is 9.03? ... you cannot only
include archetype squads without including the overhead, that paints a false narrative. you
also have to include the FTE for each of those overheads per portfolio."

He is right. `0.2 Data Config` allows every portfolio 0.7975 of overhead - half a Head of
Technology, half a Business Partner, half a Domain Architect, 0.3 of the Leadership layer -
and 0.165 for every platform it runs, being 0.3 of a Delivery Manager and 0.3 of a Technology
Manager. The 1.x tab has been adding both into its Total Cost all along. The 2.x tabs were
showing the squads only, so every portfolio's archetype was short by its own overhead
allowance and 1.8 read 9.03 where 2.8 read 7.90.

Each overhead line now carries the allowance it draws in that portfolio and the FTE behind
it. All ten working tabs tie to their own design tab exactly, and a live check on 4.0 proves
it row by row so it cannot drift again.

Two things fell out of building it:

- **A line nobody fills still draws its allowance.** Finance has no Delivery Manager, and
  listing only the lines with people in them handed that allowance to the "sitting outside"
  row. The three portfolio-drawn lines are now always listed - Finance shows a 0.084
  allowance against nil spend, which is the point of an allowance.
- **The allowance is not all drawn where the people are.** Of the 1.1275 the archetype gives
  Energy Solutions & B2B, only 0.4675 covers roles that sit in the portfolio; the Business
  Partner, Domain Architect and Leadership allowance is for people in the COEs and above the
  ledger. Netting them made the overhead subtotal read (0.01) - bang on plan - when the
  portfolio's own three lines are 0.65 over. So the subtotal covers the three lines only and
  the remaining 0.66 sits on its own line below it, named for where those people are.

### D75. 3.1's overhead allowance came off Lists; Lists and the design tabs disagree.
Lists prices the per-platform overhead lines over 30 platforms. The ten design tabs carry 21.
So 3.1 read an allowance of 6.325 while the ten tabs under it added to 4.84 - the same fact,
1.485 apart, on tabs that sit next to each other.

3.1 now reads the working tabs. Which number is right is the owner's call: 21 is what the
design tabs describe today, 30 may be a planned state. It is on the open list, and a live
check ties 3.1 to the tabs so the two can no longer drift apart silently.

### D72. The 2.x tabs carry one grand total, not five.
Owner, having looked at 2.2: "i do not see 83 total roles in customer??? why are my totals
not adding up... we should have a grand total and not incremental totals on 2.x tabs."

The totals were right. 2.2 read 83 at Total portfolio, the control read 0, and the fourteen
tabs added to exactly 525. That was not the point. The tab carried five total rows and a
filler section, the one total that mattered was at the bottom of the stack, and a reader who
cannot find it in a workbook about headcount is looking at a broken model. That is a layout
failure, and it is mine.

Three changes:

- a section with nothing in it is off the tab entirely. It used to print a label, a row
  reading "None - every group here has a figure to compare" and a subtotal of dashes: three
  rows saying nothing, on most tabs twice over;
- a section with one row gets no subtotal. "Directly funded programmes and platforms total,
  10" directly under "EGI Customer, 10" is the same number written twice;
- Total portfolio is the grand total and carries the comparison.

2.2 went from five total rows to three, and 2.9 lost a whole filler section. Every control is
still on every tab and every one still reads 0.

### D73. The totals carry the comparison. Owner's call, taken twice.
Owner: "why would we not include the archetype total vs the actual total of the portfolio.
like how are we counting it in the design tabs and not in here? im confused by the backward
logic?"

The reason it was a dash was real: the archetype does not price the overhead lines or a
programme with no funded figure, so a total variance is not a like-for-like variance. I
raised that, he asked again, and it is his model.

So every total row now states the archetype where there is one, the actual, and the
difference - on the 2.x Total portfolio row and on 3.1's ledger and grand total alike. 2.2
reads 12.06 against 17.13, 5.07 over. 3.1's ledger row reads 77.25 against 115.28, 38.03
over.

What makes it safe rather than misleading is that the composition is on the same page. Every
step the archetype does not reach is a named line directly above with a dash in that column -
the COEs, the unfunded programmes, Leadership, the overhead - so the 38.03 can be read off as
what it is rather than mistaken for overspend. The like-for-like line is still there and still
says so: "Everything with a figure to compare", 77.25 against 82.19.

### D69. A figure that restates the actual is not a comparison.
Owner, on the 2.x tabs: "we don't need the archetype cost of the strategic progs because they
are just exactly the actual cost, that's all we need to put in there."

He was right, and it was worse than a redundant column. The build forced `acost = actual` for
every EGI programme and for every COE squad. That did three things:

- printed rows whose archetype column was the actual written twice, with a variance of nil;
- **threw away a real figure**: EGI Retail has 1.52 typed on 1.1, and the column showed the
  actual 1.22 instead, hiding a 0.30 underspend the owner had set up himself;
- padded every total above it. 3.1's "everything with a figure to compare" read 108.74
  against 113.98 with 27.77 of COE and 4.01 of EGI on both sides of it, moving nothing.

One rule now, everywhere: a figure goes in the archetype column only where something prices
that cost independently of what it cost. Otherwise a dash. So a directly funded programme
compares against the amount typed on its 1.x tab, and where that is blank or zero - five of
the six EGI rows - there is nothing to compare it to and the column says so. The comparable
subtotal is 77.25 against 82.19, and every dollar in it is on both sides.

### D70. The overhead lines stay on the 2.x tabs. Asked, and answered.
Owner: "what is the purpose of having separate overhead columns? is there a reason? ... it
seems imbalanced if the overhead is not included in the archetype total and is only included
in the actual total."

They stay, because they are 43 real people costing $11.65m. The 2.x tabs are the layer that
reconciles to the ledger; take those lines off and the portfolio total stops being what the
portfolio costs, and the control at the foot of the tab fails by design. They are a separate
block because no archetype prices them: an archetype sizes a squad, and a Head of Technology
is not in a squad.

The imbalance he saw was real, and it was mine. `SUM` over a block of dashes is 0, so the
overhead subtotal printed **0.00** in the archetype column - reading as "an archetype priced
these 43 people at nothing" - beside a real $1.12m of actual. Same on the two other blocks.
A subtotal is now only a total where every row under it carries a figure; one row short and
the archetype side is smaller than the actual side by construction, and it reads "-".

Nothing on any 2.x tab now puts an archetype that covers part of a block against an actual
that covers all of it. The like-for-like comparison is on "squads priced by an archetype",
where it belongs.

### D71. A check that hardcodes a row number is the bug it is meant to catch.
Splitting 3.1's directly funded step moved its subtotal, and two checks on 4.0 broke with
`#VALUE!` because they read `'3.1 Cost Bridge'!$D$27` by row number. Both are now built from
the same anchors the tab is built from. One of them was also comparing three separately
rounded variances against one difference rounded once, and reporting the $1 residual as a
failure; both sides are now built off the same two cells.

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
2. The 8 GMs at $5.1m, a typed constant on Lists with no build-up and no presence on 0.2.
   A CFO asking "how was 5.1 derived" cannot be answered from the file. Moving it to 0.2
   as a visible input is one nod away, but 0.2 is his tab.
3. Whether to build the bridge tab that walks archetype cost to actual cost line by line.
4. Whether to restructure 1.4 so the funding block sits at the same row on every 1.x tab.
5. Whether FTE prices. Column O (FTE) never enters a cost formula: seven part-time roles
   (totalling 1.7 FTE against 7 headcount) are charged at their full base. If the base
   figures are full-time salaries, cost is overstated ~$0.36m; if they are already
   pro-rated, nothing is wrong. Only Lee knows which his source data is.
6. ~~The two Customer squads with no archetype row~~ - settled by D117: both rows are on
   1.2 and priced. Still his: the seeded Type/Size/Support on each row (cream, listed in
   D117) if the seeds are not what he would pick.

(Frozen panes came off this list: item 94 and his direct instruction settled it - none,
anywhere. The old entry here claiming they were "applied consistently" described a state
the workbook has never shipped in.)

---

## D122 - Overheads, the lights on budget, and the live register (5 August)

Nothing was built into the workbook in this round. Everything below is analysis, rulings
taken, and a register that now tracks the rest. The shipped file is unchanged since bfff80b.

### His rulings this round, verbatim where it matters

- "TDD pays for every single overhead role, let's be absolutely crystal clear on that,
  nothing is recharged." Overheads are never rechargeable. The Option B design that
  recharged overhead in proportion to squad support is dead.
- "squads don't pay for portfolio overheads, we price overheads at the platform and
  portfolio level." Delivery Manager and Technology Manager attach to the platform they
  run. Head of Technology, Business Partner, Domain Architect and the GMs attach to the
  portfolio. No per-FTE spreading into squads.
- "domain architects and business partners needs to be equally shared across portfolios."
  At actual cost, not the archetype allowance.
- "EGI does not use overheads. if they're egi they're egi, it's simple." EGI squads carry
  no overhead and sit outside the allocation base.
- "we can't price roles in which we have stated are hold." Roles on the Hold lever come
  out of the overhead pot.
- "it's not to projects it's a rechargeable amount." The word is rechargeable, everywhere.
- Per-FTE overhead rates and overhead-as-a-percentage-of-squad-cost are both banned as
  meaningless.
- No standalone 1.15 EGI tab - it double counts. Each portfolio's EGI platform belongs on
  that portfolio's own 1.x tab.
- Neil Reilly's $666,088 Business Partner cost is confirmed correct.

### The finding that matters

The lights on budget was sized on the archetype's overhead allowance of $12.366m
(including $3.0m for the GMs). The overhead that actually exists costs $24.627m across
75 roles. The gap is $12.261m.

Charging every overhead role to the lights on budget takes the position from $6.557m over
to roughly $18.7m over. Per role: Technology Manager +4.99 (22 slots priced, 25 people),
Head of Technology +3.85 (5 priced, 16 people), GMs +2.10, Delivery Manager +0.93,
Domain Architect +0.26, Business Partner +0.11, Program Management 0.00.

The archetype has the right total and the wrong shape: against the 528-role mapping it
over-prices squads by $10.66m and under-prices overhead by $10.16m, which cancels to the
$0.50m under that 3.1 reports. That is why this has never surfaced.

Budget context: 0.2 allocates $50.5m (portfolios 39.0, COEs 10.0, TDD Cyber 1.5) against
a full TDD budget of $53.8m, so $3.3m is unallocated.

### Data corrections found and not yet applied

- Head of Technology carries 16 roles; his own list has 15. The extra is Ed Tacey
  (AI Enablement Lead), on the line because of the earlier "follow the data" ruling.
  With him 5.2297, without 4.9089.
- Technology Manager carries 25; his list has 24. The extra is Shane Ker, whose title in
  the source reads "Technology Manger - Finance Platfoms" (two typos), which is why he
  drops out of a text filter. With him 6.7767, without 6.4531.
- Delivery Manager carries 10; his list has 12. The two missing are the vacant Delivery
  Assurance Manager and Delivery Excellence Manager, both $275,810, both currently coded
  as Squad roles inside COE BP&T. Correcting them takes the line to $3.3325m.
- Delivery Leads are correctly inside squads already - 44 of them, all Overhead line
  "Squad". No change needed.
- Exactly one overhead role sits on Hold: the vacant Enterprise Architect on the Domain
  Architect line, $275,810. Excluding it the Domain Architect pot is $1.3889m, so the
  equal share is $0.1389m per portfolio rather than $0.1665m. No other overhead line has
  a Hold anywhere.

### Total portfolio cost including overheads

Portfolios $95.836m (squads 61.399 + funded outside 12.274 + own overhead 13.360 +
shared overhead 8.802). Plus COEs 17.151 after charging Business Partners and Domain
Architects out, EGI 4.943 and TDD Cyber 2.138. Total TDD $120.068m, which ties to the
528-role mapping at 114.968 plus the 8 GMs at 5.100.

### The register

`docs/TDD_Model_Register.xlsx` is now the live tracker: 20 decisions outstanding,
22 items agreed and not built, 16 done and verified against the shipped file. It is
generated by `scripts/v10/build_register.py` from `docs/register_data.json`, so updating
it is a data edit and a re-run, not a rebuild. It was assembled by mining all 247 of his
messages, verifying 35 workbook facts cell by cell, and running three adversarial critics
over the result.

---

## D123 - The Lights On build shipped (5 August)

Green lit by him with a 15 point list; all 15 delivered. Pipeline scripts in
scripts/v10/lightson/ (v1 levers, v2 mapping, v3 lights on, v4 fixes, v5 role id,
v6 protection, v7 gate). Gate: 12 groups, 113 checks, green twice on the shipped file.

- His 11 lever edits ingested person keyed. His upload was a copy of the OLD 30/07
  with edits: different ledger (524 rows), no uplift column. A row-number diff showed
  73 false changes; the person keyed diff against the original 30/07 found the real 11
  (2.1 Ampol Retail 8, 2.10 Z Retail 5 incl the vacant Head of Technology marked to
  fill and Emma Natoli set to Hire cost neutral, 2.5 P&C 2). Uplift values untouched.
- Mapping corrections: Ed Tacey off the Heads line (block row moved to a Leadership
  group on 2.2; HoT 15 / 4.9089; his 0.32 now sits in Customer's rechargeable side);
  Delivery Assurance and Delivery Excellence Managers onto the Delivery Manager line
  (12 roles; line reads 3.2416 after Vanessa Allen's FTE scaling); Viren Khatri to
  TDD Group Functions (2.4 EGI TDD 5 roles / 1.3245; 1.1 EGI line now EGI Retail
  1.2214); seven part time people scaled by FTE (0.3646 reduction, names logged).
- 3.5 TDD Lights On: his fifteen columns verbatim, all live, per portfolio Other
  overheads toggle (cream, 0 to 100 in 5s), COE and EGI and TDD Cyber rows in the
  same flat list, analysis block live (why the archetype is far off, vacant
  overheads on Hire 19 / 5.36 with the already levered four called out, the dials,
  archetype 115.40 against 104.59 after levers).
- Numbers at default toggles: K total 60.93 against 50.50 (over 10.43) and 53.80
  (over 7.13); support 36.63; overheads charged 24.30; total cost incl GMs 109.69.
  Dials: hold vacant overheads 6.20, GMs above lights on 5.10, CRSO inside its
  2.00 is 4.28, unallocated budget 3.30. All three cost dials pulled: 45.35, under.
- COE netting at actual (2.12 -2.3135, 2.13 -1.3889 with the Hold DA at zero),
  3.1 netted to match, 0.2 TDD Cyber on the accurate basis 0.8384.
- 1.3 carries the EGI Ent Data platform line (1.8119 live from 2.3); 1.4's corporate
  pool line relabelled Significant Items EGI reading 2.4 live.
- Role ID R0001..R0528; 2,683 references re keyed; the shuffle test passed: the
  whole mapping randomly reordered, every control 0, every total identical.
- Protection: password Tdd123, role mapping fully locked, only levers, uplift
  cells, cream inputs and the toggles open; structure locked.
- Paused equals Hold, confirmed at zero everywhere. Register: 42 items done.

## D124 - 05/08: his master mapping is the file of record, Lights On on his eighteen columns (shipped)

- His 05/08 updates file IS the REVIEW tab: 526 rows, 29 columns, cell for
  cell, his three #N/A texts kept verbatim, and no hidden rows - the previous
  lock hid 513, his file hides none, and the gate now checks visibility
  against his file permanently.
- Every downstream block re-homed off the new master: homing by Portfolio 430,
  Division fallback 86, person-keyed override 10; 6,251 formulas repointed;
  Role IDs R0001..R0526 in his row order; each person on exactly one tab.
- Lever carry from the 30/07 model: 375 people kept their levers, 122 vacancy
  levers carried by title on their tab (his typed Hold and Offshore preserved),
  5 vacancies filled by name, 24 new people and 5 new vacancies defaulted,
  14 departed. His three Filled-vacancy edits kept.
- EGI is EGI: the whole family (six squads, 41 roles, 10.237) sits on 2.14,
  the portfolio tabs' EGI rows read 0, and 3.4 lists the six squads live.
- The cyber uplift charge is netted out of 2.11 once and charged to 1.14 once
  (0.4948); the four uplift percentages follow their people.
- 3.5 TDD Lights On rebuilt on his eighteen columns verbatim; rows are the
  0.2 Data Config set. The COE pairs carry squad truth (Strategy Architecture
  3.34, Transformation 2.88, Business Partnering 3.29, Data 1.43), price their
  charge off their planned spend lines and note the pot they hold (BP 2.31,
  DA 1.39), so Still left to fund reads 0 on every COE line. Ampol Customer
  and Z Customer split the tab with the shared slice and the overhead pool
  divided on their support costs. TDD Cyber carries an overhead share - the
  BP, DA and GM pots divide by eleven. EGI is excluded from the overhead
  engine and its row closes to zero.
- 3.6 TDD Lights On AU NZ: the same rows with AU spend, NZ spend, total and
  variance against the 0.2 AU and NZ budgets, the split live off Country.
- The numbers at default toggles: Total People cost 105.278 (= 104.783 after
  levers + 0.495 uplift charge); charged to TDD 61.043 against 50.50 allocated
  (over 10.543) and the 53.80 budget (over 7.243); support 37.870; overheads
  charged 23.173; noted in the 1.x tabs 34.382; significant items funded
  17.236; whole of TDD including the GM layer 110.378. Dials: hold the 17
  vacant overheads on Hire 4.686, the GM layer 5.100, COE Cyber above its
  allocation 3.928, the unallocated slice of the 53.80 3.300.
- Customer PCM annotations carried person keyed (82 of 83; one belongs to a
  person no longer in his file).
- Protection per his 05/08 ruling: the 0.x and 3.x tabs only, password Tdd123,
  toggles and cream inputs open, everything else unlocked, structure locked.
- Gate: 14 groups, 150 sub-checks, green twice with identical output. Deep
  verify: 50 checks green - raw identity, PCM cross-check, placement against
  the derivation, EE numbers, and a 40-person salary trace to his file.

## D125 - 05/08 review: EGI back in the portfolios, GMs across the COEs, the arithmetic made obvious, controls out

- EGI is keyed on his Platform column, not squad names: 49 roles / 12.07,
  eight of them (1.84) inside Enterprise Data squads that the squad-name rule
  had walked straight past. Each EGI squad now sits in the portfolio it names,
  so a portfolio shows its EGI cost in Total people cost and nets it out in
  Sig items funded; only the plain EGI squad, 18 roles / 5.15, stays on the
  EGI row. The EGI Ent Data grid row is built on 2.3 from his own 1.3 label,
  which is the platform line he asked for on 30/07 and which read 0 until now.
- The GM pot divides over sixteen units - the eleven portfolio units plus the
  five COE lines, 0.32 each. Business Partner and Domain Architect stay over
  eleven. Total GM charged is unchanged at 5.10. A COE line therefore shows
  its GM share in Still left to fund, because it carries a share it does not
  employ and has no business to recharge it to; the tab says so.
- A live block on 3.5 and 3.6 shows how the columns add up: 105.28 less 19.07
  funded outside equals 86.21 for TDD to fund, less 59.39 charged to lights on
  equals 26.82 to recharge, less 34.38 noted in the 1.x tabs equals (7.57)
  still to fund. A second block splits the 59.39 into support 36.22, BP 2.31,
  DA 1.39, GM 5.10 and other overheads after the toggles 14.37, against the
  53.80 budget, 5.59 over.
- 56 control rows and the 4.0 Data QA tab removed. Every reconciliation they
  proved is now checked outside the workbook by the gate.
- 0.2 Data Config and 3.5 were pricing the same portfolios differently - 4.43
  apart across 12 of 18 rows, five flipping sign - because 0.2 priced support
  at the archetype rate while 3.5 priced it at actual after levers. 0.2 now
  reads 3.5 row for row.
- One sign convention: over budget reads positive everywhere. It had been
  negative on 0.2 and the 2.x funding blocks and positive on 3.5, 3.6 and the
  1.x tabs, so a bracket meant good on one tab and bad on another.
- The eleven 1.x tabs made genuinely like for like: 1.7's block aligned with
  its siblings and a formula corrected that read the NZ budget where the others
  read the NZ variance; 1.5 given the NZ column its budget needs; the budget
  and funding blocks on one shape; one label over the block and one on the
  total.
- 3.6 now answers its own question: AU 44.96 against 33.00 is 11.96 over, NZ
  16.08 against 17.50 is 1.42 under. The duplicated variance column is gone.
- Ten further fixes on 3.5 and 3.6 from a cold read of the tabs: both over or
  under rows say which budget they measure, the four dials are labelled as
  alternatives rather than a stack, the gap analysis is all on the after-lever
  basis and its parts add to the whole, the vacant overhead count is one
  number everywhere, and both tabs carry the same nine reading notes.
- Two defects found and fixed during the build: the EGI funding line was being
  counted in both Sig items funded and Amount noted in the 1.x tabs, because
  the derived EGI grid row was never named in the funded-squads table; and the
  mapping stage's copy-through guard had let a rebuild run against its own
  previous output. Both now carry permanent checks.
- Gate: 15 groups, 189 checks. Deep verify: 50 checks green.

## D126 - 06/08: plain English forever, the cyber split approved, build order agreed

- Plain English forever: every question to Lee, every label, every note. A
  question that cannot be asked plainly gets rethought before it is asked.
  Decisions and plans are saved to the decisions log, the register and the
  session memory in the same turn they are made - context must never be lost.
- The cyber split, approved: 2.11 Cyber Risk & Service Ops becomes two COEs.
  Cyber takes the squads Cyber GRC, Cyber Risk, Cyber Sec Ops and Cyber Strat
  & Tech with a 1.5 allocation; Service Ops takes Service Op & Assurance with
  0.5. Who goes where derives from the Squad column his file already carries -
  his raw role mapping block stays untouched. The four uplift part-charge
  roles follow their squads to Cyber. The COEs carry no GM share, so the GM
  base stays 11 when the split lands (corrected 09/08 - an earlier version of
  this note said 17, superseded by D128). AU/NZ ruled in D127.
- Build order agreed: first the support percentages and budget tables move to
  the 2.x tabs (the 1.x tabs stay, fed by formulas; the archetype tables stay
  on 1.x), then the cyber split lands on the new wiring, then FY26.
- FY26: Finance actuals to June (July if it exists), split into charged to
  lights on vs recharged to the business, cut AU and NZ; the remaining months
  priced at the model's after-lever run rate divided by 12 times the months
  left, with per-vacancy hire-month overrides. All of it cream typed inputs.
  Month boundaries still to come from Lee.

## D127 - 06/08: the cyber split budgets
- Cyber COE: 1.0 AU and 0.5 NZ. Service Ops COE: 0.5 AU and 0 NZ. Together
  they equal the old line's 1.5 AU / 0.5 NZ, so the 0.2 total does not move.

## D128 - 06/08: three rulings on the audit questions
- The three Customer Programme Management people whose Platform reads EGI
  Customer (R0072, R0141, R0275) ARE EGI funded. The funded flag widens from
  Platform equals EGI to Platform beginning with EGI, so they and any future
  EGI-labelled platform land as funded outside, attract no support and no
  overhead share, and net out in Sig items funded.
- The GM pot divides by ELEVEN, permanently. The sixteen-line basis is dead:
  the formulas are right, the notes claiming a 16-line basis are stale and
  get fixed, and the COE lines carry no GM share.
- The nine offshore Enterprise Data delivery roles carry the Offshore lever
  as their true state, but their cost does not change with the levers - they
  are already priced at the offshore rate, so they take the same exemption
  WiPro roles have (before equals after, no 0.4 applied), keyed on their
  squad so future roles in it behave the same.

## D129 - 09/08: how prompts for the in-model tool are written
- Every prompt is written in first person, as if Lee drafted it himself -
  never referring to Lee in the third person inside a prompt. Super detailed
  and fully self-contained (the tool has no context between runs): exact
  cells, expected numbers, self-checks, report-back, and a hard rule to fix
  nothing beyond the brief - anything else that looks broken gets reported,
  not fixed. Each build prompt ships with a screenshot mock of its output on
  real numbers.
- Open item found while detailing the wiring move: Digital Support NZ is an
  archetype-only squad on 1.2 (typed 100% support, 0.32 planned cost, zero
  roles in the ledger), so it has no 2.2 grid row for its percentage to move
  to. It stays typed on 1.2 until Lee rules where it should live.

## D130 - 09/08: FY26 forecast v1 rejected as fabricated; corrected shape
- What was wrong with v1: it allocated Finance's leader P&L lines to model
  portfolios Finance never cuts them by (invented actuals for squads that do
  not exist), and it took budget from the FY27 model's allocation table
  instead of Finance's own lights-on budgets. v1 is withdrawn.
- The facts from Finance's file: lights on is reported at AU and NZ level
  only. AU (TDD Corporate): FY26 budget 78.90, actuals to June 40.97, their
  full-year forecast 79.31. NZ (Z-Energy): budget 52.23, actuals to June
  26.46, forecast 53.47. Nothing gives a portfolio cut of actuals, so the
  forecast never pre-fills one.
- The shape Lee specified: Portfolio | Actuals to June, charged to lights on
  ($m) | July to December forecast, one column per month. Actuals cells are
  cream typed and sit empty until real numbers are typed. Each month is the
  portfolio's FY charged to lights on divided by 12, formula-driven off the
  model. Gross people cost shown against charged to lights on, the charge
  always formula. Lee is sending an example; the rebuild waits for it.
- Formatting forever: no italics anywhere in any model, no subheaders, no AI
  look or feel, no made-up formulas to force a fit.

## D131 - 09/08: what the full May Finance file gives us (review finding)
- The May file (TDD_Finance_Report May2026 v3.1) is Finance's full working
  model; the June file Lee had was a two-tab summary print of it. It books
  people cost in the same shape as the cost model: gross internal labour,
  minus labour recharged out, equals Staff Cost - the people net cost - by
  month, actual vs budget vs forecast, AU and NZ separately, and by their GM
  lines (Digital, CISO, Strategy & Architecture, Data, Operations & Partner,
  Enterprise Tech, EGM).
- The numbers (AUD): AU people net YTD May 16.17 actual vs 15.15 budget;
  FY26 forecast 35.17 vs budget 35.23. NZ people net YTD (their actuals run
  to April in this file) 5.25 vs 4.78; FY26 forecast 19.25 vs budget 21.75.
  Combined FY26 people net: forecast 54.42 vs budget 56.98. Gross internal
  labour FY26 budget 84.00, forecast 77.42; labour recharge out FY26 budget
  (30.00), forecast (23.57).
- The story: YTD people net is 1.50 OVER budget, but gross labour is UNDER
  budget - the miss is recharge under-recovery (8.64 recovered vs 12.25
  planned). And the model's FY27 charge of 54.54 sits within 0.1 of
  Finance's FY26 people-net forecast of 54.42.
- Still no per-portfolio actuals in Finance's cut (their GM lines are not
  the model's portfolios), but their cost-centre level (Sheet1, the raw GL
  tabs, the cost-centre mapping tabs) carries lines like IT Business Systems
  - Retail, IT Above the Store, IT Security Uplift - Cyber. With a mapping
  Lee signs off, per-portfolio people actuals could be derived legitimately.
  Caveats: NZ actuals lag AU by a month in this file; Staff Cost includes
  people on initiatives (small); NZ books no labour recharge, so NZ net
  equals gross.

## D132 - 09/08: FY26 forecast v2 - the plan (Lee: come up with something; no build yet)
- Spine of the workbook: people net cost, because both files agree on it -
  Finance's Staff Cost line and the model's charged to lights on are the
  same idea, and their FY26 forecast (54.42) sits within 0.10 of the
  model's planned charge (54.54).
- Budgets: Finance's people-net budgets, AU 35.23 and NZ 21.75 (AUD, their
  tabs). Their full lights-on budgets (78.90 / 52.23) appear only as
  context, never as the track.
- Actuals: AU to June 20.01 real; NZ their 10.22 NZD June number at their
  own 0.92 rate = 9.40. No portfolio cut of actuals - cream typed cells,
  empty until real numbers exist (or until Lee signs off a cost-centre
  mapping; their raw tabs could then derive one legitimately - stage two,
  his call).
- Portfolio table: Lee's exact columns plus Charges in (AU/NZ/Both) and FY
  gross vs FY charged. Months = FY charge/12 minus vacancies not yet
  started (hire-month dropdowns, default July, option Not in FY26).
  Both-country rows split on the fixed 3.6. Four tabs total: Forecast,
  Vacancy hire months, Data - model, Data - Finance. Values snapshots with
  source cells named; zero external links; two check lines; no controls
  sprawl; no italics; no subheaders.
- Build waits for: the three model prompts to run (snapshot after them, 19
  lines), Lee's nod on this plan, and ideally Finance's June/July full
  file. Mock of the main tab sent 09/08 on real numbers.

## D133 - 09/08: FY26 forecast built on Finance's cut (Lee's ruling)
- Lee ruled: the June report has everything needed; use Finance's cut, not
  the model's portfolio cut. TDD Pack (2) gives, per line, YTD actuals,
  monthly actuals to June, Finance's own July-December monthly forecast and
  full-year totals. AU (AUD): Internal Labor, Contingent Labour, Personnel
  Costs, Labour Recharge, Staff Cost = the people cost line. NZ (NZD):
  People Costs.
- Allocated budgets come from the model's 0.2 Data Config row 27: AU 36.2,
  NZ 17.6, total 53.8 - these are the budget lines, not Finance's own
  budget column.
- The numbers: AU lands 35.61 vs 36.20 = 0.59 under. NZ lands 20.90 NZD =
  19.23 AUD at their 0.92 = 1.63 over 17.60. Total 54.84 vs 53.80 = 1.04
  over. The model's planned charge 54.54 sits within 0.30 of the landing.
  Finance's Z line carries two full-year forecasts 0.59 apart (20.31
  printed vs 20.90 summing their months); the workbook uses their months
  and says so plainly.
- Built same day by an Opus agent to a cell-level spec: three tabs
  (Forecast, Data - Finance (June), Data - model), only two typed cells
  (the cream budgets), everything else formula or sourced value, no
  external links, no italics, no subheaders.

## D134 - 09/08: AUD analysis and the two-Junes ruling
- The total TDD lights-on budget DOES exist in AUD and splits cleanly. The
  June file's Presentation sheet is entirely AUD - proven because its nine
  segment rows add exactly to its own 86.66 total, and the TDD Pack row
  labelled Lights on (AUD) carries identical Z numbers. FY26 lights-on
  budgets: AU (TDD Corporate) 78.90, Z-Energy 52.23, together 131.12 AUD;
  the whole-technology lights-on budget including business segments is
  173.12. Actuals to June: AU 40.97, Z 26.46, together 67.43 AUD.
- Finance's NZD-to-AUD convention is 0.92, proven from their own numbers:
  Z lights on 28.75 NZD (21.66 TDD-Z EBITDA + 7.09 Z-BU IT and comms)
  converts to their printed 26.46 AUD at exactly 0.92. So Z people in AUD
  = the NZD line x 0.92: landing 19.23, budget 19.81.
- TRAP: the May file's ZTDD GM_AUD sheet is an unconverted duplicate of
  ZTDD GM_NZD - identical values on both. The 21.75 NZ people budget
  quoted as AUD in the earlier review came from that sheet and is not a
  reliable AUD figure; corrected to 21.53 NZD = 19.81 AUD at their rate.
- Two-Junes ruling: the June file carries two June actuals for the AU
  lines - closed actuals (column O basis, Staff Cost June 3.84) vs their
  forecast phasing (June 2.86, the basis of their printed 35.61 full-year
  number, their own variance column showing the 0.97 gap). RULED: closed
  actuals win; their printed forecast shows as a labelled comparison row.
  Corrected landings: AU 36.58 = 0.38 OVER the 36.20 (not 0.59 under as
  provisionally reported); NZ unchanged 19.23 AUD = 1.63 over 17.60;
  total 55.81 vs 53.80 = 2.02 over; Finance printed basis 54.84; model
  charge 54.54.

## D135 - 09/08: FY26 forecast v2 shipped
- TDD_FY26_Forecast.xlsx delivered: three tabs, Finance's cut, closed
  actuals basis, LibreOffice-recalculated so every formula carries its
  value, verified twice (agent + independent QA). Only two typed cells in
  the workbook: the cream allocated budgets 36.20 and 17.60. Zero external
  links, zero italics, print set up landscape fit-to-width on all tabs.
- Landings shipped: AU 36.58 vs 36.20 = 0.38 over; NZ 20.90 NZD = 19.23
  AUD vs 17.60 = 1.63 over; total 55.82 vs 53.80 = 2.02 over; Finance's
  printed basis 54.84 and the model's planned charge 54.54 shown as
  comparison rows; TDD lights on in AUD block: 131.12 budget / 67.43
  actuals / 132.78 their forecast.
- Note precision fix during QA: the NZD back-solve figure corrected 28.75
  to 28.76 so the proven-from-their-file sentence is exactly true.

## D136 - 09/08: Wave Q rulings - FX, the 50.5 anchor, vacancy timing, the exec deck
- Live FX for NZ, typed input, 0.83 today. The 0.92 planning rate is dead
  for our outputs. FY26 v2's NZ landing restates 19.23 to 17.35 AUD at
  0.83: NZ moves from 1.63 over to 0.25 inside 17.60, total people cost
  from 55.82 to 53.93.
- The executive story tracks to the 50.5 allocation: spend today 54.54 =
  4.04 over. Hire timing and levers are the path down.
- Vacancy hire months from the Vacancy Audit drive FY26 H2 and FY27; some
  vacancies delay, some never get hired; both years can look better. The
  audit feeds a prompt to the in-model tool to land hire months in the
  2.x tabs.
- Contingent and personnel forecasts move with the operating model:
  model-driven, Finance actuals only. Never "theirs" - one company.
- Wave Q deliverable: executive confidence deck in the FY27 Budget update
  template (first 5 slides = the standard), FY26 and FY27 to the 50.5,
  vacancy story 2-3 views, offshore story (a cost reduction at the 0.4
  planning rate, never called a saving), all levers, spans and layers.
  Charts native Excel-editable. Four analysis agents launched in parallel
  (model, vacancy audit, deck deconstruction, MBB research); build agents
  follow; Fable QAs everything in loops before Lee sees it.

## D137 - 09/08: vacancy timing - the mapping and the prompt
- The audit joins to the model by title + squad + portfolio (no Role ID, no
  cost in the audit). 107 vacancies mapped cleanly to model Role IDs and
  get planned hire months written by the in-model tool; the rest are
  ruling rows for Lee (filled-marked, no-date, clones, one corrupted row).
- The timing story on the mapped set: FY26 H2 vacancy cost falls from
  7.72 (everyone starts 1 July) to 4.35 on planned months - 3.36 of
  relief before any role is cancelled; 2027 starts cut a further 0.55
  from a timed FY27. Nothing on 3.5 changes - the timing block is a lens.
- The prompt also completes the vacancy infrastructure the analysis found
  broken: register 127 to 132 (the five cyber-split roles), COE Service
  Ops rows added to the 3.7 and 3.10 portfolio tables, Exec Summary lever
  chains extended to 2.16, the 3.1 fifteen-cell guard lifted to sixteen,
  and R0273 aligned to Filled on the audit's evidence.
- Analyses all landed: model fact base (charge 54.4290 vs 50.5 = 3.9290
  over; 540 roles; 13 defect flags incl the meaningless 117.5568), deck
  deconstruction (template = kicker+headline Graphik pattern, green
  palette, no native charts anywhere - all think-cell or PNG; live deck
  numbers stale throughout), MBB style guide distilled to scratchpad.

## D138 - 09/08: FY26 v3 shipped - mechanics complete, refresh-ready
- Built by agent, independently QA'd: live FX 0.83 and the operating-model
  adjustment are the only typed cells; everything else formula. Landings
  at 0.83: AU 36.34 vs 33.0 allocation = 3.34 over; NZ 19.38 NZD = 16.08
  AUD = 1.42 under 17.5; total 52.42 vs 50.5 = 1.92 over (and 1.38 under
  the full 53.8). The FY27 planned charge 54.43 sits alongside at 3.93
  over - the FY26 timing and FX story is the better picture Lee wanted.
- Agent judgment calls accepted in QA: Internal Labor carries the
  operating-model balance (a plugged recharge line would read fake);
  vacancy relief split AU/NZ from REVIEW country data (105 AU / 2 NZ,
  3.2851 / 0.0798), NZ share shown but not taken, said in plain words.
- Refresh points named on the Sources tab: model block on tab 3 and the
  hire-month table on tab 4 re-point to the updated model when Lee runs
  the vacancy prompt and returns the file. Exec deck build still running.

## D139 - 10/08: exec deck v1 shipped after QA fixes
- 15 slides on the template identity, ten native Excel-editable charts
  with embedded workbooks, bridge-led Minto storyline, two option slides.
- QA found and fixed four things before it shipped: the FY27 table now
  carries the two cyber COEs split (3.16/0.49 and 2.83/0.00, not a
  merged pair); the spans chart and scatter gained COE Cyber at 16.8;
  the vacancy label now names COE Service Ops alone for its four roles;
  and the FY26 slide was restated from the stale 0.92 basis (54.84,
  1.04 over) to the v3 live-FX basis (52.42 vs 50.5 = 1.92 over, timing
  relief inside, FY27 charge 54.43 alongside). Also fixed in QA: a
  duplicated callout, a missing scatter label, a rounding footnote.
- Numbers refresh on the updated model once the vacancy prompt has run;
  the FY26 slide and workbook share one basis by construction.

## D140 - 10/08: deck v1 rejected - template ruling, banned word, flip-flop ask
- Deck v1 was built on Template.pptx's look. Wrong: THE template is the
  FY27 Budget Update deck itself - Mark OT font, navy 18249C, red ED0C06,
  its greys and accents. Template.pptx contributes layout ideas only,
  re-skinned. Process now: 20 template screenshots to Lee first, he picks
  the ones he likes, then the rebuild - richer tables, marimekko, MBB
  output, none of the boring-graph look.
- Banned word: charge/charged. Lee's language only: cost to TDD lights
  on, what TDD pays, landing. Existing artefacts get relabelled at the
  next touch.
- New analysis ruled in: the flip-flop story - roles GMs put on Hold
  (zero in the model) that now carry hire months in the vacancy sheet.
  FTE and dollars, the cost that comes back if they hire as stated, by
  portfolio, 2026 vs 2027. Narrative slide, not yet in the numbers.
- Updated model (1347cab8) received with the vacancy prompt applied -
  analysis agent extracting timing block, FY26 hire-month view, the
  flip-flop set, and verifying nothing on 3.5 moved.

## D141 - 10/08: post-vacancy model verified; tool scope creep needs a ruling
- The hire months landed exactly as instructed: 107 months, 4 Filled, 21
  No date, register 132 of 132, exec counts corrected, R0273 filled.
- But the tool went beyond the brief: it wrote FY27 start-quarter codes
  on 12 roles, silently re-pricing them and moving the anchor from
  54.4290 to 54.1177 (over the 50.5: 3.9290 to 3.6177). Two of the
  twelve codes are wrong (March 2027 coded as Q2), the application is
  partial, and the 3.7 FY27-reduction line now double-counts the same
  roles. Ruling asked of Lee: revert the 12 cells (recommended - one
  clean FY27 number, timing stays the overlay) or complete the quartering
  consistently across all 2027 starters. Revert prompt delivered in chat.
- Flip-flop story finalised from the model: 22 held roles now carry hire
  months - 22 FTE, 4.4709 full cost, 3.8990 comes back into FY27 if
  hired as stated, only 0.1440 into FY26 H2, 2.5389 would land on TDD
  lights on. And the two COEs furthest over their lines gave no dates at
  all for their nine vacancies.

## D142 - 10/08: FY26 v3.1 shipped on the updated model, monthly view added
- Refreshed to the post-vacancy model as it stands (the 54.12 basis,
  pending the revert ruling; one note row on tab 3 makes the revert a
  four-value retype). Landings at live FX 0.83: AU 36.44 vs 33.0 = 3.44
  over; NZ 19.28 NZD = 16.00 AUD = 1.50 under 17.5; total 52.45 vs 50.5
  = 1.95 over, 1.35 inside the full 53.8.
- New tab 6: twelve months, actuals to June then model-driven forecast
  months with vacancies excluded until their hire month; running total
  ties to the landing; Actual/Forecast labelled per month. This series
  feeds the deck's FY26 trajectory chart.
- Language sweep done: no charge/charged anywhere; Labour Recharge kept
  deliberately - it is Finance's own line name and recharge is Lee's own
  word. QA independent pass: checks Yes, two cream inputs only, zero
  italics, zero external links, banned-word sweep clean.

## D143 - 10/08: deck v2 pass one shipped, built inside Lee's file
- 17 slides in FY27_Budget_Update_v2.pptx: slides 1, 2, 6, 7, 8 untouched;
  3, 4, 5, 9, 10, 11, 12 rebuilt on current numbers; five new slides
  (exec summary SCR, hire timing, marimekko, spans bubble, decisions).
  Masters, layouts, theme and media byte-identical to Lee's file; the old
  flattened waterfall PNG replaced by a native chart. Eight native charts
  with embedded workbooks; mekko and dumbbells drawn, as his own decks do.
  MBB anatomy per slide: sub-claims, zoning, callout circles, delta chips,
  heat tints, in-table bars, so-what strips, numbered footnotes, sources.
- Build agent corrected four stale brief numbers against the live model:
  vacancy plan spend 18.12 / lands 11.12; COE Service Ops 7 vacancies;
  overhead gaps TM 4.59 / HoT 3.53; levers take 16.63 of 115.46. QA
  verified independently: zero italics, zero banned words, fonts MarkOT
  family only on touched slides, spot numbers all present.
- Pass two pending: 6-7 insight slides from the mining agent (still
  running) to reach roughly 20 content slides for Lee's cull. Anchor
  cells listed for the pending revert ruling (54.12 today, 54.43 if
  reverted). Scratchpad collision noted between agents - future builds
  work in isolated directories.

## D144 - 10/08: the mining findings - the story, the floor, and new defects
- THE STORY: TDD exits FY26 running at 50.53 annualised - level with the
  allocation. The whole FY27 overrun is the hiring pipeline turning into
  full years (+2.74 from Jul-Nov 2026 starts, +0.85 from 2027 starts =
  54.12). Robust to the pending ruling. Hire months are the control.
- THE FLOOR: holding the 19 vacant overhead roles (-4.70) plus offshoring
  the last 21 onshore roles in offshore-marked squads (-0.91 on the line;
  disjoint sets, verified) takes the position to 48.5 - under the
  allocation on either ruling basis.
- The asymmetries: Z gives back 7 percent of what TDD funds, Ampol 49;
  78.7 percent of the recharge comes from 37 percent of the cost base;
  the flat 11-way pot split costs 88.9k per head at TDD Cyber and 6.4k
  at Ampol Customer; offshoring the last 21 returns 36 cents in the
  dollar to the line. Killed narratives: big salaries (top ten = 3.8
  percent), bloated squads (dead level with allowance), offshore-cheap
  (Singapore is the dearest location per head in the book).
- NEW DEFECTS from mining: ten Z Retail roles priced at 1.00 FX and one
  TDD GF role converted the wrong way (0.21 overstatement, why R0520
  false-reads as the fourth-dearest role); the 3.7 FY27-reduction line
  double-counts entirely on the current basis; stale headline text on
  3.1 and 3.6/3.8. Folded into the rulings list for Lee.
- Deck pass two launched: six insight slides plus a slide-10 enhancement,
  target ~19 content slides for the cull.

## D145 - 10/08: deck content ruling applied; v2 final shipped
- Lee's ruling: the deck is his enumerated spec only - where we landed,
  how we got there, levers pulled, offshore (squads/roles/cost/FTE),
  vacancies (cost and FTE), vacancy timing, ramping, waterfalls, plain
  English. Killed: the recharge-terms slide, the recovery-machinery
  slide, the flat-pot-split slide, and the word "floor" (banned along
  with "terms" as jargon). Of the mining set only the annualisation
  story and the two-moves-to-under-allocation landed - the rest stays
  analysis, not slides.
- Final deck: 22 slides, 17 content, 14 native charts, in Lee's file.
  New in the last pass: the descriptive offshore slide (82 roles at the
  offshore rate, 8.39 off the base - corrected from the tab's stale
  8.30 narrative), the two-waterfall levers slide closing exactly on
  98.83 and 48.50, the hiring ramp (107 dated starts, cumulative line),
  and FY26 month by month (landing 52.45 v 50.5 and 53.8). Budget
  answer locked: FY26 hits the full budget (1.35 inside); FY27 as
  planned is 0.32 over the full budget, exits FY26 at 50.53 run rate,
  and lands 48.50 with the two decisions TDD controls.
- QA: zero banned words on all content slides (incl floor), zero
  italics, MarkOT only, 14 embedded chart workbooks, spot numbers tie.
  Cull guidance in the report: timing appears four ways (5, 14, 18,
  21); short pack 13-3-9-18-20-21-22-17.

## D146 - 10/08 pm: "saving" banned; the compliance loop's corrections
- Lee's ruling, verbatim intent: nothing in any of this is a saving - it is
  a REDUCTION OF COST. The word is banned in every form (saving, savings,
  saves, saved) across the deck, the models, the FY26 workbook, prompts and
  chat. Say cost reduction / reduction of cost / takes cost out. The
  model's own 3.10 labels (Saving ($m), Saving achieved ($m), Average
  saving per role, Potential saving ($m)) carry the word - queued for the
  next in-model fix prompt alongside the NZ FX defect, the 3.7 FY27
  double-count and the stale 3.1/3.8 headlines. The FY26 workbook's one
  occurrence (1 FY26 Forecast!D25) fixed and recalculated same turn; the
  deck sweep is in the running fix round.
- Lee also ordered an agent to read ALL his instructions across the whole
  session - deployed against the full transcript, building the complete
  instruction register, diffed against the working ledger; gaps feed the
  fix round before anything ships.
- Compliance audit verdict (18-item ledger, 18 spot checks): FY26 workbook
  passes whole; deck carries 12 misses, headline three: the offshore
  pricing split printed 13.98/5.59 where the register sums 13.59/5.19 (67
  roles; columns must add to 14.61/6.21 with WiPro 0.44 and Enterprise
  Data 0.58 unchanged); the exit-run-rate slide's 50.53 story (27.06 less
  4.15 plus 2.36 = 25.27) is not derivable from any model cell and
  contradicts the FY26 monthly slide; the Finance-table NZ column dropped
  the Delivery line (2.23 shown where O24+O25 = 2.40, total 26.94 not
  26.92). D145's "exits FY26 at 50.53 run rate" is superseded: the
  derivable exit story is December exit run rate 52.23 annualised
  (54.1177 less 1.8897 that FY27 pays for the 15 priced roles hiring in
  FY27; identical on the reverted basis, 54.4290 less 2.2010), FY26 H2
  23.96 annualising 47.91, FY27 plan 54.12. The 0.92 on the two-moves
  bridge verified exactly by the per-role walk (21 roles, five squads,
  support percentages: 0.2542 + 0.4024 + 0.0959 + 0.1052 + 0.0570 =
  0.9147); bridge end 48.50 stands.

## D147 - 10/08 pm: the FY26 landing corrected to 53.72; the full compliance loop closes; the package ships
- THE CORRECTION (found by my verification of the exit-run-rate slide,
  confirmed role by role, rebuilt by agent, recalculated and verified):
  the FY26 workbook's second half subtracted the GROSS hire-timing
  relief (3.1017) from a post-support-percentage half-year. What TDD
  pays for a squad vacancy is its cost times its squad's support
  percentage, so the relief is 1.9895 (July-start halves 4.3933 less
  planned-months cost 2.4038), with EGI-funded roles at zero and NZ at
  live FX. The half-year base is everyone-at-a-full-year 54.4290, not
  the 54.1177 anchor - the 0.3113 of FY27 start-quarter pricing (12
  roles, all proven AU-side, add-back exact) never belongs in FY26.
  Corrected landings: AU 37.6785 / NZ 16.0373 AUD (19.3220 NZD) /
  total 53.7157 - 3.2157 over the 50.5 allocation, 0.0843 inside the
  53.8 budget. H2 25.2250 annualising 50.4499; December exit run rate
  53.2722; monthly Jul-Dec 3.8265/3.9391/4.2209/4.3598/4.4394/4.4394.
  FY26 is independent of the pending start-quarter ruling. Supersedes
  the 52.4479 landing (D138-era) and D146's 52.23 exit-rate note (that
  number mixed the current anchor into FY26; the derivable exit is
  53.27). The deck's old 50.53/25.27/4.15/2.36 story and my earlier
  23.96 correction were BOTH wrong; the workbook and deck now carry one
  role-level basis and reconcile to the cent.
- The full-transcript instruction audit (Lee's order): 107 instructions
  extracted from 279 direct + 35 queued messages + 10 compaction
  summaries; the 18-item ledger was faithful for the deck window; six
  gaps closed in the fix round (saving ban, theirs-for-Finance sweep,
  em/en dash + model-era vocabulary sweeps, options cross-check,
  monthly-chart-equals-workbook check, explicit file re-attachment).
  One workbook violation found and fixed: a freeze pane on tab 4.
- Deck fix round: 16 fixes + addendum all applied and verified.
  Headlines: offshore pricing split corrected to 13.59/5.19 (columns
  now add to 14.61/6.21); exit-run-rate slide rebuilt on the derivable
  story; NZ Finance column 2.40 with total 26.94; "re-charged" wording
  on the two retained diagram slides; exec summary carries the full
  both-years budget answer including 0.32 over 53.8 (3.5 O25); agenda
  slide rewritten to this deck in MarkOT; FTE added from 3.7 col H
  (132 roles / 131.5 FTE, one half role R0526); flip-flop shares
  recomputed with the 0.85 balance named; four rounding footnotes; the
  saving family swept to zero everywhere. Formatting round before it:
  slides normalised to the slides 2-4 standard (titles 18pt - his 24pt
  cannot hold a 70-108 character action title), reordered to the story,
  exec summary at position 2, five charts fixed that PowerPoint would
  reject (dLblPos), masters/media proven content-identical to his file
  (the one flagged master diff is XML quote style from the serializer;
  image19.png left with his replaced slide 12).
- QA gates: deck - 22 slides, 14 native charts with embedded workbooks,
  zero banned words, zero dead numbers, zero italics, MarkOT-only
  rendered runs, all spot numbers present. Workbook - ten checks pass,
  two cream inputs only, no freeze panes, values verified to 4dp.
  Deck FY26 numbers equal the workbook's exactly.
- Open with Lee: the start-quarter revert ruling (54.12 v 54.43 - only
  the FY27 end moves); the eleven vacancy rows; the next in-model fix
  prompt (NZ FX on the 10 Z Retail roles + R0525, the 3.7 FY27
  double-count, 3.7 r201/r202 basis inconsistency found this round,
  stale 3.1/3.8 headlines, 3.10 Saving labels); July actuals if they
  exist; Digital Support NZ percentage home; the INTERNAL USE ONLY
  italic on his own master (left untouched).

## D148 - 10/08 night: his template file is THE vessel; the excel months must ramp
- Lee uploaded template.pptx and ruled all slides live inside it. The MBB
  grammar (chips, kickers, so-what strips) is fine and stays - the defect
  was only that the deck was never built inside his template package. The
  22 slides were transplanted byte-identical into the template (his title
  slide leads; 14 native charts and workbooks carried; validator clean bar
  the source's own comment ids), then reordered to the story.
- FY26 excel corrections after his review: tab 1's forecast months were
  flat sixths - rebuilt so Internal Labor (gross people cost) and Staff
  Cost (the lights on component) ramp month by month off the model and the
  vacancy hire months (AU 4.28 to 4.88 gross, 2.57 to 3.17 lights on; NZ
  ramps in NZD); the essays stripped to tight labels; house style matched
  to the model (Calibri, header rows white on 0F2E52). Landings unchanged:
  53.72 / +3.22 v 50.5 / 0.08 inside 53.8; all checks pass.

## D149 - 11/08: the polished deck - his cull, his storyline, the FY26 narrative on a slide
- Lee's rulings on the transplanted deck: delete the title slide, the
  agenda, both model diagrams and the divider (his numbered list), and
  the spans slide is not wanted ("that's not mekko"). The deck opens on
  the executive summary. The FY26 gross-and-lights-on narrative built
  in the excel must live on a slide. The executive summary leads FY26
  then FY27. Titles at the template's 24pt black. Storyline: exec
  summary, FY26 gross and lights on, FY26 month by month, exit run
  rate, FY26 v the FY27 model, FY27 cost by line, mekko, allocation
  bridge, overheads, levers and two moves, vacancies, hire timing,
  ramp, flip-flop, offshore, New Zealand, decisions - 17 slides.
- FY26 excel tab 1 rebuilt earlier tonight to his drawn shape: two
  tables (Gross and TDD Lights on), NZ and AU rows, monthly columns
  plus landing, budget and over/(under) allocation. Gross: NZ 18.62,
  AU 66.45 v 62.59, total 85.07. Lights on: NZ 16.04 v 17.6 budget
  (1.46 under its 17.5 share), AU 37.68 v 36.2 (4.68 over its 33.0
  share), total 53.72 v 53.8 (3.22 over the 50.5 allocation). Months
  are formulas off the vacancy table; checks pass.

## D150 - 11/08: Lee pulls the deck build; a self-contained prompt for another tool
- After repeated failed deck deliveries (files refusing to open on his
  machine, then a rebuilt deck he judged well below standard), Lee ruled:
  stop building, go through the whole conversation, and hand him one
  self-contained first-person prompt for another excel-capable tool to
  produce the slides. Deck structure he fixed: executive summary plus
  four slides each on FY26, FY27, vacancies and offshore; no NZ slide.
  All verified numbers frozen this session: FY26 lands 52.48 (+1.98 v
  50.5, -1.32 v 53.8), gross 83.46, months 3.62 to 4.23, H2 build
  27.21 - 5.63 + 2.40 = 23.99, December exit 50.80, FY27 54.12, two
  moves to 48.50, vacancies 132/131.5 FTE with holds at zero and the 17
  undated at January 2027, offshore 82 roles 14.61 to 6.21.

## D151 - 11/08: NZ priced at 0.92 for FY26; the landing moves over the full budget
- Lee's ruling: New Zealand is priced at 0.92 for FY26, superseding the
  live-rate ruling for this year. The rate stays a typed input; every NZ
  figure (actuals, forecast months, gross budget) converts at it.
- At 0.92 the FY26 story changes: lights on lands 54.20, 3.70 over the
  50.5 allocation and 0.40 OVER the 53.8 full budget (was 1.32 inside at
  0.83). NZ lands 17.61 v its 17.5 share; AU 36.59 unchanged. Gross
  lands 85.47 v the 82.37 budgets (AU 62.59 AUD + NZ 21.5 NZD at 0.92).
  December exit run rate 52.42. Reported straight, not massaged.
- Delivered as TDD_FY26_Forecast_11.xlsx on his own v10 upload lineage;
  checks rewritten to rate-independent identities, all pass.

## D152 - 11/08: the example file is the chart standard; deck cut to the presented slides only
- Lee's correction on the 48-slide example file: it is his example of the
  graphs and charts standard, NOT a formatting override and not a deck to
  ship whole. He wants only the slides that are presented.
- Presented set identified and cut: slides 13 to 28 of the example = title
  (August 2026), exec summary, FY27 by line, allocation v cost, overheads,
  FY26 month by month, FY26 exit run rate, vacancies, holds with hire
  dates, hiring ramp, hire timing, offshore, levers, New Zealand, spans,
  decisions. 16 slides. Slides 1 to 12 (earlier working set), 29 to 33
  (diagram and exec variants) and 34 to 48 (the Graphik chart library)
  dropped. Delivered as TDD_cost_update_presented.pptx.
- Residual stale numbers found in QA and fixed against the FY26 workbook
  before shipping: December exit 53.27 corrected to 52.42 everywhere; AU
  over its share 4.68 corrected to 3.59 and the total over 3.22 to 3.70;
  "just inside the full budget" corrected to 0.40 over; the H2 walk
  corrected to 28.04 less 5.65 plus 2.41 = 24.79 (annualises 49.59, 0.91
  under the allocation); hire cohorts corrected to 70 in FY26 H2 (October
  18, not 21) and 32 in FY27; the what-TDD-pays walk to 5.65, 2.41 and
  3.24; "at the live rate" reworded to "at the 0.92 rate" per D151.
- The placeholder title on the FY27 by-line slide ("Header (only use this
  formatting)") replaced with a real action title at the blessed 18pt.
  The empty exec summary slide drafted: five numbered statements on the
  ruled numbers (54.20/3.70/0.40; 54.12/3.62/0.32 with overheads 22.90 v
  11.61 and the cyber 3.49; levers 16.63 and the two moves to 48.50;
  vacancies 27.65/18.12 with 70 of 102 starts in H2 and the 52.42 exit;
  the six open decisions led by 3.49 and 4.47). Lee to strike anything
  that does not belong.

## D153 - 11/08: HIS FORMAT ruled by screenshot; deck restyled in the template vessel
- The presented-slides cut (D152) was rejected: wrong format and wrong
  vessel. Lee's screenshot defines HIS FORMAT: 24pt MarkOT title with the
  key phrase highlighted yellow, navy-header table, red numbered circle
  callouts tied to footnotes, an Insights panel in navy with white text,
  Draft for discussion chip, Ampol mark bottom-left and page number
  top-right from his template master. The kicker lines, units notes,
  sub-claim pairs and SO WHAT strips are NOT his format and were removed;
  each slide's so-what line moved into a navy Insights strip in the same
  slot. In-body chips and red circles stay per his earlier words.
- Vessel = FY26_Actuals_H2_Forecast_Platform_Model.pptx, whose master is
  canonically identical to his template.pptx; the example-file package is
  not the template and is retired as a vessel. Title slide added at the
  front replicating the template's own.
- Every number swept to the ruled basis in the same pass (54.20/3.70/
  0.40; gross 85.47 v 82.37; NZ 17.61 at 0.92; exit 52.42; walk 28.04
  less 5.65 plus 2.41 = 24.79 annualising 49.59; 70 of 102 starts in H2,
  32 in 2027; timing relief 3.24 on what TDD pays) and the monthly, ramp
  and hires charts re-pointed to the same series. QA clean on the stale
  list and banned words. Delivered under the D148 deck name.

## D154 - 11/08: new model (8fe38c44), updated gross budgets, six net-new slides
- New model upload 8fe38c44: register now 127 open vacancies (R0273,
  R0055, R0067, R0326 confirmed filled and removed); plan = 66 hire
  (14.80) + 31 offshore (2.11) + 30 delayed (zero) = 16.91 planned
  spend, 26.44 full price, 9.53 avoided; 10.10 on lights on (17 vacant
  overhead roles 4.08 + 110 squad roles 6.03). Hires 97: 70 in FY26 H2
  (cohorts unchanged 1/14/30/18/7), 27 across FY27 (12 Jan). Cyber
  roles now timed to start quarters. July-start H2 8.45, planned 4.35,
  timing takes out 4.10. Two-moves close becomes 49.12 (54.12 less
  4.08 less 0.92). 3.5/3.6 totals unchanged.
- Updated gross staff budgets (Lee): AU 64.2 AUD, NZ 25.5 NZD at 0.92
  = 23.46, total 87.66. FY26 workbook updated (K10 formula 25.5*C6,
  K11 64.2): gross lands 85.47 = 2.19 UNDER (AU 0.79 over, NZ 2.98
  under); on the old 82.37 line it read 3.10 over. Lights on unchanged
  54.20 / 3.70 / 0.40.
- The gross story ruled onto the slides: FY26 gross counts everyone
  including sig-item and EGI-funded people; FY27 gross 79.00 nets the
  19.83 out (98.83 less 19.83). Same organisation, two nets; both
  years hold the updated gross budgets, the pressure is lights on.
- Six net-new slides in his format (tables left, navy Insights and
  assumptions rail right, yellow-highlight titles): FY27 run rate (his
  mock exactly: AU/NZ/TDD Total, his column labels including "charged
  to TDD" as his verbatim headers), FY26 gross landing, vacancy stock,
  vacancy timing + defer scenarios (deferring FY26 starts does not
  move FY27; deferring the 12 January starts to July frees ~0.8
  gross), offshore squads (9 marked, 46/64 filled, 21-role redundancy
  consideration 2.52, Digital Support NZ = archetype capacity only, no
  coded roles), offshore lever roles (82 = 51 filled + 31 vacant; 52
  outside marked squads, cyber COEs carry 17). Three old vacancy-era
  slides fully superseded were dropped; register deltas swept through
  the rest. Deck 21 slides, QA clean.

## D155 - 11/08: deck cut to cover + 8; filled means named; empty markings move nothing
- Lee's rulings this round: the deck is 8 slides max (his six plus two);
  "filled" means a named person in the role (the Enterprise Data
  Delivery nine are NOT filled, they are capacity); an offshore marking
  on a squad with no coded roles (Digital Support NZ) is a label that
  moves no number and must be coded or dropped; the 0.43 difference on
  the hire line is start-quarter timing on the new cyber roles, never
  "avoided"; the held-roles slide claiming 22 was stale.
- Register facts (8fe38c44 3.7): Hire 66 / Offshore 31 / Hold 30 = 127.
  The 31 is the offshore count; holds are 30, all Australian-side.
  REVIEW named-person check: control squads carry MyHR names, the nine
  marked squads carry none, so the model's 46 "filled" there are
  modelled capacity pending HR confirmation of names.
- FY27 run-rate side split (derived, flagged on-slide): people cost AU
  78.55 / NZ 20.28 (role countries less side lever effects; NZ lever
  roles carry 0.77), sig items 19.07 / 0.76, gross 59.47 / 19.52 v
  64.20 / 23.46 budgets; overheads and support split on lights-on
  share (16.46/6.43 and 22.45/8.77).
- FY26 composition analysis (new slide, his ask): sig items 19.83 for
  the year (1.65 a month, 9.92 per half, H1 on a flagged assumption),
  vacancy cost 4.35 all in H2 ramping 0.02 to 1.10 by cohort, standing
  organisation flat 5.26 a month; H1 banked 4.17 against half the
  87.66 budget, H2 runs 1.98 over, net 2.19 under. December exits
  96.13 gross annualised, 76.30 net of sig items, marching to 79.00.
- Deck rebuilt as cover + 8 in his format (yellow-highlight titles,
  navy tables, red circles, insights-and-assumptions rail right, clean
  chart with no gridlines and labelled bars). QA clean.

## D156 - 11/08: the 0.43 is immaterial; the cyber roles just go into FY27
- Lee: the 430k hire-line difference does not matter; just put those
  roles into FY27. The vacancy stock slide now shows the hire line as
  "Hire, into FY27 at full price" with no difference column, no callout
  and no start-quarter storyline; the only trace is one footnote clause
  ("the new cyber roles sit in FY27 from their start quarters"). The
  avoided story stays on the offshore and delayed lines (9.10: 5.89
  delayed, 3.21 offshore). Model totals untouched: 127 / 26.44 / 16.91.

## D157 - 11/08: gross view includes sig items; the 540-role growth story analysed
- Lee's revised run-rate view: Gross cost = total people cost including
  sig-item people (98.83) against the 87.66 staff budget; 11.17 over
  gross before program recovery, 79.00 and 8.66 inside after. Table
  with AU/NZ filled delivered in chat (people 78.55/20.28, sig
  19.07/0.76, support 22.45/8.77, overheads 16.46/6.43, cost charged
  38.91/15.21, delta 2.71/(2.39)).
- Growth story analysis (all from the model): 540 roles 537.8 FTE (his
  503 = 540 less 30 delayed, 6 WiPro, 1 unmatched); full onshore price
  115.46 charged at 98.83 via levers (delayed 5.89, offshore-rate 67
  roles 8.39 with WiPro/EntData excluded, balance 2.34); 19.83 funded
  by programs (52 EGI roles 12.83 + line-funded 7.00) v the original
  model which had no funded-outside carve-out; vacancies priced ~4%
  under the filled average (0.208 v 0.218); 14 end-dated filled roles
  worth 3.40 (7 exit FY26, 7 FY27) still inside 98.83 until they
  leave; 44 contractors/temps at 11.42 (40 day-rated).
- NOT verifiable from held files (flagged to Lee): the 420-role and
  ~100m anchors and any contractor-to-perm conversion count; the
  original ledger predates Role IDs. Needs the original extract.
- Offshore sufficiency: 82 on the lever = 15% of FTE, 67 discounted,
  8.39 achieved = 7% of the base; next mechanical step is the 21
  onshore roles in marked squads (2.52); beyond that a marking call.
- Insights list: his four polished plus four new (levers, 53.8 v 50.5,
  the two gross readings, vacancies inside the number), in chat.
