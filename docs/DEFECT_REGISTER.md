# TDD Cost Calculator — defect register

Every finding below was reproduced with the workbook open, against
`TDD_Cost_Calc_NEW.xlsx` (Lee's uploaded copy, 51 sheets). Numbers are quoted from the
file, not from prior claims. Cached values were checked against a full LibreOffice
recalculation, so nothing here is a stale-cache artefact.

## Ground truth established

| Fact | Value |
|---|---|
| Roles in `REVIEW - Complete Role Mapping` | 525 (rows 2–528, excluding stray row 191) |
| Total cost, column AA | **$115,113,262.27** = $115.113m |
| Vacant (name contains "vacant") | 135 |
| Filled | 390 |
| Formula errors in the whole workbook after full recalc | 3 (all one root cause) |

Lee's raw columns carry the real structure: **I = Portfolio, J = Platform, K = Squad**.
Columns AC–AQ are a previous session's derivations and were not trustworthy.

---

## 1. The ledger contains a stray subtotal that breaks the workbook's own control

`REVIEW!AA191` = `SUBTOTAL(9,AA108:AA190)` = **$17,126,867.14**, sitting inside the data
range with no name in column B. `SUBTOTAL` is skipped by other `SUBTOTAL`s but **not** by
`SUM`, so any un-criteria'd sum over the ledger double counts it:

* `SUM(AA2:AA530)` = **$132,240,129.41** — not $115.113m.
* `3.2 Total Cost!D25`, labelled *"Restatement vs the REVIEW ledger ($m) - must be 0"*,
  therefore reads **−17.127** and ships red on the face of the deliverable.

## 2. Column AP invented squad groupings that contradict Lee's own column K

| AP (MSquadC) said | n | Lee's column K says | n |
|---|---|---|---|
| Z Energy Martech | 16 | Z Energy Martech | **2** |
| AU CRM & Martech | 7 | AU CRM & Martech | **2** |
| Z Energy Apps | 5 | *(no such squad)* | — |
| — | — | Z Loyalty & Martech | **13** (dropped) |
| — | — | Ampol Loyalty & Martech | **7** (dropped) |
| — | — | Z App and Web | **5** (dropped) |

Lee's "Energy" squad (1 role) had no AP value at all, so `2.2 Customer` row 13 referenced
`AP="Energy"`, matched nothing, and displayed a phantom squad with zero cost.

## 3. Column AQ's leadership flag was wrong, and Lee's data already has the answer

`AQ` flagged 13 Customer roles as leadership including a **Business Analyst** (Christian
Lazano), an **Experience Designer** (Tim Corin), a **Quality Assurance** (Waheed Malek) and
an **Engineer Data** (Jawad Hassan). Meanwhile column K already contains a literal
`Leadership` squad.

Derived properly — `K = "Leadership"` OR `J = "Leadership"` — leadership is
**53 roles, $14.250m**. `3.1 Group Summary!C57` depends on `AQ` and reads $13.339m, so the
entire "Overhead coverage" block (C55–C58) rests on the fabricated column.

## 4. The 2.x levers could not work by construction

Three independent defects, all reproduced on `2.2 Customer`:

1. **Lever ranges pointed at the wrong rows.** Every squad row read a fixed range, e.g.
   `H6 = COUNTIF($E$26:$E$35,"Hire")` for *Ampol Web* — but the role list is sorted
   **alphabetically by name** with **no squad column at all**, so rows 26–35 are
   Aaron Lu…Anthony Gen. The ranges were arbitrary.
2. **The cost maths had the wrong sign.**
   `O6 = M6 + (SUMIFS(cost,lever,"Hire") + 0.4*SUMIFS(cost,lever,"Offshore"))/1e6`
   *adds* 40% of an offshored role **on top of** its full cost already inside `M6`. So
   offshoring a filled role made the portfolio **more** expensive, and `Hold` appeared in
   no formula at all, so it did nothing.
3. **Costs and statuses were hardcoded literals.** `G26:G108` were typed numbers and
   `D26:D91` were the literal text `"Filled "` (with a trailing space). Nothing followed
   REVIEW.

Consequences visible on the face of the tab: `E16` (total roles) was the hardcoded literal
`83`; `G16` and `I16` were hardcoded `13` and `10`; column **F** (variance) and column **P**
(cost after overhead) had Lee's headers and **no formulas**; `I13` showed **−1** vacancies.
`SUM(M6:M14)` came to $12.845m against a portfolio total of $17.127m, so **$4.28m was
unattributed** — the total tied while the build-up did not.

## 5. `3.3 FTE View` — one dead reference takes out the Exec Summary

`E38` and `F38` are `='1.3 Enterprise Data'!#REF!`. A squad row was deleted from `1.3` and
the reference died. `G38`'s `IFERROR` masks it as `"-"` in the cached value, but on recalc
it is `#VALUE!` and it propagates:

```
3.3!E38/F38  →  G38  →  G42  →  G43  →  G97  →  G4 and K97
     → Exec Summary C5, C36, C37, C38, C39
     → 3.2 Total Cost H7, H23
     → Lists K4
```

**17 cells error from this one cause**, including the Exec Summary headline.

Also on `3.3`:
* **111 of 525 roles are missing.** `J97` = 414. COE Cyber (46), COE BP&T (24),
  COE SA&D (24) and EGI (17) have no squad rows. `N97` = $87.841m against a ledger of
  $115.113m — **$27.3m absent**.
* **Portfolio totals disagree with the sum of their own platforms.** `J22` = 70 (a
  `COUNTIFS` on REVIEW) while `J11+J13+J15+J17+J19+J21` = 65.
* **Column O mixes counts and dollars, then sums them.** Most rows are a vacancy count,
  but `O29 = N29-M29 = -0.217` and `O95 = -9.898` are dollars. `O97` = 176.03 is meaningless.
* **Off-by-one wiring into 2.x.** Row 23 is *Ampol App* but `P23` reads
  `'2.2 Customer'!$O$6`, which is *Ampol Web*. Row 24 (*Ampol Web*) reads `O7`
  (*EGI Customer*). The two tabs list squads in different orders and are joined by fixed
  cell refs.
* **Rows 99–135 reference individual REVIEW rows absolutely** (`='REVIEW…'!$B$359`), so any
  row insert silently repoints all of them.

## 6. `3.2 Total Cost` — base case ties, the model does not respond

Good: `D23` = $115.113m, `N23` = 525, `M23` = 135, `I23` = 390 all tie to the ledger exactly.

Broken:
* **`F23` = $78.963m ("impact after vacancy") is *less* than `K23` = $86.579m (filled-only
  cost).** The cost after decisions cannot be below the cost of people who exist today.
  It is lower because the squad joins lose roles.
* **Column J is a duplicate of column M under a wrong name.** `J5 = N5-I5` = 22 = `M5`
  (vacant count), but the header says "FTE variance".
* **`H15:H20` are the text `"-"`**, so `H23` = 310.5 excludes every COE, yet it is presented
  as a variance against 525 roles.
* **`B22` is a note row with `0` in every numeric column, inside the `SUM(C5:C22)` range.**
* Each 2.x tab is read at a *different* cell (`2.1!M17`, `2.2!O16`, `2.3!M13`, `2.5!M9`) —
  the tabs have inconsistent layouts and are all hand-wired.

## 7. `3.1 Group Summary` — the columns Lee asked for are mostly empty

* **`G` ("Actual allocated for support") and `H` ("Variance to allocation") are blank for
  8 of 10 portfolios.** The two that are populated use *different* formulas:
  `H14 = C14-G14` but `H15 = C15-F15`.
* **`F17` has no total** — every portfolio's actual cost is listed but never summed.
* **`M5` is labelled "Total Cost" but holds the archetype cost**, not the actual.
* `C19` stitches the Cyber budget from two Data Config rows (`E7 + E23`). That single line,
  −$6.398m, is 65% of the −$9.823m group gap.
* `B47` and `B49` are stranded headers with nothing under them.

## 8. `0.2 Data Config` — a missing variance and two mixed bases

* **`G13` is empty.** Every other portfolio has a variance; Ampol Customer does not.
  `E13` = 2.5 but `F13` = 7.6495, because `F13` carries the whole `1.2 Customer` cost while
  the Customer budget is `E13 + E14` = 6.5. `E14` (Z Customer, $4.0m) has no spend or
  variance at all.
* **Column F, labelled "Spend", mixes bases.** `F11`–`F22` are archetype design costs from
  the 1.x tabs; `F6`/`F8`/`F9`/`F10` are actual COE costs; `F23` is the actual Cyber people
  cost ($9.898m against a $3.5m budget). That one row is **74% of the −$8.674m total gap**.
* `B20` Legal, `B24` EG and `B25` EGI all carry zero budget, so EGI's 17 roles and $4.943m
  are invisible in the allocation ladder.

## 9. Wiring bugs on the 1.x design tabs

The AU/NZ overhead split is sound — `IF(NZ>AU, 0, OH)` / `IF(NZ>AU, OH, 0)` puts the whole
overhead on one side, and `F` always totals $0.7975m. No overhead is lost. But:

| Tab | Cells | Defect |
|---|---|---|
| `1.14 TDD Cyber` | `C6`,`D6`,`C7`,`D7` | Reads Data Config **row 15 (Commercial Fuels)** instead of row 23. Cyber's overhead split is driven by another portfolio's budget. |
| `1.2 Customer` | `C7` | Sums `H34,H42,H49` — the overhead sits in column **I**, so the AU branch would return 0 if it ever went live. `D7` correctly uses `I`. |
| `1.2 Customer` | `C6`,`D6` | The only tab that splits the portfolio overhead 50/50. Both branches are the *same* `IF`, so if AU ever exceeded NZ, Customer would lose the whole $0.7975m. |
| `1.5 P&C` | `D6`,`D7`,`D9` | **No NZ column at all** (empty cells), despite a $1.0m NZ budget. |
| `1.10 Z Retail` | `C7` vs `D7` | Asymmetric ranges: `SUM(I27,I34)` vs `SUM(I27,I34,I40)`. |
| `1.3 Enterprise Data` | `C8`,`D8`,`E8`,`F9` | The **EGI Data** platform ($1.7m, row 36) is excluded from the portfolio summary. |
| `1.3 Enterprise Data` | `B38` | Labelled "Group Data Total" but it is the *EGI Data* total — copy-paste. |
| `1.14` vs `1.13` | — | Two Cyber tabs. `1.14`'s $1.2925m is referenced by nothing; the workbook uses `1.13`'s $9.898m. `1.14` is orphaned and shows a different answer. |

## 10. Overhead is charged twice — and this is arithmetic, not opinion

`0.2 Data Config` builds a $0.7975m portfolio overhead from four roles and charges it to
each of the 10 portfolio tabs = **$7.975m**. Every one of those roles is **already a named
person in the ledger**:

| Overhead line | Charged (×10) | Already in REVIEW |
|---|---|---|
| Head of Tech | $1.38m | 11 roles, $3.383m |
| Business Partner | $2.20m | COE BP&T, 24 roles, $6.401m |
| Domain Architect | $1.40m | 7 roles, $1.665m |
| Leadership Overhead | $3.00m | 53 roles, $14.250m |
| **Total** | **$7.975m** | **60 distinct roles, $15.915m** |

**Therefore overhead can only load the archetype (designed) side.** Adding it to actual
cost double counts 100% of it, because the actual already contains those people. No netting
lines are needed and the group total stays exactly $115.113m.

## 11. Design tabs and the ledger name the same squads differently

The archetype comparison only works where a squad name matches. Today
**138 roles ($31.695m — 28% of the organisation) sit in a squad with no archetype**, and
11 designed squads have nobody in them. Most are near-miss spellings:

| Ledger (col K) | Design tab (1.x) | Roles |
|---|---|---|
| Z Loyalty & Martech | Z Energy Martech | 13 |
| Cloud, Network & Infra Ops | Network & Infrastructure | 12 |
| Integration & Process Automation | Integration | 11 |
| Manufacturing Group Projects | Manufacturing & Group Projects | 9 |
| SAP ERP | AU Finance | 9 |
| DevOps & QE | DevOps & Engineering | 8 |
| Ampol Loyalty & Martech | AU CRM & Martech | 7 |
| Network & QSR | Network / QSR | 6 |
| P&C RTA | P&C - RTA | 6 |
| Z App and Web | Z Energy Apps | 5 |

`Leadership` (44 roles in portfolios) has no archetype by design — under the archetype
model it is covered by the portfolio overhead allowance, not a squad rate.

## 12. Data-quality items inside Lee's own columns

* **Squad typos creating false squads:** `AmPos`/`AmPOS` (9+1),
  `Manuacturing`/`Manufacturing Group Projects` (2+7),
  `Integration & Process automation`/`Automation` (9+2), `Technology Suport`,
  `Customer, AI`, `Data - AU`/`Data AU`, `Data Platform`/`Data Platforms`.
* **10 "ring fenced" roles ($2.34m)** — 4 SA&D architects, 2 NZ Finance JDE specialists,
  4 TDD Cloud architects. Not named "Vacant", so under Lee's rule they score as Filled.
* **70 roles carry `NA` as their squad** (COE Cyber 46, COE BP&T 24) — no squad split
  exists in the source for either COE.
* **29 roles have a blank squad**, 23 of them in COE SA&D.
* **1 role has no squad and no department:** Jens Tom, SAP Solution Architect, Finance.
  Left visible as `Unassigned` rather than guessed.
* **EGI is both a portfolio and a squad.** 17 roles in portfolio `EGI` plus 23 roles in
  `EGI Retail`/`EGI Customer`/`EGI TDD`/`EGI Finance`/`EGI P&C` inside other portfolios.

---

## Derivation now used (all from Lee's raw columns)

```
Portfolio   col I, normalised: Retail / RETAIL / Ampol Retail -> Ampol Retail
                              Ampol Customer -> Customer;  Z -> Z Retail
                              TDD -> TDD Group Functions
                              P&C, Finance & Legal -> P&C          [confirm]
Squad       col K, typo-merged. Where col K is NA/blank, col G (Department) stands in —
            it foots exactly for all three COEs.
Status      col B: Vacant iff the name contains "vacant".
Leadership  col K = "Leadership" OR col J = "Leadership"  ->  53 roles, $14.250m
Cost        col AA, untouched.
Lever       Filled 1.0x · Hire 1.0x · Hold 0.0x · Offshore 0.4x
```

Result: **525 of 525 roles on exactly one 2.x tab. Zero orphans, zero duplicates.
$115,113,262.27 to the cent.**

## Lever proof — `2.2 Customer`, Aaron Lu (a *filled* role), full cost $235,334

| Lever | Role effective $ | Squad new $m | Squad roles | Portfolio new $m | Portfolio roles | Full cost of model $m |
|---|---|---|---|---|---|---|
| Filled | 235,334 | 1.5912 | 8 | 15.4747 | 70 | 17.1269 |
| Hire | 235,334 | 1.5912 | 8 | 15.4747 | 70 | 17.1269 |
| Hold | 0 | 1.3559 | 7 | 15.2394 | 69 | 17.1269 |
| Offshore | 94,134 | 1.4500 | 8 | 15.3335 | 70 | 17.1269 |

`235,334 × 0.4 = 94,134`. Squad falls by exactly the role's cost on Hold and by 60% of it
on Offshore; the portfolio total falls by the same amount; role counts move on Hold and
correctly do **not** move on Offshore; and "Full cost of model" stays fixed because it is
the un-levered ledger cost.

## Not yet done — why this build is not shipped

Rebuilding the 2.x tabs moved their rows, which invalidates **126 references** that reach
into them from elsewhere: `3.3 FTE View` (90), `4.0 Data QA` (14), `Exec Summary` (11),
`3.2 Total Cost` (11). Shipping the file in this state would be worse than the original, so
it has not been shipped. Those 126 references must be rewired onto named anchors before the
workbook is usable end to end.
