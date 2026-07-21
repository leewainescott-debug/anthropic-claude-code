# TDD Cost Calculator - requirements brief (what the user has explicitly demanded)

Audience: CTO / General Managers at Ampol Group. Owner: Lee (strategic/financial modeller).
Workbook: TDD_Cost_Calc_v8.xlsx. All $ figures in AUD $m unless a cell is formatted as whole dollars.

## Core model
- Roll-up: Squad -> Platform -> Portfolio -> whole operating model.
- Cost split: TDD-funded (support %) vs Other/funded-outside-TDD (1 - support %); compared to TDD
  lights-on budgets per portfolio; must show "what is funded, then what is left to fund, then total cost".
- Squads chosen by dropdowns: Squad Type, Size (XS/S/M/L), Onshore/Offshore, Support % (0/20/100%).
- Offshore cost = 40% of onshore cost. This applies to squads AND to every COE role toggle.
- Platform overhead $0.165m per platform; portfolio overhead $0.7495m per portfolio.
- Strategic Programs (AmPOS, CTRM, EGI x5): hard-coded yellow cost, no size, always onshore,
  editable support %; their non-TDD share flows to Significant Items (EGI to "Significant Items EGI").
- 2.0 Group Summary must ladder to: allocations 43.5, total TDD people budget 53.8.
- The whole-org actual cost must tie to $120.038m (Added data full cost, excluding junk stub row 550).

## The user's latest (angry) feedback - every point must hold
1. Offshore toggle on role tabs must yield 40% of full cost, NOT zero.
2. Summary blocks at the TOP of tabs, not the bottom.
3. 1.11 TDD Cyber must be single-sourced from 2.5 Cyber Roles (no double count of cyber money).
4. FTE view must show platform costs and portfolio costs (with overheads), not just squads.
5. "Archetype cost 60.58" must be explained/contextualised - the FTE view grand total now includes
   platform + portfolio overheads so it is NOT presented as "total cost of everything".
6. SA&D tab must NOT include portfolio data-squad roles (data science, reporting & analytics,
   data platform etc. belong to Enterprise Data portfolio, not the COE).
7. BP (0.4 x 11 portfolios) and Domain Architect (11 portfolios) overheads: show HOW MANY TIMES
   applied, and do NOT double count them in totals (they are funded by portfolio overheads,
   not COE budget) - 2.1 Total Cost carries an explicit de-duplication row.
8. All tabs reorganised and renumbered; NO duplicate tabs; Exec Summary exists as the first tab
   and tells the GM story ("what is the point of this exercise, what is it telling me").
9. Formatting reviewed on every tab; no AI-looking notes or grey-italic explainers
   (the user's own rough notes/writing must be left alone).
10. Portfolio drill-down (the user likes it) retained on the Exec Summary.
11. Business Partnering & Transformation split: Commercial dept goes in Business Partnering,
    Transformation dept in Transformation. Vacant roles highlighted. Full cost in AUD visible.
12. Cyber: roles/vacant/filled/cost vs budget; cyber capex monitoring $0.5m visible.
13. Positives-not-negatives convention on "Total to fund / TDD Variance / Other Variance" blocks.

## Tab map (29 tabs, exact order)
Exec Summary | 0.0 Data Config | 0.1 Squads | 0.2 FY26 Budget | 0.3 For Presentation Pack (2) |
0.4 Budget Table (Fin) | 1.1 Ampol Retail ... 1.11 TDD Cyber | 2.0 Group Summary | 2.1 Total Cost |
2.2 COE | 2.3 BP&T | 2.4 SA&D | 2.5 Cyber Roles | 3.0 FTE View | 3.1 Data QA | squad mapping |
raw data | Added data | Lists

## Review ground rules for panel agents
- Judge against THIS brief, not personal preference. A finding is a DEFECT only if it violates the
  brief, is objectively wrong (bad formula, broken tie, mislabeled number, formatting bug,
  inconsistent convention), or would visibly confuse a GM.
- Suggestions beyond the brief are welcome but must be flagged severity=suggestion, not defect.
- Do NOT propose new tabs, charts, or scope. Do NOT propose changing the user's own rough notes.
- Cite tab + cell/range for every finding.
