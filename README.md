# TDD Cost Calc — portfolio cost calculators

**File:** `TDD_Cost_Calc.xlsx`

One calculator tab per portfolio, all following the same pattern: dropdown‑driven
squad selections roll up **Squad → Platform → Portfolio**, split into **TDD‑funded**
vs **Funded outside TDD**, compared against the TDD lights‑on budget, with a
finance funding block per portfolio reconciled to `0.4 Budget Table (Fin)`.

## Workbook map

| Tab | Purpose |
|---|---|
| `0.0 Data Config` | TDD budget allocation by portfolio; role build‑up of portfolio ($0.7495m) & platform ($0.165m) overheads |
| `0.1 Squads` | Cost lookup: `Squad Type \| Size` → onshore / offshore cost |
| `0.2 FY26 Budget` | FY26 budget by business segment |
| `0.3 For Presentation Pack (2)` | Finance presentation pack extract (values frozen — external links removed) |
| `0.4 Budget Table (Fin)` | Finance budget table incl. the people component of each portfolio's lights‑on |
| `1.1 Ampol Retail` … `1.11 TDD Cyber` | Portfolio calculators (see below) |
| `2.0 Group Summary` | All portfolios side by side + reconciliation to the Data Config allocation |
| `squad mapping` | Source mapping of portfolio → platform → squad → type/size |
| `Lists` (hidden) | Dropdown lists (named ranges `SquadTypes`, `SquadSizes`, `OnOff`, `SupportPct`) |

## Portfolio tabs

| Tab | Platforms (squads) | TDD budget source |
|---|---|---|
| 1.1 Ampol Retail | Store Operations (POS, Payments, Retail Operations, Deployment) · Merch/Supply Chain · Pricing & WFM · AmPOS · Network/QSR · Data AU | Data Config `E11` (2.5) |
| 1.2 Customer | Ampol Digital (App, Web, Digital Ops) · Customer Z (Z Energy Apps, Z Energy Martech) · Group Customer Platforms (Loyalty & Martech) | `E13+E14` (5.0) |
| 1.3 Enterprise Data | Group Data (Data Science, Reporting & Analytics, Data Platforms, Enterprise Data Delivery) | `E22` TDD Data (3.5) |
| 1.4 TDD Group Functions | Workplace & Ent Tooling, Network & Infra, DevOps & Eng, Integration | `E21` TDD (5.5) |
| 1.5 P&C | P&C, P&C – RTA | `E18` (2.0) |
| 1.6 Finance | AU Finance, NZ Finance | `E19` (2.0) |
| 1.7 Infrastructure | Distribution · Manufacturing (M&GP, Tech Support) · Data & Insights | `E17` (2.5) |
| 1.8 Energy Solutions & B2B | Energy Solutions (Energy, EVCI) · B2B | `E16` (2.5) |
| 1.9 Commercial Fuels | Trading & Shipping (T&S, T&S Data) · Supply · CTRM | `E15` (2.5) |
| 1.10 Z Retail | Z Supply · Z Customer (Site Systems, Z Retail Backend) | `E12` (2.5) |
| 1.11 TDD Cyber | TDD Cyber | `E23` (1.5) |

## How each tab works

Per squad (yellow cells = your dropdowns): **Squad Type × Size** looks up the cost in
`0.1 Squads` (Offshore = onshore × 0.4); **TDD Cost = Total × Support %**;
**Funded outside TDD = Total × (1 − Support %)**. Each platform adds a $0.165m
overhead; the portfolio adds $0.7495m. The funding block pulls that portfolio's
Lights‑On / Initiatives / CapEx / Sig Items / Depreciation from `0.4 Budget Table
(Fin)` and nets off the "amount that can be allocated to people" (editable) to show
**Left to fund**.

New tabs include a **Depreciation** line so `Total budget` reconciles to the Finance
total; `1.1` retains its original 4‑line block (Retail depreciation = 0).

## QA notes

- All 1,188 formulas verified — zero errors; every platform cross‑foots; `2.0`
  reconciles portfolio budgets (32.0) + COE allocations (10.0) = Data Config total (42.0).
- ±0.1 deltas between a portfolio's budget components and Finance's total on
  1.1 / 1.4 / 1.8 / 1.9 are roundings inside the Finance source table — shown, not hidden.
- Cell comments (marked "QA") flag every normalised squad type/size default and
  budget‑line approximation. Shorthand from `squad mapping` was normalised:
  Config → Configuration / Integration · Data → Enterprise Data and Insights ·
  Ops → Operations · Support & Maintain → Build and Run · Strat (CTRM) → Configuration / Integration.
- `0.3` previously linked to the original Cost_calculator workbook; those cells are
  frozen at their cached values so nothing breaks when the source file isn't present.
- ~18,600 broken legacy defined names removed (file 281KB → 86KB).

`scripts/` contains the generation/QA tooling.
