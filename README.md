# Ampol Retail — TDD Cost Calculator

An interactive, portfolio‑level cost calculator for **Ampol Retail**. It builds cost
from the bottom up — **Squad → Platform → Portfolio** — splits every dollar into
**TDD‑funded** vs **Other**, and compares the implied TDD cost against the TDD
lights‑on budget to reveal the funding gap.

**File:** `Ampol_Retail_TDD_Cost_Calculator.xlsx`

---

## How it works

### 1. Squad level (your inputs)
For every squad you choose four things from dropdowns (yellow cells, blue text):

| Input | Options |
|---|---|
| **Squad Type** | Engineering · Configuration / Integration · Product · Operations · Enterprise Data and Insights · Build and Run |
| **Size** | XS · S · M · L |
| **On/Off** | Onshore · Offshore |
| **Support %** | 0% · 20% · 100% |

The engine then derives:

- **Total Squad Cost** — looked up from the **Squads** sheet by *Squad Type × Size*.
  Onshore uses the onshore rate; Offshore uses `onshore × 0.4` (offshore efficiency).
- **Cost TDD** = `Total Squad Cost × Support %`
- **Cost Other** = `Total Squad Cost × (1 − Support %)`

### 2. Platform level
Each platform rolls its squads up and adds a fixed overhead:

- **Platform Overhead (TDD‑funded)** = **$0.165m** — from *Data Config* (Delivery Manager + Tech Manager build‑up). Sits entirely in the TDD column.
- **Support Cost** = Σ of the squads’ *Cost TDD* and *Cost Other*.
- **Platform Total** = Overhead + Support, split TDD / Other.

Ampol Retail contains five platforms:

| Platform | Squads |
|---|---|
| Store Operations | POS · Payments · Retail Operations · Deployment |
| Merchandising / Supply Chain | Merchandising & Supply Chain |
| Pricing & WFM | Pricing & WFM |
| AmPOS | AmPOS |
| Network / QSR | Network & QSR |

### 3. Portfolio level
The summary at the top adds the portfolio overhead and totals everything:

- **Total Portfolio Overhead (TDD)** = **$0.6945m** — from *Data Config* (Head of Tech + Business Partner + Domain Architect + Leadership build‑up).
- **Total Platform Overhead (TDD)** = Σ of the five platform overheads.
- **Total Support Costs** = Σ of every squad’s TDD and Other.
- **TOTAL COST** = the three above, split into **Y (TDD total)** and **E (Other total)**.

### 4. Budget & variance
- **TDD Lights‑On Budget (Ampol Retail)** = **$2.5m** — pulled from *Data Config → TDD Budget Allocation*.
- **Variance** = `Budget − Total TDD Cost`. Negative ⇒ a shortfall that has to be funded.

With the illustrative default selections the model shows **TDD cost $4.20m vs. $2.5m
budget → a ($1.70m) shortfall**. Change the dropdowns and every roll‑up, the variance
and the verdict line update live.

---

## Sheets

| Sheet | Purpose |
|---|---|
| **Ampol Retail** | The calculator — summary, budget/variance, FY26 context, and the five platform blocks. |
| **Squads** | Cost library: onshore/offshore cost per *Squad Type × Size*. Drives the lookups. |
| **Data Config** | TDD budget allocation by area, and the role‑based build‑up of portfolio & platform overheads. |
| **FY26 Budget** | FY2026 technology budget by business segment (Lights On, Initiatives, Depreciation, Significant Items, CapEx), pulled from the presentation pack. |

---

## Assumptions & notes

- **Squad types** are the six *costed* types in the Squads library; **AI** is excluded per brief.
- Not every *Size* exists for every *Type* (e.g. only Configuration / Integration has an **XS**). An invalid combination returns $0 — pick a valid one.
- The default squad selections are **illustrative** so the model shows live numbers; overwrite them with the real shape of each squad.
- **Portfolio overhead is taken as $0.6945m** (the Data Config *Portfolio overhead subtotal*). If the intended figure is different, change it in *Data Config* and it flows through.
- All figures are **$m per annum, FY2026**.

## Colour key
- **Yellow fill / blue text** — your inputs (dropdowns)
- **Green text** — a value pulled from another sheet
- **Black text** — a formula
