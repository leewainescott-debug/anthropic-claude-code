# Non-labour mapping analysis, first pass (18/08)

Source: Lee's TDD_AU_Consolidated_2027_budget.xlsx (SAP BEx export) and the
FY26 Enterprise Technology Spend Dashboard PDF. Read only, nothing built.
This document records what the data is, what it supports, and the triage.

## What the workbook actually contains

| Tab | What it is | Size |
|---|---|---|
| SW Line Items | 12 months of software actuals, Jul-2025 to Jun-2026, SAP CO document level | 18,748 amount rows |
| HW Line Items | Same for hardware | 52,914 amount rows |
| 2026 Budget | TDD Corporate's own FY26 budget by cost centre x cost element, enriched with vendor, contract Y/N, expiry, detail, impact-of-not-doing, IT Service code. The cost optimisation work | 516 populated rows, 88.81m (44.67 labour + 44.14 non-labour) |
| Platform Model | Empty except one cell ("Old") | nothing usable |

Line columns: Cost Centre + name, Cost Element + name, CO document number,
user, vendor, purchase order, posting date, document text, 12 monthly amounts,
total. Portfolio and Platform columns exist on the SW tab only and are filled
on just 784 of 18,748 rows (4%, 1.41m): Commercial Fuels | Supply 1.38m,
B2B & Energy Solutions | B2B 0.02m, Commercial Fuels | Trading & Shipping
0.001m. HW has no mapping at all. The mapping work has barely started.

## Structural findings (all verified to the cent)

1. **The export contains a cost-centre hierarchy, parents and children both.**
   Naive summing double counts. Proven exactly:
   - APPLHCORP = BRNDCOMS + H_INFOTECH + HALT + HFIN + HLSA + HHR + GENERAL + SIG_ITEMS
   - APPLBULKFUEL(SW) = B2B + CCO + SUPPLYOPS + SUPFPO + TDDIST + FIFINANCE + INTEGM
   - APPLMARKETING = MTOTAL + RCALSTORES
   - APPLSIG_ITEMS(SW) = APOLLO + CTRMSIG + INTEMERSIG + CRPOSSIG + TDDEINTSIG + CUSTLOYSIG + HEMERALD + PCTECHKCM + PCTECHMYHR
   - In HW, BULKFUEL and SIG_ITEMS are leaves (no children exported).
2. **True totals at leaf level: SW 51.288m, HW 25.469m, together 76.756m**,
   about 49,900 leaf lines. These tie exactly to the export's own result rows
   (which reflect the six top parents: AETOTAL, BULKFUEL, FIREFINING, HCORP,
   LUBRICANTS, MARKETING).
3. One CO document can carry line items in several cost centres (e.g. one
   Fenwick invoice split across four). Line items are real, not duplicates.
4. The fiscal window is Jul-2025 to Jun-2026, a rolling 12 months of actuals,
   not the calendar FY26 the dashboard shows.

## The old model (from the dashboard PDF)

FY26 view, Jan-Jul YTD: portfolio pages Commercial Fuels 19.5 budget FY,
Retail 15.2, Infrastructure 4.4, Energy Solutions 5.2, Finance & Other 3.7,
P&C 4.1, Customer 1.7, TDD Corporate 87.7, Strategic Programs 47.6.
Ampol opex 189.1 + capex 48.1 = 237.3. Z Energy opex 64.8.
**237.3 + 64.8 = 302.1m. That is the ~302m.** (313.8 if Z capex 11.7 counts.)
TDD Corporate's cost centre groups are the old functional org (CISO, EGM
Digital & Data, Strategy/Arch/Data & AI, Tech CFE&I, Tech Customer, Tech Group
Functions, Tech Retail, Transformation & Partnering, Enterprise Tech &
Software Delivery). The Budget Detail page carries vendor and detail lookups
(Azure accrual 14.0, MS EA licences, SAP BAU, Intelligent Workplace, Retail
L1 Service Desk...).

## Triage at leaf level (76.756m, ~49,900 lines)

| Class | $m | share | lines | What it is |
|---|---|---|---|---|
| Direct, no questions | 18.6 | 24% | ~22,000 | Cost centre names one new portfolio: B2B 4.13, Ampol Energy 2.90+0.35, Refining 2.40+2.20, Supply 0.65+0.04, Lubricants 0.02+0.17, Distribution 1.02, Finance HFIN 2.68, P&C HHR 1.83, Legal 0.04, Calstores 0.17, F&I Finance 0.04, plus HW twins |
| Direct to programs (sig items) | 14.9 | 19% | ~450 | CTRM 1.36, Apollo 1.24, Emerald/EGI bits, P&C Tech (KCM, MyHR), plus HW SIG_ITEMS 11.98 which is the laptop fleet on monthly device billing. Recharged, same treatment as labour sig items |
| Rule-based split | 39.7 | 52% | ~22,400 | H_INFOTECH SW 25.31 (Azure, network, security, data platforms), MTOTAL 4.65 SW + 6.99 HW (store systems vs loyalty/martech), HW BULKFUEL 2.46, H_INFOTECH HW 0.34 |
| Decision / digging | 3.5 | 5% | ~5,100 | CCO 0.69, Brand & Comms 0.85+0.19, GENERAL 0.92 (HighRadius, OneStream: Finance systems), HALT, INTEGM, corp-function laptops HW 1.0 |

Classes re-add to 76.756m exactly.

Tell coverage inside the rule class: document text carries a recognisable
product or vendor on 66% of HCORP dollars, 83% of H_INFOTECH, 98% of HW
BULKFUEL (device serials), Salesforce alone appears on ~7,600 lines. Vendor
column is filled via purchase orders on only 17% of SW lines, 1.5% of HW, so
document text is the primary tell, vendor second, IT Service codes from the
2026 Budget tab third.

## How much moves between models

Old TDD Corporate held HCORP = 47.9m of the 76.8m (62%). The new model breaks
that block up. Movement answer: 24% keeps its old home (direct, name already
points there), 19% follows programs as today, 52% changes hands via rules
(nearly all of it out of old TDD Corporate into platforms and COEs), 5% needs
a ruling before it can land.

## Portfolios with no direct cost source (Lee's concern, legitimate)

Enterprise Data, Ampol Customer, EGI, Z, and every COE (Cyber above all: old
CISO budget 26.3m) receive cost ONLY through the rule-based splits, because no
cost centre names them. The cost exists, it sits inside HCORP/H_INFOTECH/
MTOTAL lines. Until the rules are agreed those portfolios show zero, which is
why the rules must come before any portfolio view.

## What is NOT in this workbook (needed for ~302)

- Network cost elements (Communications-Data, Mobile): only 7 SecurID rows
- Outside Services (31.1m B2026 enterprise level)
- Depreciation (5.8m), Other
- All of Z Energy / NZ (64.8m opex)
- Capex (48.1m budget)
- Labour (comes from the labour model, 98.83m gross AU+NZ)

## Approach (proposed, not started)

1. Freeze the leaf ledger: strip the 9 parent nodes, keep ~49,900 leaf lines,
   76.756m, tie to the export's result rows. Give every line an ID.
2. Mapping dictionary, three layers: (1) cost centre to portfolio for the 28
   leaf centres, lands 33.5m immediately; (2) product/vendor dictionary over
   document text for the shared centres (Salesforce, Azure, SAP, security
   products, POS, data platforms...), cross-checked against the 2026 Budget
   tab's IT Service codes and the dashboard's detail lookups; (3) named
   line-by-line digging for the remainder, quantified as it shrinks.
3. Decision list to Lee/GMs (short): EUC device policy (laptops ~13m across
   sig items + corp functions), Marketing split Retail vs Ampol Customer,
   CCO placement, Brand & Comms placement, Distribution to Infrastructure
   confirmation, GENERAL ownership, sig-items presentation.
4. Gates: portfolio x platform totals re-add to 76.756 with a zero check;
   every line keeps its ID; nothing transposed, raw tabs ride along verbatim.
5. Only then the enterprise bridge: add network, outside services, depn,
   capex, Z, and labour from the labour model, reconciling to ~302.
