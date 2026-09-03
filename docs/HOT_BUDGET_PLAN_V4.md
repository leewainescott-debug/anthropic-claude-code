# Portfolio cost bases, version 4: the plan (03/09, after Lee's review of version 3)

Source of truth: his latest mapping file TDD_Consolidated_0725_0626_Actuals.xlsx
(scratchpad hotbudget/actuals_0725_0626.xlsx, read only). Transactions July 2025
to June 2026, carried unchanged to FY27. His instructions I29 to I38; findings in
scratchpad hotbudget/v4/review_changes.md and review_reconcile.md.

## What version 4 is
1. Read me: ten lines. Basis, the tab list, where his three typed inputs live.
2. TDD corporate: his 0725_0626 Summary grid rebuilt by formula (Portfolio by
   Platform; Software, Hardware, Network; BU paid and TDD paid) with a labour
   column beside it and the totals: total cost base, labour, non-labour, paid
   by TDD, paid by BU.
3. One tab per portfolio, fourteen, in the order of his Summary tab (Ampol
   Customer, Ampol Retail, Commercial Fuels, B2B and Energy Solutions, Fuel
   Infrastructure, Finance, P&C, TDD, Enterprise Data, Cyber, COE Business
   Partnering, COE Strategy Architecture and Enterprise Data, COE Cyber Risk
   and Service Operations, TDDLT). Each tab: a band (total cost base, labour,
   non-labour, paid by TDD, paid by BU); left, the platform table (rows = his
   platforms plus Not yet mapped; columns = labour, non-labour paid by TDD,
   non-labour paid by BU, total; a total row); right, the top 25 applications
   or vendors (application, vendor, platform, paid by TDD, paid by BU, total,
   share, cumulative share; an everything else row; a total row); two charts,
   platform totals split TDD and BU, and TDD against BU. Every name in column
   B by formula from the Data tab. One line of text above each table.
4. Architecture map, rolled up: applications ranked by total cost to 80% of
   the non-labour base. Columns: application, vendor, Platform (typed, blue),
   Owner (typed, blue), paid by TDD, paid by BU, total, portfolios paying,
   share, cumulative share.
5. Architecture map, full: every application his file names, same columns.
   The count is reported at build (he asked for at least 50 more than v3's
   470).
6. Worklist: the items with no platform, largest first, per portfolio, so the
   platform mapping can be finished in the file.
7. Bridges (his inputs, typed blue, everything else formula): application to
   platform (pre-filled from his file where the row carries one); squad to
   platform for labour (six names match today, the rest blank); labour model
   line to his portfolio (obvious matches pre-filled).
8. Data: compact aggregates the faces read (portfolio by platform by paid-by
   by sheet; application and vendor by portfolio; a few thousand rows, never
   the 103,575 leaf rows). Ties to his Summary grid cell for cell and to the
   leaf certificate.
9. Raw: his sheets verbatim (zip level transplant, values as they are).

## Numbers basis
Non-labour = his cost types on the four sheets, leaf level, 97,582,004.31:
his grid 92,097,777.60 plus 5,484,226.71 whose Portfolio:Platform pair is not
one of his 22 rows (these land on the portfolio their row carries, platform
shown as not on his list). Labour = the FY27 model, 104,079,587.22, joined by
the line to portfolio bridge. Recoveries as his file books them (TDDLT and
TDD). The four version 3 toggles are dropped. Unclassified falls back to the
vendor where the application is blank, N/A or Review Required.

## Workflow, three agents in sequence, one pass
- Data (Sonnet, about 45 min): one script reads the new file once, builds the
  aggregates and the bridges, ties to his Summary grid and the certificate,
  writes JSON plus the Data tab rows.
- Build (Opus, about 90 min): faces as formulas over the Data tab, names by
  formula, his words only, the word scan built in, raw tabs transplanted.
- QA (Sonnet, about 45 min): checker (expected numbers from the Data step,
  not from the build), recalc on a copy, Excel openability gate, banned word
  scan, one fix loop at most.
Fable plans, briefs, gates and reports; never analyses, never builds.

## Gates before it reaches Lee
Ties to his Summary grid cell for cell; labour ties to the model's 3.5;
every portfolio tab foots to TDD corporate; recalc on a copy clean; opens in
Excel; zero banned words; zero typed names on faces.

## Defaults awaiting his go (register DEC-50 to DEC-53)
Tab set = his 14 portfolios; non-labour base = his cost types with the
toggles dropped; recoveries as his file books them; platform list = his 15
plus Not yet mapped. Out of scope until he says otherwise: Z, capex,
ownership decisions (his, from the map).
