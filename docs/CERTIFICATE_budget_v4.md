# Structural integrity certificate

File checked: TDD_AU_Consolidated_2027_budget_v4.xlsx (working copy budget_v4.xlsx)
Checked with: sap_integrity_check.py, run 2026-09-01
Result: PASS. All four SAP tabs tie to the cent.

## What this checks and why it matters

This file's four SAP export tabs (HW Line Items, SW Line Items, NW Line Items, TDD Line Items) do not use one row per posting. SAP has displayed some postings more than once, under more than one cost centre group, on the same tab. Adding up all the rows on a tab therefore counts some postings twice, or more. SAP's own Overall Result row at the top of each tab is not affected by this, only a plain sum of the rows underneath it is.

This certificate proves, for every tab, exactly which rows are the extra ones, ties the total left after removing them to SAP's own Overall Result row to the cent, and names the groups responsible. Never total a tab by adding up its rows. Use the leaf figures on this page, or rerun sap_integrity_check.py against the live file. A hand sum will overstate every tab, in one case by more than 77m.

## The four tabs

| tab | raw rows | raw sum | leaf rows | leaf sum | overstatement | ties to SAP total |
|---|---|---|---|---|---|---|
| HW Line Items | 46,908 | 15,612,492.37 | 35,775 | 13,153,224.79 | 2,459,267.58 | yes |
| SW Line Items | 12,777 | 28,690,965.30 | 8,583 | 22,115,008.29 | 6,575,957.01 | yes |
| NW Line Items | 3,343 | 5,165,501.23 | 2,394 | 3,585,662.27 | 1,579,838.96 | yes |
| TDD Line Items | 71,173 | 173,629,478.52 | 56,823 | 95,717,330.09 | 77,912,148.43 | yes |
| total | 134,201 | 223,098,437.42 | 103,575 | 134,571,225.44 | 88,527,211.98 | 4 of 4 |

Leaf sum is the true cost on this tab. Overstatement is what a plain sum of every row would have added on top, wrongly, for that tab alone.

## What is causing it, tab by tab

### HW, SW and NW: one group, Bulk Fuels, is a straight copy

On all three tabs, the cost centre group APPLBULKFUEL is a duplicate of rows already counted under other groups. Its rows are not extra spend. Dropping the whole group, and nothing else, brings each tab back to SAP's own total to the cent.

- HW Line Items: APPLBULKFUEL, 11,133 rows, 2,459,267.58. 99.7% of its rows, 99.8% of its dollars, are byte for byte the same as rows already kept under eight other groups (APPLSUPPLYOPS, APPLSUPFPO, APPLATOTAL, APPLCCO, APPLTDDIST, APPLB2B, APPLFIFINANCE, INTEGM). Its total is exactly those eight groups added together, to the cent.
- SW Line Items: APPLBULKFUEL, 4,194 rows, 6,575,957.01. 99.2% of its rows, 100.0% of its dollars, are the same as rows kept under seven other groups (APPLB2B, APPLCCO, APPLFIFINANCE, APPLHFIN, APPLSUPFPO, APPLSUPPLYOPS, APPLTDDIST). A small number of rows on each side are not exact copies, so the group does not close to those seven totals precisely, but the row by row evidence leaves no doubt it is the same mechanism as HW.
- NW Line Items: APPLBULKFUEL, 949 rows, 1,579,838.96. 94.6% of its rows are the same as rows kept under three other groups (APPLB2B, APPLFIFINANCE, APPLTDDIST).

This is the same group causing the same problem on all three tabs. It carries no portfolio in the file's own Portfolio Mapping tab, and the client's own 0725_0626 Summary tab already reproduces the SAP totals without it, which is independent confirmation.

### TDD: six groups are parent rollups of other groups already on the tab

On the TDD tab, six cost centre groups are parent nodes. Each one's total is exactly the sum of a specific, separate set of other groups already counted on the tab (its children). Drop the parent, keep the children, and the tab ties to SAP's total to the cent.

| parent group dropped | rows | total | equals the sum of |
|---|---|---|---|
| APPLHTDDETS | 3,442 | 28,850,770.85 | APPLHTDDICD, APPLHITBSCORP |
| APPLHITSEC | 2,765 | 27,557,168.70 | APPLHITINFOP, HITSEC, APPLITSECUPLIF |
| APPLHITBSRT | 3,191 | 8,204,481.31 | ITRETSUP, ITABOVE, ITMCONS, ITBUSOL |
| APPLHTDDFSC | 2,249 | 6,086,818.44 | ITTS, ITMB2B, ITSUPSYS, APPLHTDDEDEL, APPLHTDDCPQA |
| APPLHTDDSA | 1,821 | 5,547,699.74 | TDDSA, APPLHTDDDAA |
| APPLHITMGT | 882 | 1,665,209.39 | HTDDOPEX, HITMA, ITSTRAT |
| total dropped | 14,350 | 77,912,148.43 | (the six rows above) |

One of the six, APPLHTDDFSC, turns out on closer check to be a clean copy of its five listed groups in full (94.2% of its own rows are byte for byte the same as rows kept under exactly those five), so it reads as the same Bulk Fuels style duplicate as HW, SW and NW, not only as a rollup. The other five hold together only as an exact sum, without that row for row copy evidence, which is the parent and children pattern in its clean form.

## Ties checked outside each tab

The leaf figures above were also checked against SAP's own Overall Result row on each tab (the pass or fail test) and, separately, against every other tab in the file, for any cell carrying the same figure to the cent.

- HW leaf 13,153,224.79 is also the figure in 0725_0626 Summary, cell E30 and its own check row E31.
- SW leaf 22,115,008.29 is also the figure in 0725_0626 Summary, cell C30 and its own check row C31.
- NW leaf 3,585,662.27 is also the figure in 0725_0626 Summary, cell G31 (the Summary's own grid total for network, C30, is 3,584,962.27, a separate and already known gap of 700.00 caused by 28 rows with no platform mapped; the check row itself carries the correct figure).
- TDD leaf 95,717,330.09 does not appear on any other tab in this file. No portfolio grid in the Summary tab totals the TDD tower on its own, so there is nothing else in the file to check it against. The tie to TDD's own SAP Overall Result row still holds to the cent, which is the pass or fail test.

## Certification

Every leaf figure on this page was computed twice, once by taking the raw total and subtracting the dropped groups, and once by adding up only the kept rows directly, and the two agreed to the cent in every case. All four tabs tie to their own SAP Overall Result row to the cent.

FILE STATUS: PASS, 4 of 4 tabs.

The same script, unchanged, was also run against the prior AU HW and SW file and reproduced its already proven leaf total of 76,757,132.79 over 49,910 rows to the cent, finding the same style of duplicate group there too. This is now a standing check, not a one off build for this file.
