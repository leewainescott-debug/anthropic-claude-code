# The complete instruction register

Every instruction you have given across this engagement, from the first message to today.
Compiled from all 60+ of your messages on file. Items marked NEW or CHANGED came from today.
Items marked SUPERSEDED are earlier asks you later replaced. Nothing else is superseded.

This is the checklist the build executes against. Strike anything wrong, add anything missing.

## A. The model itself (your first brief, still the spine of everything)

1. A calculator per portfolio: squad rolls to platform, platform rolls to portfolio.
2. Squad rows: Squad Type dropdown (engineering, configuration/integration, product, operations, enterprise data and insights, build and run), Size dropdown (XS, S, M, L), Onshore/Offshore dropdown, Support % dropdown (0, 20, 100).
3. Squad cost comes from the squad archetype sheet by type and size. Offshore price is the offshore column, 40% of onshore.
4. TDD cost = squad cost x support %. Other = the remainder. Both shown.
5. Platform overhead per platform (the 165K from Data Config), portfolio overhead on top (the 606,945 line and the leadership shares).
6. Budget per portfolio from Data Config allocation (Ampol Retail 2.5). Variance = budget minus cost. Then work through TDD CapEx, project OpEx, lights on, to what is now left over and how we fund it.
7. Budgets pull from the finance pack (OpEx initiatives, depreciation, significant items, CapEx) for Z Energy, commercial fuels, retail, energy solutions, finance, other, core, P&C, customer.
8. Strategic programs (AmPOS, CTRM, EGI): no size, always on, editable support %, hard typed total cost in a yellow input, flow to Significant Items, EGI on its own line.
9. Group summary ladders to the full TDD budget 53.8 (AU + NZ) and the 43.5 allocation.
10. All formulas must pull. Dropdowns must work. Everything logical, presentable, formatted.

## B. Structure and flow (CHANGED today, this supersedes all earlier numbering)

11. NEW: 1.x = work the archetypes. Your ten portfolio design tabs plus the COE role tabs (BP&T, SA&D, Cyber) move into this group.
12. NEW: 2.x = the working tabs (the current 4.x become 2.x). This is where vacancy decisions get made.
13. NEW: 3.x = the summaries after the working tabs: Group Summary, Total Cost, Squad Detail, COE Summary, and your FTE view restored as its own tab.
14. Flow is one direction: 1 into 2 into 3. Numbering = flow. Separator tabs between groups, colored tab groups.
15. NEW: never redesign a tab. Change only what you specify, exactly how you specify it. The proof mock I sent violated this; it is dead.
16. A guide tab that explains: if I update this, it updates that. Edit map and change propagation.
17. What is the purpose of each tab, are they grouped correctly, what are we missing: the workbook has to answer this on its face.

## C. The 1.x design tabs (your design, untouched)

18. The archetype mechanic stays exactly as you built it: squad rows, type/size/on-off/support dropdowns, INDEX/MATCH into the archetype library, platform blocks, platform overhead lines, platform totals, portfolio roll up.
19. AU/NZ toggle on every squad row, done the way your example screenshot showed (AU/NZ column in the squad table), driving whether the squad is AU or NZ funded. Crucial. Z Retail and Z platform squads default NZ.
20. Variance to AU budget and variance to NZ budget shown.
21. Fix the broken total to fund on the 1.x tabs: it is not coming through clear.
22. Fix the missing gap between rows F and G (the crammed budget box).
23. The TDD funded vs business funded split lives here and must surface downstream so you can see how much more is left to fund.
24. Positives and negatives aligned consistently for total to fund and variance across sheets.

## D. The working tabs (become 2.x)

25. Title: "[Portfolio] working copy". Section header: "[Portfolio] FTE".
26. Lever column header: "Vacancy lever". Dropdown: Hire, Hold, Offshore. Offshore = 0.4x the role cost.
27. CHANGED today: the column set per your screenshot: Archetype cost ($m), Actual cost ($m), Variance, Cost after vacancy decisions, New Variance. The old header wording is banned (see K).
28. CHANGED today: archetype size gets its own column, separate from type. Not joined.
29. Show vacancy impact: vacant, planning to hire, vacancies remaining.
30. No costs shown against filled people, anywhere. Only vacant roles carry a cost.
31. Vacant roles priced at standard title rates, noted as indicative.
32. Every single role (536 Squads rows plus the Sheet2 rosters) appears on exactly one working tab. No gaps, no duplicates. "Massive gap if we miss anything."
33. Working tab results must flow: toggle a lever and vacancies remaining, cost and left to fund move on Squad Detail and the summaries. The tool is pointless otherwise.
34. The budget impact of the decisions surfaces: how much funded by TDD, how much by the business, how much more left to fund, consistent with the 1.x split.

## E. The summaries (become 3.x)

35. Group Summary in your pasted structure: Portfolio, TDD Lights On Budget ($m), Archetype Support Cost ($m), Variance ($m), Cost of FTE non TDD funded, Amount identified as rechargeable ($m), Left to fund outside TDD ($m), Total still left to fund ($m), Total Cost ($m). Archetype view note stays.
36. Total Cost in your pasted structure: Archetype cost ($m), Actual cost ($m), Variance, Cost after vacancy decisions, New Variance.
37. The Total to fund block: TDD Variance, Other Variance, Total.
38. One consolidated table on Total Cost: leadership inside portfolio cost, COEs in the same table, dollars and FTE, total at the end.
39. No squad style FTE numbers against BP&T / SA&D / COE rows (they have no squads).
40. Your FTE view restored as its own summary tab.
41. Squad Detail keeps the squad level drill (portfolio, platform, squad) and reacts to the working tabs.
42. Cyber appears once: one COE with three funding buckets (COE allocation + TDD Cyber budget + 0.5 CapEx yellow input). No separate cyber design tab. Portfolio count is 10.
43. An exec summary that tells the story in plain lines: the archetype contract, what was raised, what it costs, the vacancy decision. CTO and GM facing.

## F. The COE role tabs (move into 1.x group)

44. BP&T: all P&T roles listed live from Sheet2. Funding with real formulas: portfolios funded x 0.4 BP allocation, plus both 2.0 budgets (Business Partnering and Transformation).
45. SA&D: portfolios x domain architect rate with formulas, plus both 2.0 budgets (Strategy Architecture and Data). All roles mapped, your ruthless 29 to the COEs / 31 stay in the Group Data portfolio (Kina Birkby stays in squad).
46. Cyber: all 52 roles from Sheet2 with your updated costing, roles/filled/vacant/spend/budget/left to fund.
47. Onshore/offshore toggle restored on the COE roles (it was in your original).
48. Group by Department, no invented Category column. Commercial rolls into Business Partnering. Cyber & Risk vs Service Operations split by the Service Op & Assurance department.
49. Fix the formatting on all three role tabs, and redesign COE Summary so it is presentable.
50. Show how many times the 0.4 BP / DA overhead is applied, no double counting.
51. Paused / ring fenced statuses from Sheet2 shown and excluded from planned spend with a memo.

## G. Data, validation and reconciliation

52. Raw data is never rewritten. Squads drives the portfolio squads, Sheet2 drives the COE rosters, Added data is the cost ledger. Any mapping cell change is logged in full on the Data QA tab (currently 51 logged cells).
53. Validate the 166 vacancies: a live check that the working tabs cover exactly the raw count, split by portfolio.
54. Triple check SA&D and cyber numbers pull through to total cost.
55. All numbers reconcile. Check cells read zero. The two population snapshots (Squads counts vs Added data dollars) get a plain disclosure on the summary pages.
56. Data QA tab explains what it is telling you.

## H. Your audit register from today (all confirmed against the live file)

57. Data Config AU budget: replace the +1 plug with the live Energy Solutions reference (=G5+G7).
58. 1.8 ES & B2B: the 5.2 B2B CapEx has no Finance line. Keep your number, label it as not yet in Finance, add a check cell. Your open questions stay visible.
59. 1.9 Commercial Fuels: stop pulling the central TDD CapEx pool into CF's own budget; tie to Finance at 40.5.
60. 1.10 Z Retail: the 1.08 plug matched to a live output, the stale =D9 reference, and the MAX clamp that no other tab has. Fix the reference, standardize the treatment, replace the plug only with your agreed figure.
61. Strategic program squads showing 0 while real people sit in them (about 7.7m): surface the actual people cost beside each yellow input for you to set. I do not invent the numbers.
62. Hardcoded row references into the raw sheets (about 4,300): hardened so a sort or insert cannot silently corrupt the model.
63. XS size selectable for types with no XS archetype: dependent dropdowns plus a visible flag so a squad can never silently drop out of the roll up.
64. Hardcoded people allocations: keep your numbers, add who agreed / when beside each.
65. Totals that sum text dashes and contradict their own columns: fixed.
66. The COUNTA fixed window driving BP/DA money: named range, one source, de-dup tied to it.
67. Data Config inputs get the yellow input formatting so an EGM can see what is editable.
68. Stale hidden sheets (squad mapping, FY26 ref, empty Sheet1): labeled superseded or removed, your call recorded either way.
69. Your working notes (COE side notes, the J column notes on 1.x): moved to a notes area, never deleted, never front of house.
70. Freeze panes on every long table. Red/green conditional formatting on variances, consistent. One money format. No mixed decimal places.

## I. Formatting and design rules (standing)

71. Nicely formatted, presentable, consistent throughout. No stray blue/white cells, no white gaps, no orphan headers.
72. Inputs are yellow. Only true inputs are yellow.
73. No fonts under 10pt in anything we author.
74. No grey italics, no AI note style.
75. Simple language, your vocabulary: archetype, TDD Cost, Lights On, left to fund. No consultant jargon.

## J. Process rules (standing)

76. Confirm what I am doing before proceeding. Play back the approach. Timebox it and tell you the delivery time.
77. Ask questions, never assume.
78. Do not come back unless 100% certain the ask is delivered. Use the instruction register (this document) to check compliance every time.
79. Full transparency into build progress.
80. Your cells and your agreed numbers are never overwritten by me. Ever.

## K. Banned, permanently

81. The word "call" or "calls" in any label, header or note. It is the Vacancy lever and vacancy decisions.
82. The word "roster". It is FTE.
83. The word "seats".
84. An invented "Category" column.
85. Em dashes and en dashes in anything we write, workbook or chat.
86. Hardcoded values that look AI generated: names, words, numbers. Formula driven or live references only, except your yellow inputs.
87. Possessive AI phrasing ("your squads, your people").
88. Costs displayed against filled people.
89. "GM working copy" in titles. It is "[Portfolio] working copy".

## Superseded (for the record, not live)

S1. The original 4.x numbering for working tabs (now 2.x per B above).
S2. The separate 1.11 cyber design tab (cyber is one COE).
S3. "Cost after calls" wording (now Cost after vacancy decisions, and the word is banned).
S4. Earlier ask to show cost for filled and vacant on COE tabs (replaced by: no filled costs anywhere).
S5. The joined "type - size" archetype column on working tabs (size is its own column).
