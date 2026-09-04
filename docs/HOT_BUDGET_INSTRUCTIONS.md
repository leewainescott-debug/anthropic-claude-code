# HoT cost budgets engagement: instruction register

Lee's instructions, 28/08, numbered so every deliverable can be audited
against them end to end. The instructions agent owns this file: at every
gate it checks each item as met, partly met with what is missing, or not
yet started. Nothing ships while an item sits unmet without Lee's say so.

## The ask
- I1. MBB thought partner at the table: a technology and recharging
  specialist and a cost specialist.
- I2. Agentic workflow: Fable plans and orchestrates and can NEVER be
  an agent (ruled 28/08). Agents are Sonnet mostly, Opus where
  genuinely needed. Quant and cost stay distinct lanes.
- I3. Agent roster he named: quant, CFO, cost, research, instructions,
  cost QA, technology QA, operating model, MBB thought partner; my own
  suggested additions welcome, duplicates to be called out.
- I4. Establish what TDD pays for and recharges back to the business
  units: the dollar amounts, what the costs are, which business units
  they recharge out to.
- I5. Establish what the business pays for directly, where the business
  controls the spend.
- I6. Answer the central question: if the business pays the vendor
  directly, how does the head of technology become responsible for the
  cost of the software or the platform?
- I7. From the budgets: who is responsible for each cost, and an
  allocation approach. Who owns the software platforms, who is
  responsible for the cost, how the cost is managed going forward.
- I8. Analyse the budget spreadsheet: transaction level costs for
  software, hardware, tdd and network.
- I9. Analyse the labour dashboard: people, their costs, portfolios,
  platforms and squads; this is the operating model.
- I10. TDD owns TDD costs.
- I11. Review whether the portfolio and platform mapping on the HW, SW,
  TDD Line Items and NW tabs makes sense, and the cost type, vendor and
  application fields with it.
- I12. Understand the vendor payment logic, Salesforce as the worked
  example: only a few platforms or portfolios pay Salesforce directly;
  do they pay the vendor then recharge out to business units such as
  Finance, or how does it work?
- I13. All costs should net. Total cost for each of the HW, NW, SW and
  TDD line items.
- I14. End goal: a total cost budget for each head of technology,
  labour plus non labour.
- I15. Plan the work, replay the ask in plain English, say what he has
  missed, bring the agents' findings and my own view.

## Standing rules that bind every deliverable in this engagement
- I16. Plain English everywhere; his standing banned words; sentence
  case; no em or en dashes; raw data untouchable; never build a
  deliverable without his permission; every ruling logged the turn it
  is made; accurate numbers independently checked, never agreement for
  agreement's sake.

## Rulings added 28/08 (second instruction set)
- I17. The cost bases build off the FY26 period in the file (July 2025
  to June 2026) and carry the same costs to FY27; FY27 is assumed the
  same as FY26.
- I18. Where platform or portfolio mapping is missing, use the SAP
  transaction detail on the line (vendor, application, text, cost
  centre) to place it.
- I19. The head of technology question is parked. Every portfolio gets
  its total cost base, labour plus non-labour.
- I20. Z and capex come in later; take the data as it is today.
- I21. Build authorised; come back only when done, and the plan must
  cover everything asked.

## Rulings added 28/08 (after the retrospective: "fix everything, do some detailed planning, ensure agents run")
- I22. On every SAP sourced file, a structural integrity check runs
  first and alone, tying leaf level truth to every control total before
  any lane trusts a sum.
- I23. Mapping rules are written twice, independently, and diffed; the
  confidence grade comes from agreement, never from the author.
- I24. Technology QA runs before the build, cost QA after it, the
  instructions audit last; QA findings are claims to verify, never
  figures to build on.
- I25. Interim contact exception to I21: any single judgement above five
  percent of the total base earns one short question the moment it is
  found; nothing smaller interrupts.
- Plan of record: docs/HOT_BUDGET_PLAN_V2.md.
- I26. Costs are determined only from the HW, SW, NW and TDD Line Items
  tabs (the transactions, July 2025 to June 2026, assumed stable into
  FY27). The budget tab and the FY26 dashboard are from a different
  period: a map of structure at most, never a source of dollars, ratios
  or counts (ruled 02/09).
- I27. The architecture map of the cost comes first: who pays for each
  application, how much, and how many lines pay and use it. Ownership is
  decided by Lee after he has seen the map, never proposed on a story
  page before it (ruled 02/09).
- I28. The product is a tab per portfolio in the labour dashboard's
  pattern (the Enterprise Data page is the template), plus the corporate
  view and the architecture map. No head of technology tab or naming, no
  summary or story page (ruled 02/09).

## Rulings added 03/09 (Lee's review of version 3 and the new mapping file)
- I29. His words only, everywhere in the file and in chat: no "tower"
  (the four SAP sheets are Software, Hardware, Network and TDD, as his
  Summary tab labels them), no "line" for a portfolio, no "lines pay it
  / lines use it". Who pays is "BU" and "TDD", his Summary tab's own
  headings. Portfolio, Platform, Application, Vendor, Cost Type are his
  column names and stay as they are (03/09).
- I30. The non-labour view aligns to his '0725_0626 Summary' tab: the
  same grid (Portfolio by Platform; Software, Hardware, Network; BU paid
  against TDD paid), and every portfolio tab shows what TDD pays and
  what the business pays in that shape (03/09).
- I31. Every portfolio shows its non-labour broken down by platform,
  with the remainder that has no platform shown as its own row, never
  hidden or forced. A worklist of the unmapped items, largest first,
  sits in the file so the platform mapping can be finished. He assigns
  applications to platforms in the file and the cost follows the
  assignment by formula (03/09).
- I32. Every portfolio tab shows its top 20 applications or vendors, or
  whatever number reaches 80% of the cost base. Where the application is
  blank the vendor is used. Leaving about 16m unclassified in Ampol
  Retail is not acceptable (03/09).
- I33. Names in column B of every tab come by formula from the working
  table, never typed, so the tabs read as built in Excel by a person
  (03/09).
- I34. Two architecture maps: one rolled up to about 80% of cost, one
  full end to end. Both carry a Platform column and an Owner column for
  him to type. He asked for at least 50 more applications on the map;
  the map takes every application his file names and the count is
  reported at build (03/09).
- I35. Far fewer words. No AI-sounding words. One line above a table at
  most, no essays, no explainer rails. Charts only where their meaning
  is obvious from the title, in his words (03/09).
- I36. The source for non-labour is his latest mapping file,
  TDD_Consolidated_0725_0626_Actuals.xlsx (scratchpad copy
  actuals_0725_0626.xlsx). Every mapping change in it is carried; the
  transaction dollars are unchanged from budget_v4 (03/09).
- I37. The labour view joins the same tabs: labour by portfolio and,
  once the squad to platform bridge is typed, labour by platform beside
  the non-labour (03/09).
- I38. The workflow must be time and token efficient: the fewest agents
  that do the job, in sequence, one pass, data pre-aggregated so the
  faces read a compact table; the version 3 run took too long (03/09).
- I39. Version 4 go given 03/09. The COEs have no platforms: no platform
  table and no Not yet mapped row on the COE tabs (and TDDLT). Lee fills
  the platform for the Retail and Customer side himself in the file, so
  the typed Platform column and the worklist serve those two portfolios
  first (ruled 03/09).

## Rulings added 04/09 (Lee's review of version 4: "the dashboards have digressed")
- I40. Charts and graphs at the top of every tab, and more of them. An
  exec reads the dashboard view up top, charts and graphs, then the
  detail below. The dashboard has to be advanced; his five reference
  images (dark finance dashboards with KPI tiles, donuts, bars, trend
  lines; the Xero profit and loss dashboard) set the bar.
- I41. Every chart and graph carries all its labels, and the writing on
  them is white text.
- I42. The Share and Cumulative columns do not make sense and go. The
  80% idea is shown another way or not at all.
- I43. Labour comes in the way the labour dashboard shows it: labour
  cost total, TDD overheads, total labour cost, lights on, sig items,
  initiatives, over/(under) the allocation, lights on %, FTE, squad cost
  by platform and lever. Platform costs for labour without the overheads
  is not the true narrative. Total cost base = total labour cost
  (overheads included) + non-labour; otherwise the model is half baked
  and the heads of technology do not know what their labour and non-
  labour cost is made of.
- I44. Every portfolio tab breaks its non-labour down the way the
  0725_0626 Summary tab does: for each platform, Software, Hardware and
  Network, each split into TDD paid and BU paid, plus the other columns
  on that tab. Collapsing this to TDD paid against BU paid totals was a
  digression to the old model.
- I45. Removing explainer text never meant removing depth of insight.
  The portfolio tabs must carry more insight and depth, not less.
- I46. Process: five options for the portfolio tab and five for the TDD
  corporate tab ("5 portfolio examples and 5 tdd examples!") as screenshots on real numbers, then a playback of what
  the dashboard provides and what the work is trying to achieve, before
  anything is built. Screenshots of everything before any build.
- I47. Version 5 go given 04/09 ("build that out and we can start to
  iterate"): the faithful mock (D220) built across every tab, formula
  driven as version 4, the labour dashboard content joined per
  portfolio, his v3 look with the navy family only and white labels
  inside navy bars, $m; screenshots of every tab before it ships; then
  iteration with him.
- I48. Nothing ships before Lee approves the screenshots: build, QA, a
  screenshot of every tab to him, his approval, then the file to him and
  into the repo (04/09, "there's no point rendering if you ship before I
  approve").
