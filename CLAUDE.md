> HOW MEMORY IS SPLIT (18/08, after Lee asked why language rules sit in a
> project): there are TWO kinds of rule and they must not be mixed.
> 1. UNIVERSAL - how Lee works, always, on any client: plain English, banned
>    words, no dashes, no italics, never build without permission, raw data
>    untouchable, the slide method. These live in docs/LEE_STANDARDS.md,
>    written to be portable - copy that file into any other project or into
>    central memory unchanged. They are repeated below so this file still
>    loads them on its own.
> 2. CLIENT-SPECIFIC - Ampol TDD facts: the cyber split, FY26/FY27 numbers,
>    FX rates, squad names, file versions. These belong ONLY here and must
>    never be carried to another client.
> When adding a new ruling, decide which of the two it is first. Slide/deck
> method = docs/DECK_STANDARD.md (method is universal, fonts and brand
> colours are a per-client skin).

## Standing rules from Lee (05-06/08) — permanent, load every session
- PLAIN ENGLISH FOREVER: every question to Lee, every label, every note — no
  modelling jargon, no options phrased in build-speak. If a question cannot be
  asked plainly, rethink it first.
- Track decisions and plans continuously: every ruling lands in
  docs/DECISIONS.md, docs/register_data.json (rebuild the register workbook),
  and here, in the same turn it is made. Lee cannot afford lost context.
- Cyber split (approved): 2.11 Cyber Risk & Service Ops splits into TWO COEs -
  Cyber (squads Cyber GRC, Cyber Risk, Cyber Sec Ops, Cyber Strat & Tech;
  budget 1.5) and Service Ops (squad Service Op & Assurance; budget 0.5).
  Membership derives from the Squad column; his raw REVIEW block stays
  untouched. The uplift part-charge roles follow their squads to Cyber. COEs carry no GM share, so the
  GM base stays 11 when the split lands. AU/NZ split ruled: Cyber 1.0 AU / 0.5 NZ; Service Ops 0.5 AU / 0 NZ.
- Sequencing (agreed): task 2 first (support % + budget tables move to the 2.x
  tabs, 1.x become formula views of them; archetype tables stay on 1.x), then
  the cyber split, then FY26.
- FY26 forecast (corrected 09/08 after Lee rejected v1 as fabricated):
  Finance reports lights on at AU and NZ level ONLY - their Presentation
  sheet: AU (TDD Corporate) FY26 lights-on budget 78.90, actuals to June
  40.97, their full-year forecast 79.31; NZ (Z-Energy) budget 52.23, actuals
  26.46, forecast 53.47. The budget is FINANCE'S lights-on budget, never the
  FY27 model's allocation table. NO portfolio/COE cut of actuals exists -
  never allocate Finance lines to portfolios (v1 mapped their leader P&L
  lines to portfolios = fabrication; v1 withdrawn, Lee told to bin it).
  RULED 09/08 (supersedes the portfolio-cut plan): the FY26 forecast uses
  FINANCE'S CUT, never the model's portfolio cut - "look at how finance has
  cut it, i dont want ours to be cut differently". Source = the JUNE file
  sheet TDD Pack (2), which carries per line: YTD actual (col O), monthly
  actuals Jan-Jun (AY:BD), Finance's own monthly forecast Jul-Dec (BE:BJ),
  FY forecast (BK), FY budget (X). AU lines (AUD): Internal Labor r31,
  Contingent r32, Personnel r33, Labour Recharge r34, Staff Cost r35 (= the
  people cost line, 20.01 YTD, lands 35.61). NZ (NZD): People Costs r49
  (10.22 YTD, lands 20.90 by their months; their other FY forecast col says
  20.31 - 0.59 gap, flagged not resolved). Allocated budgets = 0.2 Data
  Config ROW 27: AU 36.2 / NZ 17.6 / total 53.8 (Lee's ruling; these are
  the budget, not Finance's X-column budgets). NZ AUD conversion at their
  0.92. Landings: AU 0.59 under, NZ 1.63 over, total 54.84 vs 53.80 = 1.04
  over; model charge 54.54 within 0.30. Fair derived assumptions are
  allowed when drawn from deep analysis of both files (his permission
  09/08). Built TDD_FY26_Forecast.xlsx v2 via Opus agent on this spec.
  CORRECTED 10/08 pm (v4, supersedes every earlier landing): the second
  half runs on what TDD pays - relief = each priced vacancy's cost times
  its squad's support percentage (EGI-funded at zero, NZ at live FX),
  never the gross 3.10; the base is everyone at a full year, 54.4290
  (FY27 start-quarter pricing never leaks into FY26; all 12 coded roles
  proven AU-side). Landings: AU 37.68 / NZ 16.04 (19.32 NZD) / total
  53.72 - 3.22 over the 50.5 allocation, 0.08 inside the 53.8 budget.
  H2 25.22, annualising 50.45; December exit run rate 53.27. FY26 is
  independent of the pending start-quarter ruling.
- FORMATTING FOREVER (Lee 09/08): no italics anywhere in any model, no
  subheaders, no AI look or feel, no made-up formulas to force numbers to fit
  an outlook. Mocks and built tabs alike.
- Prompts for the in-model tool go in chat, never as files, unless Lee asks.
  Every prompt is written in FIRST PERSON, as if Lee drafted it himself —
  never refer to Lee in the third person inside a prompt. Prompts must be
  super detailed and fully self-contained (the tool has no context): exact
  cells, expected numbers, self-checks, report-back, and a hard no-scope-creep
  rule (report what looks broken, fix nothing beyond the brief). Each build
  prompt ships with a screenshot mock of what it produces, on real numbers.
- Wiring-move fact base (probed 09/08 on his 0608 file): eleven 1.x/2.x pairs
  (1.14 pairs with 2.15 TDD Cyber; 2.11-2.14 have no 1.x partner and are not
  touched). 52 typed Support % cells across the 1.x squad tables; budget block
  H4:I10 uniform on every 1.x; funding-block extents vary by tab (1.2 runs to
  row 26, 1.14 carries the cyber uplift funding block H13:J19 instead).
  Columns U:X empty on every 2.x — blocks land at V/W/X same row numbers, the
  Support % column at U. One exception: Digital Support NZ (1.2 G41, 100%,
  archetype-only squad, 0.32 planned cost, zero roles in REVIEW) has no 2.2
  grid row — its percentage stays typed on 1.2 until Lee rules otherwise.
- Audit rulings (06/08): Platform beginning with EGI = funded outside (the
  three EGI Customer people included). GM pot divides by 11 - the 16-line
  basis is dead, fix stale notes wherever seen. The nine offshore Enterprise
  Data delivery roles: Offshore lever, cost unchanged (WiPro-style exemption
  keyed on their squad).
- Wave Q rulings (Lee 09/08 pm): FX for NZ is LIVE FX, typed input cell,
  0.83 today - never the 0.92 planning rate, never hardcoded. The exec
  story anchors to the 50.5 ALLOCATION (spend 54.54 = 4.04 over today).
  Vacancy hire-month toggles are mandatory and drive FY26 H2 and FY27 -
  the Vacancy Audit workbook holds planned hire months; hire dates matter
  more than funding plan. Contingent and personnel forecasts are op-model
  dependent - model-driven, Finance for actuals only. LANGUAGE: never
  "theirs/their" for Finance - one company, one set of numbers. Latest
  model = f4ce93da upload: cyber split APPLIED (2.11 COE Cyber, 2.16 COE
  Service Ops), new 0.4 Presentation Pack, 3.9 Spans & Layers. Exec deck:
  template = FIRST 5 SLIDES of Template.pptx; FY27_Budget_Update.pptx is
  the live deck whose FY26/FY27 slides get rebuilt; every chart native
  Excel-editable, never shapes; MBB standard, Minto, action titles, short
  sharp sentences, options for key slides; agents build, Fable plans and
  QAs in loops - nothing reaches Lee below standard.
- SLIDE STANDARD (Lee 18/08, supersedes every earlier template ruling).
  THE STANDARD IS HIS METHOD, NOT THE PAINT - he corrected me on this:
  "the font etc is specific for this client so not so much the standard -
  it's the way that it flows, the language, how i design things from the
  perspective of where things are on the page, how it reads left to right,
  how the story flows, the so what." Fonts/brand colours are a per-client
  skin, swapped every time. Full standard: docs/DECK_STANDARD.md - READ IT
  BEFORE BUILDING ANY SLIDE. Worked example kept untouched at
  reference/Ampol_FY27_Budget_Update_FINAL.pptx.
  The method: (1) SO-WHAT IN THE TITLE - a full sentence with the number
  AND its consequence/condition ("tracking 320k over ... WITH vacancy and
  offshore controls available to bring TDD back within budget"); titles
  read alone must give the whole argument; bad news stated plainly with
  the response attached. (2) STORY FLOW - answer first (exec summary = the
  whole case in three NUMBERED columns), then what changed, proof it works,
  where we stand, where we are heading, then the levers; each slide answers
  the question the previous one raises; never a slide because data existed.
  (3) PAGE READS LEFT TO RIGHT - EVIDENCE LEFT (~70%: tables, charts),
  MEANING RIGHT (~25%: "Insight & Assumptions" rail); assumptions never
  hidden, they get their own named column; left-to-right also = before to
  after (From left, To right); TOP TO BOTTOM = broad to specific (gross
  table above, lights-on beneath; what above, when beneath). (4) LAYERS -
  title, then a kicker line above every table giving what it is + units +
  basis, then the numbers, then numbered footnotes, then speaker notes as
  the full spoken narrative conclusion-first; a reader can stop at any
  layer and be correct. (5) NUMBERS always paired with what they are
  measured against ("$54.2m vs a budget of $53.8m", "24% to 16%"); every
  table resolves with a labelled total row; brackets for negatives; red
  only for over; TBC where genuinely unknown, never a dash or a plug;
  numbered callouts link a cell to its insight bullet. (6) DENSE BUT
  LAYERED - substance not decoration. (7) Charts native and editable,
  never pictures - waterfall to bridge, columns+cumulative line for
  phasing, direct labels over legends, zero baselines. (8) Sentence case
  everywhere. "charged to"/"recharged to" are HIS words and are fine - the
  old ban on charge/charged is void, his finalised deck uses it throughout.
- Flip-flop analysis (Lee 10/08): roles GMs put on Hold (priced zero) that
  now carry hire months in the vacancy sheet - show FTE and dollar impact
  of the flip-flopping; not in the numbers yet, narrative only. Updated
  model post-vacancies = 1347cab8 upload; FY26 runs off it.
- LANGUAGE RULING (Lee 10/08 pm): "saving/savings/saves/saved" BANNED
  everywhere - nothing is a saving, it is a REDUCTION OF COST. Say cost
  reduction / reduction of cost / takes cost out. Applies to the deck, the
  models, the FY26 workbook, prompts and chat. The model's own 3.10 labels
  (Saving ($m), Saving achieved, Potential saving) go to the next in-model
  fix prompt. Full banned list now: charge/charged/charging (recharge and
  Labour Recharge stay - Finance's line name and Lee's word), wave, seat,
  floor, saving in every form. Standing from the model era, enforced in
  every artefact: em dashes and en dashes, call/calls as jargon, roster,
  Category as a column label, GM working copy, dashes as cell filler,
  theirs/their for Finance.
- Offshore scenario tool (Lee 12/08): v1 Excel picker rejected as unusable
  ("looks like shit", 540 dropdown rows, no toggles). His answers: keep an
  Excel; each role its own start month; three baskets to COMPARE; yes to a
  month-by-month picture. He then wanted it "slick like a coded app,
  automated". RULED: no macro book (blocked by default, breaks on
  Mac/web/phone, and VBA can't be built or tested in this env = unverified).
  Delivered BOTH: the slick coded thing = a web app
  (deliverables/offshore_scenario_planner.html, published artifact) with live
  search/filters, A/B/C toggles, per-role months, live cards + SVG run-rate
  chart + CSV export; and a clean formula Excel
  (deliverables/TDD_Offshore_Scenario_Builder.xlsx, Start here/Pick
  roles/Compare/Lists, AutoFilter, green toggles, protected Tdd123). Offshore
  maths unchanged: moved role = 40% of full onshore; vendor day-rate exempt;
  held = added cost; FY = Ampol calendar year (2026/2027), horizon
  Jul-2026..Dec-2027. Demo basket ties 0.38966/yr, 0.07118 in 2026. Watch the
  case-insensitive SUMIFS RETAIL/Retail double-count - fold with ONE SUMIFS.
- RAW DATA RULING (Lee 12/08, permanent): NEVER transpose or reshape his raw
  data. Any file built off a source tab carries that tab VERBATIM - same
  rows, same columns, same orientation, values as they are (his #N/A cells
  at REVIEW A31/C31/S525 stay). Derived working grids are additions, never
  replacements, and must trace back to the raw tab (R0001 = REVIEW row 2,
  by row order). The offshore builder always contains REVIEW - Complete
  Role Mapping in full. Never claim a data tab is in a file unless it is
  actually there - the v3 dashboard build shipped a front-page note saying
  the mapping was included when the tab had been dropped; caught by Lee.
- Non-labour mapping (Lee 18/08, "sweet build" after the analysis
  playback): deliverables/TDD_NonLabour_Mapping.xlsx via
  scripts/build_nonlabour_mapping.py. AU HW/SW SAP export has a
  cost-centre HIERARCHY (parents AND children exported - solve before
  summing; leaf truth = SW 51,288,134.47 + HW 25,468,998.32 =
  76,757,132.79 over 49,910 lines). Six open calls live as dropdowns on
  the Decisions tab (EUC central/follow, Marketing digital, CCO, Brand,
  Distribution, GENERAL), proposals loaded and marked PROPOSED - his to
  set, never hardcode. Profile moves off block subtotals (LibreOffice
  cannot gate 100k cross-sheet IFs; ledger Final columns are static under
  current proposals, refresh after rulings). The ~302m = Ampol opex 189.1
  + capex 48.1 + Z opex 64.8 per the FY26 dashboard. Not in the AU file:
  network, outside services, depreciation, capex, all of Z/NZ. Enterprise
  Data nets ~zero on current tells (flagged, never plugged). Full analysis
  docs/NONLABOUR_MAPPING_ANALYSIS.md.
- FY27 portfolio dashboard (Lee 28/08, D191/D192): the workbook is
  deliverables/TDD_FY27_Portfolio_Summaries.xlsx. Ten option mocks were
  tested with him; ruling: the ORIGINAL dashboard design stands ("nothing
  new or better"), option 7 (Lever mix) and option 10 (Cost per FTE) are
  ADDITIVE tabs only. Portfolio tabs must show portfolios pay for
  overheads: KPI band per his Enterprise Data markup (Labour cost total |
  TDD Overheads | Total {line} Labour Cost | Lights on | Sig items |
  Initiatives | Over/(under) | Lights on % | Total FTE) plus a typed
  sentence NEVER a formula: "Lights on Xm is Ym own labour charged to
  lights on plus Zm allocated overheads, P% of labour cost total."
  Platform-grouped squad chart, platform cost chart and FTE chart on
  every portfolio tab; squad working block titled "Squad cost" ("inside
  the line" BANNED as confusing). Proven numbers basis: total
  98.979587216 = own lights 45.401019872 + sig 19.832731041 +
  initiatives 33.745836303 (his model's new 3.5 col Q = C-D-L+F+G+H,
  never negative, TDD Cyber 0.17); lights on 54.203504558 = own +
  allocated overheads 8.802484686 (BP 2.3135484 + DA 1.388936286 + GM
  5.1, each divided by 11); vs 50.5 allocation 3.70 over, vs 53.8 budget
  0.40 over; FTE 537.8 (473.8 doers), levers Filled 351.5 / Hire 66.0 /
  Offshore 83.8 / Hold 36.5. REVIEW carried verbatim (helper-formula
  columns stored as computed values, three #N/A kept). Open with Lee:
  Enterprise Data 0.50 sig-tagged no-funding group; his model's stale
  bridges (3.5 r30-32, 3.8 r34-38, off by 8.80) and stale 3.8 B2 title;
  REVIEW platform hygiene (NA 75 rows, blank 20, AMPOS/AmPOS spellings)
  kept visible as raw.
- HoT cost budgets engagement (Lee 28/08, D194): assign a total cost
  budget to each head of technology, labour plus non labour, off
  TDD_AU_Consolidated_2027_budget_v4.xlsx (four transaction towers HW/
  SW/NW/TDD plus his own mapping engine tabs) and the FY26 spend
  dashboard pdf. TDD owns TDD costs; all costs must net; the central
  question is how a HoT is responsible for costs the business pays
  vendors directly. Instruction register docs/HOT_BUDGET_INSTRUCTIONS.md
  (I1 to I28) audited at every gate. Workflow: Fable orchestrates and can NEVER be an agent;
  agents Sonnet mostly, Opus where needed; quant and cost distinct
  lanes. RULED 28/08: HoT question parked, EVERY portfolio gets its
  total cost base (labour + non-labour); basis = FY26 window Jul25 to
  Jun26 carried unchanged to FY27; Z and capex later; missing platform
  mapping resolved from SAP transaction detail; build authorised,
  report only when done. Known traps carried from the v1 file: SAP
  cost centre hierarchy (parents and children both exported, leaf level
  first), case sensitive matching (RETAIL/Retail). Scratchpad:
  hotbudget/ (budget_v4.xlsx, labour_dashboard_2026.pdf, findings_*).
- RULED 02/09 (D204, I26): costs come ONLY from the four transaction tabs
  (HW/SW/NW/TDD Line Items, Jul25 to Jun26, assumed stable into FY27).
  The budget tab and the FY26 dashboard are a different period: a map of
  structure at most, never a source of dollars, ratios or counts. Fable
  orchestrates only, never probes; agents Sonnet or Opus, never Fable.
- RULED 02/09 (D207, I27): the ARCHITECTURE MAP of the cost comes first,
  who pays for each application and how much, and ownership is decided
  AFTER reading it, never proposed on a front page (Salesforce is one row
  among 470, not the lead). RULED 02/09 (D208, I28): the product is a tab
  per PORTFOLIO in the labour dashboard pattern (his Enterprise Data
  screenshot); no HoT tab, no summary or so-what page, no findings names
  on client pages. SHIPPED 02/09 (D209): deliverables/
  TDD_Portfolio_Cost_Base.xlsx v3, 31 tabs (Read me, TDD corporate,
  Architecture map of the cost, thirteen portfolio tabs, working tabs,
  seven raw tabs verbatim). Basis: total 201.11m = labour 104.08 (his
  fixed 3.5) + non-labour 97.03 from the four transaction tabs at leaf
  level (raw 223.10m, leaf 134.57m, overstatement 88.53m from parent and
  child cost centres both exported); TDD pays 58.17m net, business
  direct 38.85m. Toggles on Assumptions, typed blue, default to the data
  as is: T1 contractors 2.00m, T2 personnel 0.85m, T3 platform pools
  24.68m (Microsoft 19.34 + SAP 5.34), T4 DAA split 0.61m. Map: 470
  applications, 89 multi-payer 53.46m (the ownership calls live here),
  272 single-payer 30.95m, 109 no vendor 12.94m. Open with Lee:
  ownership of the 89, the four toggles, and the Read me register kicker
  still reads "I1 to I27" (table runs to I28; corrected in the build,
  lands with the next rebuild). Pipeline: scratchpad hotbudget/cb3
  (build3, faces3, data3, check3 140 checks, gate3 recalc on a copy,
  render3). Next design must pre-aggregate: the faces scan the
  103,575-row Assignments tab and recalc slowly.
- v3 REVIEWED by Lee 03/09 (D210, I29 to I38): HIS WORDS ONLY ("tower",
  "line" for a portfolio, "lines pay it / use it" BANNED; the four sheets
  are Software, Hardware, Network, TDD; who pays = BU / TDD, his Summary
  headings); align to his '0725_0626 Summary' grid (Portfolio by
  Platform, BU paid against TDD paid); platform breakdown per portfolio
  with the unmapped remainder visible and a worklist; top 20 applications
  or vendors or 80% (vendor fallback); names in column B by formula; two
  maps (80% and full) with typed Platform and Owner columns; far fewer
  words; labour joined; time and token efficient workflow. New source =
  actuals_0725_0626.xlsx: no dollar moved, mapping changed (TDD rows now
  carry Portfolio and Platform, software Cost Type split, Vendor "N/A"
  convention). Findings: his Summary reproduces cell for cell and sums at
  leaf level; his grid 92.10m vs our 97.03m explained to the cent (5.48m
  off his grid, one COE Cyber Risk and Service Ops line, recoveries to
  TDDLT and TDD); 29.67m of 97.58m (30.4%) has no platform; Ampol Retail
  10.15m no application, 84% of it has a vendor; top 25 clears 80%
  everywhere. Plan docs/HOT_BUDGET_PLAN_V4.md (three agents in sequence,
  one pass), defaults DEC-50 to DEC-53. GO GIVEN 03/09 (D211, I39): the
  COEs and TDDLT have no platforms (no platform table there); Lee fills
  the Retail and Customer platforms himself via the typed Platform column
  and the worklist. Pipeline scratchpad hotbudget/cb4. D212: non-labour
  follows his Hierarchy column (his Summary tab's logic) = 92,117,406.02,
  not the 97.58m quoted (mixed classifications); the 42.45m TDD sheet
  rows outside the tech view are a reconciliation block, not cost base;
  total cost base 196,196,993.24; map 346 names from his file (470 was
  padded with our fallback labels); labour from the model 98,979,587.22.
- SHIPPED 03/09 (D213): deliverables/TDD_Portfolio_Cost_Base.xlsx v4, 29
  tabs (Read me, TDD corporate, Map 80 percent, Map full, fourteen
  portfolio tabs, Worklist, Bridges, Data, Checks, seven raw). Basis
  196,196,993.24 = labour 104,079,587.22 + non-labour 92,117,406.02 (his
  Hierarchy column; 22 pairs tie his Summary row 41 to the cent;
  42,453,819.42 outside his tech view as a reconciliation block). Map
  346 names, 56 to 80% before credits; 159 platforms pre-filled under
  the strict rule. Pipeline scratchpad hotbudget/cb4 (data4, build4,
  check4_qa, render4, ship4). Open with Lee: his platform typing on Map
  full for Retail and Customer, the labour platform bridge (27 blank),
  Owner column, ownership of multi-portfolio names.
- v4 REJECTED by Lee 04/09 (D214, I40 to I46): charts and graphs at the
  TOP of every tab and more of them, exec view up top then detail; all
  chart labels, WHITE text; Share and Cumulative columns gone; the
  labour dashboard content comes in (overheads, lights on, sig items,
  initiatives, over/(under), lights on %, FTE, squads by platform and
  lever), total cost base = total labour cost + non-labour; every
  portfolio tab in the Summary tab's shape (per platform
  Software/Hardware/Network each TDD or BU paid); removing explainers
  never meant less insight. Process: five portfolio options and two
  corporate options as screenshots on real numbers, playback and purpose
  statement, NOTHING BUILT before he picks. Mocks pipeline scratchpad
  hotbudget/mock5.
- ROUND TWO 04/09 (D216): the ten mocks rejected as "the exact
  screenshot of one of the images"; options must answer DIFFERENT
  QUESTIONS with different evidence, never one layout in five coats.
  Five lenses to five Opus designers in parallel plus a judge
  (scratchpad hotbudget/mock6): A where the money goes, B who pays and
  who should own it, C run against change, D what can move, E the FY27
  budget. Each lens = an Ampol Retail page and a corporate page on real
  numbers. Nothing built before his pick.
- ROUND TWO SENT 04/09 (D218): 22 pages, five lenses judged and fixed;
  judge's usefulness ranking D what can move, B who pays and who owns, C
  run against change, E FY27 budget, A where the money goes.
  Recommendation put: portfolio tab = B with D's lever block and E's
  quarters beneath; corporate = E with B's shared applications and D's
  bridge. AWAITING HIS PICK; nothing built.
- LENSES REJECTED 04/09 (D219): "none work"; his v3 tab (light page, one
  navy family, $m, tiles, charts up top) is "aesthetically the closest";
  "wtf are all the colours for?"; "don't go off creating your own
  interpretation or own view". RULED: v3 look, navy family only, white
  labels inside navy bars, HIS LIST ONLY (labour as the labour
  dashboard; total cost base = total labour cost + non-labour; Summary
  tab shape per platform; TDD against business paid; platform breakdown
  with unmapped remainder; top 20 or 80%; cost shared with other
  portfolios and platforms; charts on top, detail below). Faithful mock
  in scratchpad hotbudget/mock7, then playback.
- FAITHFUL MOCK SENT 04/09 (D220): scratchpad hotbudget/mock7, his list
  in his v3 look; Ampol Retail 31.36m = 12.54m + 18.82m; AWAITING HIS
  GO; nothing built.
- GO 04/09 (D221, I47): version 5 = the D220 mock on every tab, formula
  driven as v4, labour dashboard joined per portfolio, navy family only,
  $m; pipeline scratchpad hotbudget/cb5; screenshots of every tab before
  ship; then iterate with him.
