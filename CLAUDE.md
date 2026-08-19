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
