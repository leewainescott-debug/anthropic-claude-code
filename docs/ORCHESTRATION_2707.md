# Orchestration plan - the 2707 consolidation round

Fable plans and orchestrates; Opus agents execute in waves. Every wave's output feeds the
next, and nothing ships until the verification wave comes back clean and the instruction
auditor confirms every task in the owner's message was done.

## What the owner asked for, itemised

| # | Instruction | Where it lands |
|---|---|---|
| T1 | Apply the review workbook's changes - 0.2 Data Config and 1.1-1.13 - to the 2707 version | Base assembly + chain rebuild |
| T2 | Load the new Customer dataset (~83 roles) into REVIEW and flow it through the model | REVIEW rebuild |
| T3 | Critically analyse the model; 1.x consistent, 2.x consistent, 3.x consistent | Verification wave |
| T4 | Consistent application of rules, design and formatting | Builders + design reviewer |
| T5 | The Hold / Hire / Offshore and On/Off levers on the COE 1.x tabs, applied through the model | 1.x spec from analysis wave + builders |
| T6 | REVIEW updated with the new cyber roles (list updated in the review workbook) and the new Customer roles | REVIEW rebuild |
| T7 | Review the finished 2707 to confirm every instruction was followed | Instruction auditor |

## The file genealogy the plan rests on

- `rev.xlsx` (to_review) descends from `base_ship.xlsx` - the same generation my build
  chain starts from. Its only extra sheet is a Claude Log. So his review-workbook edits
  transplant cleanly: rev becomes the new chain input.
- `base_2707.xlsx` is my latest shipped build plus his own hand edits (catalogued below).
  Those edits are adopted into the builders so a rebuild reproduces them instead of
  reverting them.
- `cust_new.xlsx` carries the new Customer dataset to load into REVIEW.

## His 2707 hand edits, adopted into the build

| Edit | Adoption |
|---|---|
| 3.1: section bar row deleted | 3.1 builds without the bar |
| 3.2 relabels: "Overhead & Leadership comparison", "Overheads incl. GMs", "Of which sits in the N-role ledger" (clause trimmed), plain-English allocation note | Builder constants |
| 2.x: section renamed "No archetype in 1.x tabs"; ELSEWHERE note trimmed to "...sit in the COEs" | Builder constants, applied on all fourteen tabs |
| 1.7/1.8/1.9: "Actuals" line in the budget box reading the tab's own actual total, plus "Variance to actuals" | Applied to all ten portfolio 1.x tabs, wired to the rebuilt actuals block by label |
| 1.x footers: residual row renamed "Additional costs" (kept only where non-zero), control row off the tab | Footer builder; the control moves to 4.0 where every other check lives |
| 2.7 levers: Stevani Kho (Delivery Lead) to Offshore; the vacant QA role to Filled | Re-applied by name after the rebuild |
| 1.2!I54 = 2.21, 1.8!E12 = 7.2 typed | Carried onto the assembled base |

## Waves

1. **Analysis (3 Opus agents, parallel)** - ledger delta, 1.x/0.2/Lists semantic spec,
   Customer dataset extraction. Fable catalogued the 2707 edits inline.
2. **Assembly & build (Fable, deterministic chain)** - new base from rev + 2707 edits;
   REVIEW loaded with the new Customer and cyber datasets, verified against source
   cell-by-cell; chain parametrised for the new ledger extent (every hardcoded 525/528/529
   is now measured); full rebuild; his levers re-applied; seven QA passes.
3. **Verification (3 Opus agents, parallel)** - independent QA recompute from the new
   ledger; design-consistency review of rendered tabs; instruction audit of T1-T7 plus the
   standing register. Findings come back to Fable, get fixed, and the wave reruns until
   all three agents return clean.
3a. **Fix wave F (3 Opus builders, parallel, disjoint script ownership)** - round one of
   wave 3 returned: the audit found three number-changing defects (fabricated Holds from
   `extend_lists` template-copy, DECISIONS D86; net-vs-gross budget basis on 1.11/1.12,
   D87; the Actuals column question, resolved as his own wiring, D88) plus dropped owner
   content traced to `finish.py`'s literal sweep (D89), the 0.2 EG-row clear (D90), the
   rev-inherited hidden Exec Summary (D91). The design review found theme-fill bars
   rendering black on all ten 1.x tabs, the 1.6 header hole (D92), 0.2 clipped rows, and
   a long family-consistency list. Builder A owns merge_review / repair_design / fix1x /
   fixcoe / finish / polish (instruction defects); Builder B owns final2x/3x/4x/35 /
   actuals + the QA readers (2.x/3.x family rules); Builder C owns the new
   `design2707.py`, inserted between polish and finish (visual sweep). Then chain2 +
   chainA2 rerun and all seven passes back to zero; the adversarial QA report folds into
   the same loop when it lands.
4. **Ship** - candidate to `TDD_Cost_Calc.xlsx`, docs updated, committed and pushed.
5. **Wave G (owner's post-ship round)** - his instructions verbatim: 0.3 was not to be
   changed (it is his cost library - the chain now treats it as a source tab and a parity
   check against rev enforces it); the 1.x bottom actuals/decisions table moves up top as
   a clean uniform table with zero row shifts; 3.2 states the Business Partner and Domain
   Architect role counts sitting in the COEs, adds one no-double-count row for total
   roles across COEs and portfolios, and replaces the opaque column L; and a new
   1.14 TDD Cyber portfolio-style design tab (platform TDD Cyber, one squad Cyber
   Uplift) is integrated end to end - 2.15 working tab, 3.1/3.3 rows, 4.0 ties - with
   an Opus investigator mapping the blast radius first (the ten-portfolio allocation
   counts must not inflate; EGI is the precedent for sitting outside them). Opus
   builders implement; the regression gate grows to cover all of it.

## Ground rules carried over

- REVIEW is the only source of truth; his raw columns load faithfully from his files.
- Join by label, never by row number.
- A variance needs two figures on one basis; dashes where nothing prices a row.
- No frozen panes, no "seat", cream inputs, plain English.
- Every figure a reader sees must survive a recomputation from the ledger.
