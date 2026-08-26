# Phase C build spec: the mapping deliverable (26/08)

Replaces the rejected v1 at deliverables/TDD_NonLabour_Mapping.xlsx. Built
only after QA-B reaches zero findings and the four rule files freeze.
Governing rulings: D168/D169 (deliverable shape), D172 (plan), D174
(evidence conventions), RAW DATA RULING (verbatim tabs), FORMATTING
FOREVER (no italics, no AI look), LANGUAGE RULING (banned words).

## Inputs (frozen at build start)
- The upload copy: scratchpad nonlabour/budget.xlsx (five tabs).
- Rule files: agents/phaseb_b1.json (issue 2), phaseb_b2.json (issue 3),
  phaseb_b3.json (issue 3, 13 contra entries), phaseb_b4.json (issue 3).
- Item keys: agents/items_sw_rows.jsonl, items_hw_rows.jsonl (Phase A,
  zero findings).
- Structure: docs/OP_MODEL_STRUCTURE.md (only allowed destinations).
- Dashboard controls: agents/dashboard_controls.json (sanity block only).

## Tab plan
1. Read me first. Plain English: what each tab is, how a charged business
   unit verifies one line from the named columns, what blue text means,
   what the open questions are worth. No jargon, no person references.
2. SW Line Items. Re-created from the upload cell by cell: same rows,
   same columns A..AA, same orientation, values as they are (his #N/A
   cells stay). Column A Portfolio and B Platform: his 784 typed rows
   untouched in his own formatting; every other leaf row filled from the
   rule files in BLUE text (hardcoding convention); parent roll-up rows
   stay blank. Added columns right of AA, never inside his range:
   AB Item (Phase A key), AC Rule (block and number, e.g. S2-79),
   AD Open question (blank when ruled). Header row matches his row-14
   style.
3. HW Line Items. Same treatment; A and B get headers matching the SW
   wording and every leaf mapping is blue (he typed none on HW).
4. By vendor and product. The Gate A mock's list layout on SUMIFS off
   the two tabs' Item columns: one line per single-destination item;
   split items expand one line per destination with dollars, share and
   the deciding logic in column terms. Kicker line above the table:
   what it is, units, basis. Totals row ties the grand total.
5. Mapping register. Every rule from the four files: block, rule number,
   the match in column terms, destination, lines, dollars, the evidence
   sentence. The dollars column is SUMIFS off the tabs' Rule column and
   must tie each stated figure to the cent. Then the open questions,
   priced, with their missing fact. Then the contra document listing
   (both legs, both destinations, document numbers) so transfers stay
   visible.
6. Totals. Portfolio by platform SUMIFS grid over both tabs, tab
   subtotals, and a tie row against the proven leaf totals
   (51,288,134.47 software, 25,468,998.32 hardware, 76,757,132.79
   together) showing 0.00 differences computed by formula.
7. Reconciliation. Per cost centre: SUMIFS off each tab against the
   centre's raw leaf total, difference column all 0.00. Beneath, the
   dashboard sanity block copied verbatim with the windows note (budget
   view against the export's actuals window; never cent-tied).
8. Decisions. The priced open questions as rows (Q1 device home
   6,752,768.57; Q2 prepaid references -2,196,046.41; Q3 bulk fuels
   split key -205,850.04; Q4 marketing residual 189,422.76; Q5 Apollo
   journals 4,012.94; Q6 brand centre cloud recharge 90,000.00; software
   opens APPLTDDIST 1,023,626.60, APPLFIFINANCE 35,107.27, b2's
   -147,174.37 over 123 lines, b3's 595,881.25 over 572 lines).
   Candidate treatments stated as text where a frozen file names one;
   NO dropdown controls (amended at build: in this workbook a picked
   ruling would drive nothing, because the line mappings are hardcoded
   blue cells, and a dead control reads as broken; rulings are given in
   conversation and applied by regeneration).

## Conventions
- No italics anywhere. No AI look: plain bold headers, sentence case,
  light borders, no colour fills beyond the blue hardcode text and red
  for over. Brackets for negatives. TBC where genuinely unknown, never
  a dash or a plug. No dashes as cell filler, no em or en dashes.
- Banned words enforced everywhere (saving in every form, wave, seat,
  floor, charge/charged/charging except recharge and Labour Recharge,
  roster, Category as a label, theirs/their).
- Neutral impersonal register; the 784 cited as mapping values present
  in the source file.
- Every summary number a formula off the raw tabs. A hardcoded summary
  number anywhere is a build failure.

## Build method
- openpyxl re-creation (not zip surgery: the tabs gain columns and blue
  fonts, so the sheets are rebuilt; fidelity is proven by value diff,
  not byte hash). Any formatting the rebuild cannot carry is recorded
  in the build report, never silently dropped.
- SUMIFS performance: if the engine recalc is slow over ~53k-row
  multi-criteria SUMIFS, fall back to one hidden single-key helper
  column per tab (key = portfolio|platform or item) and SUMIF over it.
  Helper columns sit right of the added columns and are labelled.
- Recalc gate (amended at build): a LibreOffice SAVE corrupts this SAP
  export's package, so the engine never touches the deliverable. The
  gate runs on a throwaway copy: LibreOffice recalculates the copy,
  every tie cell must read exactly 0.00 and the six built tabs must
  show zero formula errors; the deliverable itself carries
  fullCalcOnLoad so Excel computes on open. Inherited errors on the
  SAP tabs (dead references, add-in header cells) are pre-existing and
  out of scope.

## Gates
- C1 raw fidelity: every cell value on both line tabs identical to the
  upload (49,910 leaf rows plus parents and headers), row and column
  counts identical, #N/A cells preserved, the 784 untouched and not
  blue, every agent-filled mapping cell blue, parent rows blank.
- C2 formula integrity: zero hardcoded summary numbers; engine recalc
  reproduces every displayed value; all tie cells 0.00; register
  dollars tie each rule file figure to the cent.
- C3 language and render: register scan clean on every sheet; rendered
  screenshots of every tab eyeballed against the conventions; vendor
  view matches the Gate A mock's layout on real numbers.
- C4 orchestrator gate: independent recomputation of the grid totals
  from the rule files, final engine recalc, and a line-count and value
  sample diff against the upload, before anything ships.

## Loop
Builder agents per tab group (line tabs; vendor view + register; totals
+ reconciliation + decisions + read me), then QA agents run C1 to C3
fresh; builders fix only named findings; loop to zero; C4 is mine.
