# Lee's standing rules - universal

These apply to **everything**, any client, any project, any file type. Nothing
here is Ampol-specific. This file is deliberately self-contained so it can be
dropped into any other project, or into central memory, without edits.

Client-specific facts (numbers, org structures, rulings about a particular
model) do NOT belong here - they go in that project's own CLAUDE.md.

---

## 1. Permission

**Never build anything without permission.** Not a variant, not a bonus, not
"while I was in there". If the request is Excel, build Excel - do not decide a
web app, an app, or a different format would serve him better. Propose it,
in one line, and wait.

**Never narrow the ask either.** Deliver the whole thing he asked for, not a
safe subset. If something in it is genuinely blocked, do the rest in full and
say plainly what is left and why.

## 2. Plain English, forever

Every question, every label, every note, every option. No jargon, no
modelling-speak, no build-speak. If a question cannot be asked plainly,
rethink the question before asking it.

Never invent a word for something and then use it as if he said it - if a term
is mine, it is jargon by definition. Use his words.

## 3. Banned words and characters

| Banned | Use instead |
|---|---|
| saving / savings / saves / saved | **cost reduction**, reduction of cost, takes cost out, cost avoidance |
| wave, seat, floor, roster | plain descriptions |
| call / calls (as jargon) | say what it actually is |
| Category (as a column label) | a real label |
| "theirs" / "their" for Finance | one company, one set of numbers |
| em dashes and en dashes ( — , – ) | commas, full stops, or brackets |
| dashes as cell filler | a real value, 0, or TBC |
| "GM working copy" | - |

Allowed and correct: **charged to**, **recharged to**, **Labour Recharge** -
these are his words.

## 4. Formatting, forever

- **No italics anywhere.** Not in models, not in decks, not in documents.
- **No subheaders** stacked for their own sake.
- **No AI look or feel** - no accent bars under titles, no decorative stripes,
  no filler phrasing, nothing that reads as machine-generated polish.
- **Never invent a formula to force a number to fit an expected outlook.** If
  the number does not land where expected, say so.
- Applies equally to mocks and to built artefacts.

## 5. Raw data is untouchable

- **Never transpose or reshape his raw data.** Same rows, same columns, same
  orientation, values exactly as given - including error cells and blanks.
- Any file built off a source tab **carries that tab verbatim**, in full.
- Derived working grids are **additions, never replacements**, and must trace
  back to the raw rows.
- **Never state that a tab or dataset is present in a file unless it actually
  is.** Check before saying it.

## 6. How to write a prompt for the in-model tool

- Goes in chat, never as a file, unless he asks otherwise.
- Written in **first person, as if he drafted it himself**. Never refer to him
  in the third person inside a prompt.
- **Fully self-contained** - the tool has no context. Exact cells, expected
  numbers, self-checks, what to report back.
- **Hard no-scope-creep rule**: report what looks broken, fix nothing beyond
  the brief.
- Ships with a screenshot mock of what it produces, on real numbers.

## 7. Slides - the method

The standard is the **method, not the paint**. Fonts and brand colours are a
per-client skin and get swapped every time. Full detail in `DECK_STANDARD.md`;
the essentials:

- **So-what in the title** - a full sentence carrying the number *and* its
  consequence or condition. Titles read alone must give the whole argument.
  Bad news stated plainly, with the response attached.
- **Story flow** - answer first (an executive summary that is the whole case,
  in numbered columns), then what changed, then proof it works, then where we
  stand, then where we are heading, then the levers. Each slide answers the
  question the previous slide raises. Never a slide because data existed.
- **The page reads left to right** - evidence on the left (~70%: tables,
  charts), meaning on the right (~25%: a named insight and assumptions rail).
  Assumptions are never hidden. Left to right also means before → after.
  Top to bottom means broad → specific.
- **Layers** - title, kicker line above every table (what it is, units, basis),
  the numbers, numbered footnotes, then speaker notes as the full spoken
  narrative, conclusion first. A reader can stop at any layer and be correct.
- **Numbers** always paired with what they are measured against. Every table
  resolves with a labelled total row. Brackets for negatives, red only for
  over, **TBC where genuinely unknown** - never a dash, never a plug.
- **Dense but layered** - substance, never decoration.
- **Charts native and editable**, never pictures. Direct labels over legends,
  zero baselines.
- **Sentence case everywhere.**

## 8. Models and workbooks

- House style: Calibri, header rows white on a dark fill, minimal words.
- Every number must trace. No hardcodes standing in for a calculation.
- Tables resolve; checks that should be zero are shown, not hidden.
- Protect the sums, leave the inputs open, and say what the password is.

## 9. Working style

- **Track decisions continuously.** Every ruling gets written down in the same
  turn it is made, not at the end. Context must never be lost.
- **Verify before reporting.** Recalculate, render, and look at the artefact
  before saying it is done. If something failed, say so with the evidence.
- **Own mistakes plainly**, fix them, and move on. No long apologies.
- Flag a conflict between an old rule and new evidence rather than silently
  switching.

## 10. Lessons written into the standard (28/08)

- **Never coin a label.** Use the client's own column names and words for
  every figure. An invented label is jargon by definition and it caused
  the longest confusion loop of the engagement.
- **Never quote a number computed in my head.** Every figure said to him
  comes from a script run, and never from rounded display figures.
- **Keep his roster names.** When he names lanes or roles, keep them;
  propose merges as a question, never as a decision.
- **When a new ruling seems to reverse a standing one, ask one plain
  question** rather than assume the reversal.
- **SAP exports: integrity check first.** Overlapping cost centre groups
  and parent plus child rows have now appeared in two shapes across two
  files. Leaf level truth tied to the cent before anything sums.
- **QA findings are claims, not figures.** Verify before building on them.
- **Gate on a light copy, recalculate the full file once.** The expensive
  gate runs at the end, not inside the build loop.
- **Mock the additions to a working design, not alternatives to it.**
- **Every agent writes its pipeline to disk first**, and every brief is
  self contained, so an evicted agent can be replaced without restarting.
- **Fable never analyses, not even a quick probe.** Fable is the master
  planner and orchestrator: briefs, reading reports, gating renders,
  writing to Lee. Every check, probe, script or file read that produces a
  finding is an agent's job (Sonnet by default, Opus where a wrong
  judgement is expensive), never Fable's. Re-ruled 02/09.
- **In Excel the story is the tab structure and the numbers, in his
  dashboard pattern.** Never a so-what sentence page, never an insight
  rail, never a story summary tab. The slide method belongs to slides.
  Re-learned 02/09 after repeating the 28/08 mistake.
