# Lee's deck standard - Ampol FY27 Budget Update (FINALISED 18/08)

THE reference for every slide from here on: language, structure, formatting.
Source file kept in the repo at `reference/Ampol_FY27_Budget_Update_FINAL.pptx`
(Lee's finalised upload, untouched). When building any slide, match this.
Never invent a look.

---

## 1. Canvas, fonts, colours

**Canvas** 13.33in x 7.5in (16:9). 338 slide layouts in the package, 1 master,
think-cell present.

**Fonts** - Mark OT family only:
| Use | Font |
|---|---|
| Theme major + minor | Mark OT Book |
| Body / most text | MarkOT |
| Bold emphasis | MarkOTBold |
| Tables and charts fall back to | Arial |

**Sizes actually used** (by volume): 9pt is the workhorse (table text, dense
body), then 10pt, 12pt (insight rail bullets, kickers), 8pt (footnotes),
14-16pt (rail headers, section labels), 32pt (action titles), 36-48pt (title
slide). Body text never gets bigger than the title.

**Theme colours** - use these hex values, nothing else:
| Token | Hex | Where it is used |
|---|---|---|
| accent1 navy | `18249C` | table header rows, kickers, bars, side panels, block diagrams |
| accent2 red | `ED0C06` | over/(under) numbers, numbered callout circles, cumulative line, Ampol logo |
| accent3 grey | `5F5F5F` | secondary text |
| accent4 | `E0E0E0` | rules, light fills |
| accent5 coral | `F76D69` | dots / markers on diagrams |
| accent6 periwinkle | `747BC3` | "Draft for discussion" flag |
| lilac | `C6CAF6` | table row fill, block diagram inner fills |
| light greys | `F2F2F2`, `E3E3E3`, `E3E5F0` | panel fills, banded rows, total rows |
| blacks | `000000`, `151515`, `1A1A1A` | body text, black total rows |
| white | `FFFFFF` | header text, text on navy |

Red is rationed - it only marks a number that is over, a callout circle, or
the cumulative line. Never decorative.

---

## 2. The content slide archetype (slides 5, 6, 7, 8)

This is the workhorse. Build every data slide this way.

```
┌─────────────────────────────────────────────────┬──────────────┐
│ Action title, full width, 32pt, 1-2 lines       │  [7]  flag   │
├─────────────────────────────────────────────────┼──────────────┤
│ navy kicker line (optional)                     │ Insight &    │
│ ┌─────────────────────────────────────────────┐ │ Assumptions  │
│ │ TABLE - navy header, banded body, total row │ │              │
│ └─────────────────────────────────────────────┘ │ • bullet     │
│ navy kicker line                                │ • bullet     │
│ ┌─────────────────────────────────────────────┐ │ • bullet     │
│ │ TABLE or CHART                              │ │              │
│ └─────────────────────────────────────────────┘ │              │
├─────────────────────────────────────────────────┴──────────────┤
│ [logo]  1 footnote   2 footnote   3 footnote                   │
└────────────────────────────────────────────────────────────────┘
```

**Measured geometry (inches), copy these:**
| Element | x | y | w | h |
|---|---|---|---|---|
| Action title | 0.48 | 0.47 | 12.36 | 0.38-0.49 |
| Left content (tables/charts) | 0.41-0.48 | 1.34-1.76 | 8.67-9.20 | varies |
| Kicker above a table | 0.41 | 1.40 / 4.16 | 6.69-7.99 | 0.29-0.47 |
| Insight rail | 9.61-9.76 | 1.24-1.40 | 3.26-3.36 | 4.27-5.49 |
| Footnote strip | 1.10 | 6.64 | 12.37 | 0.43 |

Left content ~69% of the width, insight rail ~25%. Slide number top right.
Ampol logo bottom left on content slides.

**Action title** - a full sentence carrying the so-what and the number,
sentence case, black, left aligned, wraps to 2 lines. Never a label.
His actual titles:
- "TDD is currently tracking 320k over the TDD lights on budget with vacancy and offshore controls available to bring TDD back within budget"
- "FY26 forecasts TDD to land $400k above the $53.8m allocated TDD Lights on budget; assuming all vacancies are hired as expected"
- "Vacancy controls will be used to determine when and how vacancies are recruited in the new operating model"
- "82 roles and 9 squads have been identified for offshoring, providing an estimated $8.4m cost reduction to the overall TDD people cost"

**Kicker** - navy bold, sits directly above the table it describes, explains
what the table is in one line:
- "GROSS, $m AUD: the model price of everyone on board, ramping with the hire months"
- "TDD LIGHTS ON, $m AUD: what lands on lights on. Budget = the 53.8 allocation split; gross budgets are AU 62.59 AUD and NZ 21.5 NZD at 0.92 FX"
- "Nine squads are marked Offshore or Hybrid"

**Insight rail** - header "Insight & Assumptions" (or just "Insight") bold,
then 3-5 bullets, 12pt, each a complete thought explaining a number or an
assumption behind it. This is where assumptions live, never in the title.

**Footnotes** - numbered 1, 2, 3 at 8pt along the bottom, tied to superscript
numbers in the content above.

**"Draft for discussion"** - periwinkle box, top right, on slides not yet final.

---

## 3. Tables

- Header row: navy `18249C` fill, white bold text, centred, wraps.
- Body rows: white or `F2F2F2` banding; lilac `C6CAF6` where a block needs to
  read as one group; first column often gets its own tint.
- Total row: black fill with white bold text, or `E3E3E3` grey with bold text.
  Always present, always labelled ("Total", "All roles", "All open vacancies",
  "Nine squads").
- Numbers centred, 2dp for $m, negatives in brackets `(2.98)`, over-budget
  figures in red.
- Row labels left aligned, plain.
- "TBC" where a number is genuinely not known - never a dash, never a guess.

## 4. Charts

Native and editable (think-cell in his file) - never a picture of a chart.

- **Waterfall** for bridging a cost position: labelled steps
  (Full price → Hold → Offshore → Other → After levers), value label on every
  step, navy bars, y-axis in $m.
- **Combo column + line** for phasing: navy columns for the current year,
  light blue for the next, dashed vertical divider between the two years, red
  cumulative line with a data label on every point, twin axes, manual legend
  swatches underneath ("FY26 starts", "FY27 starts"), and a one-line caption
  under the chart explaining what the reader is looking at.
- Direct data labels beat legends. Axis furniture stays minimal.

## 5. The other slide types

**Title slide** - full-bleed darkened photograph, AMPOL GROUP logo top left in
white, title white lower left over 2 lines. No date, no author, no subtitle.

**Executive summary** - three columns. Each column gets a navy header block
containing a large numeral (1, 2, 3) and the column title in white; body sits
below in a light grey `F2F2F2` panel as bullets with bold lead-ins
("**FY27 Impact** – ...", "**Labour Cost Allocation in platform model** – ...").
Column titles: "TDD Operating Model", "Operating Model Implications",
"Operating Model Cost with Budget Guidance".

**From / To comparison** - left half white headed "From: ...", right half on a
grey panel headed "To: ...". Stacked navy/lilac blocks with navy outlines,
each stack captioned underneath. White annotation boxes with drop shadows
carry the "what changed" notes. Coral dots mark the blocks that move.

**Section / org slide** - full-height navy panel down the left ~30% carrying
the title and narrative in white; content tables on white to the right;
AMPOL GROUP logo top right. Tables banded lilac/white with a soft drop shadow,
labelled underneath ("SUMMARY", "BY GM").

## 6. Speaker notes

Notes carry the full spoken narrative, several paragraphs, written as he would
say it - the conclusion first, then the evidence, then what it means. Slide 4's
note runs to a dozen paragraphs with a "Conclusion:" at the end. Build notes
this way, not as bullet fragments.

---

## 7. Language and voice

**Sentence case everywhere.** Not Title Case, not caps (except deliberate
labels like SUMMARY, BY GM, GROSS).

**His vocabulary** - use these exact terms:
lights on · landing · gross · significant items (sig items) · overheads ·
vacancy controls · right shore · offshore / hybrid · spans and layers ·
management density · span of control · layers · portfolio / platform / squad ·
FTE · Op-Model · cost avoidance · cost reduction · recharge · allocation ·
delta to budget · over/(under) · archetype · TBC

**Numbers** - "$54.2m vs a budget of $53.8m", "$400k above", "24% to 16%",
"4.2 to 6.3", "from 420 FTE to 517 FTE". Always paired with what they are
measured against.

**Banned, still banned** (his standing rulings, and the finalised deck obeys
them): saving / savings / saves / saved - it is a **cost reduction** or
**cost avoidance**, never a saving. Also banned: wave, seat, floor, roster,
Category as a column label, em dashes and en dashes as sentence punctuation,
dashes as cell filler, "theirs/their" for Finance.

**Note on "charged to"** - the earlier ruling banned charge/charged. His
finalised deck uses it freely and deliberately: "directly charged to strategic
programs", "Total Overheads charged to TDD", "Costs will be recharged to the
BU", "All Portfolio initiative costs charged to business". The finalised deck
wins - "charged to" and "recharged to" are his words and are fine. Flagged to
him 18/08.
