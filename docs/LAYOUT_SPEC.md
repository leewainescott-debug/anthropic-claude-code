# Layout spec - the 2.x and 3.x tabs

What each tab says, column by column. Written before the build, not after.

Rules that apply to every tab here:

- One fact lives on one tab. If 3.2 states archetype cost, 3.1 does not.
- Every column earns its place. No working columns, no flags, no lookup helpers.
- No sentences inside cells. A note goes at the foot of the block, in plain black.
- No colour carries judgement. Negatives are in brackets, never red.
- Yellow means you can type in it. Nothing else is yellow.
- Calibri throughout, matching the owner's own tabs. Nothing under 10pt.

---

## 2.1 to 2.14 - working copies

The tab where decisions get made. One per portfolio.

**Fact bar** (row 3): Portfolio, Roles, Vacant, Cost today. Four numbers, so the GM
knows where they are before reading anything.

**Squad summary**

| Col | Header | What it is |
|---|---|---|
| B | Squad | Squad name from REVIEW |
| C | Archetype type | From the portfolio's design tab |
| D | Archetype size | Its own column, not joined to type |
| E | Archetype roles | From 0.3 |
| F | Roles | People in this squad today |
| G | Vacant | Of which vacant |
| H | To hire or offshore | Levers set to Hire or Offshore |
| I | On hold | Levers set to Hold |
| J | Roles after decisions | F minus I |
| K | Archetype cost ($m) | From 0.3, offshore priced from column H |
| L | Actual cost ($m) | What these people cost today |
| M | Variance to archetype ($m) | L minus K |
| N | Impact of decisions ($m) | O minus L. Negative is a saving |
| O | Total cost after decisions ($m) | Sum of each person's cost after their lever |

Blocks in order: delivery squads, subtotal, overhead roles, subtotal, total portfolio.
The COEs and EGI have no overhead block at all.

Two control lines under the table, both must read 0.

**FTE, grouped by squad**

The squad name is a band across the row and carries that squad's two totals, so there
is no separate subtotal row to read past.

| Col | Header |
|---|---|
| B | Name |
| C | Role |
| D | Status |
| E | Vacancy lever (yellow, dropdown: Filled, Hire, Hold, Offshore) |
| F | Cost if hired ($) |
| G | Cost after decision ($) |

---

## 3.1 Group Summary - the budget story

One line per portfolio. Budget, what it costs, what is left to fund. Nothing else.

| Col | Header |
|---|---|
| B | Portfolio |
| C | TDD lights-on budget ($m) |
| D | Actual cost ($m) |
| E | Over/(under) budget ($m) |
| F | Total cost after decisions ($m) |
| G | Left to fund after decisions ($m) |

The budget is the lights-on allocation from 0.2. It is named that way in the header
because it is not a total people budget, and comparing it to one without saying so is
how the old tab produced a variance nobody could explain.

## 3.2 Total Cost - the archetype story

| Col | Header |
|---|---|
| B | Portfolio |
| C | Archetype cost ($m) |
| D | Actual cost ($m) |
| E | Variance to archetype ($m) |
| F | Impact of decisions ($m) |
| G | Total cost after decisions ($m) |
| H | New variance to archetype ($m) |

Second block underneath, clearly headed: overhead allowance against what overhead
actually costs, line by line. Overhead is not priced by a squad archetype, so it is
stated separately rather than mixed into the portfolio rows.

## 3.3 Squad Detail - squad by squad, roles and cost

| Col | Header |
|---|---|
| B | Portfolio |
| C | Squad |
| D | Archetype type |
| E | Archetype size |
| F | Archetype roles |
| G | Roles |
| H | Filled |
| I | Vacant |
| J | Archetype cost ($m) |
| K | Actual cost ($m) |
| L | Variance to archetype ($m) |
| M | Total cost after decisions ($m) |

Every squad on every working tab appears once, under its portfolio, with a portfolio
total and a group total.

---

## Housekeeping in the same build

- 1.14 TDD Cyber is deleted. Cyber appears once, on 1.13, per the brief.
- Tab order follows the flow: inputs, 1.x in number order, 2.x, 3.x, evidence.
  1.10 sorts after 1.9, not between 1.1 and 1.2.
- Every explanatory sentence sitting in a cell is removed.
- Red text is removed everywhere it was used as commentary.
