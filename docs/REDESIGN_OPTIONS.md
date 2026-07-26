# Redesign options

Four per section. Nothing is built. Pick one per section, or tell me none of them are right.

---

## 1. The 2.x working tabs

Where you make decisions. One per portfolio.

### Option 1A - Two tables: squads on top, people grouped below

```
SQUAD SUMMARY
Squad                 Archetype type   Size  ARoles  Roles  Vac  Hire  Hold  After   Arch$  Actual$   Var$  Impact$  Total$
Data Platforms        Eng / Integ       M      7.5      9    6     6     0      9    1.70     1.79   0.09     0.00    1.79
Data Science          Product           S      4.0      4    0     0     0      4    0.90     0.81  -0.09     0.00    0.81
Reporting & Analytics Eng / Integ       M      7.5     12    3     3     0     12    1.70     2.76   1.06     0.00    2.76
  Delivery squads                      19.0     25    9     9     0     25    4.30     5.36   1.06     0.00    5.36

  Overhead roles
Head of Technology    Overhead            -       -      1    0     0     0      1       -     0.41      -     0.00    0.41
Technology Manager    Overhead            -       -      2    0     0     0      2       -     0.65      -     0.00    0.65
  Overhead total                            -      3    0     0     0      3       -     1.06      -     0.00    1.06

  TOTAL PORTFOLIO                        19.0     28    9     9     0     28    4.30     6.42   1.06     0.00    6.42


ENTERPRISE DATA FTE
Name                  Role                     Status   Lever      Cost if hired   After decision
Data Platforms                                            9 roles      1,786,160        1,786,160
  Dinesh Chandra..    Lead Engineer            Filled  [Filled]          243,985          243,985
  Vacant              Engineer Data            Vacant  [Hire]            202,853          202,853
Data Science                                              4 roles        806,397          806,397
  Carol Kuang         Senior Data Scientist    Filled  [Filled]          184,871          184,871
```

The squad band carries that squad's totals, so there is no separate subtotal row to read past.

### Option 1B - One table, people indented under their squad, collapsible

```
   Squad / person        Role                Status  Lever     Roles  Vac   Actual$    Total$
-  Data Platforms        Eng / Integ  M                            9    6  1,786,160 1,786,160
     Dinesh Chandra..    Lead Engineer       Filled  [Filled]                243,985   243,985
     Vacant              Engineer Data       Vacant  [Hire]                  202,853   202,853
+  Data Science          Product  S                                4    0    806,397   806,397
+  Reporting & Analytics Eng / Integ  M                           12    3  2,764,812 2,764,812
   Delivery squads                                                25    9  5,357,369 5,357,369
```

One continuous table. The `+` and `-` are Excel's outline buttons: collapse every squad to
see only the summary, expand the one you care about. No scrolling between two tables.

Trade-off: the archetype columns and the person columns share the same grid, so some
columns are blank on person rows.

### Option 1C - Squad summary here, all people on one separate tab

The working tab holds only the squad summary, about 20 rows, fits on a screen. Every
person in TDD sits on one "FTE" tab you filter by portfolio.

Trade-off: the lever is no longer on the same tab as the squad it moves.

### Option 1D - Side by side: squad summary left, its people right

```
Squad                Roles  Act$  Total$  |  Name              Role            Status  Lever
Data Platforms           9  1.79    1.79  |  Dinesh Chandra..  Lead Engineer   Filled  [Filled]
Data Science             4  0.81    0.81  |  Vacant            Engineer Data   Vacant  [Hire]
Reporting & Analytics   12  2.76    2.76  |  ...
```

The right block shows the people for whichever squad you click.

Trade-off: needs a helper cell to drive the selection.

---

## 2. 3.1 Group Summary

### Option 2A - Budget, cost, left to fund. One line per portfolio

```
Portfolio            Lights-on budget$  Actual$   Over/(under)$  After decisions$  Left to fund$
Ampol Retail                     5.50    14.01           8.51            14.01           8.51
Customer                         6.50    17.13          10.63            17.13          10.63
...
Total                           50.50   115.11          64.61           115.11          64.61
```

Simplest. Six columns.

### Option 2B - Funding bridge

```
                                            $m
TDD lights-on budget                      50.50
Cost of the organisation today           115.11
  Over budget                             64.61
Impact of the vacancy decisions            0.00
Cost after decisions                     115.11
  Still to fund                           64.61
```

A vertical walk instead of a table. Reads as a story, but you lose the per-portfolio view.

### Option 2C - Two blocks: portfolios, then COEs

The ten portfolios with their own subtotal, then the COEs and EGI with theirs, because
they are funded differently: portfolios draw a lights-on allocation, the COEs draw a
named COE budget.

### Option 2D - Budget split AU and NZ

```
Portfolio          AU budget$  NZ budget$  Total budget$  Actual$  Over/(under)$  Left to fund$
```

The budgets are held AU and NZ separately on 0.2, and your register asked for both
variances. This surfaces that split rather than hiding it in one number.

---

## 3. 3.2 Total Cost

### Option 3A - Archetype against actual, overhead in a second table

```
Portfolio            Arch$   Actual$   Var$   Impact$   Total$   New var$
Ampol Retail          9.88     14.01    4.13      0.00    14.01      4.13
...
Total                64.20    115.11   50.91      0.00   115.11     50.91

OVERHEAD - ALLOWANCE AGAINST ACTUAL
Line                  Roles   Rate$  Units  Allowance$  Actual$  Over/(under)$
Head of Technology       15   0.138     10        1.38     4.87          3.49
...
```

### Option 3B - Overhead as columns on the same row

One table, wider: each portfolio row carries its squad cost, its overhead cost and its
total, so you never leave the row to see the whole portfolio.

### Option 3C - Three column groups: Design | Today | After decisions

```
                    ---- DESIGN ----   ---- TODAY ----   -- AFTER DECISIONS --
Portfolio           Roles     Cost$    Roles    Cost$      Roles       Cost$
Ampol Retail         47.5      9.88       70    14.01         70       14.01
```

Roles and cost sit together under each heading, so the comparison is like for like on
both dimensions at once.

### Option 3D - Cost bridge, group level

```
                                          $m
Squad archetype cost                    64.20
Delivery squads over the archetype      39.26
Overhead roles                          11.65
Cost of the organisation today         115.11
Impact of decisions                      0.00
Cost after decisions                   115.11
```

Explains the whole gap in six lines. Loses the per-portfolio detail, which would move
to 3.3.

---

## 4. 3.3 Squad Detail

### Option 4A - Flat list, grouped under portfolio with subtotals

```
Portfolio        Squad                 Type        Size  ARoles  Roles  Filled  Vac   Arch$  Actual$   Var$  Total$
Enterprise Data  Data Platforms        Eng/Integ    M      7.5      9       3    6    1.70     1.79   0.09    1.79
Enterprise Data  Data Science          Product      S      4.0      4       4    0    0.90     0.81  -0.09    0.81
Enterprise Data total                                     19.0     28      19    9    4.30     6.42   1.06    6.42
```

### Option 4B - Same, but portfolios collapse

Excel outline grouping, so you can shut every portfolio and see just the fourteen
subtotals, then open the one you want.

### Option 4C - Two tables: delivery squads, then overhead lines

Delivery squads compared to the archetype. Overhead lines compared to the allowance.
Two different comparisons, so two tables rather than dashes down the archetype columns
on every overhead row.

### Option 4D - Roles block and cost block side by side

```
                              ------ ROLES ------      ------ COST ($m) ------
Portfolio  Squad              Arch  Act  Fill  Vac     Arch   Actual   Var   Total
```

Same rows, but the two things you compare are visually separated.

---

## 5. Exec Summary

### Option 5A - Four blocks and a drill-down
The organisation today, against the archetype, against the budget, the vacancy decision.
Then a yellow dropdown to pick a portfolio and see its numbers.

### Option 5B - One page of numbered lines
The story in eight numbered lines, no blocks, CTO-facing. Everything else lives on 3.x.

### Option 5C - Three questions
What does it cost, what should it cost, what can I change. Each answered in four lines.

### Option 5D - Dashboard
The headline numbers as large tiles across the top, the portfolio table underneath.

---

## 6. Workbook structure

### Option 6A - Flow order, separators kept
Exec, inputs, 1.x, 2.x, 3.x, evidence. Separator tabs between the groups.

### Option 6B - Flow order, no separators, coloured tab groups
Same order, but the group is shown by tab colour rather than by a blank sheet.

### Option 6C - Flow order, retired sheets moved to the end and hidden
`Squads`, `Added data`, `Sheet2`, the two superseded sheets and `Portfolios` are hidden
at the end. Nothing is deleted.

### Option 6D - Flow order, retired sheets deleted
The model no longer reads them. Deleting them takes the file from 49 sheets to 42.

---

## 7. Formatting

### Option 7A - Keep your current look
Navy header, white bold text, Calibri 11, pale blue group bands, grey subtotals.

### Option 7B - Lighter
No fills on headers, just bold with a rule underneath. Less ink, more white space.

### Option 7C - Banded rows
Navy header, alternating white and very light grey rows for long tables.

### Option 7D - Print-ready
Everything in 7A, plus set print areas, repeat header rows on every page and fit each
tab to one page wide.
