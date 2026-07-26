# What is actually wrong with the design

I rendered every tab to an image and looked at it. I had never done that. Every layout
decision I made was taken from reading cell values, which is why the numbers tied out and
the workbook still looked like this.

---

## 1. There was never one builder per family

This is the root cause of "how do the COE tabs look so different".

The tabs in a family were written by different scripts at different times, and each script
set its own widths, its own header, its own total row. Nothing ever normalised a family
as a whole.

| Family | Tabs | Distinct column-width profiles | Distinct title styles | Header row |
|---|---|---|---|---|
| 1.x | 14 | **14** - every tab different | 1 | row 5 |
| 2.x | 14 | **9** | **2** | row 4 on twelve, row 5 on two |
| 3.x | 5 | **5** - every tab different | **3** | row 5 on four, row 4 on one |

2.12 BP&T and 2.13 SA&D carry a 14pt blue title. The other twelve carry a 15pt black one.
That is not a formatting slip, it is two different scripts.

## 2. Colour is being used for nothing

| Family | Fill colours per tab | Should be |
|---|---|---|
| 1.x | 7 to 8 | 3 |
| 2.x | 4 to 6 | 3 |
| 3.x | 3 to 6 | 3 |

On 1.1 alone: navy headers, a cream input fill striping five columns of every squad table,
an orange box floating on the left for the budget reconciliation, yellow blobs mid-table,
blue text on some numbers and black on others, and red text on one label.

## 3. Notes are written into cells because there was nowhere else to put them

Three long sentences sit inside cells on every 2.x tab. Four to five on the COE design
tabs. On 3.2 the note is white italic text inside a navy box, which is the same mistake
as black text in a blue box.

## 4. Blank tables are styled empty rows

This is the "what the fuck are the blank tables below" question.

| Tab | Blank rows carrying fills or banding |
|---|---|
| 3.3 FTE View | 28 |
| 1.14 TDD Cyber | 18 |
| 3.1 Group Summary | 13 |
| 1.12 SA&D | 6 |
| 2.3 Enterprise Data | 6 |

They are rows that used to hold something. The content was removed, the styling was not.

## 5. Headers are truncated because the columns were never sized

On 2.1 the archetype type reads "nfiguration / Integrat" and "rchetype for this sq". The
flag column collides with the number to its left, so the sheet shows "11chetype after
decisions". On 3.1 the header reads "archetype ($m)" because "Variance to" is cut off, and
"decisions ($m)" because "Cost after" is cut off.

Nineteen columns were written onto tabs whose widths were set when there were fourteen.

## 6. Conditional formatting is decorating, not informing

3.3 carries seven rules. They paint a navy band with white text across the first row of
each portfolio, and salmon blocks down 28 empty rows below the table. 3.1 paints green on
one variance column and salmon on another. None of it means anything a reader could state.

## 7. Red means the opposite of what it should

The money format is `#,##0.00;[Red](#,##0.00)`. Variance is actual minus budget, so over
budget is positive and prints black, under budget is negative and prints red. The alarm
colour is on the good outcome.

---

---

## The format you want is already in the file

1.11 BP&T and 1.13 Cyber Roles are clean, and they are clean in the same way as each
other. Nothing else in the workbook follows them.

What they do:

- black bold title, nothing else on the row
- a navy bar naming the section: `Summary`, `Funding buckets to draw down`, `Roles`
- under each bar, a header row: navy fill, white bold, centred, every column sized to
  its contents so nothing truncates
- data rows with light borders, a shaded bold total row at the end
- one blank row between blocks, and blocks stack down the page in one column
- label / value pairs for the funding block, not a wide table
- yellow only on the one cell you type in
- the two notes sit at the very bottom, under everything, in plain black

That is a complete, consistent, readable pattern, and it is yours. The right move is to
make every tab in the workbook look like 1.11, rather than invent anything new.

Applied to a 2.x tab it would read: title, `Squad summary` bar, the squad table, a shaded
total, blank row, `Overhead roles` bar, that table, blank row, `Total portfolio`, blank
row, `FTE` bar, the people under a band per squad. Same bars, same header treatment, same
total row, on all fourteen.

## What this means for the rebuild

The fix is not tab-by-tab. It is:

1. **One builder per family.** All fourteen 2.x tabs come out of one function, so they
   cannot differ. Same for 1.x, same for 3.x.
2. **A fixed grid.** Column count, widths and headers defined once per family and applied
   to every member, COE tabs included.
3. **Three fills, no more.** Header, group band, total. Yellow only on an input.
4. **No sentences in cells.** If something needs explaining, the header says it.
5. **Nothing empty is styled.** A rebuilt tab is wiped, then written, so no orphan
   banding can survive.
6. **No colour carries judgement.** Negatives in brackets, never red.
