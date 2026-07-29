"""Say it the way the owner says it.

"Ledger" is not his word. It is mine, and it reached 87 labels - every control row on
every working tab, the Exec Summary's headline lines, half of 4.0 Data QA. He calls the
same thing the role mapping, and the tab is literally named "REVIEW - Complete Role
Mapping". A technology GM opening this file has a role mapping in front of them, not a
ledger.

Everything else the label scan turned up is his own vocabulary and stays: archetype is
his word and he ruled on it in as many words ("Do not call it design. It is the
archetype."); COE, EGI, BP&T, SA&D and Lights On are Ampol's; "Vacancy lever" is his.
Replacing those would be me overwriting his language with mine, which is the mistake this
file exists to undo.

Run last, on the finished workbook, so it catches every label whoever wrote it.
"""
import re
import sys

import openpyxl

# ordered: the longest phrase first, so a short rule cannot eat a long one
SAY = [
    ("the 531-role ledger", "the 531 roles in the role mapping"),
    ("531-role ledger", "531 roles in the role mapping"),
    ("Cost of the 531 roles in the ledger", "Cost of the 531 roles in the role mapping"),
    ("roles in the ledger carry", "roles in the role mapping carry"),
    ("Roles in the ledger", "Roles in the role mapping"),
    ("roles in the ledger", "roles in the role mapping"),
    ("outside the ledger", "outside the role mapping"),
    ("Above the ledger", "Above the role mapping"),
    ("above the ledger", "above the role mapping"),
    ("against the ledger", "against the role mapping"),
    ("in the ledger", "in the role mapping"),
    ("the ledger plus", "the role mapping plus"),
    ("the ledger", "the role mapping"),
    # "Budget to draw down" was on this list too, swapped for "Budget available". It should
    # not have been. "Ledger" is mine and he never wrote it; "Budget to draw down" is his,
    # it is in his own review workbook on 1.11, 1.12 and 1.13, and D83 settled that his
    # labels win. Rewording his own heading to one I preferred is the exact mistake this
    # file was written to undo, committed inside the file undoing it.
    # applied after the swap above: three labels that the longer wording pushes past
    # their column. The count they drop is stated on the same page either way.
    ("outside the 531 roles in the role mapping", "outside the role mapping"),
    ("above the 531 roles in the role mapping", "above the role mapping"),
    ("against every overhead role in the role mapping",
     "against overhead roles in the role mapping"),
]


# His exact wording, put back where a pass normalised it for no reason but tidiness.
# tab -> {cell: (what the build now says, what he wrote)}. The build's text is stated so a
# rename that lands somewhere else fails loudly instead of overwriting the wrong cell.
HIS = {
    "0.2 Data Config": {
        # his own abbreviation, on his own config tab. M5 four rows up says "Allocation %"
        # in his book too - the two headings differ because he wrote them that way.
        "M13": ("Allocation %", "Alloc %"),
    },
}


def restore(wb, out):
    for tab, cells in HIS.items():
        if tab not in wb.sheetnames:
            out.append(f"{tab} is not in the workbook - nothing restored")
            continue
        for cell, (now, his) in cells.items():
            cur = wb[tab][cell].value
            if cur == his:
                continue
            if cur != now:
                out.append(f"{tab}!{cell} holds {str(cur)[:40]!r}, expected {now!r} - "
                           f"left alone")
                continue
            wb[tab][cell] = his
            out.append(f"{tab}!{cell} {now!r} -> {his!r} - his wording")


def run(src, dst):
    wb = openpyxl.load_workbook(src)
    n = 0
    hit = []
    for ws in wb.worksheets:
        # the owner's own source tabs are never reworded
        if ws.title in ("0.1 Budget Table (Fin)", "0.4 Presentation Pack",
                        "0.3 Squad Archetypes"):
            continue
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not isinstance(v, str) or v.startswith("="):
                    continue
                new = v
                for old, rep in SAY:
                    if old in new:
                        new = new.replace(old, rep)
                if new != v:
                    c.value = new
                    n += 1
                    if len(hit) < 6:
                        hit.append(f"{ws.title}!{c.coordinate}: {new[:60]}")
    # and inside the formula-built labels, which are strings the reader still reads
    nf = 0
    for ws in wb.worksheets:
        if ws.title in ("0.1 Budget Table (Fin)", "0.4 Presentation Pack",
                        "0.3 Squad Archetypes"):
            continue
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not (isinstance(v, str) and v.startswith("=") and '"' in v):
                    continue
                new = v
                for old, rep in SAY:
                    if f'"{old}' in new or f'{old}"' in new or f' {old} ' in new:
                        new = new.replace(old, rep)
                if new != v:
                    c.value = new
                    nf += 1
    back = []
    restore(wb, back)
    wb.save(dst)
    return [f"{n} labels and {nf} formula-built labels now say role mapping, not ledger",
            *[f"  e.g. {h}" for h in hit], *back]


if __name__ == "__main__":
    for x in run(sys.argv[1], sys.argv[2]):
        print("  ", x)
