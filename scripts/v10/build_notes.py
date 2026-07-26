"""Two notes still pointed at column AB.

The agreed cost override started life in column AB. AB is "MyHR ee no" and holds 27
employee numbers, which the cost formula read as dollars and which dropped $1.87m out of
the model. The override moved to column AU and the formula followed it, but the two notes
that tell a reader where to find it did not. A note that sends someone to the wrong column
is worse than no note, because AB contains real-looking numbers.
"""
import openpyxl
import wbio

OVERRIDE_COL = "AU"


def notes(wb):
    ws = wb["REVIEW - Complete Role Mapping"]
    ws["AV172"] = (f"Banded rate agreed for this role; its own components give $321,135. "
                   f"Clear {OVERRIDE_COL}172 to price it from the components.")
    wb["Lists"]["AF16"] = (
        "Full Cost AUD = base x (1 + STI + payroll + pension + CPI) + medical, "
        f"or day rate x days x (1 + CPI). Column {OVERRIDE_COL} holds any agreed override.")
    return [f"REVIEW AV172 and Lists AF16 now point at column {OVERRIDE_COL}, "
            f"where the agreed cost override actually lives (they said AB, which is "
            f"MyHR ee no)"]


def run(src, dst):
    wb = openpyxl.load_workbook(src)
    out = notes(wb)
    wb.save(dst)
    return out


if __name__ == "__main__":
    import sys
    src, mid, dst = sys.argv[1], sys.argv[2], sys.argv[3]
    for x in run(src, mid):
        print("  ", x)
    rc, st = wbio.build(mid, dst)
    print("  injected", st)
