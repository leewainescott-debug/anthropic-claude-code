"""What of his did not make it into the build.

He asked a fair question - "what else have i not requested that you've got rid of?" - and
the only honest way to answer it is mechanically, not from memory.

Method: take every non-empty cell he typed in his own workbooks, and ask whether that
content still appears anywhere on the same tab of the build. Not the same cell - rows
moved, tables were relocated, the 1.x actuals block went from the foot of the tab to the
top - so a cell-by-cell diff would report hundreds of false losses. Content survival is
the question that matters: is what he wrote still in the file.

Typed content only. Formulas are the model's plumbing and are rewritten by design; a
formula that changed is a decision, and the decisions are in DECISIONS.md. A label or a
number he typed that is simply gone is a loss, and that is what this prints.

    python3 whatsgone.py base_2707.xlsx cand_A.xlsx
    python3 whatsgone.py base_2707.xlsx cand_A.xlsx --later rev.xlsx

With --later, a second workbook of his that came after the first, the output splits in two:
what his own later book already replaced, which is not a loss, and what the build changed,
which is the list he actually wants.
"""
import re
import sys
from collections import defaultdict

import openpyxl

# tabs the chain builds from scratch: their labels are the model's, not his
BUILT = re.compile(r"^(2\.\d+ |3\.\d+ |4\.0|Exec |Lists$|REVIEW)")


def typed(ws, fws=None):
    """Every value he typed on this tab, as (coordinate, value).

    With fws - the same tab read formulas-first - a cell whose formula view is a formula
    is dropped: its value is something the model computed, not something he wrote, and it
    is meant to move when the data moves."""
    out = []
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if v is None:
                continue
            if fws is not None:
                f = fws[c.coordinate].value
                if isinstance(f, str) and f.startswith("="):
                    continue
            if isinstance(v, str):
                v = v.strip()
                if not v or v.startswith("="):
                    continue
            out.append((c.coordinate, v))
    return out


def norm(v):
    """Compare the way a reader would: case and spacing do not count, an en dash and a
    hyphen are the same character to anyone reading the page, and a number is its value."""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return round(float(v), 4)
    s = str(v).lower()
    for d in "‐‑‒–—―":
        s = s.replace(d, "-")
    return re.sub(r"\s+", " ", s).strip()


def run(his_path, build_path):
    a = openpyxl.load_workbook(his_path, data_only=True)
    af = openpyxl.load_workbook(his_path)
    b = openpyxl.load_workbook(build_path, data_only=True)
    gone = defaultdict(list)
    tabs_gone = []
    for t in a.sheetnames:
        if t not in b.sheetnames:
            tabs_gone.append(t)
            continue
        if BUILT.match(t):
            continue
        # the build's side keeps computed values: his typed 5.5 is not lost if the build
        # now computes 5.5 in the same place
        have = {norm(v) for _, v in typed(b[t])}
        for coord, v in typed(a[t], af[t]):
            if norm(v) not in have:
                gone[t].append((coord, v))
    return gone, tabs_gone


def later_has(later_path, build_path):
    """Every value typed in his later workbook, per tab, for the split below."""
    lat = openpyxl.load_workbook(later_path, data_only=True)
    bld = openpyxl.load_workbook(build_path, read_only=True)
    out = {}
    for t in lat.sheetnames:
        if t in bld.sheetnames:
            out[t] = {norm(v) for _, v in typed(lat[t])}
    return out


def report(gone, build, later=None, cap=40):
    """One line per distinct value, most-repeated first, with what the build says now."""
    b = openpyxl.load_workbook(build, data_only=True)
    seen = defaultdict(list)
    for t, items in gone.items():
        for coord, v in items:
            seen[str(v)].append((t, coord))
    mine, superseded = [], []
    for v, places in sorted(seen.items(), key=lambda kv: -len(kv[1])):
        # his later book still says it and the build does not -> the build changed it
        bucket = mine
        if later is not None and not any(norm(v) in later.get(t, ()) for t, _ in places):
            bucket = superseded
        bucket.append((v, places))
    return mine, superseded, b


if __name__ == "__main__":
    his, build = sys.argv[1], sys.argv[2]
    later = None
    if "--later" in sys.argv:
        later = later_has(sys.argv[sys.argv.index("--later") + 1], build)
    gone, tabs_gone = run(his, build)
    n = sum(len(v) for v in gone.values())
    print(f"{his} -> {build}")
    print(f"{n} typed values on his own tabs are not in the build, across {len(gone)} tabs")
    if tabs_gone:
        print(f"tabs of his that are not in the build: {tabs_gone}")
    cap = 40
    mine, superseded, b = report(gone, build, later)
    for title, group in (("the build changed it", mine),
                         ("his own later workbook replaced it - not a loss", superseded)):
        if later is None and title.startswith("his"):
            continue
        print(f"\n=== {title}: {sum(len(p) for _, p in group)} cells, "
              f"{len(group)} distinct ===")
        for v, places in group[:cap]:
            t, coord = places[0]
            print(f"{len(places):>3} x  {v[:62]:<62} | {t}!{coord} -> "
                  f"{str(b[t][coord].value)[:36]!r}")
        if len(group) > cap:
            print(f"     ... {len(group) - cap} more distinct")
