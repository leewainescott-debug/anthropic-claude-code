"""Adversarial QA. Assumes the workbook is wrong and tries to prove it.

An error count only catches #REF!/#VALUE!. It does not catch a formula that points at a
cell the rebuild emptied and quietly returns 0, which is exactly what the Exec Summary
was doing. These checks look for wrong answers that Excel reports as valid.
"""
import re, collections
import openpyxl

ERRS = ("#REF!", "#N/A", "#VALUE!", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!", "Err:")
REVIEW = "REVIEW - Complete Role Mapping"
SHEET_RE = re.compile(r"(?:'([^']+)'|([A-Za-z0-9_.&\- ]+))!\$?([A-Z]{1,3})\$?(\d+)")


def load(path):
    return (openpyxl.load_workbook(path, data_only=False),
            openpyxl.load_workbook(path, data_only=True))


def refs(formula):
    """Every explicit sheet!cell reference in a formula (ranges give their anchors)."""
    out = []
    for m in SHEET_RE.finditer(formula):
        name = m.group(1) or m.group(2)
        out.append((name.strip(), f"{m.group(3)}{m.group(4)}"))
    return out


def check_dangling(wf, wv):
    """A formula referencing a cell that is empty in BOTH formula and value views."""
    bad = []
    names = set(wf.sheetnames)
    for sn in wf.sheetnames:
        for row in wf[sn].iter_rows():
            for c in row:
                if not (isinstance(c.value, str) and c.value.startswith("=")):
                    continue
                if "SUMIFS" in c.value or "COUNTIFS" in c.value or "SUMIF(" in c.value \
                        or "COUNTIF(" in c.value or "INDEX" in c.value or "MATCH" in c.value:
                    continue          # criteria ranges legitimately span blank rows
                if c.value.startswith("=N(") or "IF(N(" in c.value:
                    continue          # an empty target is handled explicitly here
                for tgt, coord in refs(c.value):
                    if tgt not in names:
                        bad.append((sn, c.coordinate, f"unknown sheet {tgt!r}", c.value[:60]))
                        continue
                    tf, tv = wf[tgt][coord].value, wv[tgt][coord].value
                    if tf is None and tv is None:
                        bad.append((sn, c.coordinate, f"points at empty {tgt}!{coord}",
                                    c.value[:60]))
    return bad


def check_silent_zero(wf, wv):
    """A formula whose value is 0 and whose every direct reference is an empty cell."""
    bad = []
    names = set(wf.sheetnames)
    for sn in wf.sheetnames:
        for row in wf[sn].iter_rows():
            for c in row:
                f = c.value
                if not (isinstance(f, str) and f.startswith("=")):
                    continue
                v = wv[sn][c.coordinate].value
                if v not in (0, 0.0):
                    continue
                if f.startswith("=N(") or "IF(N(" in f:
                    continue          # an empty target is handled explicitly here
                rr = [(t, k) for t, k in refs(f) if t in names]
                if not rr:
                    continue
                if all(wf[t][k].value is None and wv[t][k].value is None for t, k in rr):
                    bad.append((sn, c.coordinate, "reads 0 from empty cells", f[:60]))
    return bad


def check_errors(wf, wv):
    bad = []
    for sn in wf.sheetnames:
        for row in wf[sn].iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    v = wv[sn][c.coordinate].value
                    if isinstance(v, str) and any(e in v for e in ERRS):
                        bad.append((sn, c.coordinate, v, c.value[:60]))
    return bad


def check_sum_ranges(wf, wv):
    """A SUM whose range contains a cell that is itself a SUM of part of that range,
    or contains a labelled note row - both double count or drag in stray values."""
    bad = []
    for sn in wf.sheetnames:
        ws, vs = wf[sn], wv[sn]
        for row in ws.iter_rows():
            for c in row:
                f = c.value
                if not (isinstance(f, str) and f.startswith("=")):
                    continue
                for m in re.finditer(r"SUM\((\$?[A-Z]{1,3})\$?(\d+):(\$?[A-Z]{1,3})\$?(\d+)\)", f):
                    col = m.group(1).replace("$", "")
                    if col != m.group(3).replace("$", ""):
                        continue
                    lo, hi = int(m.group(2)), int(m.group(4))
                    # a row-total (a SUM across columns) inside a column SUM is not a
                    # subtotal; only flag when the inner SUM covers the same column
                    if sn == "0.1 Budget Table (Fin)":
                        continue
                    if c.row in range(lo, hi + 1):
                        bad.append((sn, c.coordinate, "SUM includes itself", f[:60]))
                        continue
                    for r2 in range(lo, hi + 1):
                        inner = ws[f"{col}{r2}"].value
                        if isinstance(inner, str) and inner.startswith("=") and "SUM(" in inner:
                            im = re.search(r"SUM\(\$?[A-Z]{1,3}\$?(\d+):\$?[A-Z]{1,3}\$?(\d+)\)",
                                           inner)
                            if im and lo <= int(im.group(1)) and int(im.group(2)) <= hi:
                                bad.append((sn, c.coordinate,
                                            f"SUM range contains subtotal {col}{r2}", f[:60]))
                                break
    return bad


def check_hardcoded(wf, wv):
    """Numbers typed into a column whose other rows are formulas - a stale literal."""
    bad = []
    skip = {"0.1 Budget Table (Fin)", "0.2 Data Config", "0.3 Squad Archetypes",
            "Lists", REVIEW, "Squads", "Added data", "Sheet2", "FY26 Budget (superseded)",
            "squad mapping (superseded)", "0.4 Presentation Pack", "Portfolios",
            "3.5 Source Reconciliation"}
    for sn in wf.sheetnames:
        if sn in skip or sn.startswith("-"):
            continue
        ws = wf[sn]
        bycol = collections.defaultdict(list)
        for row in ws.iter_rows():
            for c in row:
                if c.value is None or c.column_letter in ("A", "B"):
                    continue
                bycol[c.column_letter].append(c)
        for col, cells in bycol.items():
            fs = [c for c in cells if isinstance(c.value, str) and c.value.startswith("=")]
            ns = [c for c in cells if isinstance(c.value, (int, float))]
            if len(fs) >= 4 and ns:
                for c in ns:
                    # a yellow cell is a declared input, not a stale literal
                    fill = c.fill
                    if fill and fill.patternType and \
                            str(fill.start_color.rgb or "").upper() == "FFFFFF00":
                        continue
                    bad.append((sn, c.coordinate, f"literal {c.value} in a formula column",
                                f"{len(fs)} formulas in col {col}"))
    return bad


def check_family_consistency(wf):
    """Every tab in a family must be built the same way."""
    out = []
    fam = {
        "1.x": [s for s in wf.sheetnames if re.match(r"^1\.\d+ ", s)],
        "2.x": [s for s in wf.sheetnames if re.match(r"^2\.\d+ ", s)],
    }
    # 2.x: identical header row and identical formula shape per column
    shapes = {}
    for sn in fam["2.x"]:
        ws = wf[sn]
        hdr = tuple(str(ws[f"{c}5"].value or "")[:28] for c in "BCDEFGHIJKLMNOPQRST")
        shapes.setdefault(hdr, []).append(sn)
    if len(shapes) > 1:
        for h, tabs in shapes.items():
            out.append(("2.x headers differ", ", ".join(tabs), h[1][:40]))
    # 2.x: same formula skeleton on the first squad row
    ARCH = {f"2.{i} " for i in range(1, 11)}
    for family in ("archetype", "coe"):
        sk = {}
        for sn in fam["2.x"]:
            is_arch = any(sn.startswith(p) for p in ARCH)
            if (family == "archetype") != is_arch:
                continue
            ws = wf[sn]
            s = tuple(re.sub(r"\d+", "#", re.sub(r"'[^']+'", "'T'", str(ws[f"{c}6"].value or "")))[:70]
                      for c in "EGHIJMOPQ")
            sk.setdefault(s, []).append(sn)
        if len(sk) > 1:
            out.append((f"2.x {family} squad-row formulas differ",
                        " | ".join(",".join(v) for v in sk.values()), ""))
    return out, fam


def check_1x_consistency(wf, wv):
    """The 10 portfolio design tabs should have the same summary block."""
    out = []
    tabs = [s for s in wf.sheetnames if re.match(r"^1\.(10|[1-9]) ", s)]
    layout = {}
    for sn in tabs:
        ws = wf[sn]
        rows = {}
        for r in range(4, 14):
            lab = str(ws[f"B{r}"].value or "").strip()[:26]
            if lab:
                rows[lab] = r
        key = tuple(sorted(rows))
        layout.setdefault(key, []).append(sn)
    if len(layout) > 1:
        for k, v in layout.items():
            out.append(("1.x summary block differs", ", ".join(v), " / ".join(k)[:90]))
    return out


def check_cross_tab_facts(wv):
    """The same fact stated on two tabs must agree."""
    out = []
    def g(sn, cell):
        return wv[sn][cell].value
    facts = [
        ("group cost", [("3.2 Total Cost", "F21"), ("3.1 Group Summary", "D20"),
                        ("Exec Summary", "C7"), ("Exec Summary", "C26"),
                        ("0.2 Data Config", "F26")]),
        ("group roles", [("3.2 Total Cost", "C21"), ("3.1 Group Summary", "J20"),
                         ("3.3 FTE View", "I117"), ("Exec Summary", "C40")]),
        ("filled", [("3.2 Total Cost", "D21"), ("3.3 FTE View", "G117"),
                    ("Exec Summary", "C41")]),
        ("vacant", [("3.2 Total Cost", "E21"), ("3.3 FTE View", "H117"),
                    ("Exec Summary", "C42")]),
        ("cyber cost", [("1.13 Cyber Roles", "F8"), ("1.14 TDD Cyber", "C12"),
                        ("3.4 COE Summary", "K10")]),
        ("budget variance", [("0.2 Data Config", "G26"), ("3.1 Group Summary", "E20")]),
    ]
    for name, cells in facts:
        vals = [(f"{s}!{c}", g(s, c)) for s, c in cells]
        nums = [v for _, v in vals if isinstance(v, (int, float))]
        if not nums:
            out.append((name, "no numeric value anywhere", str(vals)))
        elif max(nums) - min(nums) > 1e-6:
            out.append((name, "disagree", "; ".join(f"{k}={v}" for k, v in vals)))
    return out


def run(path):
    wf, wv = load(path)
    report = collections.OrderedDict()
    report["formula errors"] = check_errors(wf, wv)
    report["dangling references"] = check_dangling(wf, wv)
    report["silent zeros"] = check_silent_zero(wf, wv)
    report["bad SUM ranges"] = check_sum_ranges(wf, wv)
    report["stale literals"] = check_hardcoded(wf, wv)
    fam, _ = check_family_consistency(wf)
    report["2.x inconsistency"] = fam
    report["1.x inconsistency"] = check_1x_consistency(wf, wv)
    report["cross-tab disagreements"] = check_cross_tab_facts(wv)
    return report


if __name__ == "__main__":
    import sys
    rep = run(sys.argv[1] if len(sys.argv) > 1 else "finalp.xlsx")
    total = 0
    for k, v in rep.items():
        total += len(v)
        print(f"\n{'='*78}\n{k.upper()}: {len(v)}")
        for x in v[:18]:
            print("   ", " | ".join(str(y)[:62] for y in x))
        if len(v) > 18:
            print(f"    ... {len(v)-18} more")
    print(f"\n{'='*78}\nTOTAL FINDINGS: {total}")
