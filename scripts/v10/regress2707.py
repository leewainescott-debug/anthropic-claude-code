"""The regression gate for the 2707 consolidation: every finding from the verification
wave (instruction audit, design review, adversarial QA), asserted dead on the candidate.

Run against the injected candidate: python3 regress2707.py cand_A.xlsx
Exit code 0 only when every check passes. Each line reports PASS/FAIL and the fact.
"""
import os
import re
import sys
import zipfile

import openpyxl

REVIEW = "REVIEW - Complete Role Mapping"
ARCH = "0.3 Squad Archetypes"
REV = "rev.xlsx"
OUT = []


def check(name, ok, detail=""):
    OUT.append((bool(ok), f"{'PASS' if ok else 'FAIL'}  {name}" + (f" - {detail}" if detail else "")))


# ---------------------------------------------------------------- 0.3 is his, untouched
def _fill_of(c):
    """The cell's fill as a comparable string, '' for no fill."""
    try:
        if not (c.fill and c.fill.patternType):
            return ""
        s = c.fill.start_color
        return f"{s.type}:{s.rgb or s.theme or s.indexed}:{s.tint or 0}"
    except Exception:                                       # noqa: BLE001
        return "?"


def _col_w(ws, k):
    """The width column k actually renders at, falling back to the sheet's own default."""
    if k in ws.column_dimensions and ws.column_dimensions[k].width:
        return float(ws.column_dimensions[k].width)
    return float(ws.sheet_format.defaultColWidth or 8.43)


def _row_h(ws, r):
    """The height row r has been given of its own, or None where it takes the default.

    Read through the sheet's own default on purpose. LibreOffice recalculates the workbook
    mid-chain (chain2.sh, w1 -> w1r) and that conversion writes an explicit height on every
    row while moving the sheet default from 14.5 to 14.25. A row that inherited the default
    still inherits it - the engine has only spelled it out - so the question worth asking is
    whether a row was given a height of its own, and whether that height matches his.
    """
    h = ws.row_dimensions[r].height if r in ws.row_dimensions else None
    if not h:
        return None
    default = float(ws.sheet_format.defaultRowHeight or 0) or None
    if default and abs(float(h) - default) < 0.05:
        return None
    return float(h)


def archetypes_parity(path, rev=REV):
    """0.3 in the candidate against 0.3 in the owner's workbook, cell for cell.

    The tab arrives from rev.xlsx through assemble_base and no step is allowed to lay it
    out, so this is an equality, not a tolerance. What is compared: every value, every
    fill, every column width and every row height over the used range of both copies.

    Two things are deliberately not compared, because the chain does not cause them and
    cannot fix them. Alignment: the LibreOffice recalc writes the defaults out explicitly
    (None becomes 'general'/'bottom'), which is the same rendering spelled differently.
    Width to the last decimal: the same recalc rounds 28.26953125 to 28.27, so widths match
    to 0.05 rather than to the bit.
    """
    if not os.path.exists(rev):
        near = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            os.path.basename(rev))
        rev = near if os.path.exists(near) else rev
    if not os.path.exists(rev):
        return [f"cannot read {rev} - the parity check needs the owner's own workbook"]
    a = openpyxl.load_workbook(rev)[ARCH]
    b = openpyxl.load_workbook(path)[ARCH]
    bad = []
    # the one declared addition to his tab: the hybrid input he asked for, two cells
    # beside his offshore rate. Everything else stays an equality; the gate's own
    # hybrid checks assert what these two cells must hold.
    EXEMPT = {"K7", "K8"}
    rows = max(a.max_row, b.max_row)
    cols = max(a.max_column, b.max_column)
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            x, y = a.cell(r, c), b.cell(r, c)
            if x.coordinate in EXEMPT:
                continue
            if x.value != y.value:
                bad.append(f"value {x.coordinate}: {x.value!r} -> {y.value!r}")
            if _fill_of(x) != _fill_of(y):
                bad.append(f"fill {x.coordinate}: {_fill_of(x)!r} -> {_fill_of(y)!r}")
    for i in range(1, cols + 1):
        k = openpyxl.utils.get_column_letter(i)
        if abs(_col_w(a, k) - _col_w(b, k)) > 0.05:
            bad.append(f"width {k}: {_col_w(a, k)} -> {_col_w(b, k)}")
    for r in range(1, rows + 1):
        u, v = _row_h(a, r), _row_h(b, r)
        if u is None and v is None:
            continue
        if u is None or v is None or abs(u - v) > 0.05:
            bad.append(f"height row {r}: {u or 'default'} -> {v or 'default'}")
    return bad


def run(path):
    wb = openpyxl.load_workbook(path)                 # formulas
    wv = openpyxl.load_workbook(path, data_only=True)  # cached values
    z = zipfile.ZipFile(path)

    # ---- audit: the fabricated Holds are gone, his three stand (D86)
    ws, vs = wb["1.12 SA&D"], wv["1.12 SA&D"]
    holds = [r for r in range(20, 50) if str(vs.cell(r, 8).value or "").strip() == "Hold"]
    check("1.12 carries exactly his three Holds (rows 31/43/44)", holds == [31, 43, 44], f"holds at {holds}")
    check("1.12 rows 45-47 Onshore, notes empty",
          all(str(vs.cell(r, 8).value or "") == "Onshore" and vs.cell(r, 9).value is None
              for r in (45, 46, 47)))
    w13 = wb["2.13 COE SA&D"]
    lev = [str(wv["2.13 COE SA&D"].cell(r, 5).value or "") for r in range(1, w13.max_row + 1)]
    check("2.13 has exactly 3 Holds, 0 Offshore", lev.count("Hold") == 3 and lev.count("Offshore") == 0,
          f"{lev.count('Hold')} holds, {lev.count('Offshore')} offshore")

    # ---- audit: net-vs-net budget basis (D87)
    check("1.11!C15 is his =C14", wb["1.11 BP&T"]["C15"].value == "=C14",
          repr(wb["1.11 BP&T"]["C15"].value))
    check("1.12!C15 is his =C14", wb["1.12 SA&D"]["C15"].value == "=C14")
    all_text = []
    for t in wb.sheetnames:
        for row in wb[t].iter_rows():
            for c in row:
                if isinstance(c.value, str):
                    all_text.append(c.value)
    joined = "\n".join(all_text)
    check("no 'Offshore discount' label anywhere", "Offshore discount" not in joined)
    check("1.11 funding line labelled and beside its value",
          str(wb["1.11 BP&T"]["B18"].value or "").startswith("Business Partner funding met by")
          and str(wb["1.11 BP&T"]["C18"].value or "").startswith("="))
    check("1.12 funding line labelled and beside its value",
          str(wb["1.12 SA&D"]["B19"].value or "").startswith("Domain Architect funding met by")
          and str(wb["1.12 SA&D"]["C19"].value or "").startswith("="))

    # ---- audit: 1.11 lever vocabulary
    # the lever price is no longer spelled out in the engine - it is looked up from
    # Lists!AC:AD, the one place the four factors live - so what this asserts now is that
    # the lookup is there and reaches the whole table, Hold row included
    eng = [str(wb["1.11 BP&T"].cell(r, 20).value or "") for r in range(21, 45)]
    eng = [f for f in eng if f.startswith("=")]
    check("1.11's cost engine prices every lever off the Lists table",
          eng and all("Lists!$AC$2:$AC$5" in f and "Lists!$AD$2:$AD$5" in f for f in eng),
          f"{len(eng)} engines, first {eng[0][:60] if eng else ''!r}")
    dvs = [dv.formula1 for dv in wb["1.11 BP&T"].data_validations.dataValidation
           if dv.formula1 and "Onshore" in dv.formula1]
    check("1.11 dropdown offers Hold", any("Hold" in f for f in dvs), str(dvs[:2]))

    # ---- audit: visibility (D91)
    hidden = sorted(s.title for s in wb.worksheets if s.sheet_state != "visible")
    check("hidden set is exactly 0.1 / 0.4 / Lists (Exec visible)",
          hidden == ["0.1 Budget Table (Fin)", "0.4 Presentation Pack", "Lists"], str(hidden))

    # ---- audit: his content survives (D89, D90)
    keep = [("1.8 Energy Solutions & B2B", "E17", "What is the B2B initiatives #?"),
            ("1.8 Energy Solutions & B2B", "E18", "What is the B2B CapEx?"),
            ("1.11 BP&T", "B9", "Planned spend is net of the Business Partner"),
            ("1.12 SA&D", "B9", "Planned spend is net of the Domain Architect"),
            ("1.11 BP&T", "B47", "Commercial roles sit in the Business Partnering"),
            ("1.11 BP&T", "B49", "On/Off: set a role to Offshore"),
            ("1.12 SA&D", "B53", "Squad-based SA&D roles sit in their portfolio squads"),
            ("1.12 SA&D", "B55", "On/Off: set a role to Offshore"),
            ("1.13 Cyber Roles", "B73", "Roles and costs come straight from the REVIEW"),
            ("1.2 Customer", "L15", "5m mobile, 4.5m for loyalty"),
            ("1.2 Customer", "L18", "4.5m? Flagged"),
            ("1.2 Customer", "M41", "CPI actuals; pull through"),
            ("0.2 Data Config", "B24", "EG"),
            ("0.2 Data Config", "K24", "Notes"),
            ("0.2 Data Config", "L24", "Reallocated 7m across Ampol & Z Retail"),
            # his K-column squad-row note rides the house relocation to the M note column
            ("1.10 Z Retail", "M39", "No Overhead required"),
            ("1.13 Cyber Roles", "B2", "Cyber, Risk & Service Operations roles and funding")]
    for tab, ref, want in keep:
        v = str(wb[tab][ref].value or "")
        check(f"owner content at {tab}!{ref}", v.strip().startswith(want[:40]), repr(v[:50]))

    # ---- audit: ledger hygiene (D93)
    R, Rv = wb[REVIEW], wv[REVIEW]
    check("REVIEW AR1/AT1 labelled",
          R["AR1"].value == "Overhead line" and R["AT1"].value == "Squad or overhead line")
    stray = [f"{r}:{c}" for r in (191, 192) for c in range(1, 53)
             if R.cell(r, c).value is not None]
    check("spacer rows 191-192 fully empty", not stray, str(stray[:6]))
    for bad in ("Project Manger", "Portfolio Mnager", "michelle Siegman", "vijay Solanki",
                "DeveloperSAP ECC", "EnterpriseProcess Analyst", "Support Analyst -Retail",
                "Siginificant", "acorss"):
        check(f"typo gone: {bad!r}", bad not in joined)

    # ---- QA: the ledger anchors hold
    names = [r for r in range(2, R.max_row + 1) if str(Rv.cell(r, 2).value or "").strip()]
    total = sum(v for r in names if isinstance((v := Rv.cell(r, 27).value), (int, float)))
    check("531 roles in the ledger", len(names) == 531, str(len(names)))
    check("ledger totals 115,589,735.11", abs(total - 115589735.11) < 0.05, f"{total:,.2f}")
    cust = sum(v for r in range(108, 191) if isinstance((v := Rv.cell(r, 27).value), (int, float)))
    check("Customer block 16,522,075.33", abs(cust - 16522075.33) < 0.05, f"{cust:,.2f}")

    # ---- QA H1: the filter hazard is dead
    xml = b"".join(z.read(n) for n in z.namelist()
                   if n.startswith("xl/worksheets/") and n.endswith(".xml"))
    check("no filter criteria anywhere", b"filterColumn" not in xml)
    check("no sortState anywhere", b"sortState" not in xml)
    afs = re.findall(rb'<autoFilter[^>]*ref="([^"]+)"', xml)
    check("exactly one autoFilter, full-width on REVIEW",
          len(afs) == 1 and afs[0].startswith(b"A1:AX"), str(afs))
    hid = [r for r in range(2, R.max_row + 1)
           if R.row_dimensions[r].hidden]
    check("no hidden rows on REVIEW", not hid, str(hid[:5]))

    # ---- QA H3 / M1: platform overhead
    v12 = wv["1.2 Customer"]
    check("1.2!F7 = 0.495 (double-count dead)", abs((v12["F7"].value or 0) - 0.495) < 1e-9,
          str(v12["F7"].value))
    check("1.2!F9 = 15.5625", abs((v12["F9"].value or 0) - 15.5625) < 1e-6, str(v12["F9"].value))
    v10 = wv["1.10 Z Retail"]
    check("1.10!F7 = 0.330 (his no-overhead note honoured)",
          abs((v10["F7"].value or 0) - 0.330) < 1e-9, str(v10["F7"].value))
    check("1.10!I40 stays unpriced", wb["1.10 Z Retail"]["I40"].value is None)
    check("1.10!D7 aligned to his two platforms",
          "SUM(I27,I34)" in str(wb["1.10 Z Retail"]["D7"].value) and
          "I40" not in str(wb["1.10 Z Retail"]["D7"].value))

    # ---- QA M2 + B: one allowance basis
    v31, v32 = wv["3.1 Cost Bridge"], wv["3.2 Overhead & Leadership"]
    d_over = next((v31.cell(r, 4).value for r in range(4, 50)
                   if str(v31.cell(r, 2).value or "").startswith("Overhead roles in the portfolios")),
                  None)
    # found by label on both sides: 3.2's bands moved when the section bar came off
    oh32 = next((r for r in range(5, 30)
                 if str(wb["3.2 Overhead & Leadership"].cell(r, 2).value or "").strip()
                 == "Of which sits in the portfolios"), None)
    f13 = v32.cell(oh32, 9).value if oh32 else None
    check("3.2 and 3.1 state one overhead allowance",
          d_over is not None and f13 is not None and abs(f13 - d_over) < 1e-6,
          f"3.1 {d_over} vs 3.2 {f13}")
    ah5 = wv["Lists"]["AH5"].value
    n_rows = 0
    for t in [x for x in wb.sheetnames if re.match(r"^1\.(10|14|[1-9]) ", x)]:
        for r in range(1, wb[t].max_row + 1):
            if str(wb[t].cell(r, 2).value or "").strip() == "Platform Overhead" \
                    and wv[t].cell(r, 9).value:
                n_rows += 1
    check("Lists platform count == priced overhead rows on 1.x", ah5 == n_rows,
          f"Lists {ah5} vs 1.x rows {n_rows}")

    # ---- QA: 3.4 on his country basis
    w34, v34 = wb["3.4 COE Detail"], wv["3.4 COE Detail"]
    check("3.4 has no elsewhere plug column", "Cost - elsewhere" not in joined)
    ctl = [v34.cell(r, c).value for r in range(20, 35) for c in range(3, 12)
           if isinstance(w34.cell(r, c).value, str) and "ROUND(" in w34.cell(r, c).value
           and "-$G" in w34.cell(r, c).value.replace(" ", "")]
    check("3.4 control is live and reads 0", ctl and all(abs(x or 0) < 1e-9 for x in ctl), str(ctl))

    # ---- QA: live SUMIFS instead of =0
    for tab, cells in (("2.3 Enterprise Data", ["O14"]), ("2.6 Finance", ["O14", "O15"])):
        for ref in cells:
            f = str(wb[tab][ref].value or "")
            val = wv[tab][ref].value
            check(f"{tab}!{ref} live SUMIFS evaluating 0",
                  f.startswith("=SUMIFS") and abs(val or 0) < 1e-9, f"{f[:24]} -> {val}")

    # ---- decisions after the fix: 8 holds, after-cost
    ex = wv["Exec Summary"]
    holds_n = next((ex.cell(r, 3).value for r in range(20, 40)
                    if "hold" in str(ex.cell(r, 2).value or "").lower()), None)
    check("Exec hold count is 8", holds_n == 8, str(holds_n))
    after = next((v31.cell(r, 7).value for r in range(30, 60)
                  if str(v31.cell(r, 2).value or "").startswith("Everything on the ledger")
                  or "after decisions" in str(v31.cell(r, 2).value or "").lower()
                  and v31.cell(r, 7).value), None)

    # ---- design: no theme fills on 1.x bars, cream levers, no banned words
    theme = []
    for t in [x for x in wb.sheetnames if x.startswith("1.")]:
        for row in wb[t].iter_rows(max_row=90, max_col=20):
            for c in row:
                try:
                    if c.fill and c.fill.patternType and c.fill.start_color.type == "theme":
                        theme.append(f"{t}!{c.coordinate}")
                except Exception:
                    pass
    check("no theme fills left on 1.x tabs", not theme, str(theme[:6]))
    low = joined.lower()
    check("the word 'seat' appears nowhere", "seat" not in low)
    check("no en dashes", "–" not in joined)
    frozen = [s.title for s in wb.worksheets if s.freeze_panes]
    check("no frozen panes", not frozen, str(frozen))
    red = 0
    for s in wb.worksheets:
        for row in s.iter_rows():
            for c in row:
                if "[Red" in (c.number_format or "") or "[RED" in (c.number_format or ""):
                    red += 1
    check("no [Red] number formats", red == 0, str(red))

    # ---- design: 2.x family rules
    bare_total = dup = 0
    for t in [x for x in wb.sheetnames if x.startswith("2.")]:
        s, sv = wb[t], wv[t]
        labels = [str(s.cell(r, 2).value or "").strip() for r in range(1, s.max_row + 1)]
        if labels.count("Total portfolio") != 1:
            dup += 1
        for r, l in enumerate(labels, 1):
            if l == "Total" and sv.cell(r, 15).value is not None:
                bare_total += 1
    check("no bare 'Total' subtotals on 2.x", bare_total == 0, str(bare_total))
    check("exactly one Total portfolio per 2.x tab", dup == 0, f"{dup} tabs off")

    # ---- design: 1.5/1.6 family repairs
    check("1.5 headers TDD AU / TDD NZ",
          wb["1.5 P&C"]["C5"].value == "TDD AU ($m)" and wb["1.5 P&C"]["D5"].value == "TDD NZ ($m)",
          f"{wb['1.5 P&C']['C5'].value!r}/{wb['1.5 P&C']['D5'].value!r}")
    w6 = wb["1.6 Finance"]
    check("1.6 scratch moved to S25", str(w6["S25"].value or "").startswith("Nbr Archetype"),
          repr(w6["S25"].value))
    check("1.6 K/L carry the Actual/Variance pair",
          "ctual" in str(w6["K25"].value or "") or "ctual" in str(w6["K24"].value or "")
          or any("ctual" in str(w6.cell(r, 11).value or "") for r in range(20, 30)))

    # ---- 0.3 is the owner's cost library, and the chain does not lay it out
    par = archetypes_parity(path)
    check(f"{ARCH} matches rev.xlsx cell-for-cell (values, fills, widths, heights)",
          not par, f"{len(par)} differences: " + "; ".join(par[:4]) if par else "")

    # ---- wave G: the owner's round - 1.14/2.15, the 3.2 redesign, the table up top
    check("1.14 TDD Cyber exists", "1.14 TDD Cyber" in wb.sheetnames)
    if "1.14 TDD Cyber" in wb.sheetnames:
        w14, v14 = wb["1.14 TDD Cyber"], wv["1.14 TDD Cyber"]
        names = wb.sheetnames
        check("1.14 sits directly after 1.13",
              names.index("1.14 TDD Cyber") == names.index("1.13 Cyber Roles") + 1)
        check("1.14 platform overhead priced, F7 = 0.165",
              abs((v14["F7"].value or 0) - 0.165) < 1e-9, str(v14["F7"].value))
        check("1.14 Cyber Uplift awaits its inputs",
              v14["H26"].value == "check size", repr(v14["H26"].value))
        check("1.14 draws no portfolio overhead",
              (v14["C6"].value or 0) == 0 and (v14["D6"].value or 0) == 0)
        e14 = [c.coordinate for row in v14.iter_rows() for c in row
               if isinstance(c.value, str) and c.value.startswith("#")]
        check("1.14 carries no error cells", not e14, str(e14[:4]))
    check("2.15 TDD Cyber exists", "2.15 TDD Cyber" in wb.sheetnames)
    if "2.15 TDD Cyber" in wb.sheetnames:
        w15, v15 = wb["2.15 TDD Cyber"], wv["2.15 TDD Cyber"]
        ctl15 = [v15.cell(r, 3).value for r in range(1, w15.max_row + 1)
                 if str(w15.cell(r, 2).value or "").startswith("Control -")]
        check("2.15 controls read 0",
              len(ctl15) == 2 and all(abs(x or 0) < 1e-9 for x in ctl15), str(ctl15))
        hot15 = any(str(w15.cell(r, 2).value or "").strip() == "Head of Technology"
                    for r in range(1, w15.max_row + 1))
        check("2.15 draws no Head of Technology line", not hot15)
        check("0.2!F23 includes the 1.14 spend",
              "'1.14 TDD Cyber'" in str(wb["0.2 Data Config"]["F23"].value))
    # ---- wave H/J: 3.2 in the owner's own layout
    v32g = wv["3.2 Overhead & Leadership"]
    w32g = wb["3.2 Overhead & Leadership"]
    HDR32 = ["Overhead roles", "Applied to",
             "Archetype allocation (per portfolio or platform) ($m)",
             "# of times applied in archetypes", "Roles priced for in archetype",
             "Actual number of leadership roles", "# of roles not applied in archetype",
             "Total Archetype cost ($m)", "Actual cost of leadership roles",
             "Variance between archetype and actuals", "Where they sit",
             "Allocation applied"]
    # the header row is found, not assumed: he took the section bar off this tab, so
    # every row below moved up one and a fixed row number would only be right by luck
    h32 = next((r for r in range(2, 12)
                if str(w32g.cell(r, 2).value or "").strip() == "Overhead roles"), None)
    check("3.2 header found", h32 is not None)
    if h32:
        got = [w32g.cell(h32, c).value for c in range(2, 14)]
        check("3.2 carries his headings, in his order", got == HDR32, str(got[:5]))
        lo = h32 + 1
        check("3.2 Times applied is his to set, cream and seeded from the model's count",
              [v32g.cell(r, 5).value for r in range(lo, lo + 6)] == [10, 10, 10, 22, 22, 10]
              and all(w32g.cell(r, 5).fill.patternType
                      and str(w32g.cell(r, 5).fill.start_color.rgb).upper() == "FFFFF2CC"
                      for r in range(lo, lo + 6)),
              str([v32g.cell(r, 5).value for r in range(lo, lo + 6)]))
        check("3.2 states every line's roles in the organisation",
              [v32g.cell(r, 7).value for r in range(lo, lo + 6)] == [15, 6, 7, 10, 24, 8],
              str([v32g.cell(r, 7).value for r in range(lo, lo + 6)]))
        check("3.2 states every line's roles not priced for",
              [round(v32g.cell(r, 8).value or 0, 1)
               for r in range(lo, lo + 6)] == [10.0, 1.0, 2.0, 3.4, 17.4, 5.0],
              str([v32g.cell(r, 8).value for r in range(lo, lo + 6)]))
        check("3.2 HoT row splits its fifteen roles",
              str(v32g.cell(lo, 12).value or "") == "10 in the portfolios, 5 in the COEs",
              repr(v32g.cell(lo, 12).value))
        check("3.2 BP row places all six of its roles in the COEs",
              str(v32g.cell(lo + 1, 12).value or "") == "All 6 in the COEs",
              repr(v32g.cell(lo + 1, 12).value))
        check("3.2 DA row places all seven of its roles in the COEs",
              str(v32g.cell(lo + 2, 12).value or "") == "All 7 in the COEs",
              repr(v32g.cell(lo + 2, 12).value))
        check("3.2 says the allocation in words",
              str(v32g.cell(lo, 13).value or "") == "50% across 10 portfolios",
              repr(v32g.cell(lo, 13).value))
    ALL32 = ("Roles in the organisation, all lines and squads: 531 - portfolios 412, "
             "COEs and EGI 119, each counted once")
    allrow = next((r for r in range(6, 30)
                   if str(v32g.cell(r, 2).value or "").startswith(
                       "Roles in the organisation, all lines and squads")), None)
    check("3.2 all-roles row present", allrow is not None)
    if allrow:
        check("3.2 counts each role once (412 + 119 = 531)",
              str(v32g.cell(allrow, 2).value or "") == ALL32
              and v32g.cell(allrow, 7).value == 531,
              f"{v32g.cell(allrow, 2).value!r} G={v32g.cell(allrow, 7).value}")
        check("3.2 all-roles control reads 0",
              abs(v32g.cell(allrow + 1, 7).value or 0) < 1e-9,
              str(v32g.cell(allrow + 1, 7).value))
    totrow = next((r for r in range(5, 30) if str(w32g.cell(r, 2).value or "").strip()
                   == "Overheads incl. GMs"), None)
    check("3.2 overheads total row present", totrow is not None)
    if totrow:
        check("3.2 totals 70 roles carried against 31.2 priced for",
              v32g.cell(totrow, 7).value == 70
              and abs((v32g.cell(totrow, 6).value or 0) - 31.2) < 1e-6,
              f"{v32g.cell(totrow, 7).value} / {v32g.cell(totrow, 6).value}")
        check("3.2 the two gaps total 38.8 roles and 11.68m",
              abs((v32g.cell(totrow, 8).value or 0) - 38.8) < 1e-6
              and abs((v32g.cell(totrow, 11).value or 0) - 11.682053) < 1e-6,
              f"{v32g.cell(totrow, 8).value} / {v32g.cell(totrow, 11).value}")
    ohrow = next((r for r in range(5, 30) if str(w32g.cell(r, 2).value or "").strip()
                  == "Of which sits in the portfolios"), None)
    check("3.2 'of which sits in the portfolios' band present", ohrow is not None)
    if ohrow:
        check("3.2 archetype cost where the people sit = 5.005",
              abs((v32g.cell(ohrow, 9).value or 0) - 5.005) < 1e-6,
              str(v32g.cell(ohrow, 9).value))

    # ---- wave H: the owner's Actuals-vs-archetype table on every 1.x tab
    BARH = "Actuals vs archetype"
    WANTH = [BARH, "What the cost covers",
             "Actual portfolio", "Archetype portfolio", "Variance"]
    # retired labels, matched EXACTLY: "Actual cost after decisions ($m)" is still the
    # head of every squad table's own actual column and always will be. The bar went.
    DEAD_EXACT = ("Actual cost after decisions", "Squads priced by an archetype",
                  "Squads with no archetype to price them",
                  "Overhead roles in this portfolio", "Additional costs",
                  "Total actual cost after decisions")
    DEAD_START = ("Archetype against actual",)
    tops, bots, shape, stale, wired = [], [], [], [], []
    for t in [x for x in wb.sheetnames if re.match(r"^1\.(10|14|[1-9]) ", x)]:
        ws = wb[t]
        r0 = next((r for r in range(3, 9)
                   if str(ws.cell(r, 11).value or "").strip() == BARH), None)
        for row in ws.iter_rows():
            for c in row:
                s = str(c.value or "").strip()
                if s in DEAD_EXACT:
                    stale.append(f"{t}!{c.coordinate} {s!r}")
                elif any(s.startswith(x) for x in DEAD_START):
                    bots.append(f"{t}!{c.coordinate} {s!r}")
        if r0 is None:
            tops.append(t)
            continue
        got = [str(ws.cell(r0 + i, 11).value or "").strip() for i in range(5)]
        if got != WANTH:
            shape.append(f"{t}: {got}")
        if any(ws.cell(r, c).value is not None
               for r in range(r0 + 5, r0 + 8) for c in range(11, 15)):
            shape.append(f"{t}: K{r0 + 5}:N{r0 + 7} is not clear under the block")
        tot = next((r for r in range(1, 20)
                    if str(ws.cell(r, 2).value or "").strip() == "Total Cost"), None)
        if tot and str(ws.cell(tot, 7).value or "").strip() != f"=$N${r0 + 2}":
            wired.append(f"{t}!G{tot} = {ws.cell(tot, 7).value!r}")
        if tot and abs((wv[t].cell(tot, 7).value or 0)
                       - (wv[t].cell(r0 + 2, 14).value or 0)) > 1e-9:
            wired.append(f"{t}!G{tot} value {wv[t].cell(tot, 7).value!r}")
    check("the actuals table sits up top on every 1.x tab", not tops, str(tops))
    check("every 1.x actuals table is bar + header + Actual / Archetype / Variance, "
          "with nothing under it", not shape, str(shape[:4]))
    check("no retired actuals label or bar text left on a 1.x tab", not stale,
          str(stale[:6]))
    check("no old bottom block remains on any 1.x tab", not bots, str(bots[:4]))
    check("the summary Actuals cell reads the table's Actual portfolio cost", not wired,
          str(wired[:4]))

    # ---- wave H: the hybrid rule
    a03, a03v = wb["0.3 Squad Archetypes"], wv["0.3 Squad Archetypes"]
    check("0.3 hybrid input present, labelled and 2",
          a03["K7"].value == "Onshore roles in a hybrid squad" and a03["K8"].value == 2,
          f"{a03['K7'].value!r} / {a03['K8'].value!r}")
    check("0.3 hybrid input is cream and not a percent",
          (a03["K8"].fill.patternType and
           str(a03["K8"].fill.start_color.rgb).upper() == "FFFFF2CC"
           and "%" not in (a03["K8"].number_format or "")))
    h19 = str(wb["1.9 Commercial Fuels"]["H26"].value or "")
    check("hybrid branch prices 2 onshore plus the rest offshore (1.9 probe)",
          "MIN('0.3 Squad Archetypes'!$K$8,INDEX('0.3 Squad Archetypes'!$F$5:$F$23,"
          in h19 and "))/2," not in h19)
    n_hyb = bad_hyb = 0
    for t in [x for x in wb.sheetnames if re.match(r"^1\.(10|14|[1-9]) ", x)]:
        for row in wb[t].iter_rows(min_col=8, max_col=8):
            for c in row:
                f = str(c.value or "")
                if f.startswith("=") and "0.3 Squad Archetypes" in f:
                    n_hyb += 1
                    if "$K$8" not in f or "))/2," in f:
                        bad_hyb += 1
    check("all 40 squad formulas carry the hybrid rule, none the old midpoint",
          n_hyb == 40 and bad_hyb == 0, f"{n_hyb} formulas, {bad_hyb} old-shape")

    # ---- wave J: the simplification sweep, asserted so it cannot creep back
    idx_join = []
    direct = 0
    JOIN = re.compile(r"INDEX\('REVIEW - Complete Role Mapping'!\$[A-Z]{1,2}:"
                      r"\$[A-Z]{1,2},\d+\)")
    DIRECT = re.compile(r"^='REVIEW - Complete Role Mapping'!\$[A-Z]{1,2}\$\d+$")
    for t in [x for x in wb.sheetnames if re.match(r"^2\.\d+ ", x)]:
        for row in wb[t].iter_rows():
            for c in row:
                v = str(c.value or "")
                if JOIN.search(v):
                    idx_join.append(f"{t}!{c.coordinate}")
                elif DIRECT.match(v):
                    direct += 1
    check("no 2.x cell finds a person by hardcoded row number", not idx_join,
          f"{len(idx_join)} left, e.g. {idx_join[:3]}")
    check("the 2.x ledger join is the insert-safe direct reference", direct > 2000,
          f"{direct} direct references")
    lst = wb["Lists"]
    an = [lst.cell(r, 40).value for r in range(2, 12)]
    check("the override table is keyed on the person, not a row number",
          not any(isinstance(x, (int, float)) for x in an),
          str([x for x in an if x is not None][:4]))
    ar = str(wb[REVIEW]["AR2"].value or "")
    check("REVIEW's overhead-line formula lost its dead branch",
          len(ar) < 520 and "$AQ" not in ar, f"{len(ar)} chars")
    dead = [f"{L_}{r}" for L_ in ("AL", "AM", "AN", "AS")
            for r in (2, 100, 400)
            if wb[REVIEW][f"{L_}{r}"].value is not None]
    check("the four dead ledger columns are gone", not dead, str(dead[:6]))
    ex = wb["Exec Summary"]
    exv = wv["Exec Summary"]
    wide = [ex.cell(r, 3).value for r in range(20, 40)
            if isinstance(ex.cell(r, 3).value, str)
            and re.search(r"\$[DE]:\$[DE]", ex.cell(r, 3).value)]
    check("Exec's vacancy counts read bounded ranges, not whole columns", not wide,
          f"{len(wide)} whole-column")
    e17 = str(wb["3.1 Cost Bridge"]["E17"].value or "")
    e27 = str(wb["3.1 Cost Bridge"]["E27"].value or "")
    check("3.1's archetype subtotal is a plain SUM", e17.startswith("=SUM("), e17[:40])
    check("3.1 keeps the gate on the step that needs it",
          "SUMPRODUCT(--ISNUMBER" in e27, e27[:40])
    cfg = wb["0.2 Data Config"]
    homes = [cfg.cell(r, 10).value for r in range(11, 26)]
    check("0.2 states each portfolio's home country",
          cfg["J5"].value == "Home country"
          and all(h in ("AU", "NZ") for h in homes if h is not None)
          and any(h == "NZ" for h in homes), f"{cfg['J5'].value!r} {homes[:6]}")
    geo = []
    for t in [x for x in wb.sheetnames if re.match(r"^1\.(10|14|[1-9]) ", x)]:
        for r in range(5, 10):
            for c in (3, 4):
                v = str(wb[t].cell(r, c).value or "")
                if "'0.2 Data Config'!$D$" in v and ">" in v:
                    geo.append(f"{t}!{wb[t].cell(r, c).coordinate}")
    check("no tab decides its country by which budget is bigger", not geo, str(geo[:4]))
    lev = []
    for t in ("1.11 BP&T", "1.12 SA&D", "1.13 Cyber Roles"):
        for row in wb[t].iter_rows(min_col=20, max_col=20):
            for c in row:
                if '"Offshore",0.4' in str(c.value or ""):
                    lev.append(f"{t}!{c.coordinate}")
    check("the lever price lives only on Lists", not lev, str(lev[:4]))
    check("1.13's CapEx line follows its input cell",
          "C13" in str(wb["1.13 Cyber Roles"]["F11"].value or ""),
          str(wb["1.13 Cyber Roles"]["F11"].value))
    check("0.2's position figure is computed, not typed",
          isinstance(cfg["L23"].value, str) and cfg["L23"].value.startswith("="),
          repr(cfg["L23"].value))

    # ---- 4.0 all zero
    v40 = wv["4.0 Data QA"]
    fails = []
    for r in range(1, v40.max_row + 1):
        val = v40.cell(r, 5).value
        if isinstance(val, (int, float)) and abs(val) > 1e-6:
            fails.append(f"r{r}={val}")
    check("4.0 Data QA all zero", not fails, str(fails[:5]))

    ok = all(o for o, _ in OUT)
    for _, line in OUT:
        print(line)
    n_bad = sum(1 for o, _ in OUT if not o)
    print(f"\n{len(OUT)} checks, {n_bad} failing")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(run(sys.argv[1] if len(sys.argv) > 1 else "cand_A.xlsx"))
