# TDD Cost Calculator - validation evidence pack

QA harness result: FAILS: 0 (exit 0). Three passes:
  PASS 1 - independent Python recompute straight from raw data / Added data (no Excel formulas)
  PASS 2 - full formula-engine evaluation of every cell: zero formula errors
  PASS 3 - the assertion battery below, plus a second engine run with the offshore
           toggle flipped to prove Offshore = exactly 40% of onshore cost

## Key numbers from this run
  PASS1: grand=120.038 squads+strat=83.980 lead=16.067 coe=10.828 unmapped=9.163 | BPT 23@6.126 SAD 11@2.576 CYB 48@9.837 unspec 7@1.595
  KEY: actual grand=120.03817995310001 model grand=84.32636977033333 squads model=67.01664827733333 FTE model grand=79.05614827733334 FTE actual=83.9799506261

## The 84 named assertions (every one passed)
    1. pass1 classes sum to grand
    2. pass1 BPT count = build
    3. pass1 SAD count = build
    4. pass1 CYB count = build
    5. pass1 unspecified count = build
    6. 2.1 actual grand = Added data total
    7. 2.1 check cell = 0
    8. 2.1 subA actual = independent squads+strat
    9. 2.1 subB actual = independent leadership
   10. 2.1 subC actual = independent COE
   11. 2.1 subD actual = independent unmapped
   12. 2.1 de-dup row = -(11*L7+11*L8)
   13. 2.1 subC model = COE gross - overhead-funded BP/DA
   14. 0.1 Squads K5 offshore rate
   15. 2.3 BP overhead applied 11x
   16. 2.3 BP funding/portfolio = L7
   17. 2.3 BP total funding = 11*L7
   18. 2.3 BP FTE-equivalents 4.4
   19. 2.4 DA overhead applied 11x
   20. 2.4 DA funding/portfolio = L8
   21. 2.4 DA total funding = 11*L8
   22. 2.4 DA FTE-equivalents 5.5
   23. 1.11 G24 = 2.5 TDD COE planned
   24. 1.11 G25 = 2.5 TDD Cyber planned
   25. 1.11 capex 0.5 visible
   26. 2.5 budget = people bucket only (capex separate)
   27. 2.5 tie row = 1.11 total to fund (one cyber number)
   28. 2.2 D8 = SA&D Strategy planned
   29. 2.2 D10 = BP&T Transformation planned
   30. 2.2 D11 = BP&T Business Partnering planned
   31. 2.2 D12 = SA&D Data planned
   32. 3.0 actual squad cost = 2.1 squads actual
   33. 3.0 EGI platforms carry no overhead (5 exempt)
   34. 3.0 model grand = squads model + non-EGI platform OHs + portfolio OHs
   35. 3.0 model grand TIES 2.0 J17 delivery total
   36. 3.0 cross-check = 0
   37. 3.0 KPI org roles
   38. 3.0 KPI model cost = grand M
   39. 3.0 KPI actual cost = grand N
   40. 3.0 Cost var grand = 2.1 squads variance (like-for-like)
   41. 3.0 seats-vs-model grand counts archetype squads only
   42. 3.0 cyber row model single-sourced from 1.11/2.5
   43. 3.0 leadership count = 53
   44. 2.0 J26 net = 2.1 model grand
   45. 2.2 unspecified roster = 7 (actual cells)
   46. 2.2 unspecified total = independent mapping
   47. visible sheets = 37
   48. Exec: budget 53.8
   49. Exec: allocated 43.5
   50. Exec: TDD Cost = 2.0 D24
   51. Exec: funded outside = 2.0 G24
   52. Exec: left to fund = 2.0 I24
   53. Exec money block FOOTS to the archetype total
   54. Exec archetype total = 2.1 grand
   55. gap identity: vacant - filled-underspend = gap
   56. Exec COE line = 2.2 F13
   57. Exec cyber line = 2.5 tie
   58. Exec people: archetype seats = FTE G4
   59. Exec people: raised seats = G+K
   60. Exec people: raised beyond = K
   61. Exec people: outside-archetype seats = J-G-K
   62. Exec people: vacant 166
   63. Exec people: filled 370
   64. Exec people: squad-lever vacancies 125
   65. Exec people: non-lever vacancies 41
   66. Exec lever cost = sum of GM hire-all (mapped, all 125 seats)
   67. Exec money: TDD net after double-count
   68. Exec B49 = filled minus archetype (positive = over)
   69. Exec drill: archetype seats lookup (Ampol default)
   70. Exec decision: over-archetype = G grand
   71. 2.2 F8 = 2.4 net
   72. 2.2 F10 = 2.3 net
   73. 2.2 F11 = 2.3 net
   74. 2.2 F12 = 2.4 net
   75. 2.0 I24 = portfolio left-to-fund + net COE
   76. 2.2 E13 sums the draw-downs
   77. Exec drill budget = 2.0 C6
   78. Exec drill TDD cost = 2.0 D6
   79. Exec drill total cost = 2.0 J6
   80. 2.0 C30 allocations 43.5
   81. 2.0 C32 check = 0
   82. GM tabs cover every squad seat (425)
   83. 3.1 raw records 536
   84. 3.1 added records 548

## Structural guards (fail the build if violated): 40
  - tab order/list wrong
  - duplicate tab names
  - stale ref
  - 1.11 G24 not live ref
  - 1.11 G25 not live ref
  - 1.11 total-to-fund block shape unexpected
  - 3.0 merged cyber row missing
  - 3.0 leadership count not live COUNTIF
  - 2.0
  - 2.0 C35 not 0% format
  - negative-zero residue at
  - 0.1 H
  - SupportPct range wrong
  - SupportPct missing
  - 2.3 BP FTE-equivalents not referencing K7
  - 2.4 DA FTE-equivalents not referencing K8
  - 2.4 row
  - header-fill bleed remains at
  - gridlines still on
  - 3.1 live check cells missing
  - 3.1 not labelled a snapshot
  - 2.1 G5 header not in archetype vocabulary
  - 2.1 C5 header not in archetype vocabulary
  - Exec Summary is not the first sheet
  - Exec title-rate caveat line missing
  - Exec vacancy/ring-fenced basis line missing
  - Exec section missing
  - Exec section missing: The people
  - 2.0 E5 header not updated
  - 1.1 TDD Lights On label not disambiguated
  - Exec drill-down selector not found
  - Exec drill TDD Variance negative
  - Exec drill TDD Variance not numeric
  - 3.0 Other unmapped
  - italic/tiny at
  - grey font at
  - en/em dash in displayed values
  - dash in
  - italic/tiny font at
  - comment at
