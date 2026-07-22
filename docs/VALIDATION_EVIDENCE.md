# TDD Cost Calculator - validation evidence pack

QA harness result: FAILS: 0 (exit 0). Three passes:
  PASS 1 - independent Python recompute straight from raw data / Added data (no Excel formulas)
  PASS 2 - full formula-engine evaluation of every cell: zero formula errors
  PASS 3 - the assertion battery below, plus a second engine run with the offshore
           toggle flipped to prove Offshore = exactly 40% of onshore cost

## Key numbers from this run
  PASS1: pt={'Business Partnering': 3.293, 'Transformation': 3.108} sad={'Strategy & Architecture': 3.043, 'Data': 0.942} cy={'TDD Cyber': 4.387, 'TDD COE': 6.529} egi=4.666 BP_budget=3.892 SA_budget=3.540 ledger=120.038
  KEY: model 87.089255418 actual 119.59554360310001 BPbudget 3.8920000000000003 SAbudget 3.54 cyber 10.915989044 egi 4.665919350000001 restate -0.443

## The 49 named assertions (every one passed)
    1. p1.divisions
    2. p1.sad_coe_n
    3. p1.sad_covers_techstrategy
    4. p1.sad_covers_architecture
    5. p1.nport
    6. p2.zero_engine_errors
    7. bpt.nport
    8. bpt.fte_funded
    9. bpt.draw
   10. bpt.budget_both
   11. bpt.g6_budget
   12. bpt.g7_budget
   13. bpt.check0
   14. sad.nport
   15. sad.fte_funded
   16. sad.draw
   17. sad.budget_both
   18. sad.h6_budget
   19. sad.h7_budget
   20. sad.paused_memo
   21. sad.check0
   22. sad.roles_total
   23. cy.total52
   24. cy.check0
   25. cy.tie_111
   26. cy.flows_to_21
   27. coe.e8
   28. coe.e11
   29. coe.d8_ref
   30. coe.d12_ref
   31. coe.d10_ref
   32. coe.d11_ref
   33. t21.dedup
   34. t21.total_c
   35. t21.total_i
   36. t21.restate
   37. t21.egi_memo
   38. t30.egi_total
   39. t30.egi_n
   40. t30.xcheck0
   41. exec.model
   42. exec.actual
   43. exec.filled
   44. exec.vacant
   45. exec.dedup
   46. exec.unmapped
   47. exec.lever
   48. order.49_410
   49. off.flip

## Structural guards (fail the build if violated): 0
