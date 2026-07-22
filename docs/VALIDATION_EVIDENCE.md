# TDD Cost Calculator - validation evidence pack

QA harness result: FAILS: 0 (exit 0). Three passes:
  PASS 1 - independent Python recompute straight from raw data / Added data (no Excel formulas)
  PASS 2 - full formula-engine evaluation of every cell: zero formula errors
  PASS 3 - the assertion battery below, plus a second engine run with the offshore
           toggle flipped to prove Offshore = exactly 40% of onshore cost

## Key numbers from this run
  PASS1: pt={'Business Partnering': 3.293, 'Transformation': 3.108} sad={'Strategy & Architecture': 4.321, 'Data': 2.233} cy={'Service Operations': 4.387, 'Cyber & Risk': 6.529} egi=4.666 BP_budget=3.720 SA_budget=3.400 ledger=120.038
  KEY: model 89.054848478 actual 118.19870640309999 BPbudget 3.72 SAbudget 3.4000000000000004 cyber 10.915989044000002 egi 4.665919350000001 restate -1.839

## The 64 named assertions (every one passed)
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
   25. cy.coe_row_26
   26. cy.buckets
   27. cy.grid_24
   28. cy.group_row19
   29. coe.grid_bp
   30. coe.grid_tr
   31. coe.grid_sa
   32. coe.grid_data
   33. coe.grid_budget_bp
   34. coe.grid_budget_sa
   35. t21.dedup
   36. t21.total_c
   37. t21.total_i
   38. t21.restate
   39. t21.egi_memo
   40. t30.egi_total
   41. t30.egi_n
   42. t30.xcheck0
   43. exec.model
   44. exec.actual
   45. exec.filled
   46. exec.vacant
   47. exec.dedup
   48. exec.unmapped
   49. exec.lever
   50. recon.c32
   51. cy.grouping_hdr
   52. cy.no_stale_hdrs
   53. order.49_410
   54. raw.Squads.only_logged_changes
   55. raw.Added.owner_cols_untouched
   56. sadpart.all_accounted
   57. cover.squads_all_on_4x
   58. cover.sheet2_all_on_4x
   59. aunz.au_total
   60. aunz.nz_total
   61. aunz.covers_squads_sanity
   62. t20.k24_net
   63. lever.offshore_04
   64. lever.offshore_04

## Structural guards (fail the build if violated): 0
