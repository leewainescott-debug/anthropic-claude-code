# Portfolio cost bases, version 2: the plan (28/08, after the retrospective)

Ruled by Lee: "fix everything, do some detailed planning, ensure agents run."
Fable plans, orchestrates and gates only. Agents: Sonnet by default, Opus
only where a wrong judgement is expensive to unwind. Every agent writes its
pipeline to disk first and gets a self contained brief, so an evicted agent
can be replaced without starting over.

## What changes from version 1, and why
1. Structural integrity check runs FIRST and alone, as a reusable script,
   before any lane trusts a sum (v1 found the 88.5m overstatement inside a
   bigger brief; the mechanism has now appeared in two shapes across two
   files, so it becomes a standing five minute check).
2. Mapping rules are written twice, independently, then diffed: the v1 set
   was written from the transaction detail outward; a fresh agent writes a
   second set from the FY26 budget tab bridge outward (columns AY to BB,
   cost centre to portfolio, platform and squad), blind to the first.
   Agreement sets the confidence grade automatically; disagreement IS the
   grade, resolved by a named note (v1 had one agent grading its own work,
   which needed a correction round).
3. Technology QA runs on the merged rule set BEFORE the builder starts (v1
   ran it alongside the build and forced a pause and a face rebuild).
4. The labour boundary is proven by named role against the FY27 labour
   model, not by vendor dollars alone: the 26.08m contractor question and
   the contingent and personnel elements get a two scenario dollar table.
5. The dropped thread closes: the head of technology and delivery manager
   reconciliation gap (about 1.85m between the title basis and the
   overhead block basis) proven by named role.
6. The two largest judgements become typed toggles in the workbook that
   drive the face by formula (contractors: non labour or labour; the
   enterprise platform pools: with the payer, or pooled and allocated by a
   stated key), so Lee's ruling flips the numbers without a rebuild.
7. Build gating splits: a light formulas only copy for every iterative
   check; the slow full file recalculation runs exactly once at the end
   (v1 lost about ninety minutes to a crashed check and a killed
   recalculation mid build).
8. One interim contact, under the exception Lee approved: any single
   judgement above five percent of the total base earns one short question
   the moment it is found. Checkpoint 1 carries the certificate, the three
   payment piles and the two scenario tables for the two largest
   judgements; the build continues on the data as it is, with the toggles
   ready for his ruling.

## Sequence (S = Sonnet, O = Opus)
1. S1 Integrity certificate (S). Reusable sap_integrity_check.py; run on
   budget_v4.xlsx and on the earlier v1 file if present. Output: one page
   certificate tying every tower to the cent and naming the overlap
   mechanism. Gate: nothing downstream trusts a sum until this exists.
2. S2 Labour boundary and dropped thread (S), in parallel with S1
   (different data). Output: labour_boundary.json with the two scenario
   tables, hot_dm_reconciliation.md.
3. O1 Mapping from the budget tab bridge outward (O), starts when S1
   lands, blind to the v1 rule set. Also derives any split keys the bridge
   supports for HITSEC, Data Analytics and AI, BP&T and SA&D.
4. S3 Rule diff and confidence grading (S): v1 set versus O1 set, row
   level; produces mapping_rules_v2.json and leaf_assignments_v2.jsonl with
   grades from agreement; every disagreement carries a named note.
5. S4 Technology QA (S) on the merged set, adversarial, before any build.
   Findings are claims: the diff agent verifies and applies, the builder
   never builds on unverified QA numbers.
6. CHECKPOINT 1 to Lee: certificate, three piles, two scenario tables.
7. O2 Build v2 (O): the v1 pipeline extended with the toggles, the graded
   rules, the role based labour boundary note, the closed thread, the
   certificate summary on Read me. Iterative gates on a light copy; full
   recalc once; Excel openability gate; renders read.
8. S5 Cost QA (S), independent re derivation including both scenario
   positions. S6 Instructions audit (S) against I1 to I25 and this plan.
9. Fable gate on renders, ship, log, one final report.

## Standing rules added to the registers by this plan
- I22. Integrity check first on every SAP sourced file.
- I23. Mapping rules written twice and diffed; confidence from agreement.
- I24. Technology QA before build; cost QA after; instructions audit last;
  QA findings are claims to verify, never figures to build on.
- I25. Interim contact exception: one short question for any single
  judgement above five percent of the total base, nothing else interrupts.
