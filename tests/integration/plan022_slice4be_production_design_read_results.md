# Slice 4be Production designer load/refresh observations

Last verified:2026-09-30 UTC. Architecture v4.11 D18 discovered-control refinement
was committed as documentation `e3adb72` before test/runtime work. The isolated
`deploy/validation-production-design-reads` candidate implements catalog19:
Process Refresh, View Process, Edit as New Version, Recipe Refresh and Load.
It registers103 global IDs and42/68 constructed Production controls;26 remain
unregistered. Catalog18 and17 comparison packages remain frozen. Registration and
focused GREEN do not establish operational promotion or user acceptance.

The contract observes local REFRESHED/PRESENTED/STAGED completion only. It preserves
existing empty-on-read-failure behavior, parser acceptance of empty record arrays,
draft replacement, normalization and next-version fallback. It does not validate
a design, establish source availability/status, reserve a version or change saved
authority. Any change to those algorithms requires a separate architecture decision.

## Focused packaged RED

The unchanged `deploy/validation-production-recipe-structure` baseline completes
**166 PASS/455 FAIL across621 unique checks**,07:19:49.3300025--07:22:25.7192720 UTC.
Five instrumented packages compile. Admin-generated disposable authority fixtures
contain a saved/released Process and Recipe. The adapter invokes all
five original private Click handlers; injected read responses exercise the
existing parser/population code, not replacements for those handlers.

Failures are the intended missing catalog entries/observations and guards:

- 99 catalog-extension/history checks and five fixed-control metadata checks;
- 273 attempt/outcome/context/redaction/integrity/terminal checks;
- 10 loading/busy, five nested-read, three nested local-result and five permission
  read/mutation checks;
- 40 stale target/session/sign-out/closed-workbook checks;
- five visible unavailable-tracking notices; and
- 10 handled read/partial-failure checks.

Normal loading and reuse preserve existing behavior. All empty-array/list and
malformed/unavailable-read preservation assertions pass. Injected partial failures
leave local effects, so rollback is not asserted. Nested calls on the baseline
can duplicate local rows; those three failures support the required nested-action
guard. Disabled/unavailable tracking preserves authorized local behavior. Saved
authority, unknown workbook columns, operator file bytes and earlier records
remain unchanged. No unexpected failure or harness exception remains.

Controller:
`reports/runtime/production-design-read-controller/211381b06aa742c49bad4bd73950d88d`.
Result:
`reports/runtime/slice4be-production-design-reads/8dcdc3d927fc4dde806e0571e5ef7d34/red.json`.
The controller's `verification.json` confirms settings/five package preservation,
Excel closure and zero Excel Application1000/1001/1002 events over the full interval
plus12 seconds. The ordinary harness closes/quits without a termination path;
desktop probes remain healthy with no error5.

Refreshed `reports/runtime/production-design-read-red-static` retains265 components,
6093 procedures,134045 lines,9 literal/45 unresolved dynamic calls,191 duplicate
groups and all28 exact size caps. Three schemas and324 PowerShell parses pass.
Runtime source is unchanged and unrelated user-document byte hashes are preserved.

## Focused packaged GREEN

The exact621 ordered RED identities pass621/621 on catalog19,
07:35:02.8027679--07:38:23.2365139 UTC. Controller
`reports/runtime/production-design-read-controller/90779516512c4f07b400c107f964157c`,
result `reports/runtime/slice4be-production-design-reads/365e59a3cf2e4fb390365bbcb0c4e11e/green.json`.
Five instrumented compiles, settings/five-package preservation, unassisted closure
and zero Excel Application1000/1001/1002 events over the interval plus12 seconds
pass. Controller source snapshots preserve the exact focused test implementation.
No desktop error5 is observed.

The candidate build/compile/cold-load is retained in
`reports/runtime/production-design-read-build`,07:34:22.5029870--07:35:00.7978035 UTC.
Compiled components increase258 to260: three changed and two new owners. Only
Core's catalog/evaluation matching and Operations' form delegation change;
new typed action/catalog owners contain the orchestration and fixed vocabulary.
The form decreases11697 to11685 lines, within its existing cap. Existing loaders
and selection conversion semantics are preserved. Settings and the frozen
catalog18 packages are unchanged; normal closure and the Application audit pass.

## Earlier preliminary attempts

These are retained with their actual limits, not combined into the focused RED:

- Controller `43cd3cc6074b4a4e97db6edeb44c6586`:39 PASS/two failures, before the
  actual cases. The fixture incorrectly read status from the three-column Run
  picker. Status belongs to the six-column Recipe list. This is harness setup
  failure, not product RED.
- Controller `86fba5b8a4594a37923b796675fe1570`:93 PASS/283 FAIL, then a stale
  worksheet reference interrupts preservation checks. Reacquire that reference
  after reopening the disposable workbook. This run is incomplete.
- Corrected preliminary controller `6bea9e617e52494d920c2706f316a68b`:
  100 PASS/282 FAIL across382 checks, five compiles, preservation, closure and zero
  Application failures. Result `slice4be-production-design-reads/7be91b1252594d44b189358f39e98ce5`.
  Its source snapshot/pins are retained with the controller. The621-check gate
  above adds the required context/permission, optional-tracking, partial-failure
  and nested-action coverage before runtime changes.

Controller names in this subsection are under
`reports/runtime/production-design-read-controller/`.

Static evidence `reports/runtime/production-design-read-static` records267 components,
6097 procedures and134205 lines. The two bounded owners account for the added
implementation; all28 oversized-module ratchets do not grow, the form shrinks,
dynamic calls remain9 literal/45 unresolved and duplicate candidates decrease191
to190. Three schemas and325 PowerShell parses pass. This is the pre-editor-fix
checkpoint; regenerate after any further runtime change.

## Recording/editor integration finding

The first separate paired-path attempt records85 PASS/two real missing-outcome
choice failures, followed by one harness interruption. Both original five-action
recordings, expected local drafts, publication and Event Detail pass. Core's
`modExpectationDraft.Choices` fixed list omits PRESENTED, so the actual Expected
Conclusion editor cannot choose either load control's positive outcome. This
contradicts the already specified D18 guide/diagnostic contract; adding the term
to that catalog-filtered choice list is implementation of the existing contract.

Controller `production-design-read-controller/19ceb4401d7b4e4ab038120cb6c316f2`,
result `slice4be-production-design-read-paths/34f5d66fdb624cdf9cd552287df37fd9`,
07:38:32.1688770--07:43:36.5623085 UTC (both roots under `reports/runtime/`).
Attempting to write the absent value into the dropdown raises visible VBA380.
An identity-verified owned-fixture termination releases the blocked worker;
settings and five packages restore, and the Application audit has zero Excel
failures. Closure is assisted, not normal. This is not desktop error5 or the
historical native c0000028 failure.

The clean actual-editor RED completes90 PASS/two expected failures across92 unique
checks,07:43:55.0168104--07:47:07.6118497 UTC. Only
`DesignReadPaths.ExpectedOutcome.PROCESS_LOAD` and `.RECIPE_LOAD` fail; Refresh
and Reuse choices pass. Five instrumented compiles, settings/five-package
preservation, unassisted closure and zero Excel Application failures pass.
Controller `production-design-read-controller/aa295fcacd924ab8a785a5862b0d73a0`,
result `slice4be-production-design-read-paths/e765f030ec6647d586e23d135feacaf8/red.json`
(both under `reports/runtime/`). The test cancels and returns on missing choices;
it never writes an unsupported dropdown value. Source snapshots/pins are retained.
`modExpectationDraft.Choices` is unchanged at this RED checkpoint. Next add
PRESENTED to its catalog-filtered fixed list, build a distinct candidate, retain
these92 checks in the full path GREEN and rerun621 focused checks plus regressions.

## Remaining work

The typed owner/catalog implementation and exact621 focused GREEN are recorded.
Require separate original
recording/publication/How-To/Diagnostic/Compare evidence, all established packaged
regressions and visible captures. Preserve catalog18 and catalog17 comparisons.
The earlier intermittent native Excel failure remains an independent open finding;
this group must not claim to repair it. Slice4be/Release1 acceptance remains open.
