# Slice 4be Production lifecycle recording and diagnostics

Status: RED checkpoint; candidate build blocked by Windows cleanup, 2026-09-25 UTC.
Architecture v4.11 D5/D18 and Plan022 govern
this work. This record does not accept Slice4be or Release1, replace the preceding
615 lifecycle checks, or waive visible, presentation, regression and full-chain gates.

## Discovered D5 processor read breach

The packaged fixture retains the accepted absence of optional Config `Timezone`
and an unknown operator column. After Admin saves capture policy version1, the
actual Process Save handler applies its Designs event but loses its completion
observation. Its processor calls `AddConfiguredInboxTargets`, which previously
opened Config through `OpenOrCreateConfigWorkbookRuntime`. That provisioning path
requires optional headers, normalizes the worksheet set and saves Config. It
removes the tracking-policy sheets, returning policy version0/defaults. The
completion guard correctly refuses a changed policy; weakening that guard would
violate D18 and conceal the D5 cause.

The focused packaged check is
`LifecyclePaths.Process.Save.Confirmed.ConfigReadPreservesPolicyAndBytes`, paired
with `ExactOriginalOccurrence`. Both fail on the frozen lifecycle package after
the real command succeeds. Private controller
`reports/runtime/production-lifecycle-paths-controller/e442f999bc55459e92b1289a33764051`
closes normally, restores settings and preserves all five packages. Its report
`reports/runtime/slice4be-production-lifecycle-paths/324aa20923d74262b6e95db23104a66b/red.json`
contains144 PASS/11 FAIL: eight separate missing Designs evaluator assertions,
these two behavioral failures and the subsequent prerequisite-stop marker. The
stop marker is not behavioral RED and the evaluator UI was not reached in that run.

The correction shares Core's existing `ResolveExistingConfigForRead` with the
processor using a direct typed call. The processor opens existing Config read-only,
preserves a borrowed workbook and closes only its own transient read. Ordinary
processing no longer invokes provisioning at this boundary. No normative contract,
configuration authority, capability or cross-package object bridge changes.

The replaced private `FindOpenWorkbookByNameProcessor` helper has no remaining
callers: repository-wide VBA/tooling/root-registry search finds only its definition
and return assignment after the replacement. The read-corrected candidate compiles
all five instrumented packages and preserves Config through all24 owning actions.
That reviewed reachability and protecting evidence supports removing this obsolete
name-only lookup; the final candidate still requires its own compile/regression gates.

## Required evidence being exercised

`Test-Slice4beProductionLifecycle.ps1 -ActionPaths` records24 real form actions:
BeforeAppend, AfterAppend, Pending and Confirmed for each of the six lifecycle
controls. Isolated queues prevent uncertain/pending fixture submissions from being
processed by later cases. The stopped journal must retain all48 original records
in order, and real Admin publication must retain exact Activity/Designs identities.
Actual Viewer selection and detail handlers protect reachability and identity.

The real expectation editor and Evaluate button check command completion separately
from Designs application. Repeated failures must match distinct original occurrences.
Ordered six-command expectations preserve the18 extra attempts explicitly. Saved
source evidence must retain the exact four owner-reference fields, full line count
and hash, and empty SystemKeys for Designs. Evaluation must preserve original
activity/journal, saved authority and the loaded publication.

Supplemental `Slice4beDesignsSourceEvidence.ps1` checks complete multi-line evidence,
integer text, uncertain versus submitted missing IDs, coverage omissions, malformed
application evidence and saved-result validation. Existing Inventory key and optional
warehouse semantics remain protected. These pure checks supplement the packaged
owner/UI route and cannot replace it.

The complete evaluator RED on `deploy/validation-production-policy-read` records
668 PASS/40 expected FAIL across708 checks. Controller
`reports/runtime/production-lifecycle-paths-controller/02aec33d88644412bff06fcf6b1dfa2a`
and report
`reports/runtime/slice4be-production-lifecycle-paths/fdbe45d1bed1455d8ae27457c3cc2e5c/red.json`
retain the private evidence. Failures comprise eight supplemental Designs checks,
six confirmed-command conclusions, twelve confirmed-source conclusion/evidence
checks, twelve pending-source conclusion/evidence checks and two ordered-series
conclusions. All24 original action pairs, published identities, editor delivery,
matched occurrences and read-only preservation checks pass. The five protecting
test files remain pinned for the candidate GREEN replay.

## Shutdown and candidate status

All708 unique results are saved with zero harness failures and unchanged protecting
test hashes. This does **not** establish normal shutdown: the worker reaches its
terminal RED output after Close/Quit, but Excel continues consuming CPU. Exact
owned-process termination and then completed-worker termination are requested;
the original outer controller remains alive with the pre-test settings snapshot.
Windows subsequently reports HasExited=True/no main window while still enumerating
one thread and65069 handles. Taskkill reports no running instance. At01:37 UTC the
controller still has no closure receipt, so final settings restoration is **pending**.
Keep that controller alive; do not start another Excel gate or infer restoration
from the worker's terminal output. Private recovery receipts accompany the controller.
The unstarted build waiter is stopped so that no build begins unattended when the
Windows process finally disappears. Resume the build explicitly after the original
controller verifies restoration and package preservation.

The final source correction is not yet built or GREEN. The intermediate package
proves the D5 preservation cases; it deliberately retains the old evaluator for RED.
Do not label either package as the final candidate. A read-only cursor probe briefly
succeeds at01:10 UTC, then returns error5 again at01:29; no new native capture is
accepted, and Group Policy remains an unproven hypothesis. No Windows policy changes
were made. These observations do not establish a shared cause for cursor and shutdown
problems.

Final source static evidence is `reports/runtime/production-paths-final-static`:
253 components,6058 procedures,133166 lines (+12),9 literal/45 unresolved dynamic
calls unchanged,191 duplicate groups (one fewer), all28 existing size caps held.
Three evidence schemas,285 PowerShell parses and15 source-layout checks pass.
Receipts: `production-paths-static-verification.json`, `production-paths-red-verification.json`
and `production-paths-test-pins.json` under reports/runtime. Static evidence cannot
replace compile, packaged GREEN, regressions or visible acceptance.

Pending: final package build/compile, evaluator GREEN, current-candidate lifecycle/
draft/Settings/smoke/full-chain/reusable regressions, How-To/Diagnostic comparison,
native cancellation and visible operator acceptance. Release1 and Slice4be remain
incomplete.
