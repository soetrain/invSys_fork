# Slice 4be Production lifecycle recording and diagnostics

Status: packaged functional GREEN708/708; assisted controller cleanup,
current-candidate regressions and visible extensions pending, 2026-09-25 UTC.
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

## Final candidate GREEN and remaining gates

The previous retained process disappears before02:06 UTC without an agent-requested
restart. Desktop cursor access succeeds. The five final packages are built in
`deploy/validation-production-paths`; Operations cold start and all five compile
checks pass. Comparison of246 compiled components with the lifecycle predecessor
finds exactly four changed Core modules: `modConfig`, `modProcessor`,
`modEvaluationMatches`, and `modEvaluationSources`; no component is added.

The unchanged protecting gate now records **708 PASS/0 FAIL**,708 unique checks,
zero harness failures and the exact RED check set. All40 expected RED failures
are corrected. The five protecting test hashes remain unchanged through the GREEN
result. Both ordered six-action conclusions, all24 original action pairs, original
publication/detail identities, saved Designs evidence and read-only preservation
pass. Controller:
`reports/runtime/production-lifecycle-paths-controller/21b2eeea63354e18a4d8973e03c96b31`.
Result:
`reports/runtime/slice4be-production-lifecycle-paths/df9129b233fd490b977db6aff397facf/green.json`.
Verification: `reports/runtime/production-paths-green-verification.json`.

This functional GREEN does not establish normal shutdown. After the last lifecycle
checks at02:42:46 UTC, cleanup delays the remaining general checks and terminal
output until about02:47:53. Windows then marks Excel exited while retaining one
kernel-busy cleanup thread and65024 handles. No Excel termination is requested in
this run. The completed PowerShell worker remains in COM teardown and is terminated
only after verifying the saved708-check result, five package hashes and no live
Excel. The previously proven temporary process-local Get-Process filter excludes
HasExited entries, preserves live processes, and lets the original controller's
unchanged restoration/comparison finish. Its02:52:53 closure verifies settings
restored and all packages preserved; ExitCode=-1 records worker termination.
Controller and debugger exit. ExcelClosed in that receipt means no live Excel,
not disappearance of the retained Windows entry. No settings values are exported,
and the repository cleanup guard is unchanged.

Read-only resource samples show growing Section/Event handle counts during the
long gate while GDI/USER counts remain broadly steady. At cleanup, user-code CPU
time largely stops increasing while kernel CPU time continues. Application and
Display event queries find no corresponding failure through02:47:53. These facts
do not prove a graphics-driver, RDP, Group Policy or application root cause.
Private receipts: `production-paths-green-recovery.json`,
`production-paths-resources.jsonl`, `production-paths-handle-types.jsonl`,
`production-paths-cleanup-cpu.jsonl`, and `production-paths-application-events.json`.
Ordinary Excel gates remain blocked until the retained entry disappears.

After verifying the original five GREEN pins, test-only extensions add isolated
`-NativeCancellation` and `-Presentation` routes. The first invokes all four real
Release/Obsolete handlers, declines only their exact owned native questions, and
checks CANCELLED facts, source/draft preservation and captured-workbook identity.
The second records two independent six-action series, authors a guide from one,
explicitly evaluates the other, and checks How-To/Diagnostic/Compare both with
original identities and immutable evidence. These extensions are prepared and
parsed, **not yet packaged-tested or visually accepted**. No runtime or normative
behavior is changed by these additions.

Refreshed maintenance evidence in `reports/runtime/production-paths-visible-static`
retains253 components,6058 procedures,133166 lines,9 literal/45 unresolved dynamic
calls,191 duplicate groups and28 existing caps. All three generated evidence schemas
and289 PowerShell parses pass. Earlier15 source-layout checks
remain applicable because no runtime/form source changes follow the final build.
Native captures, paired-view captures,615 lifecycle/390 draft/202 Settings/86 smoke,
full-chain/live-role and reusable current-candidate regressions remain required.

## Historical RED cleanup qualification

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

Subsequent recovery at01:45:23 UTC resolves the settings/controller risk. A
disposable controller first proves that a temporary process-local Get-Process
wrapper can exclude HasExited=True objects while preserving in-memory state.
PowerShell's supported Enter-PSHostProcess/Debug-Runspace then reaches the original
controller's wait loop. The same temporary wrapper keeps live processes visible,
reports zero live Excel processes and lets the unchanged restoration/comparison
logic finish. Its closure.json verifies **SettingsRestored=True** and
**PackagesPreserved=True**. The controller and debugger exit. No settings values
or credentials are exported; the wrapper changes no repository script or global
Windows setting. ExitCode=-1 reflects the earlier worker termination; ExcelClosed
in this recovered receipt means no live Excel process, not normal unassisted
shutdown or disappearance of the retained Windows process entry.

Windows still enumerates that exited Excel process with its cleanup thread. The
ordinary build/test workflow remains paused pending host recovery; a host restart
is requested from the user after restoration is verified. No new Excel is started.
Private receipts: `test-exited-controller-recovery-result.json` and
`production-paths-controller-recovery.json` under reports/runtime. This supersedes
the earlier pending-settings state, not the recorded assisted-cleanup limitation.

At that RED checkpoint, the final source correction was not built or GREEN. The
intermediate package proves D5 preservation and retains the old evaluator for RED.
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

The subsequent final-candidate section above supersedes those build/GREEN blockers
and identifies the current remaining gates. Release1 and Slice4be remain incomplete.
