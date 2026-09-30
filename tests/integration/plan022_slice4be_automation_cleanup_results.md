# Slice 4be isolated automation cleanup

Last verified:2026-09-30 UTC. This is test-harness work under unchanged Architecture
v4.11/Plan022, not an invSys runtime repair or product behavioral RED. The five
catalog16 packages and VBA source were unchanged at the original checkpoint.
The later catalog17 timing comparison below also changes no runtime code or XLAMs.
All runtime paths below are
under ignored `reports/runtime/`; only sanitized findings are committed.

## Recipe structure candidate: failed-host cleanup receipts

The unobserved run-only attempt at06:33:56--06:34:39 UTC crashes at
`mProduction.RunReusableProductionRunActionContractTest`, with RPC0x800706BE and
Excel/ntdll c0000028. Afterwards, reading a missing `IsAddin` property escapes
the validator's finally block before workbook closure, final reference release
and final cleanup receipts are written. This secondary harness defect does not
explain the native crash. Exact failure evidence remains in
[structure results](plan022_slice4be_production_recipe_structure_results.md).

`Test-ReusableFailedHostCleanup.ps1` executes the actual final cleanup block
against disposable host doubles: missing workbook metadata, missing package
metadata, and a successful empty host. It opens no Excel process. Before the
correction,7 checks pass and14 fail; afterwards21/21 pass with the same identities.
The tests require the original workflow failure to remain intact, a truthful
redacted closure state, an explicit failing cleanup row, and continued Quit,
reference-release and final-receipt bookkeeping. These are harness calibration
results, not product behavioral RED/GREEN. Exact roots:
`reusable-failed-host-calibration/357efb5cbfe848d0812cf21189230cea` and
`reusable-failed-host-calibration/8b100521a8424b5e8648119d915c4d64`.

The ordinary validator now writes `Completed=False` and only the failure stage,
exception type and HRESULT if owned workbook/package closure fails. It retains
the original workflow evidence and adds `HARNESS_CLEANUP` as a failed check, so a
cleanup failure cannot become a passing result. It continues the existing
best-effort closure, reference release and cleanup receipts; ownership validation
is unchanged. Successful closure still clears owned references as before.
A receipt that a process is absent is not proof of normal exit after a crash.
Application-event evidence and the primary workflow result remain required.

The existing live ownership/closure calibration retains7/7, including whole-set
validation before closing anything, foreign fixture preservation, unassisted exit
and no forced termination: `reusable-cleanup-calibration/5a1012fafaca47b3834e623c8f2e35ca`.
The native-observer generator retains12/12 against the changed validator:
`native-attach-calibration/789c057889d2403d98f43c09eadcc8bf`.
No runtime VBA, package, control contract or Windows policy changes.

The standard run-only gate with no debugger, tracing or VBE preparation retains
one aggregate and all67 prior Boolean observations in exact order:
`production-recipe-structure-regression/reusable-25e3f0ba041d4964b42cda948ce7a1c1`,
06:52:45.0267790--06:58:45.7021887 UTC. The new closure receipt reports
Completed=True, no failure, three fixture workbooks and four loaded packages
closed. Final release has zero failures, waits2616ms for unassisted exit, and
requests no termination. Settings and all five package hashes are preserved;
the full interval plus12-second Application audit finds zero Excel failures.
Desktop probes record no error5. This verifies the final-cleanup path with the
real packaged workflow; it neither reproduces nor repairs the earlier native
crash. No unchanged broad retry is claimed as a repair.

Refreshed `production-failed-host-cleanup-static` retains265 components,6093
procedures,134045 lines,9 literal/45 unresolved dynamic calls,191 duplicate groups
and all28 exact size caps. Three schemas and321 PowerShell parses pass. The
existing candidate's compile/layout evidence retains its original scope; no
runtime rebuild is warranted by this harness-only change.

## Smoke shutdown deadline on the catalog17 candidate

The ordinary smoke gate retains86/86 but its Final host exceeds1000ms and requests
termination: `production-recipe-order-regression/smoke-50b1bc571fe042f386f9b91969adf20c`,
02:42:13--02:42:34 UTC; shutdown root
`packaged-smoke-closure/c0848cb9486c4abf812d6ec6fded1ac2`. Initial exits normally,
both phases release19 references with zero failures, settings/packages/report
restore, and zero Excel Application1000/1001/1002 events occur. This is an assisted
cleanup result, not product RED or proof of a native crash.

A disposable validator copy changes only the wait and reported deadline to30000ms.
It retains86/86 and both hosts exit unassisted; Final requires2323ms. Controller
`production-recipe-order-regression/smoke-e0544173c25f496aa4fdb31a4d69d13b`,
02:44:09--02:44:31 UTC; shutdown root
`packaged-smoke-closure/3c73c10c8954447c98ed09327bd48647`. Generation receipt
`production-recipe-order-smoke-wait/generation.json` proves reversing those two
constants reproduces the original validator; its source hash was verified intact
after this comparison and before the ordinary change. All prior assertions and
owned-workbook/process guards remain; no other behavior is changed.

The ordinary validator now uses the same30-second bound. It again retains86/86,
all prior check names, settings/five packages/tracked report and zero Application
failure events. Initial/Final exit unassisted, with Final observed after2315ms:
`production-recipe-order-regression/smoke-86beb00f94b84131abd7d7aa9fe95da9`,
02:45:17--02:45:40 UTC; shutdown root
`packaged-smoke-closure/b0609a0ac3ac4d3d8d734128acb3ced0`. Both phases release19
references with no failures; ordinary source matches the calibrated comparison.
Each controller has a strict `verification.json`. Earlier1000ms evidence keeps
its original deadline and result. This measured timing correction does not explain
the older native crashes or complete Slice4be acceptance.

## Defect and protecting calibration

Smoke retains86/86 but its newly observed Initial and Final cleanup both request
termination after the existing1000ms wait. Controller
`production-components-regression/smoke-46bbf224c98e47aa809dc9100711423c`; receipt
`production-components-smoke-baseline-verification.json`.

The initial attempt to reuse completed-host reference release fails74/1 after
confirming zero workbooks; it is removed and safely restored. That attempt does
not identify the primary failure line. Its exact preservation/cleanup evidence
is retained in [component results](plan022_slice4be_production_component_results.md).

`Test-IsolatedAutomationCleanup.ps1 -IncludeExpiredNestedReferences` adds a closed,
released blank workbook to a list and dictionary alongside aliases, a matrix and
a cycle. The unchanged helper reproduces `InvalidComObjectException` at its
two-object `ReferenceEquals` call, line21; calibration
`isolated-automation-cleanup/27416b8e1f664590a7355febda2622dc`.
A self-comparison guard is insufficient: `247479d39ead4cebbca4d2f17166868b` fails
at the same comparison, shifted to line25. Neither failed approach is retained.

The helper now tracks visited object identities through .NET
[ObjectIDGenerator](https://learn.microsoft.com/en-us/dotnet/api/system.runtime.serialization.objectidgenerator?view=netframework-4.8.1),
which identifies already-seen object references. No serialization or object IDs
are written. Packaged observation then identifies another invalid-wrapper failure
at the initial null comparison, line17. Fixed metadata receipt:
`packaged-smoke-closure/f1050cc3f1b34d949f7eb728ed793dd6/Initial-reference-failure.json`,
controller `production-components-regression/smoke-42b228fb06e04224bc64723f77ac4efe`.
The helper also skips `InvalidComObjectException` during retrieval/classification.
It retains alias/cycle handling, unique-reference accounting, repeated-release
safety and fixed counts only. It remains restricted to owned, completed hosts.

The final original calibration is **8/8**, root
`isolated-automation-cleanup/dbeaaf012cbb451a9970252612af50fd`; the expanded calibration
is **9/9**, root `isolated-automation-cleanup/ab92a4316e6c4cba9586065ff94011a5`.
Both exit unassisted with zero release failures. The original five-reference
expectation is unchanged. The added fixture identifies six references; its initial
five-reference expectation was corrected only for that opt-in case. The intermediate
`2307af8e2eaa4c65a9e9baf8e32b4f85` already exited normally but failed that stale count.

## Packaged smoke outcome

With all old workbooks closed and Quit returned, the smoke harness releases its
completed-host variables before the unchanged1000ms wait; replacement Excel is
created afterward. It records release counts, normal/forced exit and bounded
failure metadata (phase, exception type/HResult, source filename/line), never
fixture values, credentials or raw exception text in these sidecars.

Controller `production-components-regression/smoke-d3dde992de54428cbc3914a076ace473`,
01:04:32--01:04:53 UTC, retains **86/86** and every prior check identity. Both hosts
exit unassisted, with no termination request. Each releases19 references with zero
release failures. Settings, all five packages and the tracked report are preserved;
no Excel Application1000/1001 event occurs. Receipt
`production-components-smoke-corrected-verification.json`; shutdown evidence
`packaged-smoke-closure/e22e4091148649aaab409872bc3ad1fc`.

The preceding traversal-only packaged attempt
`production-components-regression/smoke-902db509280a455ea67d5b87ae585b12` fails74/1
before a primary-location receipt exists; the instrumented attempt above also
fails74/1. Both restore settings/packages/reports and leave no Excel process.
These are harness failures, not product RED or acceptance results.

The existing Settings restart consumer retains **202/202** and every preceding
identity, five instrumented compiles, unassisted internal restart/final exit and
settings/package preservation. Controller
`production-components-regression/settings-0d7b5cb67645400fb6378b45cc8d1e13`,
01:05:41--01:09:32 UTC; result
`slice4be-tracking-settings/98bd8df222e84d4faaeb779f6961e1b1/green.json`. Internal
cleanup releases10 references with zero failures and no termination request.
No Excel Application failure event occurs. Receipt
`production-components-settings-cleanup-regression.json`.

All310 tracked PowerShell scripts parse. The smoke workflow-preservation proof
retains every original validator statement and the1000ms wait; only completed-host
release/metadata are inserted (`packaged-smoke-cleanup-workflow-preservation.json`).
Normal reusable Production shutdown, earlier native crashes, comprehensive
coverage, guide transfer and human/NAS acceptance remain separate open requirements.

## Reusable Production comparison

The repaired helper alone does not resolve reusable Production shutdown. A cold
standard run-only diagnostic preserves every original workflow statement and
passes its aggregate without VBE preparation, tracing or a native debugger.
After Quit it finds no live references to release (twelve already released
references and five already released variables), with no helper failure. Excel
remains alive for the30-second observation and requires the original termination
fallback. Settings, five frozen packages and the original validator are preserved;
zero Excel Application events occur. Root
`production-batch-boundary/620289cb150b475db2a4d4d33e30f1fa`,
01:13:32--01:18:08 UTC. This diagnostic's redacted aggregate is not a new67-value
comparison. See [boundary evidence](plan022_slice4be_production_batch_boundary_results.md).

An independent cold standard run-only control closes the exact captured disposable
operator workbook before the add-ins, using ordinary Workbook.Close. Events are
enabled, exactly one owned workbook matches and the workbook count decreases by
one. The workflow aggregate passes, but final cleanup still requests termination.
Thus workbook-first closure alone is not a sufficient repair. The diagnostic
neither changes the default validator nor unloads forms through an injected macro.
All original workflow statements, settings and five packages remain preserved;
zero Excel Application events occur. Root
`production-batch-boundary/74240f34fafd442f80b3e156fce9f714`,
01:20:13--01:24:22 UTC; fixed result `operator-first-cleanup.json`.
Generation-only calibration `production-batch-boundary/817e077ff6cc4bae9e45264cd0714379`
and the execution's generation receipt both preserve original statements and parse.
No runtime, permission, authority or architectural change follows from this control.

Subsequent untraced, cold scoped comparisons retain the same seven launcher/scale
checks, original workflow statements, settings and five packages. Every comparison
records zero Excel Application events. All roots below are under
`production-batch-boundary/`; none is a full reusable acceptance run.

| Diagnostic root | New observation/control | Shutdown result |
|---|---|---|
| `7f8a41814a134033abcfea77c2aaf56f` | Observe existing close calls: three workbooks remain before Quit; three close calls return and four fail; Quit returns. | Forced. |
| `107ca6a08c2e4cc1a41c7825e383a228` | Validate every workbook against isolated roots before closing any; all three owned fixture workbooks close, leaving zero before Quit. | Forced; the same add-in close failure remains. |
| `f35ac1f629dd44788f796565dabf1474` | Allowlisted package attribution identifies `invSys.Core.xlam` with HRESULT0x800A03EC; three old fixture references report0x80010108. | Forced. |
| `1dfacd71d24544f0b5e01fbb673c02bc` | Close the existing opened list in reverse order after the owned workbooks. All four package closes now return; only the old fixture-reference errors remain. | Forced within the original500ms allowance. |
| `ebf63b1acb874066a1e2f27d59cc9988` | Add the completed-host reference helper/GC after those closures. It finds zero live references,19 already released references and five already released variables, with no release failure. | Unassisted within the diagnostic30-second observation; this alone does not validate the default500ms allowance. |

Observers write counts, flags, allowlisted package names and HRESULTs only.
Original swallowed close errors remain swallowed by the diagnostic; the observer
does not turn them into workflow failures. The owned-workbook control validates
all paths before closing any workbook and never disables events. Reverse closure
changes only declared post-report cleanup order, and generation reverses that
change when checking original-source preservation. Native-crash causation remains
unresolved; these are cleanup controls, not a native runtime repair.

The no-extra-wait control `1b3e1b96506b45749617aa518c242b3b` retains7/7 but still
requests termination after the original500ms allowance; do not count the preceding
30-second control as a500ms pass. The complete cold standard run-only control
`e1b86f46ad1d4bdaa63ea677c72831e4` then passes its aggregate and exits unassisted
after an observed **2514ms** wait (01:37:38.521--01:37:41.040 UTC). It closes all
three workbooks and all four packages, preserves original statements/settings/five
packages and records zero Excel Application events. Run interval
01:33:28--01:37:41 UTC. This supports a bounded test-harness exit wait; no runtime
latency or native-crash repair is asserted.

`Test-ProductionReusableCleanup.ps1` passes **7/7** in
`reusable-cleanup-calibration/eb8dd5cb8ae041d5845a247bf571e139`. A saved workbook
outside the expected runtime causes refusal before either workbook closes. The
owned workbook remains open and the outside fixture's sentinel value is unchanged.
After explicitly closing that separately owned calibration fixture, normal owned
closure and process exit pass without forced termination or release failure.

## Ordinary reusable validator

The ordinary `ProductionReusable` route now closes its remaining owned fixture
workbooks before add-ins, closes the four add-ins in reverse dependency order,
releases completed-host references and waits for exit with a30-second upper bound.
The internal restart preserves its existing workbook-save loop and applies the
same package closure/release/wait. The existing750/500ms fallback checks and
termination receipts remain after the bounded wait. Other workbook-state modes
and the separate palette diagnostic keep their existing cleanup paths.

The change inserts22 lines into the validator and removes no original statement;
`production-reusable-cleanup-workflow-preservation.json` verifies exact restoration
and unchanged runtime source. New helper `ProductionReusableCleanup.ps1` writes
fixed counts/exit flags/times and bounded failure metadata only. There is no
architectural contract change or new product behavioral RED claim.

The full ordinary workbench/export/restart gate passes **2 aggregate assertions**
and preserves the exact ordered **171 Boolean observations /166 distinct pairs**
from the assisted baseline, including all67 run-only observations. Both hosts exit
unassisted, with no termination request or reference-release failure: restart waits
**541ms**, final exit **2465ms**. Settings and five packages remain preserved;
zero Excel Application1000/1001/1002 failures are found. Controller
`production-components-regression/reusablefull-6c2580352bb541dd94207a159b6da7dc`,
01:39:04--01:45:32 UTC; receipt
`production-components-reusable-full-cleanup-verification.json`.

The ordinary run-only route also passes its aggregate with all **67 prior Boolean
observations in the exact same order**. Final exit is unassisted after a2657ms
wait, with no release failure or termination request. Settings/five packages are
preserved and zero Excel Application failures occur. Controller
`production-components-regression/reusable-e2c93ba3f4d1482abf6235cbadb2c153`,
01:46:21--01:50:34 UTC; receipt
`production-components-reusable-run-cleanup-verification.json`.

Fresh `production-reusable-cleanup-static` retains259 components,6078 procedures,
133764 lines,9 literal/45 unresolved dynamic calls and191 duplicate groups. All28
module caps are identical, all three generated schemas pass and all312 PowerShell
files parse. The frozen runtime source/packages have no changes, so prior build,
compile, layout, full-chain, live-role and visible evidence retains its scope.
Earlier native crashes remain unexplained; comprehensive coverage, guide transfer
and human/NAS acceptance remain open.
