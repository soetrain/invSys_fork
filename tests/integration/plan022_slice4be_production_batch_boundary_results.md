# Slice 4be Production batch-boundary diagnostic

Last verified: 2026-09-30 UTC. Architecture v4.11 D12/D13/D18 and Plan022 remain
unchanged. This is diagnostic tooling for the native failure recorded in
[lifecycle continuation evidence](plan022_slice4be_production_lifecycle_visible_results.md),
not an invSys runtime repair or complete Production/Release1 acceptance.

## Designer read candidate: full-flow observation

The final catalog19 candidate fails the ordinary full reusable gate at its first
batch-scale call with RPC800706BE and the retained ntdll/c0000028/offset12d2f
signature. See [designer read evidence](plan022_slice4be_production_design_read_results.md).
The existing native diagnostic forced `-ProductionRunOnly`, omitting full-only
workflows and restart. A bounded tooling extension adds `-FullProductionFlow`
with early native observation, preserving every standard validator statement and
requiring both original aggregate results. The short control retains its exact
argument selection. No runtime package or architectural contract changes.

Offline calibration records19 PASS/two expected failures before implementation,
then21/21 with the same identities. It executes the helper's actual argument
selection and verifies the original full flow, both required aggregates, the
unchanged short control, observer placement, original statements and no VBE
preparation. Roots under `reports/runtime/native-attach-calibration/`:
`9c047e1ef3834cd6803229dbfb684306` (RED) and
`5bf210e40bd142bf8c8d324e79444e5a` (final GREEN). No Excel is opened by calibration.

The observer covers the initial Excel process, including the failing batch-scale
boundary; it does not reattach to the restarted process. The full validator and
Application audit still cover restart. Redacted stack handling is unchanged.
An observed pass cannot establish a native repair or replace the ordinary failed
execution.

The observed full flow passes both aggregates,08:25:49.1094946--08:34:12.6444525
UTC, with zero Excel Application failures. Root:
`reports/runtime/production-batch-boundary/d90653b4ed3b4fa5a953eccb6aef44c9`.
The initial observer is ready at08:25:51.6475630 UTC and observes process exit
at08:33:45.7337990 UTC. There is no c0000028 or second-chance exception and no
stack file; one handled first-chance000006BA occurs during shutdown. Restart
closes four packages and exits unassisted; final closure completes three fixture
workbooks/four packages and exits unassisted. Both release receipts have zero
failures and neither cleanup requests termination. Settings, five package pins
and the original validator remain unchanged. Desktop probes have no error5.

This redacted diagnostic retains the original two aggregate assertions; it does
not independently compare the171 original Boolean observations. No failed stack,
native cause or repair is established. The ordinary failed full gate remains
open. Continue independent regressions rather than repeat this unchanged observed
control as repair evidence.

Refreshed `reports/runtime/production-full-native-static` preserves267 components,
6097 procedures,134205 lines,9 literal/45 unresolved dynamic calls,190 duplicate
candidates and all28 exact size caps. Three schemas and325 PowerShell parses pass.
Runtime source and frozen packages are unchanged; existing compile/layout evidence
retains its scope. The extension introduces no build or form-layout change.

## Recipe structure candidate: fault-stack diagnostic calibration

The unobserved run-only gate at06:33:56--06:34:39 UTC again fails at
`mProduction.RunReusableProductionRunActionContractTest` with RPC0x800706BE and
Excel/ntdll c0000028/offset12d2f. See
[structure evidence](plan022_slice4be_production_recipe_structure_results.md).
The earlier full reusable GREEN does not explain this failure. Post-crash cleanup
also fails while reading `IsAddin`; absence of its final receipts is not successful
cleanup. Desktop probes remain healthy during this failure.

The existing external native observer now captures at most32 redacted stack frames
for c0000028, second-chance exceptions and its disposable calibration exception.
It records only allowlisted module names and relative offsets; other modules and
unmapped frames have unavailable offsets. No absolute addresses, arguments,
symbols, paths, debug strings, process/thread identifiers or memory dumps are
written. Routine filtered first-chance exceptions do not trigger stack capture.
The stopped thread's context is read, never written, and ordinary exception
handling and non-killing detach remain unchanged. Runtime packages are unchanged.

Microsoft identifies c0000028 as an invalid stack encountered during unwind;
that classification does not identify the invSys root cause. The observer uses
the documented x64 context and stack-walking interface:
[status reference](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-erref/596a1078-e883-4972-9bbc-49e60bebca55),
[StackWalk64](https://learn.microsoft.com/en-us/windows/win32/api/dbghelp/nf-dbghelp-stackwalk64),
[x64 CONTEXT](https://learn.microsoft.com/en-us/windows/win32/api/winnt/ns-winnt-context).

Test-first tool calibration, not product behavioral RED/GREEN:

- Stack capture:34 existing checks pass and nine new checks fail before the
  implementation;43/43 pass afterwards, including ordinary exception handling,
  detach, exact output fields, bounded frames and redaction. Roots:
  `reports/runtime/native-exception-calibration/4d2a57656cbe4ab1b6c667640b9dd5ad`
  and `reports/runtime/native-exception-calibration/ceae8593b5b24a38aa595fcc9af01fb7`.
- Late attachment:9 PASS/3 FAIL before implementation becomes12/12. Both default
  early attachment and explicit `-NativeBeforeRun` placement preserve every
  original validator statement, parse without opening Excel, install one observer
  and avoid VBE preparation. Roots:
  `reports/runtime/native-attach-calibration/bdfdcbe08de34ea5b5fcb3b21044ca78`
  and `reports/runtime/native-attach-calibration/be45a80ef12e4d3187c9c8ce30e0b915`.

`-NativeBeforeRun` requires the native observer and attaches immediately before
the reusable run-action macro, after setup and the batch-scale test. The default
early attachment remains available for explicit comparison. An observed pass
cannot establish repair of an unobserved failure; debugger attachment can affect
timing and exception handling. Architecture v4.11, Plan022 behavioral contracts
and the controls catalog's acceptance boundary are unchanged.

The late-attached `-NativeFaultsOnly` run on the frozen structure candidate passes
one aggregate,06:42:55.4224300--06:49:03.2034177 UTC. Controller:
`reports/runtime/production-batch-boundary/9ee73cf343ec4f18b3105ca7c40e5511`.
The observer is ready at06:43:22.6703796 UTC and observes normal process exit at
06:49:02.4369841 UTC. There is no c0000028 or second-chance exception and no stack
file is produced. One handled first-chance000006BA appears during shutdown; it
does not become a second-chance failure. The Application audit records zero Excel
1000/1001/1002 events. Settings, all five candidate hashes and the original
validator are preserved. Closure receipts confirm three fixture workbooks and
four loaded packages closed, zero release failures, unassisted exit and no
termination request. Desktop probes record no error5.

This report deliberately redacts Boolean details, so it does not independently
re-establish the67 prior observed values. It establishes only its original
aggregate assertion under late native observation. No failed native stack was
captured and no cause or repair is established. The ordinary failed run remains
open; do not repeat this unchanged control as proof of repair.

Refreshed `reports/runtime/production-native-stack-static` retains265 components,
6093 procedures,134045 lines,9 literal/45 unresolved dynamic calls,191 duplicate
groups and all28 exact size caps. Three schemas and320 PowerShell parses pass.
No runtime/form source changes, package rebuild or new layout claim is involved.
Unrelated user-document byte hashes remain unchanged.

The later standard run-only execution, after a failed-host cleanup harness
correction, retains the original aggregate and67 exact prior Boolean observations
without a debugger. It closes unassisted and records no Application failures.
This does not explain the intermittent native crash or turn the preceding
observer-assisted result into repair evidence; see
[failed-host cleanup results](plan022_slice4be_automation_cleanup_results.md#recipe-structure-candidate-failed-host-cleanup-receipts).

## Earlier component-candidate controls

**Current component-candidate observation,2026-09-30 UTC:** The unchanged standard
run-only gate now passes its aggregate and retains all67 prior Boolean values;
the full workbench/export/restart gate passes both aggregates with171 observations
(166 distinct pairs), including all67 run-only values. Full chain/live roles/Create
Warehouse also passes32/48/15 with unassisted closure. All preserve settings/packages
and record zero Excel Application failures. Reusable run-only final cleanup and
both full-gate shutdowns explicitly request termination. These scoped GREENs do
not explain earlier native crashes or establish normal reusable shutdown. Exact
controllers, receipts and qualifications are in
[component evidence](plan022_slice4be_production_component_results.md).

**Later cleanup outcome:** The bounded closure/release comparisons lead to an
ordinary-validator correction. The full reusable gate now preserves both aggregate
assertions and the exact171 Boolean observations with unassisted restart/final
exit, settings/package preservation and zero Excel Application failures. This
supersedes the earlier assisted shutdown status, not the unresolved native cause.
See [cleanup evidence](plan022_slice4be_automation_cleanup_results.md) for each
control, the failed500ms comparison and the measured exit times.

**Cold standard-flow reference-release control,2026-09-30:** The repaired isolated
helper is also tested without tracing, VBE preparation or a native debugger.
`-StandardRunFlow -ReleaseAutomationForTest` preserves every original statement
and completes the one standard run-only aggregate PASS. After the original Quit,
it finds zero live references, twelve already-released references and five
already-released variables, with no release failure. Excel remains alive after
the diagnostic30-second observation and the unchanged final fallback terminates
it. Settings, all five component-candidate packages and the original validator
are preserved; zero Excel Application events occur. This does not establish
normal reusable shutdown. Generation calibration:
`reports/runtime/production-batch-boundary/0c6e33e4b8d6436e8f90694fbea0e671`;
execution `620289cb150b475db2a4d4d33e30f1fa` in that same parent,
01:13:32--01:18:08 UTC. The diagnostic report redacts observations, so its single
aggregate must not be presented as a new comparison of the67 Boolean values.

**Later boundary observation,2026-09-29 20:02 UTC:** The isolated
`validation-production-uom-staging` candidate changes only Operations
`modProductionUomCatalog` among248 compiled components. Its full-chain run passes
projection recovery and two Production batches, then fails at **Run Shipping
BtnShipmentsSent** with RPC0x800706BE and the same ntdll c0000028/offset12d2f
signature. Chain5/1, live roles40/1, Create Warehouse15/15; the recovery process
required assisted termination before settings/package/report restoration. This
extends the observed boundaries; it does not establish a common cause or implicate
the UOM correction. Do not infer that the native failure is confined to the batch
adapter or rerun the unchanged chain as a repair. Exact evidence is linked in
[UOM preservation results](plan022_slice4be_production_uom_staging_results.md).

The failed full gate stopped at `mProduction.RunProductionBatchScaleContractTest`
with RPC0x800706BE and ntdll.dll0xc0000028. That adapter re-enters the public
launcher, reads the form's visibility, then calls the typed scale test. Its single
progress marker did not distinguish those operations.

## Current instruction candidate: standard-flow comparison

The frozen `validation-production-instructions-typed` candidate fails the
unmodified standard run-only gate at the same batch-scale boundary: RPC0x800706BE,
Excel ntdll.dll0xc0000028, and no requested termination. The earlier67 reusable
observations are not re-established by that failed execution.

`Test-ProductionBatchBoundary.ps1 -StandardRunFlow` retains the complete standard
run-only validator instead of cutting after scale. Initially it required either
`-TraceBoundaries` or the separate `-CompileOnly` VBE preparation control; later
explicit native-observer and post-report cleanup controls are also supported.
Generation reverses declared insertions/redaction and verifies every original
validator statement is preserved; both modes parse before Excel starts.
`-GenerateOnly` calibrates this without invoking Excel. The existing31-stage
placement/redaction calibration retains70/70.

| Standard run-only control | Result | Limitation |
|---|---|---|
| Tracing and VBE preparation |1 aggregate PASS;139 allowed markers; original55-entry boundary sequence retained exactly. |Four in-memory compiles; automatic final termination requested. |
| VBE preparation without tracing |1 aggregate PASS; zero trace entries; all four compile commands execute. |Automatic final termination requested. |

Both preserve the original validator, settings and all five frozen packages and
record zero Excel Application events. These diagnostic passes neither repair the
cold failure nor establish normal shutdown or full/restart acceptance. VBE
preparation and compilation remain possible influences; causation is unproven.
`Test-PackagedVbaCompile.ps1` opens packages read-only and closes without saving:
its compile gate verifies compilability, not a persisted compiled artifact.
A separate saved-compile experiment also fails; saving compilation is not an
established repair. No runtime behavior or Architecture v4.11 contract changes.

The isolated `validation-production-instructions-persisted` experiment builds,
compiles and saves all five packages. All248 component source hashes match the
frozen typed candidate exactly; all five package byte hashes change. The original
candidate, runtime source, compile tool and settings remain preserved. Its cold,
unmodified run-only validator fails0 PASS/1 HARNESS at the initial
`mProduction.BtnOpenProductionForm` invocation, before batch scale, with
Excel ntdll.dll0xc0000028 at offset0000000000012d2f. No HRESULT is recorded for
this attempt. Final cleanup requests no termination; this is a native crash,
not normal shutdown or a behavioral RED. Do not promote the experimental set.
The cause remains unresolved; continue a bounded native/COM investigation without
adding VBE preparation, retries or sleeps as runtime workarounds.

- Cold failure: `reports/runtime/production-instructions-reusable-attempt-verification.json`.
- Placement calibration: `reports/runtime/production-batch-trace-placement/bbf4ff2b8ebb4e8eb0b84d01ee51327e`.
- Traced standard generation: `reports/runtime/production-batch-boundary/bf6cd95b28144b2bba6b5d445f9651ba`.
- Scoped generation retained: `reports/runtime/production-batch-boundary/66defd7c93214e5eb37718b2ed4cb2a9`.
- VBE-only generation: `reports/runtime/production-batch-boundary/61ee7f92b1b74771820d96f33730a275`.
- Traced standard execution: `reports/runtime/production-batch-boundary/d01d9b1b41394f60bb6eaa28290e2c82` (18:39:19--18:43:46 UTC).
- VBE-only execution: `reports/runtime/production-batch-boundary/2884dd7554454dc38a9adce47843ca69` (18:46:00--18:49:33 UTC).
- Each execution retains generation, closure, result and final-cleanup receipts.
- Saved-compile build: `reports/runtime/production-instructions-persisted-build/verification.json` (18:52:36--18:53:18 UTC).
- Saved-compile cold failure: `reports/runtime/production-instructions-persisted-cold-verification.json` (18:53:52--18:54:14 UTC).

## Native exception observation

`NativeExceptionObserver.cs` attaches only to the owned disposable process selected
by the diagnostic. It writes exception codes, first-chance status, allowlisted
module names and relative offsets. Unknown modules are redacted. It does not write
memory dumps, exception parameters, debug strings, absolute addresses or paths.
It handles the initial attach breakpoint and otherwise preserves ordinary exception
dispatch; detach does not kill the target. The implementation follows Microsoft's
[debug attach](https://learn.microsoft.com/en-us/windows/win32/api/debugapi/nf-debugapi-debugactiveprocess)
and [exception continuation](https://learn.microsoft.com/en-us/windows/win32/api/debugapi/nf-debugapi-continuedebugevent)
contracts. Attach changes timing and is diagnostic only.

The disposable-process calibration passes34/34: known first-chance delivery,
continued target handling, module-offset resolution, redaction, normal exit,
non-killing detach and declared filtering. The optional `-NativeFaultsOnly` mode
omits three observed first-chance codes only (`E06D7363`, `40080201`, `00000005`);
every second-chance exception and all other codes remain observable. The last of
these is a raised exception code, not a failed desktop-access API result.
Two initial compiler-setup failures are harness preparation, not product RED.
Existing trace placement retains70/70; all301 PowerShell scripts parse.

The first native-observed standard run passes1 aggregate workflow check with no
VBE preparation or source instrumentation, no cut and all original validator
statements retained. It records368 first-chance exceptions, no second-chance or
`C0000028` exception, and zero Excel Application events. Final cleanup still
requests termination. The original validator, settings and five package hashes
remain preserved. This does not reproduce the cold crash or prove normal shutdown;
explicit VBE preparation is not necessary for this diagnostic pass. The filtered
comparison also passes1 aggregate check, with three first-chance `80010108`
records and no second-chance/native fatal exception. It preserves all original
statements, validator/settings/package hashes and zero Excel Application events,
but again requests final termination. Neither observer mode reproduces the cold
failure. Do not repeat these passing diagnostic variants as evidence of a repair.

Separately, a disposable blank-workbook calibration retains five child COM
references after Quit and application release/GC. Excel remains alive after two
seconds, then exits without termination after those references are released.
No packages or saved workbook are involved. This proves a harness reference-lifetime
mechanism on this host, not the cause of the Production crash or Settings retention.
Subsequent packaged cleanup controls below retain this distinction.

- Final calibration: `reports/runtime/native-exception-calibration/337246656f1f4ee6b6e1caa8c4eb9158`.
- Trace placement: `reports/runtime/production-batch-trace-placement/3486fb84cb9a49478ce340c024bde556`.
- Native generation: `reports/runtime/production-batch-boundary/2adbe728aa7a4c3993b1ce3e91d0a57b`.
- Native observed run: `reports/runtime/production-batch-boundary/c4910cec7d9947b9977590cc66828ae6` (19:03:39--19:08:36 UTC).
- Filtered generation: `reports/runtime/production-batch-boundary/47e1fcea86044debac00f07004715ea5`.
- Filtered observed run: `reports/runtime/production-batch-boundary/f96a29127e2e4372948ca4f2dcbfe6c8` (19:09:00--19:13:41 UTC).
- Blank-workbook reference calibration: `reports/runtime/excel-reference-release/6c22f8e0e0f5459090d267197c1d351f`.

## Automation cleanup controls: no packaged repair

`IsolatedAutomationCleanup.ps1` is restricted to completed isolated workers that
own every supplied COM reference. It traverses script variables and collections,
deduplicates aliases, handles cycles/multidimensional arrays, skips already released
references and reports counts only. The disposable blank-workbook calibration
passes8/8, including ordinary process exit without termination. Expanded cases
first exposed unsupported two-dimensional indexing and a released RCW; these are
tool calibration failures, not invSys behavioral RED.

`-ReleaseAutomationForTest` runs after the original workflow report and Quit,
then observes exit for30 seconds before the unchanged termination fallback. The
optional `-ClearErrorReferencesForTest` clears only the disposable worker's error
records after assertions/reporting. It never suppresses workflow errors. Generated
validation reverses these declared insertions and preserves every original
statement. Exceptions in the diagnostic step produce sanitized type/code/line
metadata and still reach the original cleanup fallback.

| Packaged scoped control | Workflow | Cleanup result |
|---|---|---|
| Initial helper / matrix correction |7/7 in each attempt |No release receipt; explicit assisted cleanup. Original worker error transport truncated. |
| Sanitized failure receipt |7/7,55 trace entries |Released RCW encountered while collecting variables; automatic fallback termination. |
| Released-reference correction |7/7,55 trace entries |Zero live references found,12 already released references and5 already released variables; still alive after30 seconds, automatic termination. |
| Post-report error-record control |7/7,55 trace entries |Same reference counts;33 error records cleared; still alive after30 seconds, automatic termination. |

All retain settings, five frozen packages and the original validator. The last
three retain zero Excel Application events. No cleanup control establishes normal
packaged shutdown or repairs the cold native crash. Do not promote either cleanup
variant into the ordinary validator on these results. All303 PowerShell scripts
parse; runtime VBA/source packages remain unchanged. Continue independent coverage
with these release gates explicitly open rather than repeating these comparisons.

- Final cleanup calibration: `reports/runtime/isolated-automation-cleanup/f22bd80633fd4d0d92a9024ab8759ab4`.
- Initial scoped attempts: `reports/runtime/production-batch-boundary/77cfa5e07d704128ae29f5a0d8c94320` and `479a333a4bf947838b2f7c6d26124786`.
- Sanitized failure: `reports/runtime/production-batch-boundary/febd43f45a4e4c099b0148edc35e0d58`.
- Corrected reference control: `reports/runtime/production-batch-boundary/b5a92fe7f6154a76ad2e016cc3158aec` (19:25:07--19:26:01 UTC).
- Error-record control: `reports/runtime/production-batch-boundary/0120112c573b492cb17d580e6e8724cb` (19:27:19--19:28:13 UTC).

## Earlier scoped boundary evidence

`ProductionBatchBoundaryTrace.ps1` inserts31 fixed markers in unsaved Operations
memory, across that adapter, `BtnOpenProductionForm`, `ShowProductionForm` and
`frmProduction.TestBatchScaleContract`. It resolves all anchors before mutation,
compiles the four loaded workflow packages before forms, and never saves an XLAM.
It records no runtime values. Its logger rejects unknown stage names before any
write; the live transport calibration proves this with an unregistered value.

`Test-ProductionBatchTracePlacement.ps1` passes70/70: exact statement placement,
fixed unique stages, unchanged original statements, VBA case normalization,
missing/duplicate-anchor rejection and the bounded logger allowlist. This is
tool calibration, not product behavioral RED/GREEN.

`Test-ProductionBatchBoundary.ps1` derives a scoped copy of the actual standard
validator. It retains initial/repeated public launches and the original scale
assertion, stops immediately afterward, redacts report details, and retains the
existing cleanup decisions. Seven checks cover setup, launcher calls, saved
station-local workbook, reuse and exact scale results. The scope excludes later
Production workflows and restart.

| Control | Result | Cleanup and limits |
|---|---|---|
| Traced |7/7, four instrumented compiles,55 allowlisted stage entries; all three launcher entries and scale calculation return. |Automatic final termination requested. Settings/packages/validator preserved; zero Excel Application events. |
| Untraced |7/7 with the exact same check identities and no VBA instrumentation. |Automatic final termination requested. Settings/packages/validator preserved; zero Excel Application events. |

The55 traced entries match the complete expected ordered success sequence,
including all three public launcher entries, with no exception branch recorded.

Both controls observe three newly opened workbooks, exactly one station-local
operator workbook and zero new workbooks on the second launch. An earlier
diagnostic incorrectly required one new workbook overall and recorded6 PASS/two
failures, including its prerequisite-stop marker. The original validator instead
requires one station-local operator workbook; correcting the diagnostic preserves
that existing contract. This was a harness assertion error, not product RED.
Earlier JSON-array, anchor and output-path setup failures occurred before product
callbacks and are likewise not meaningful RED.

The traced and untraced controls run16:53:19--16:54:32 UTC. Their successful
boundaries do not explain the preceding crash or establish normal unassisted
shutdown. Keep the earlier failure and automatic-termination qualifications.
The subsequent standard full run at16:55:16--16:55:39 UTC stops0 PASS/one harness
failure at the same batch-scale adapter, with RPC0x800706BE and an Excel
ntdll.dll0xc0000028 Application Error. Full-only steps and restart are unreached.
The new cleanup receipt records TerminationRequested=False; the native failure
precedes cleanup. Settings/packages restore and Excel is closed. This is not
normal shutdown or full acceptance. Stop unchanged broad retries: the pair did
not reproduce or explain the failure. Further diagnostics must distinguish the
failed execution itself rather than infer a repair from successful scoped runs.

A read-only native-callback source review finds no call site for
MouseScroll.EnableMouseScroll outside its own unused declarations/comments,
consistent with the earlier Receiving-denial investigation. That lead does not
explain this failure and does not authorize deleting the module.

## Exact evidence

- Offline calibration: `reports/runtime/production-batch-trace-placement/2d2ca67c5db44d50b5e395b9bc36e891`.
- Invalid first workbook assertion: `reports/runtime/production-batch-boundary/c0b1a9daa23f44689f6041bdbaecd82d`.
- Traced control: `reports/runtime/production-batch-boundary/d18ba441a91f4bdfb619a11b3b734425`.
- Untraced control: `reports/runtime/production-batch-boundary/40e254ac78c6477fadf045865836333a`.
- Pair verification: `reports/runtime/production-batch-boundary-pair-verification.json`.
- Exact ordered stages: `reports/runtime/production-batch-boundary-stage-verification.json`.
- Full follow-up: `reports/runtime/production-paths-regression/reusablefull-3e7d9de859b442b4ba7be665d6ff8fea` and `reports/runtime/production-batch-full-followup-verification.json`.

The candidate remains `deploy/validation-production-paths`, pinned to the five
hashes in `production-lifecycle-native-controller/65e8e27585ca47289cf27e87e147e83e/package-pins.json`
under runtime reports. No Windows policy, normative contract, runtime VBA or
operational workbook changes. The user's conditional instruction remains: stop
the goal if desktop error5 returns. No such error is observed in these runs.

Refreshed `reports/runtime/production-batch-boundary-static` retains253 components,
6058 procedures,133166 lines,9 literal/45 unresolved dynamic calls,191 duplicate
groups and the exact28 previous size limits. All three schemas and295 PowerShell
parses pass. Runtime/form source is unchanged; no product rebuild or geometry
change follows from these diagnostic helpers. Prior compiled and layout evidence
retains its scope. Unrelated user files and the five package hashes are preserved.
