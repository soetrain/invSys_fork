# Plan 022 Slice 4be: Settings editor observations

Architecture v4.11 D18 specifies 26 Settings controls in catalog 11: ten Tracking,
eight Event Detail and four personal-preference controls on each of Admin and
Operations. Documentation commits `889022f` and `3fb0f7a` precede implementation.
The latter names exact diagnostic terminal facts; those evaluator changes remain
separate from the observation implementation described here.

This is an **isolated implementation checkpoint**, not completed Slice 4be or
Release 1 acceptance. Diagnostic mapping passes 780/780, and the compiled legacy
Settings regression passes 202/202 with ten accepted captures. Broader role,
comparison and release gates remain pending. The large diagnostic run's assisted
cleanup remains explicit. D8-A and Event Detail unlocking remain unapproved.

## Observation RED and first GREEN

The frozen `deploy/validation-owner-command-completion` predecessor supplies
actual callbacks, existing owner services and disposable Admin-generated fixtures.
Probes are installed and all five packages compiled before forms are created.
Input assignment suppresses automatic form events; each tested private callback
then executes exactly once. Core setup and initialization remain unobserved.

The corrected expanded RED is **179 PASS / 141 expected FAIL**, 2026-09-22
**17:20:39.2213336--17:28:48.2219557 UTC**, at
`reports/runtime/slice4be-settings-activity/497c06807af84490826fc0d53ca177e9/red.json`.

The failures comprise 37 catalog checks, 66 missing observation-pair/fact checks,
one complete-run check, one policy-save request check and 36 stale-editor checks.
All 34 independent owner checks and both real Excel save-cancellation checks pass.
The five frozen package hashes remain unchanged. Source edits following the
completed initial callback RED do not alter this frozen test predecessor.

The first implementation candidate, `deploy/validation-settings-editor-activity`,
passes **320/320**, preserving every corrected RED check identity and resolving
all 141 failures, **17:31:31.5619921--17:41:19.9843656 UTC**. Its report is
`reports/runtime/slice4be-settings-activity/8afc2b35e02644fd8f7012d1b9e22d49/green.json`.
Operations cold start and all five package compiles pass. All 299 runtime-source,
186 test-source and five package pins are unchanged at gate completion.
Excel closes normally without force or an additional Quit; build-through-focused
Application events 1000/1001/1002 total zero. All five owned-form captures are
directly reviewed, fully painted and readable.

The actual handlers protect selections, unsaved staging, Reset, Reload, appended
profile versions, changed and unchanged preferences, stale expected versions,
uncertain cancelled saves, and a signed-in Operations user without ADMIN_MAINT.
Twenty-five actions produce 50 observations in a complete stopped run. Successful
policy save retains its original eligible REQUESTED observation and interrupts
recording with lifecycle **Incomplete**, reason **POLICY_CHANGED**. It cannot
manufacture an ordinary COMPLETED observation across policy versions.

All 24 stale-editor cases cover Reset/Reload in four sections after sign-out,
same-user reauthentication or another warehouse selection. They preserve held
staging and require reopening, with no cross-context activity or Config writes.

## Additional boundary evidence

The first expanded boundary run is **415 PASS / 26 incorrect test assertions**
at `reports/runtime/slice4be-settings-activity/7f4e4ec9e2b74af4bb917b26b856d28d/red.json`.
The older-policy assertion mistakenly expected a successful read for an absent
control. Existing policy authority correctly returns `False|False|False|0`:
unavailable, not collected or visible. The actual preference save still succeeds
without retroactive collection. Correct the test; do not change this contract.
All seven captures from that run are reviewed and readable.

Unsupported profile schema correctly uses the disclosed display-default path,
disables saving, preserves Config bytes and records only the deliberate failed
Reload. It does not reproduce a synthetic-selection defect.

A separate actual-handler diagnostic retains the same session while revoking
ADMIN_MAINT in disposable authority, then performs two denied saves and Reload
during explicit recording. It supplies **166 PASS / five behavioral FAIL** at
`reports/runtime/slice4be-settings-activity/ea71e4623630497c9082594ac5b04a0c/diagnostic-settings-safety-red.json`.
All five failures concern the extra synthetic selection emitted when the denied
Detail Reload clears its list. They are not credential disclosure or a Config
write. The fix suppresses form events around that internal clear and retains the
deliberate Reload's FAILED/Unchanged observation. This diagnostic closes Excel
immediately and normally, preserves all five frozen packages, and has zero
matching Application errors. Three supporting captures are reviewed.

The follow-up candidate `deploy/validation-settings-editor-safety` passes
**450/450**, retaining all initial 320 and expanded 441 check identities and
resolving the five synthetic-selection failures. Its report is
`reports/runtime/slice4be-settings-activity/ae71a06b56f04bf892ddb6ddf46d1704/green.json`,
**18:01:29.9158277--18:12:05.2720254 UTC**, 2026-09-22. Operations cold start and
all five compiles pass. All 299 runtime, 187 test and five package pins are
preserved at completion. Eight captures are directly reviewed and readable,
including denied Reload, unavailable profile and separate owner/tracking status.
Excel closes normally without force or an additional Quit after a delay; matching
Application events 1000/1001/1002 total zero from build through completion.
The diagnostic terminal-map gate is separate and remains pending.

## Diagnostic terminal-map RED

The frozen safety candidate supplies **753 PASS / 27 expected FAIL / 780 checks**
at `reports/runtime/slice4be-settings-activity/3ca342d2cc6d4ad4bf4f2b0057a7b1c9/red.json`.
All 450 observation GREEN identities remain passing. Forty-six actual
expectation-editor/Evaluate cases cover 25 positive control mappings, two
UNCHANGED preference saves and 19 negative cases. Only the 27 positive terminal
classifications fail: their exact original facts, authored intent, selected
journal version/hash and loaded publication all match. All 54 captures are
directly reviewed and readable. Evaluations preserve original fixture bytes
and append only 46 separate result files. All 299 runtime, 188 test and five
package pins are unchanged at completion.

The negative cases retain REQUESTED/DENIED/REJECTED/FAILED, policy interruption,
missing later Save and empty Domain references. A cancelled Tracking Policy
write retains the existing pre-write interruption with only its REQUESTED fact
in the incomplete journal. It is not converted into a complete stopped run.

**Shutdown limitation:** checks complete before the test's Excel process stalls
after Quit. Read-only inspection finds no native windows or active COM object,
but cannot verify a typed workbook count. After eight minutes and exact process,
receipt and pin checks, the identified completed Excel process and its waiting
test host require forced cleanup. No operational process is targeted, and no
normal-shutdown acceptance is claimed. The wrapper exits -1; zero matching
Application errors occur through the cleanup audit. The receipt records start
**18:21:54.2917485 UTC**, wrapper completion **18:59:50.5552366 UTC** and final
audit **19:01:26.0984071 UTC**, 2026-09-22. Forced cleanup is separate from the
27 verified behavioral failures.

The initial diagnostic attempt, root `09fdc18531e64a34b358f3483b525db8`, is
268 PASS / one harness failure: it tries to start a new recording through stale
Viewer controls after policy interruption. Reopen Viewer for the next fixture;
this is not product RED. The corrected run above preserves that interruption.

Core's explicit map now names exactly the approved Settings controls and their
outcomes, with an exact owner check. It adds no generic success classifier,
policy-save completion or Domain evidence. The fresh
`deploy/validation-settings-diagnostic` candidate builds, passes Operations cold
start and all five compiles. Its actual-handler/expectation-editor/Evaluate run
passes **780/780** at
`reports/runtime/slice4be-settings-activity/8fc21e0c92e24e748ffdb7f68f8c1281/green.json`,
resolving all 27 terminal-classification failures and retaining all 450 prior
GREEN checks. All 54 captures are directly reviewed and their hashes verified.
The completion audit at **19:33:18.6810122 UTC, 2026-09-22** verifies all 299
runtime, 188 test and five package pins before subsequent regression-fixture
reconciliation. Original evidence bytes remain unchanged; only the 46 separate
evaluation results are appended.

**GREEN shutdown limitation:** read-only observations verify zero workbooks at
**19:32:18.9478453 UTC**, then Quit and FinalRelease return. Excel and the test
host still linger after the final 780/0 receipt. After five minutes, exact
identity/restoration/receipt checks permit cleanup of only the completed test
host. The wrapper exits -1 at **19:38:13.7060201 UTC**, reporting Excel closure
unavailable. A subsequent exact-process stop is requested; Windows reports no
running instance while a residual Excel entry remains. A read-only native check
at 19:41:50 UTC returns exit code zero but an unsignaled process handle; neither
that exit code nor .NET HasExited is accepted as completed shutdown. By the final
audit at **19:43:10.4509230 UTC**, the residual process is gone, with zero matching
Office Application 1000/1001/1002 events. No reboot or guard bypass is used.
Normal unassisted shutdown is not claimed. This operational limitation is
separate from the verified behavioral GREEN; broader regression/release gates
remain pending.

The older regression fixtures now explicitly accept catalog 11 while retaining
historical catalog checks. Programmatic assignments suppress automatic events
before one deliberate native callback. The policy-change assertion retains the
original General Save pair and exact Capture REQUESTED/STAGED and Policy Save
REQUESTED observations: three actions, five observations, seven journal versions,
Incomplete/POLICY_CHANGED. These test-only corrections await regression execution.

The first broader Settings attempt retains **191/191 prior GREEN checks**, then
adds one harness failure during final screenshot capture at
`reports/runtime/slice4be-tracking-settings/306e386aa1604ab4904d1419a10809e0/green.json`.
The injected `OperationsViewerWindowForTest` Repaint/DoEvents/native-handle getter
returns 0x80020009, followed by unavailable Excel readiness. A new empty Excel
child appears; no Application 1000/1001/1002 event establishes a crash cause.
All eight available captures are reviewed, including fully painted General and
native maximized/restored Settings. Source/package pins are preserved. An exact
native/creation-bound, typed-zero-workbook inspection permits normal Quit of the
remaining empty instance; the recovery dialog is explicitly set to retain files.
The final audit at 19:51:11 UTC verifies no remaining Excel process. This is not
acceptance GREEN. A fresh retry uses the existing process-bound native-caption
capture helper, retains all 191 assertions, and refreshes diagnostic process
identity after the intentional preference restart; runtime packages are unchanged.

That retry, root `adf6a10075784c8284a70c7f3fa9ce07`, passes 179 checks before a
separate foreground-capture failure during native maximize. No behavioral check
fails; final Excel closure is immediate and unassisted. Its six captures are
reviewed: five accepted, one unpainted background Tracking-page image rejected.
All runtime/test/package pins are verified before correcting the remaining
legacy capture calls. Admin Settings captures now use the existing unique owned
caption helper; native maximize/restore uses its explicit owned-handle helper.
No handler/assertion or runtime package changes. A fresh complete 191-check run
remains required; these attempts are preserved, not combined into acceptance.

The owned-capture attempt at `1afbd53c6e6a43c7a68418c4d3d37a5e` passes 57 checks
before RPC becomes unavailable at `TestD5Commands.ShowSettings`. Its two initial
captures are fully painted and reviewed. Source/package pins remain unchanged;
the replacement empty Excel instance is identified by native owner/creation,
verified to have zero workbooks, and closed by normal Quit. No Application
1000/1001/1002 event establishes a cause. The failure remains a harness/runtime
environment investigation, not a business-contract RED.

The next test setup explicitly compiles all injected probes before forms, both
in the initial five-package Admin session and the restarted four-package
Operations-only session. Default-Close probe installation moves into initial
setup. Exact package/reference-root checks remain enforced, and all 191 prior
assertions remain. Nine compile checks and two zero-loaded-form checks bring the
expected total to 202. This tests an unproven instrumentation-state hypothesis;
no runtime implementation is changed to accommodate the failures.

The compiled setup passes **202/202**, including all **191 prior GREEN checks**,
at `reports/runtime/slice4be-tracking-settings/d66694130e7d440688ec1a61aba6d23d/green.json`,
**20:02:40.5827096--20:06:22.7802905 UTC, 2026-09-22**. All ten captures are directly
reviewed and accepted, including General Settings, Tracking, Detail, personal
preferences, Viewer entry and native maximize/restore. The 20:07:28 UTC audit
verifies 299 runtime, 188 test and five package hashes, immediate unassisted final
Excel closure and zero Application 1000/1001/1002 events. This completes this
Settings regression; it does not establish a root cause for prior RPC failures
or waive the large diagnostic run's shutdown limitation. Full Release 1 and other
role/comparison regressions remain required before completing the implementation
step.

## Full-chain attempt awaiting successful regression

The frozen diagnostic candidate's first full-chain attempt passes Create
Warehouse 15/15, then stops at Receiving's packaged
`modTS_Received.RunReceivingConfirmWritesFormActionForTest` with HRESULT
`0x80020009`: live-role evidence is 14 PASS / 1 harness FAIL and the enclosing
chain is 5 PASS / 1 harness FAIL. This is not behavioral RED or release acceptance.
The same callback failed intermittently on the preceding owner-completion
candidate before its separately preserved successful chain; root cause remains
unproven.

The 20:17:10 UTC audit preserves 299 runtime, 246 tooling and five package hashes,
verifies restoration of all three tracked reports, no remaining Excel instance,
and zero Application 1000/1001/1002 events. Cleanup required normal Quit of a
freshly identified instance with typed workbook count zero and direct review of
the recovery dialog's retain-files choice. Recovered files were retained.
Evidence: `reports/runtime/settings-diagnostic-chain-attempt-verification.json`;
the failed reports remain under the same prefix. A subsequent successful gate
must retain this failed attempt and cannot establish its root cause by itself.

The unchanged candidate's retry passes **32/32 chain, 48/48 live-role and 15/15
Create Warehouse** checks, including the formerly interrupted Receiving action.
All preceding GREEN identities are retained. The 20:47:29 UTC audit preserves
299 runtime, 246 tooling and five package hashes and verifies all three report
restorations. Runtime and static package checks also pass.

**Clean shutdown is not accepted:** the strict audit detects Application Error
1000 (`EXCEL.EXE`, `combase.dll`, `c0000005`) at 20:45:23 UTC and Windows Error
Reporting 1001 (`OFFICE_MODULE_VERSION_MISMATCH`) at 20:45:29 UTC. The affected
instance had been observed read-only with typed workbook count zero. A later
guarded Quit was withheld when COM attachment was unavailable; no forced
termination occurred. The final recovery instance was independently verified
empty, quit normally, and its recovered files retained after direct dialog
review. No Excel instance remains. These observations do not prove root cause.
The strict gate remains failed; the separate
`reports/runtime/settings-diagnostic-chain-retry-behavior-verification.json`
records behavioral evidence only. Full release acceptance is explicitly false.

A temporary, ignored harness copy adds stage/process/count-only lifecycle tracing
without changing any package or assertion. That run again passes 32/32, 48/48
and 15/15, with original source/package pins and restored reports verified at
21:00:05 UTC. The trace binds the repeated access violation to the reconciliation
instance: its original Quit returns with typed workbook count zero; a later
fresh read-only attachment confirms zero; an additional guarded normal Quit is
followed about five seconds later by the access violation. This is correlation,
not proof of cause. Avoid reattachment or repeated Quit to a reconciliation
instance after its original Quit; a further lifecycle diagnostic must observe
passively to distinguish shutdown behavior from that interaction.

The traced run records two Application events and **no clean-shutdown acceptance**
(`reports/runtime/settings-diagnostic-chain-trace-behavior-verification.json`).
Its final recovery instance is separately verified empty and quit normally.
The retain-files choice is clicked, but the intermediate screenshot does not
paint the radio options before confirmation; therefore retention is requested,
not visually verified. The corrected cleanup receipt explicitly records that
limitation. No process is forcibly terminated, and no Excel instance remains.

## Owner-regression capture attempt

The first owner regression on the diagnostic candidate finishes with 453 PASS /
7 capture FAIL. All 92 owner checks, including the twelve actual expectation
editor/Evaluate cases, pass. The failed checks are the four Boxing action
screenshots and three Action Paths layout screenshots; these failures remain
failures, not behavioral RED. Of sixteen produced images, fifteen are accepted;
the General Settings image has blank list content and is rejected. Its earlier
separately accepted Settings images do not turn this gate GREEN.

The 20:28:34 UTC audit preserves 299 runtime, 188 test and five package hashes,
immediate unassisted Excel closure and zero matching Application failure events.
Evidence: `reports/runtime/settings-diagnostic-regression-owner-attempt-verification.json`.
A read-only native observation finds visible forms with the workbook window in
the foreground. The capture helper uses title activation only; the follow-up
harness uses the existing ownership-checked caption activation helper for Boxing
and General Settings captures. It changes no package behavior or expectations.
A fresh complete owner gate then passes **460/460**, retaining every preceding
owner check. All **23 images are accepted**, including the complete General
Settings form and all seven formerly failed Boxing captures. Five instrumented
projects compile; 33 resource samples stay bounded (maximum GDI 801, USER 669).
The 20:29:00--20:39:20 UTC run closes Excel immediately without assistance and
has zero matching Application failure events. Runtime/test/package pins remain
unchanged. Evidence:
`reports/runtime/settings-diagnostic-capture-regression-owner-verification.json`.
This accepts the focused owner regression, not the broader Shipping gate or
whole Slice 4be.

The same candidate's separate Admin UOM regression passes **228/228**, preserving
all preceding GREEN identities and all **11 accepted images**. The
20:47:48--20:50:32 UTC run closes Excel immediately without assistance, with no
Application error events and preserved runtime/test/package pins. Evidence:
`reports/runtime/settings-diagnostic-uom-verification.json`.

## Broader Boxing/Shipping capture failure and DPI correction

The first broader gate on this candidate stops at **1,681 PASS / 8 FAIL**: seven
are the preserved unapproved D8 missing-auth-file failures, and one is a capture
harness exception (`Capture focus cursor positioning failed`). Twenty-six prior
checks are unreached. Nineteen of twenty images are accepted; the older-policy
UNBOX image is misframed, includes the background editor and clips the form.
The 21:19:53 UTC audit verifies all source/package pins, immediate unassisted
Excel closure and zero Application failure events. This is incomplete evidence,
not a passing broader gate. The attempt is preserved under
`reports/runtime/settings-diagnostic-regression-boxing-attempt-verification.json`.

Read-only native inspection finds virtualized bounds of 1,307 by 1,133 pixels
versus physical bounds of 1,961 by 1,700 for the same form. Separately, Windows
denies input-desktop access with error 5; no causal equivalence between those
two findings is claimed. Visible packaged checks must wait for desktop access.

The new `tests/tooling/Test-SettingsCaptureGeometry.ps1` exercises the actual
shared capture helper against an owned borderless solid-color window and the
independent physical DWM bounds. Its harness RED is 3 PASS / 1 FAIL: a required
480 by 240 image is incorrectly saved as 320 by 160. A scoped thread DPI context
around capture/focus operations fixes the coordinate mismatch without changing
process DPI or any package. The focused GREEN is **4/4**, including restoration
of the caller's DPI context after success and invalid-window failure. RED root:
`reports/runtime/capture-geometry/98099c19edfd4060bf44fd8f3ec07da0`; GREEN root:
`reports/runtime/capture-geometry/12a43dda386944e7ab0f1204bfc9fca3`.
This is capture-harness evidence, not product behavioral RED/GREEN or a substitute
for the still-required visible packaged rerun. Windows documents the
[DPI virtualization of GetWindowRect](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getwindowrect)
and [scoped thread-context restoration](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setthreaddpiawarenesscontext).

## Observation maintenance review

### Resumed desktop check, 2026-09-23 UTC

Input-desktop access is available at 02:10:48 UTC. The disposable foreground
calibration completes all three cases (`hidden-first`, `visible`,
`hidden-restored`) with Captured results, directly reviewed readable images,
normal Quit and verified process closure. Its report and image review are under
`reports/runtime/capture-foreground-calibration/3c3b3b534b344599a55a1960c5ee6f7d`.
This establishes capture calibration only, not packaged product acceptance.

The following fresh broader Boxing/Shipping run stops at **17 PASS / 1 harness
FAIL**, after all five instrumented package compiles pass. No product image is
produced: caption activation fails with `Capture focus cursor positioning failed`.
A new desktop check at 02:13:01 UTC returns NULL/error 5. The reason desktop
access changed is unknown; do not infer an application regression or repeat
visible runs without restored access. Excel closes immediately without assistance.
The 02:14:11 UTC audit preserves 299 runtime files, 189 test files and five package
hashes, with zero matching Application failure events. The unchanged accepted
gates remain valid; this attempt adds no broader acceptance. Evidence:
`reports/runtime/settings-diagnostic-capture-regression-boxing-attempt-verification.json`;
report root `reports/runtime/slice4be-shipping-activity/a8c172b9928a43a7905885c659a74530`.

### Recorded static results

The follow-up source inventory is 250 components, 6,045 procedures and 132,881
lines before the terminal map: one bounded Core vocabulary module, nine procedures and 326 lines above
the accepted owner-completion baseline. Literal/unresolved Application.Run counts
remain **9/45**. Four duplicated context helpers were removed.

**Bounded duplicate-metric exception:** the scanner reports 193 groups versus
192 previously because it normalizes away the different string literals in 26
one-line, actual MSForms callbacks. Each adapter passes a distinct specified
ControlId to its form's typed dispatcher. Retain these required event bindings;
they are not duplicated authority or storage implementations. The exception
applies only to that reviewed adapter group, not future duplication or relaxed
module limits. All three regenerated evidence schemas and all 28 existing
oversized-module limits pass; unrelated user-document pins are preserved.
The terminal map adds 12 lines only, for final totals of **250 components, 6,045
procedures and 132,893 lines** (+338 versus the owner-completion baseline), with
the same 9/45 dynamic-call counts and 193 duplicate groups. All three schemas and
28 oversized-module limits pass again. Broader regression and release gates
remain required before completing this step.

## Preserved failed attempts

- Initial callback RED: 168 PASS / 126 FAIL at
  `reports/runtime/slice4be-settings-activity/b152f5298e4d49058ae7d5829dc23cba/red.json`.
  One assertion used the nonexistent lifecycle `Interrupted`; the existing
  contract is `Incomplete`. The other 125 failures are behavioral. All five
  captures are reviewed, including a complete General Settings image.
- Expanded RED attempt: 131 PASS / 97 FAIL at
  `reports/runtime/slice4be-settings-activity/e8b3348d5de845e88493e42b5b834374/red.json`.
  The cancellation fixture left Excel events disabled, so its BeforeSave observer
  could not cancel. This is a harness failure, not product RED. Enable events only
  around the real cancellation and restore their prior value afterward.
- The corrected expanded RED's General Settings image is partially painted and
  rejected. Its four Settings-section images are accepted. The first GREEN's
  fresh General image passes after using the existing owned-foreground capture
  helper. Earlier rejected images and failed attempts remain preserved.
