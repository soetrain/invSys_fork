# Slice 4be.1 Process instruction observations

Last verified: 2026-09-29 UTC. Architecture v4.11 D18 specifies these five
discovered controls before implementation; Plan022 and the controls catalog carry the
same contract. This is local instruction editing under the existing Production
or Admin permission, with no new saved-definition authority.

The actual Add, Update, Remove, Up and Down form handlers now use a typed
Operations owner. It retains the captured workbook/session, rejects stale or
closed bindings, guards loading and nested actions, and preserves existing
trim/empty-Update/ordinal/selection behavior. Catalog14 adds exactly five fixed
control identities while preserving1-13. REQUESTED precedes authorization and
validation; STAGED is local success, REJECTED is invalid input/selection/movement,
DENIED prevents editing, and FAILED retains uncertainty. Only STAGED satisfies
CommandCompleted; empty source references cannot establish Domain application.
Instruction text and draft identity never enter activity records.

## Packaged D13 evidence

| Gate | Result | Scope |
|---|---|---|
| Initial RED |69 PASS/228 FAIL|The five actual edits and saved-authority checks pass; missing catalog/observations/guards fail. |
| Tightened RED |93 PASS/278 FAIL|Valid mutation preconditions prevent busy-guard no-ops from hiding defects; terminal and unsupported-outcome checks added. |
| Initial GREEN |371/371|Same tightened identities, five instrumented compiles; normal closure and preservation. |
| Expanded RED |105 PASS/306 FAIL|Actual nested list-event entry, initialization guard and unavailable storage coverage on the old candidate. |
| Expanded GREEN |411/411|Exact expanded identities and all prior GREENs retained on the typed candidate. |
| Recording/publication/paired views |105/105|Two distinct original five-action runs, owning publication, guide author/reader separation and six reviewed principal images. |

The expanded RED and GREEN both compile all five instrumented packages before
callbacks. No harness failure or Excel Application Error occurs in either window.
Actual nested events reach the second handler; the GREEN suppresses its edit and
extra records. Unavailable/disabled optional tracking preserves authorized edits;
unavailable storage supplies a visible fixed notice without a local fallback.
Target/session/sign-out/closed-workbook guards, unknown columns, local workbook
bytes, canonical files, redaction and record hashes pass. Controllers restore
settings, preserve all five candidate hashes and observe normal unassisted Excel
closure. These are disposable fixtures, not operational inventory.

## Build and maintenance

Final candidate: `deploy/validation-production-instructions-typed`. Five builds,
five compiles and cold-start dependency validation pass. Among248 compiled
components, exactly five are changed/added relative to `validation-production-paths`:
Core modActivityCatalog, modEvaluationMatches and modProductionInstructionCodes;
Operations modProductionDesignerActions and frmProduction. No other component
changes. The first string-dispatch candidate passed371/371 but introduced a
normalized duplicate group; typed action values remove it without a scanner
exception or runtime contract change.

Static maintenance records255 components,6065 procedures,133334 lines,9 literal
and45 unresolved dynamic calls, and191 duplicate groups. All28 previous module
limits are retained or reduced; frmProduction falls11743 to11739 lines. Static
layout8/8 and Run List layout7/7 pass. Three report schemas and all298 tooling
PowerShell parses pass. Current-version expectations in existing tests advance
13 to14 (policy control count74 to79); historical catalog1-13 checks remain intact.
An initial layout invocation omitted RepoRoot and failed during parameter setup;
the explicit-root run passes. That setup failure is not product RED.

The first instruction-path extension reaches83 PASS/one harness failure after
recording, publication, Viewer identity and command-versus-Domain checks. Its
fixture omitted ACTION_PATH_MAINT for the guide author, so the actual Create Guide
entry is unavailable. The corrected fixture uses the existing explicit grant and
checks author/reader separation before claiming presentation evidence. No runtime
permission changes follow from that fixture correction. The first controller
`production-instruction-controller/43bae888da394fd38cf7d7cfbad9db11` under runtime
reports restores settings/packages and closes Excel; it is not meaningful RED.
The next run reaches100 PASS/two failures: its authority pin incorrectly includes
snapshot projections that publication owns, and an exclusive hash read cannot
read the open operator workbook. Canonical-only pins and the existing shared-read
hash pattern correct those test assertions without changing the candidate. That
controller is `production-instruction-controller/b644e51ffd914f23bc54af8f429c14f8`.

The corrected path gate passes105/105, five instrumented compiles, zero Excel
Application events and normal unassisted closure. Two five-action recordings
retain all20 original REQUESTED/STAGED records and their exact journal order.
Admin publication preserves every original line; Viewer finds all five source
identities. The actual expectation editor/evaluator concludes CommandCompleted
and leaves SourceEventsApplied incomplete. A separately authored five-step guide
retains its original source while a reader without author permission evaluates
the distinct observed run. How-To, Diagnostic and Compare retain the same evidence,
five exact matches and the explicit local-only conclusion. Original journals,
activity, canonical bytes and unknown operator columns/workbook bytes are preserved.
Snapshot projections are owned publication output, not canonical authority pins.
Settings and all five package hashes restore.

Six final principal captures were reviewed: instruction editor/status, Event
Detail, How-To, Diagnostic, Compare, and the scrolled saved conclusion. The final
detail capture shows REQUESTED with both pair members listed; the earlier
`4e3908560f134ccaa9903bbe2410f149/instruction-event-detail.png` shows STAGED on the
same candidate. Images remain ignored runtime evidence and establish no human
or NAS acceptance.

## Evidence locations and remaining gates

Current-candidate regression checkpoint:

| Gate | Result | Preservation |
|---|---|---|
| Release1 full chain |32/32|Exact prior identities; normal unassisted closure. |
| Live roles |48/48|Exact prior identities, including accepted operational workflows. |
| Create Warehouse |15/15|Exact prior identities. |
| Settings, tracking, detail and preference |202/202|Exact prior identities; five instrumented compiles and normal unassisted closure. |
| Production draft/Action Paths |390/390|Exact prior identities, five compiles, original journals and saved authority preserved; normal closure. |
| Production lifecycle |615/615|Exact prior identities, five compiles and normal unassisted closure. |
| Packaged Production layout |PASS|Three sizes x five pages, native minimize/restore/maximize/restore and three complete reviewed captures after capture-tool correction. |
| Settings command observations |467/467, qualified|All450 prior identities plus17 historical exclusions; eight reviewed captures. Assisted cleanup; normal shutdown remains open. |
| Native lifecycle cancellation |94/94|Exact prior identities; four complete reviewed question captures and normal unassisted closure. |
| Packaged smoke |86/86|Exact prior identities; automatic termination is possible but unobserved, so normal exit is not established. |
| Reusable Production run-only |0 PASS/1 harness failure|RPC/native failure at the batch-scale adapter; no termination requested, no reusable result. |

These controllers restore settings and all five package hashes, retain tracked
report bytes, and have zero Excel Application events. Chain runs17:41:14--17:46:16
UTC, Settings17:46:34--17:50:14 and draft17:50:33--17:57:55 UTC. Other regressions
remain open below.

Lifecycle runs17:58:10--18:03:10 UTC with the same preservation guarantees.
Initial layout runs18:03:30--18:03:45 UTC and closes normally with settings/packages
preserved. Its minimum/default images clip content at their edges despite passing
internal geometry assertions. The actual capture helper mixed DPI-virtualized
bounds with physical PrintWindow pixels. A disposable independent physical-window
test records5 PASS/2 dimension FAIL before the fix, then7/7 including caller-context
restoration after success, invalid handle and write failure. The helper now uses
a scoped physical-pixel context and restores it. No runtime/package contract changes.
The unchanged packaged candidate reruns18:17:23--18:17:38 UTC, passes geometry and
native window actions, and closes normally with preservation. All three complete
captures were reviewed; minimum/default are byte-identical. The earlier clipped
captures remain retained but are superseded for visible layout evidence.

Settings command observations initially finish449 PASS/1 fixture FAIL: the
catalog-10 fixture removed only26 Settings controls, leaving17 later Production
controls in a table labeled catalog10. All450 prior identities remain, including
context/preservation checks. Eight captures were reviewed. Hidden Excel required
explicit assisted cleanup after the report completed; this run proves neither
full GREEN nor normal shutdown. Settings/packages restore, with zero Excel
Application events. The fixture correction retains exactly the36 catalog-10 IDs
and checks exclusion of every later ID; the failed attempt is harness evidence,
not D13 runtime RED. The corrected rerun passes467/467, retaining all450 prior
identities and adding17 exclusion checks. Its eight captures were reviewed.
Hidden Excel again remains beyond120 seconds after report completion and requires
explicit assisted cleanup. Settings/packages restore and zero Excel Application
events are recorded. This is qualified functional GREEN; normal shutdown remains
open. Do not repeat the unchanged full Settings gate to investigate closure;
use a bounded shutdown diagnostic retaining actual callback behavior.

Refreshed maintenance retains255 components,6065 procedures,133334 lines,
9 literal/45 unresolved Application.Run sites,191 duplicate groups and all28
size caps. Three schemas and300 PowerShell parses pass. This follow-up changes
developer capture and fixture construction only; all runtime packages stay fixed.

Native cancellation runs18:31:38--18:33:29 UTC and smoke18:33:48--18:34:07 UTC.
Both restore settings/packages with zero Excel Application events; smoke also
restores its tracked report. Native closure is unassisted. Smoke's validator can
Stop-Process after its wait and emits no termination receipt, so do not infer
normal closure from its controller's ExcelClosed field.

Reusable run-only fails18:34:41--18:35:05 UTC at
`mProduction.RunProductionBatchScaleContractTest`, with RPC0x800706BE and Excel
ntdll.dll exception0xc0000028 (offset0000000000012d2f). Its receipt explicitly
records TerminationRequested=False. Settings/packages restore; no desktop error5
occurs. This matches the earlier native signature but does not prove the cause
or preserve the prior67 reusable observations on this candidate. Both run-only
and full Production/restart remain open. Do not repeat unchanged broad runs or
infer a repair from passing shortened diagnostics; trace the failing standard flow.

- Expanded RED controller: `reports/runtime/production-instruction-controller/1efdc1da85fb4757bd4d3cbce1867f88`.
- Expanded RED results: `reports/runtime/slice4be-production-instructions/02a0cd3f2f18464192acafa7392a63d2/red.json`.
- Expanded GREEN controller and five package pins: `reports/runtime/production-instruction-controller/babcce28be534358ba247e93d1c044df`.
- Expanded GREEN results: `reports/runtime/slice4be-production-instructions/12e52e8ea1734498b9740aade91ba25d/green.json`.
- Verification: `reports/runtime/production-instructions-focused-verification.json`.
- Compiled source evidence: `reports/runtime/production-instructions-typed-compiled.json`.
- Final maintenance: `reports/runtime/production-instructions-final-static` (same metrics and limits).
- Path controller: `reports/runtime/production-instruction-controller/ee3ece2928c7436191bf431d96e901aa`.
- Path results and six principal captures: `reports/runtime/slice4be-production-instruction-paths/097b8f69aabb497ba4301038d3c4f1a6`.
- Path verification: `reports/runtime/production-instructions-paths-verification.json`.
- Chain controller: `reports/runtime/production-instructions-regression/chain-a948d08fa0264712bbcc5bbc779fab9a`.
- Chain verification: `reports/runtime/production-instructions-chain-verification.json`.
- Settings controller: `reports/runtime/production-instructions-regression/settings-b35c9ac8ce4e48bba6b2b078dfd580d8`.
- Settings verification: `reports/runtime/production-instructions-settings-verification.json`.
- Draft controller: `reports/runtime/production-instructions-regression/draft-88979ee8cdb44e5f9f5ce59f9f8a74ff`.
- Draft verification: `reports/runtime/production-instructions-draft-verification.json`.
- Lifecycle controller: `reports/runtime/production-instructions-regression/lifecycle-bbc51bd70263425389486f915deaeadf`.
- Lifecycle verification: `reports/runtime/production-instructions-lifecycle-verification.json`.
- Layout controller, geometry report and three captures: `reports/runtime/production-instructions-regression/layout-2da6a60afa1a4997b52edbeaaf816b28`.
- Capture-tool RED: `reports/runtime/production-layout-capture/adc115be56854b5bb784f4aa6ace946f/red.json`.
- Capture-tool GREEN: `reports/runtime/production-layout-capture/463a52b436e7475faf350d0142cab8f2/green.json`.
- Corrected packaged layout and complete captures: `reports/runtime/production-instructions-regression/layout-0c9bf9f961b64221b53e03601026f209`.
- Layout verification: `reports/runtime/production-instructions-layout-verification.json`.
- Settings observations failed fixture/assisted cleanup: `reports/runtime/production-instructions-regression/settingsactivity-2813f5e975b44928b5dfbebbb61756c8`.
- Settings failed attempt verification: `reports/runtime/production-instructions-settingsactivity-attempt-verification.json`.
- Corrected Settings controller and assisted cleanup receipt: `reports/runtime/production-instructions-regression/settingsactivity-99d039af47864415806157f4994531d3`.
- Corrected Settings results and eight captures: `reports/runtime/slice4be-settings-activity/9f58e277a36c4556ae579aaab3390a0c`.
- Corrected Settings verification: `reports/runtime/production-instructions-settingsactivity-verification.json`.
- Refreshed maintenance and verification: `reports/runtime/production-instructions-validation-static` and `reports/runtime/production-instructions-validation-static-verification.json`.
- Native controller: `reports/runtime/production-instructions-regression/native-a5196f84527e48af984d78353b7dfe42`.
- Native results/four question captures: `reports/runtime/slice4be-production-lifecycle-native/e8376b936b4e4e7abd7512b9cd6277b4`.
- Native verification: `reports/runtime/production-instructions-native-verification.json`.
- Smoke controller: `reports/runtime/production-instructions-regression/smoke-141d94d73e2948338c1004b6412ce0d6`.
- Smoke verification: `reports/runtime/production-instructions-smoke-verification.json`.
- Reusable attempt: `reports/runtime/production-instructions-regression/reusable-5c174e72a164412b90d24240462c11a3`.
- Reusable failure verification: `reports/runtime/production-instructions-reusable-attempt-verification.json`.

Normal Settings shutdown and reusable Production gates remain pending. The earlier
full Production native failure remains independently open. This focused checkpoint
does not accept the five controls, comprehensive Slice4be, Release1, human or NAS
acceptance. The user requires the goal to stop if desktop Windows error5 recurs;
none occurred during these gates.
