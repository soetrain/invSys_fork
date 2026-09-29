# Slice 4be.1 Process instruction observations

Last verified: 2026-09-29 UTC. Architecture v4.11 D18 specifies these five
discovered controls before implementation; Plan022 and controls1.259 carry the
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

Current-candidate regression gates remain pending. The earlier
full Production native failure remains independently open. This focused checkpoint
does not accept the five controls, comprehensive Slice4be, Release1, human or NAS
acceptance. The user requires the goal to stop if desktop Windows error5 recurs;
none occurred during these gates.
