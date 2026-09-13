# Plan 022 Slice 4be.1 Shipping submission evidence

Last verified: 2026-09-12. Runtime remains **22e14b6** in
`deploy/validation-shipping-capability-guard`. This test-only continuation follows
the [558-check access matrix](plan022_slice4be_shipping_access_results.md).
D18's owner-fact, exact-reference and captured-workbook rules govern discovery;
pending D8-A is not implemented or treated as approved.

Focused diagnostics pass **136/136**, proving pending source identities and their
survival through lost acknowledgment. Two revised full runs fail early with native
Excel faults; the focused result does not establish full-matrix acceptance.

## Real action and controlled faults

`Slice4beShippingSubmission.ps1` runs after the preserved Shipping matrix in a
separate Admin-generated fixture. Box Designer/Maker handlers create its inventory;
the actual Add handler calls the ordinary commit owner and Core submission APIs.
Unsaved observers retain the commit result/reserve ID and actual Core return
values. They never generate an ID, grant a permission, manufacture successful
persistence or substitute a fake business owner.

| Case | Controlled fault | Evidence needed |
|---|---|---|
| ServerUnavailable | Server API returns unavailable before queue entry; local API remains real. | Local acceptance generates one exact ID retained by Shipping. |
| LostAcknowledgment | Actual server queue returns success, then its public acknowledgment is forced false. | Real local fallback preserves that same ID. |
| ExceptionalAcknowledgment | Actual server queue succeeds, then the public API raises a fixed fixture exception. | The typed cross-project call and Shipping fallback retain the same ID. |
| UncertainAcceptance | Actual server queue succeeds, but acknowledgment and authorized fallback both return unavailable. | Shipping reports owner failure while retaining the possibly submitted ID. |

The fallback-unavailable fault occurs after the existing current-target permission
check. Each case distinguishes owner result, source identity and actual Domain
application. Ordinary Remove clears successful local staging between cases; the
uncertain case is last. Reports contain fixed check names and booleans, never raw
payloads, entered values, credentials or returned error text.

## Initial calibration and isolation correction

Initial discovery completes **552 PASS/53 FAIL across605 checks**: all558 prior
identities/505 GREENs remain, with no duplicate or lost checks. All47 new checks
pass. Both public submission boundaries execute once per fault. Every known ID
survives fallback and remains available inside the actual commit owner. All four
Add cases have zero Domain log applications for their exact submitted ID at the
observation point. The uncertain case's owner returns False after server acceptance;
the other three owners return True. Thus neither a Boolean result nor absence from
the applied inventory log supplies a complete account of submission.

This run also exposes a fixture-isolation omission after editing the unsaved Core
project: the new fixture does not reestablish bootstrap paths. It generates a
default repository template and a fixture operator under the default user-profile
operator location. The operator is identified by its generated filename, creation
window, Shipping Extra column, fixture box name and fixture description. The
generated repository template is moved to ignored diagnostic storage. Automatic
approval review rejects the combined cleanup command as "blocked by policy";
that rejected command executes nothing. Subsequent separately verified single-file
moves succeed for the generated operator and its fixture-only staging file, keeping
both in ignored evidence and leaving surrounding folders and unrelated files alone.
Do not treat initial discovery as proving isolated fixture paths or source-file
read-back; it records actual owner returns and Domain-log observations only.

The revised helper sets the accepted template and a generated operator root
explicitly before generation. An unsaved local-staging root override confines the
real file writer to this fixture as well. Additional checks require the operator
and submission paths inside the generated root, read back exact persisted source
IDs from server inbox/local JSONL, preserve source bytes during inspection and
verify the accepted template's bytes. No production path behavior is changed or
claimed by this test-only root override.

The first isolated run stops **67 PASS/1 harness exception** before reaching the
new submission helper: Excel RPC fails at the existing Shipping setup/launcher
stage. Windows records ntdll.dll/c0000028 at2026-09-13T03:33:51Z. Its empty recovery
instance is verified twice with zero workbooks; Quit does not terminate it, then
only that verified process is stopped. A fresh unchanged-test retry is justified
by the preceding completed605-check run and retained separately. This is not
behavioral RED or a native-crash repair. That retry also stops67/1 at the same
existing setup stage before the new helper, with ntdll.dll/c0000028 at03:39:59Z.
Its separate empty recovery instance is likewise inspected twice; Quit does not
terminate it, and only that verified empty process is then stopped.
The full matrix remains unverified for the revised helper; do not merge partial
runs or keep restarting that unchanged full path without new evidence.

`-ShippingSubmissionOnly` now provides a separately named focused diagnostic
report. It installs the same packaged form-handler facades, then runs only the
new submission fixture after the shared configuration/activity prerequisites.
It preserves the default full-matrix route and cannot establish its558-check
regression result. A focused pass would calibrate the revised fixture/source
observations only; native stability and the failed full runs remain unresolved.

## Completed focused result and limits

The focused route completes **136/136**, exit0:70 shared configuration/activity
prerequisite checks and66 submission checks. All47 initial submission checks remain
GREEN; the19 additional path/persistence checks pass, with no duplicate identities.
The server inbox and local staging contain the same exact ID after successful
fallback; the uncertain case retains its server ID with no successful local write.
All four cases still have no Domain application for that exact ID. Reads preserve
source bytes, all new paths remain inside the disposable fixture, and the accepted
template's bytes remain unchanged. No Windows Application1000 Excel native fault
is observed during this focused run's window; this does not clear earlier crashes.

This result does **not** prove the revised624-check full route. The preceding
605-check complete run and the focused136-check run retain separate scope; their
counts must not be combined into a full GREEN. The46 missing-activity assertions
and seven pending D8-A findings remain open on the default route.

All50 package/eight current source pins match after Excel closes. Static evidence
is regenerated; all three JSON schemas and28 preceding module limits pass. Runtime
counts remain177 components,5511 procedures,8 literal Application.Run targets,
45 unresolved dynamic calls and195 duplicate bodies. Three edited PowerShell
scripts parse; no runtime source or accepted package changes. Existing broader
candidate gates retain their documented scope and failures rather than being rerun
or replaced by these diagnostics.

## Evidence locations

All raw paths below are relative to ignored
`reports/runtime/slice4be-shipping-activity/`:

- Initial: `submission-discovery.log`,
  `f893a3f3eedb49528789cf54e20d9305/green.json`,
  `submission-discovery-comparison.json`.
- Generated template: `submission-generated-inventory-template.xlsb`.
  `submission-generated-operator.xlsm` and `submission-generated-staging.jsonl`
  retain the verified initial fixture artifacts; none is staged for Git.
- Early native failure: `submission-isolated.log`,
  `8043b04fa4854fa19dddca732248a238/green.json`,
  `submission-isolated-native-faults.json`.
- Isolated retry: `submission-isolated-retry.log`.
  `3045476e33304adf9e2f1bd94a266e42/green.json` and
  `submission-isolated-retry-native-faults.json` retain the second early failure.
- Focused: `submission-focused.log`,
  `29783901ebec40dda439b3c1e206b725/diagnostic-submission-green.json`,
  `submission-focused-comparison.json`, `submission-focused-native-faults.json`.
- Final verification: `submission-final-pin-verification.json`,
  `submission-focused-static.log`.

Exact invocation for each run, with its own output log:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-shipping-capability-guard -Phase GREEN -CheckActivityEvidence -CheckActivityFoundation -CheckShippingActivity
```

The136-check focused command adds `-ShippingSubmissionOnly` to that invocation;
its report is explicitly named `diagnostic-submission-green.json`.

Disposable authority fixtures, including temporary unexpected Auth copies, are
removed by the outer harness's existing protected cleanup. Retained evidence is
the sanitized result/log, not raw Auth workbooks. This qualifies the preceding
access report's helper-local preservation language.

## Consequence for comprehensive activity

The normal-action activity test originally used newly applied inventory-log IDs as
its reference expectation. These Add observations prove that this misses pending
owner submissions. The subsequent [owner-reference RED](plan022_slice4be_shipping_owner_reference_results.md)
corrects that expectation and completes650 checks (597 PASS/53 FAIL), preserving
every preceding GREEN and missing-activity assertion. Submitted references mean confirmed
queue acceptance, Unknown references retain uncertain known IDs, and neither means
Domain application. Do not parse report wording to recover these facts.

Register Shipping controls/outcomes and mixed per-reference states under D18 with
synchronized Plan022/controls, then implement observation through the actual
handlers. Update/Remove multi-event or partial-release behavior, Shipments Sent
pending processing, tracking/policy failures and all remaining Shipping/Boxing and
Operations/Admin controls remain required. No full Slice4be, native stability,
physical deployment or human comparison acceptance is claimed here.
