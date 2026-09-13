# Plan 022 Slice 4be.1 Shipping mutation permission

Last verified:2026-09-12 PDT. This advances the full D18 Events/Action Path goal;
it is not comprehensive activity or Release1 acceptance. Runtime repair follows
normative clarification **0cc61c9**, committed/pushed with Plan022 and controls1.94
before implementation. The clarification constrains existing SHIP_POST eligibility
under approved D18 semantic inheritance; it adds no capability or activity outcome.

## Protecting packaged test

`Test-Slice4beConfigCommands.ps1 -CheckActivityEvidence -CheckActivityFoundation
-CheckShippingActivity` retains all379 checks from the preceding
[workbook-close checkpoint](plan022_slice4be_native_validation_results.md).
`Slice4beShippingCapability.ps1` adds56 checks using the actual Add, Update Row,
Remove, Send Hold, Return, To Shipments and Shipments Sent handlers.

Start from actual authorized active/held staging. Calibrate each handler at its
mutation-owner boundary using the existing stopped-owner probe. Revoke only
SHIP_POST in the disposable Auth fixture, then require Core's real role-access
check to deny it while the same invSys session remains signed in. Repeat the
seven handler probes and require pre-mutation denial and a visible notice.
These probes prove dispatch placement, not Domain behavior. A separate real Hold
action runs the actual owner to protect local movement after capability loss.
Preserve captured workbook, keys/unknown values, canonical bytes and unrelated
workbook. Restore only the disposable Auth fixture after the case. No auth bypass,
synthetic user activity, operational workbook change or raw credential report.

## RED before implementation

On unchanged `deploy/validation-shipping-timer-rebuild`, runtime0aa0084:
**374 PASS /61 FAIL**,435 checks. All379 old checks and333 GREENs remain, without
duplicate/missing checks. The46 missing-activity failures remain. Fifteen new
failures are seven pre-owner checks, their seven permission notices, and real
Hold's staging-preservation check. Core denial and same-session calibration pass.
The real Hold moves local staging despite revoked SHIP_POST. No canonical
mutation or submission is asserted: those independent preservation checks pass.

Ignored evidence under `reports/runtime/slice4be-shipping-activity/`:
`capability-red.log`, `7c44ca24413d45c8b22a841a641c1f22/red.json`,
`capability-red-comparison.json`. All45 old package/eight then-current source pins
match before the runtime edit; Excel is closed. This is meaningful behavior RED,
not the native/fixture failures retained in the preceding checkpoint.

## Bounded candidate repair

`modShippingFormContext.CanAct` validates captured context, checks current Core
SHIP_POST eligibility, then revalidates context after that authorization work.
The existing form guard delegates to it for mutations, including its existing
post-yield and per-row calls. Failed permission verification takes the same
cancel-sync/status path, with the fixed specification notice. Core retains its
authorization/security decisions and the owners retain their existing guards.
No raw authorization error is classified into a new activity outcome.

The automatic-sync caller explicitly uses the existing context-only path. Form
reuse still tests captured context; Close/recovery and source identities remain.
The form stays3019 lines; the bounded context helper grows17 lines to56. No Core,
Domain, event catalog, service, Ribbon or authority schema changes.

Candidate: `deploy/validation-shipping-capability-guard`. All five builds,
explicit compiles and Operations cold start pass (`capability-build.log`,
`capability-compile.log`, `capability-compiled-source.json`). Five package hashes
and the two changed source hashes are pinned in `capability-package-hashes.json`
and `capability-source-hashes.json`. The compiled-source comparison retains170
components, with exactly the two intended Operations components changed.

Focused GREEN is **389 PASS /46 missing-activity FAIL**,435 checks. All15 new
permission failures are fixed; all56 added permission checks pass, retaining every
old check/GREEN without duplicates, missing checks or non-activity failure.
Evidence: `capability-green.log`,
`9dceaf96d32345bba2aa2e85b7d49ead/green.json`, `capability-green-comparison.json`.
That run's `shipping-permission-denied.png` was inspected: the fixed notice is
legible, with staging retained. Existing key-editor/button and header crowding
remains; this automated capture is not human layout acceptance.

Static evidence is regenerated at177 components/5511 procedures/123772 lines:
one new procedure and17 lines, with8 literal/45 unresolved Application.Run calls
and195 duplicate-body candidates unchanged. All28 preceding module-size limits
hold, including the form's3019 lines. Candidate smoke passes86/86; remaining
same-package gates are in progress and must be recorded independently.

The first ordinary live-role run fails14 PASS /1 harness exception at Receiving
ConfirmWrites (`DISP_E_EXCEPTION`,0x80020009). Its report and log are
`capability-live-role-results.md` and `capability-live-role.log`; the time window
is recorded in `capability-gate-exits.json`. No Windows Application1000 Excel
fault was observed for that run when inspected. It is not a Shipping permission
RED or proven native crash. The queue stops with Excel open; the remaining
instance has zero workbooks, receives Quit and exits before the next inspection.
The queue resumes at full-chain, retaining the failed live-role entry rather than
overwriting or treating it as passed. Current live-role acceptance remains open.

The resumed queue passes full-chain30/30, Viewer and combined launchers3/3.
Their exact reports/logs use `capability-full-chain`, `capability-viewer` and
`capability-launchers` under the ignored evidence root; timestamps/exit codes are
in `capability-gate-exits.json`. Full-chain and Viewer raw tracked reports were
copied there, then their original worktree versions restored. A Windows Excel
combase.dll/c0000005 fault at2026-09-13T02:32:50Z falls within the passing chain's
window (`capability-native-faults.json`). The business checks do not clear native
stability. Full reusable Production/restart passes2/2 without reduced flags, and
Shipping layout passes1/1. Their reports are
`capability-production/production-reusable-production.md` and
`capability-shipping-layout/shipping-layout.md`, with matching named logs.
The resumed queue returns failure overall because it retains the earlier failed
live-role gate, even though all subsequent gates pass. A fresh ordinary live-role
run, justified by the same candidate's later full-chain Receiving success, uses
`capability-live-role-retry.log` and `capability-live-role-retry-results.md`.
It retains the complete48-check scope and passes **48/48**, exit0. The first14/1
Receiving exception remains recorded; the fresh pass does not establish its
source cause or clear the separate native fault during full-chain validation.

Final verification matches all50 package pins across ten preserved sets and eight
current Receiving/Shipping source pins. The two changed source files match their
candidate pins; six unaffected source pins retain their original hashes.
`capability-final-pin-verification.json` records the result. All test handles are
terminal and Excel is closed. All three static JSON contracts and the24 launcher
source checks pass; no module-size, dynamic-call or duplicate-body exception is
needed. Raw reports remain ignored; accepted deployment, operational/NAS workbooks
and unrelated user changes are preserved.

## Remaining acceptance

Keep all435 focused checks, missing activity failures and every accepted regression.
The same-package gate outcomes above retain their individual scope and failures;
they do not establish native stability or human acceptance. Native crashes from
preceding candidates and the current chain remain unresolved unless supported
causal evidence clears them. Capability change during
a UI yield, unavailable Auth/Config and uncertain/pending submissions need their
own precise evidence; pre-entry revocation is not proof of those branches.

Then refine Shipping control/outcome/source-reference catalog definitions under
D18 and implement comprehensive activity. Full4be.1-4be.6, physical deployment and
human comparison/acceptance remain required; no narrower completion is claimed.

The [subsequent access matrix](plan022_slice4be_shipping_access_results.md) now
proves the three actual-yield revocations and seven unavailable-Config cases on
the unchanged candidate. Its558-check result retains all435 checks/389 GREENs.
Missing Auth rejects all seven mutations but recreates the file; the explicit
pending D8-A proposal records the resulting architecture decision. Neither this
extension nor the proposed decision clears pending/uncertain submission coverage.
