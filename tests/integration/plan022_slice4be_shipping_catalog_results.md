# Slice 4be.1 Shipping catalog and source eligibility

Last verified: 2026-09-12; UTC test timestamps fall on September 13.
This is a partial implementation checkpoint toward comprehensive Operations/Admin
Events and How-To/Diagnostic/Compare Action Paths. It does not complete Slice 4be.

## Authority and implemented behavior

Normative commit `37825b5` precedes runtime implementation and synchronizes
Architecture v4.11 D18, Plan 022 and controls v1.95. Its Shipping clarification
inherits the approved comprehensive coverage and owner-fact rules; it changes no
Shipping business mutation, capability, Domain event schema or authority store.
D8-A remains pending and is not implemented.

Core catalog 8 adds SHIPPING_ADD, SHIPPING_UPDATE, SHIPPING_REMOVE, SHIPPING_HOLD,
SHIPPING_RETURN, SHIPPING_STAGE and SHIPPING_SEND, with exact captions, the logical
SHIPPING_WORKFLOW owner and existing SHIP_POST eligibility. Versions 1-7 retain
their definitions. Saved catalog-7 policies remain valid and cannot enable the new
controls without an explicit policy update; policy reads preserve saved bytes.

`modShippingActivityCodes` supplies allowlisted outcomes and fixed messages.
`modActivityCatalog` registers those definitions; `modActivityReferences` validates
Shipping source eligibility. STAGED means confirmed local staging only. PENDING
means accepted source submission with unknown inventory effect. CONFIRMED is
restricted to Send's owner-confirmed processing/refresh path and still has unknown
inventory effect. FAILED retains exact Submitted/Unknown references, including
mixed states. Hold/Return never have references. Pre-write rejection/denial and
STAGED have empty references; PENDING/CONFIRMED require nonempty Submitted references.
Duplicate, cross-warehouse, invalid-identity, unknown-field/source/state and
uncertain-success references are rejected.

No Shipping form or owner implementation changes in this checkpoint. Definitions
alone do not create user activity, prove an action's result or replace the actual
handler tests. Core and Domain remain headless; captured-workbook binding, exact
System_Key and unknown-column obligations remain in force.

## D13 evidence and limits

The supplemental `Slice4beShippingCatalog.ps1` probe is installed unsaved into Core
before disposable fixture setup. It reads the real packaged catalog, outcome and
reference boundaries; later policy checks read an explicit catalog-7 fixture.
The main harness retains the complete Shipping form-handler route and all 650
preceding assertion identities. The probe never submits inventory or creates fake
handler activity. Direct Core checks supplement the required actual handler tests.

Pre-change package: `deploy/validation-shipping-capability-guard`, runtime `22e14b6`.
The RED run executes all **204 catalog checks: 115 PASS / 89 FAIL**. Failures include
absent catalog 8, missing Shipping definitions/outcomes and unsupported legitimate
source references. These are meaningful behavioral RED before source edits.
Afterward the fixture bootstrap crashes before actual Shipping handler execution;
the overall result is **183 PASS / 90 FAIL**, including one harness failure.

Candidate: `deploy/validation-shipping-catalog-eight`. All **213 catalog/reference/
policy checks PASS**, including every former RED and nine added existing-policy
regression checks. All 204 RED identities and their prior GREENs are retained with
no duplicates. The later fixture bootstrap again crashes before handler execution;
the overall result is **281 PASS / 1 harness failure**. This is focused catalog
GREEN, not a full Shipping route pass. The preceding 650-check result retains its
original package/scope and cannot be transferred to this candidate by arithmetic.

Both failures explicitly name `modAdminConsole.BootstrapWarehouseLocalAdmin` with
six values. Windows Excel Application event 1000 records `ntdll.dll` / `c0000028`
at `2026-09-13T04:43:25.6448677Z` (RED) and
`2026-09-13T04:49:37.9994827Z` (candidate). Fixture roots were verified before
generation. Neither crash is catalog RED, and no native repair is claimed.
Each recovery instance was inspected as empty, asked to Quit, then stopped only
after rechecking its exact process/start identity and zero workbook count.

## Package and regression gates

All five isolated XLAM builds, explicit VBE compiles and Operations cold start
pass. Exported compiled-source comparison finds exactly two changed Core components
and the added modShippingActivityCodes; no other package component differs.
The candidate's five package hashes and three changed Core source hashes are pinned.
Accepted deployment and prior candidate sets remain untouched.

Package smoke passes 86/86. Live-role validation stops 39 PASS / 1 harness exception
after Production.Form.CheckIn.NoPrematureLog, during the completion portion. The
ordered full chain stops 4 PASS / 1 harness exception after the Admin source
integration check, while invoking its ordered live validator. Windows records the
same ntdll.dll/c0000028 signature at `04:52:38.6370856Z` and `04:56:08.5967747Z`.
All four native failures in this checkpoint share fault offset `0000000000012d2f`;
that identifies a recurring signature, not its cause. No process memory was collected.

Viewer returns FAIL after an unsuccessful fixture sign-in. Its saved flags show
AuthLoaded=True, TargetSelected=True, TargetPathsSet=True, SignedIn=False and
SnapshotCreated=True. It waits on the real sign-in information dialog, whose
capture was inspected. Only that exact message on the verified fixture-owned Excel
process was acknowledged/closed; no credentials, permissions or sign-in result
were changed. The initial OK request and subsequent close requests allowed the
test to return its failure. Viewer content/reuse/Events acceptance is not established.
Do not mistake a live modal wait for a terminated process or a rendering pass.

The independent queue initially stopped after trying to git-restore the ignored
live-role report. Its result was already preserved. The queue-only cleanup was
corrected and execution resumed at full-chain, then Viewer after guarded cleanup
of an empty recovery process. No failed gate was overwritten or retried unchanged.
Combined launchers pass Receiving and then stop with one harness failure (1 PASS /
1 RED); Windows records ntdll.dll/c0000028 at `05:05:15.4257799Z`. The separate full
reusable Production/restart gate passes 2/2 without reduced flags, and Shipping
layout passes 1/1. No Excel Application1000 fault was observed in those two passing
gate windows or the Viewer window. These passes do not clear the other failures.
The queue completes with exit 1, preserving every failed gate. Full handler
preservation, native stability, visible/human comparison and Release 1 acceptance
remain unproven. Receiving's full activity regression is still required on this
candidate; definition equality is not a substitute for that behavior test.

Static regeneration passes all three JSON contracts and all 28 preceding module
limits. Counts are 178 components, 5,514 procedures and 123,883 lines (+1/+3/+111),
with eight literal Application.Run targets, 45 unresolved dynamic calls and 195
duplicate bodies unchanged. Both edited PowerShell scripts parse, 46 local Markdown
targets resolve, and diff/status review preserves unrelated user edits.

All test sessions are terminal and Excel is closed. Verification matches all 55
package pins, six preserved Receiving/Shipping source pins and three current Core
pins. The old verifier correctly rejects the intentionally changed catalog; the
new verifier replaces only the two superseded Core source expectations and adds
the new module while retaining every other pin. Raw reports remain ignored, and
the accepted deployment is not replaced by this isolated candidate.

## Reproduction and raw evidence

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-shipping-capability-guard -Phase RED -CheckActivityEvidence -CheckActivityFoundation -CheckShippingActivity
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-shipping-catalog-eight -Phase GREEN -CheckActivityEvidence -CheckActivityFoundation -CheckShippingActivity
```

The nine policy regressions were added after the recorded RED, without changing
policy implementation. Both recorded full invocations exit 1 for the stated
failures; no test flags omit the actual handler route.

Raw files are ignored under `reports/runtime/slice4be-shipping-activity/`:

- `catalog-eight-red.log`, `985ec7a1eede4e988cceaea49b5f2440/red.json`,
  `catalog-eight-red-native-faults.json`.
- `catalog-eight-green.log`, `7e1dcf1a37ce4553bc61612de89f14c8/green.json`,
  `catalog-eight-green-native-faults.json`, `catalog-eight-comparison.json`.
- `catalog-eight-build.log`, `catalog-eight-compile.log`,
  `catalog-eight-compiled-source.json`, `catalog-eight-compiled-comparison.json`,
  `catalog-eight-package-hashes.json`, `catalog-eight-source-hashes.json`.
- `catalog-eight-final-static.log` regenerates maintenance evidence after the
  final policy-test additions; `catalog-eight-static.log` is the earlier generation.
- `catalog-eight-final-pin-verification.json` and `verify-catalog-eight-pins.ps1`
  verify all preserved packages and the explicit current source expectations.
- `catalog-eight-gate-exits.json` and the matching named gate logs retain each
  independent release gate's status and time window.
- `catalog-eight-live-role-native-faults.json`,
  `catalog-eight-full-chain-native-faults.json`, `catalog-eight-native-signatures.json`,
  and `catalog-eight-viewer-signin-warning.png` retain sanitized failure observations.
- `catalog-eight-launchers-native-faults.json` retains the later launcher fault;
  the matching Viewer/Production/Shipping-layout native-fault files record empty windows.
- `catalog-eight-production/production-reusable-production.md` and
  `catalog-eight-shipping-layout/shipping-layout.md` retain the separate passing gates.

Only sanitized evidence is committed. No credential, raw authority workbook,
operational payload or machine/session report is added to source control.

## Required continuation

Implement typed owner-reported facts and real Shipping handler observations under
the normative outcomes. Preserve every preceding GREEN and test mixed/partial
Update, Remove and Stage outcomes, Send processing, denial/staleness and optional
tracking failure. Do not derive results from report text or newly applied catch-up
IDs. Full candidate handler validation remains blocked by the recorded fixture
bootstrap fault until stronger evidence establishes a valid route.

Source review also identifies a remaining ROW compatibility slot in reachable
`PersistShipmentRowsLocal` -> `HoldRowField`. D14 prohibits it; a focused packaged
save/reopen test must cover exact identity and unknown columns before correction.
This discovery is recorded in Plan 022 and is not approved as a compatibility
contract. It does not alter the catalog implementation or supersede the goal's
remaining Settings, comprehensive Viewer, recording, guide/comparison and release work.
