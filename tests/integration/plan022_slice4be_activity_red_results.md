# Plan 022 Slice 4be.1: first packaged activity RED

Last verified: 2026-09-07. Architecture v4.11 D18 and the detailed shared
Events/How-To/Diagnostic contract are approved. Plan 022 remains the current
implementation plan; user-supplied critique 023 is advisory.

## Scope and result

The first activity test adds optional assertions to the proven packaged D5
fixture. It invokes the actual Admin Settings Save Value and Production UOM
Send/Retrieve form handlers through unsaved test instrumentation in the loaded
XLAMs. No VBA runtime implementation or package was changed.

On the unchanged `b4ce6d9` packages: **19 PASS / 12 FAIL**, with no harness
exception. All **18 existing D5 assertions pass**. The additional direct-service
case also passes: invoking the Core UOM service does not masquerade as a user
control action. Three real-handler cases lack their required activity records,
causing four evidence assertions per case to fail:

| Case | Existing workflow result | Missing D18 evidence |
|---|---|---|
| Admin Settings Save Value | Changed value saved and read back | Requested/completed records, stable correlation, owner/context/data effect, redacted payload |
| Production UOM Retrieve | Catalog version incremented and staging removed | Requested/completed records, stable correlation, owner/context/data effect, redacted payload |
| Denied Production UOM Retrieve | Staging preserved; Config file unchanged | Requested/denied records, stable correlation, owner/context/data effect, redacted payload |

The absence of `Training\Activity\<WarehouseId>` evidence after a successful
fixture/action is the missing product behavior. It is not a missing warehouse
fixture or a failed command. Twelve assertions do not mean twelve independent
root causes: all three cases currently have no activity implementation.

Fixtures enter through packaged Admin warehouse generation, use disposable
credentials and isolated roots, and are removed afterward. Only check names
and booleans are reported; payloads and credential material are never printed
or persisted in reports. The assertions explicitly reject credential values,
paths, input values and backend handler names if records become available.

## Baseline and harness checks

- Cold-start Operations dependency location: PASS.
- Explicit VBA compilation: **5/5 PASS**.
- Packaged smoke: **81/81 PASS**.
- Both changed/new PowerShell test files parse successfully; diff/whitespace
  checks pass. No changes to deployed packages or runtime source are included.
- One initial focused launch stopped at the Excel-closed precondition after
  smoke validation left an empty process. Read-only COM inspection confirmed
  zero workbooks. That isolated process was closed before the successful RED
  run; the precondition failure is not behavioral RED.

## Reproduce and next gate

With Excel closed:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -DeployRoot deploy/current -Phase RED -CheckActivityEvidence
```

This is an intentionally failing opt-in D18 test. Ordinary D5 invocation keeps
its existing assertions and separate report directory. Raw check output is
ignored at `reports/runtime/slice4be-activity/red.json`.

Next implement the approved headless observation boundary and the catalogued
handlers against this RED, adding required policy/context, hashing/atomic-write,
re-entrancy, failure-isolation and additional coverage cases before claiming
4be.1 GREEN. Do not satisfy only record existence while omitting the approved
storage, ownership or security contract. Comprehensive control coverage,
Settings, Viewer publication, Action Paths and final Release 1/UAT gates remain
unfinished. No runtime completion, fresh full-chain/live-role result or human
acceptance is claimed by this evidence.
