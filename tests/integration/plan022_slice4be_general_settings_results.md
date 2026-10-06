# Slice 4be-A General Settings observations

Architecture v4.11 D18 catalog31 refines comprehensive recording for nine existing
General Settings controls. Runtime remains catalog30. The packaged test-first
checkpoint is **105 PASS / 86 FAIL across191 unique checks**; implementation and
acceptance remain open.

## Contract and protecting test

Cover config Reload/selection, carrier Add/Remove/Reset/selection, UOM selection
and connection-option choice/save through actual form callbacks. Carrier and
connection settings retain per-Windows-user storage. Recorded Changed outcomes
cannot imply warehouse inventory mutation. Preserve captured context, existing
Admin access, owner results, redaction and optional tracking failure behavior.
No new storage, permission, replay or business authority is introduced.

`Test-Slice4beGeneralSettings.ps1 -DeployRoot deploy/validation-guide-transfer-notice-01 -Phase RED`
uses generated warehouses, five compiled disposable probes, the existing D5
baseline and actual Settings handlers. The native Reset observer matches the
owned process and exact carrier question before choosing Yes/No. Its existing
UOM behavior remains the default; run the UOM regression after implementation.

## Verified RED

UTC2026-10-06 **08:48:15.0444652--08:52:00.7786003**. Ignored receipts under
`reports/runtime/`:

- Controller: `general-settings-controller/59ceb7ec87d94aa98a81b6c33dad0ce8`.
- Worker: `slice4be-general-settings/8b106ab8425c4d5491a9a6e5d08634dd`.
- Exact failure-set and preservation proof: controller `verification.json`.
- Static evidence: `general-settings-red-static-01/ratchet-verification.json`.

| Failure class | Assertions | Evidence |
| --- | ---: | --- |
| Missing catalog31 definitions/extension | 11 | Nine definitions and two extension/preservation checks fail on catalog30. |
| Missing observations | 64 | All16 ordinary actions pass independent owner checks but lack their expected request/outcome, metadata, redaction and integrity proof. These are missing records, not evidence of leaked values. |
| Incomplete expected recording | 1 | The16 unobserved actions cannot supply the expected32 observations. |
| Stale held form | 10 | Carrier Add and config selection remain active after sign-out, reauthentication and warehouse change; connection Save remains active after reauthentication/warehouse change; Reload changes staging after sign-out/reauthentication. |

Signed-out connection Save and cross-warehouse Reload preserve their tested
state. Do not generalize those two passes into full stale-form protection.
All16 ordinary owner-result checks, native Reset Yes/No/unchanged, warehouse
config preservation and prior activity preservation pass. Runtime source and
the frozen package hashes are unchanged. The candidate's earlier transfer499,
policy429, transfer71 and native-edit58 evidence retains its original scope.

Five instrumented compiles, restored settings, normal Excel closure and delayed
zero Excel Application failures pass. Two captures were directly reviewed:
the exact carrier Reset question and loaded General Settings layout. Other
generated captures are not claimed as additional direct reviews; no human
acceptance is claimed. Static metrics remain317 components,6328 procedures,
138507 lines,9/45 dynamic calls,190 duplicate groups and28 oversized caps.
All three static schemas validate and447 PowerShell scripts parse.

## Next action

Implement catalog31 vocabulary and actual-handler observations with explicit
local-owner outcomes and captured-context guards, then retain all191 checks in
GREEN. Add policy/storage and live permission-loss coverage, independent recorded
conclusions, affected UOM/Settings regressions and final acceptance evidence.
Do not change carrier/connection storage to repair missing observations.
