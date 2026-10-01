# Slice 4be Production Complete Run

## D15 selected-Process prerequisite

Architecture v4.11 D15 requires one selected Process at a time. Source review
found that `frmProduction.CompleteProductionRun` instead called
`CompleteReusableRun` when `ActiveRunProcess()` was empty. The test exercises the
actual `mBtnManagerApplyOutput_Click` through an unsaved adapter in the packaged
Operations XLAM. Its completion-owner counters observe entry without replacing
the owner, writer, processor or inventory reads.

On the unchanged `validation-production-check-in-activity-03` candidate,
the focused RED is55 PASS/4 FAIL/59 unique checks. Controller
`production-run-local-controller/47cf6f7264754377a97062c4fccf386c`, result
`slice4be-production-complete-baseline/418bd17ee6a2429fbb77a1a1fcc76834/red.json`,
2026-10-01 23:18:29.2097244-23:20:54.7677237 UTC. The selected positive case
completes, consumes the exact allocated entities and creates a fresh output key
with the expected quantity. With no Process selected, the actual handler enters
the whole-run owner, changes staging, consumes input inventory and presents batch
success. The four failures protect owner non-entry, unchanged staging, unchanged
exact input balances and the selection-required message. Shared42, five compiles,
package/settings preservation, unassisted closure and the delayed zero native
Excel failure audit pass. Both captures were individually reviewed.

An earlier attempt stopped because the test had not loaded `RestartPins`:
controller `207bf758a8414587ab15f7598c8fb954`, result
`slice4be-production-complete-baseline/f73e6f86b0a944ab91f0b5db674828de/red.json`.
Its39 PASS/1 harness failure is not product RED. Settings/packages were restored,
Excel closed unassisted and its delayed native audit was zero. The helper now
loads the established fixture utilities inside the test's fixture scope.

The correction follows existing D15: reject an empty Process selection before
batch-note synchronization, Actual Output staging or completion-owner entry,
using the existing owner wording `Choose one Process before Complete Run.`
The selected owner remains `CompleteReusableProcess`; the form's whole-run
fallback is removed. No new event catalog or tracking contract is introduced.

New unpromoted candidate `validation-production-complete-selection-01` was built
and cold compiled under `complete-selection-build-01`,23:21:53.8385555-
23:22:31.4543385 UTC. All five compiles and cold Operations startup pass. Exactly
`invSys.Operations.xlam/frmProduction` changes among283 compiled components;
282 remain identical. Frozen activity03, settings and normal shutdown are
preserved; the delayed native audit is zero.

Focused GREEN passes59/59 with all59 ordered identities and all55 prior passes
preserved. Controller `production-run-local-controller/0b52b368e4e441e18ef5b788d5065873`,
result `slice4be-production-complete-baseline/b97fc38d64124c58813d9ba37adb0fee/green.json`,
23:22:51.0809958-23:25:10.2943940 UTC. Shared42, five instrumented compiles, new
candidate pins, settings restoration, unassisted closure and the delayed zero
native audit pass. Both new captures were individually reviewed: the refusal
message is legible and selected completion remains visible. Long fields still clip.
Raw captures/reports stay ignored.

`complete-selection-static-01` retains290 components/6175 procedures,9 literal/
45 unresolved calls and190 duplicate candidates. Runtime lines decrease by one
to135521; all28 module caps are non-growing. Three schemas and377 PowerShell
parses pass. The RED baseline is `complete-baseline-static-01` (135522 lines).

The activity03 Check In629-check evidence remains historical until rerun on this new candidate. This
one-Process positive fixture does not establish multi-Process completion,
interruption behavior or Complete Run observations. Those, packaged smoke/layout,
live-role/full-chain and full human/NAS Release1 acceptance remain open.
