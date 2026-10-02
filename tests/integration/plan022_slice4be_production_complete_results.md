# Slice 4be Production Complete Run

## Broader gates on the corrected package

The unchanged, unpromoted complete-selection01 candidate passes the following
gates on its canonical five-package pins. Each controller restored settings,
preserved packages and restored any tracked test reports before terminal exit.
Excel was closed and the delayed Application1000/1001/1002 audit found zero
Excel failures for each gate. Raw reports and captures remain ignored.

| Gate | Evidence under reports/runtime/complete-selection01-regression | Result |
| --- | --- | --- |
| Packaged smoke | `smoke-60f4b5a4d14d432f8013082fbad2a3e2` | 86/86; exact prior check order; initial/final unassisted closure verified |
| Packaged layout | `layout-2d3ca5d270224d49867a7b5d384db208` | 18 requested size/page pairs, six activated/maximized pages, five native transitions, two distinct actual sizes; representative geometry preserved |
| Full Release1 chain | `chain-625b549f08094ec786e94c0ccecaf6fe` | 32/32 chain,48/48 live-role,15/15 warehouse creation; exact prior ordered PASS results retained |

Smoke ran2026-10-01 23:50:38.8715725-23:51:01.4770427 UTC; layout ran
23:51:39.4175655-23:51:56.6063880 UTC; chain ran
23:52:43.1959440-23:58:07.1783001 UTC. The chain harness completed its ordered
warehouse/Seed/Receiving/Production/Boxing/Shipping/restart sequence.

All six new layout captures were individually reviewed and their hashes recorded
in the ignored capture-review record. No overlap or out-of-bounds regression was
found. These are empty Run List and Settings views; they do not establish populated
list scrolling, multiline rendering or human usability acceptance. No runtime,
test or contract change was needed for this evidence-only checkpoint; no new RED
is claimed. Owner interruption/partial-submission tests, Complete Run observations
and full human/NAS Release1 acceptance remain open.

## Selected Process with an unallocated second Process

Test-only expansion on unchanged complete-selection01 passes68/68, preserving
all59 prior ordered identities/PASS results and adding nine checks. Controller
`production-run-local-controller/a9cfe91e4add40be81c59084cba2e839`, result
`slice4be-production-complete-baseline/508474226bc24e3bb0689568bff23bc7/green.json`,
2026-10-01 23:44:56.7100635-23:48:34.2268026 UTC. The fixture creates two released
Processes and a released Recipe through the existing form workflow, allocates
and checks in only the selected Process, then invokes the actual Complete Run
handler. The selected owner alone is entered; its exact allocations are consumed
and its new output key has the expected quantity. The second Process remains
unallocated, visibly NEEDS ALLOCATION, without an output key or completed state.
The batch remains incomplete. Captured-workbook custom values/formula, decoy,
saved operator bytes and other warehouse are preserved.

Shared42, five instrumented compiles, canonical pins, restored settings,
unassisted closure and the delayed zero native audit pass. All three captures
were individually reviewed. Long fields still clip; the multi-Process capture
shows the second Production Output row only partly visible at the captured size.
Scrolling/populated-list usability remains unverified; this observation does not
establish a new minimum-row contract. `complete-multi-static-01` retains290
components/6175 procedures/135521 lines,9 literal/45 unresolved calls,190 duplicate
candidates and28 non-growing caps; three schemas/377 parses pass. There is no
runtime change or new product RED in this expansion. Complete Run tracking,
interruption/partial-submission evidence and full acceptance remain open; the
broader gates above have now passed.

## Check In regression on the corrected package

On unchanged complete-selection01, Check In passes629/629 with all629 prior
activity03 identities and PASS results retained in order. Controller
`production-run-local-controller/cb210925e13843ac8bc7e4fb003aa362`, result
`slice4be-production-check-in-activity/ab2d6ec127e44ffb82b6e5a2a5edbd1e/green.json`,
2026-10-01 23:28:03.3323285-23:44:38.8953672 UTC. Shared42, five instrumented
compiles, canonical new-candidate pins, settings restoration, unassisted shutdown
and the delayed native audit with zero Excel failures pass. Excel shutdown again
took several minutes; the controller was retained until natural exit. Captures
were regenerated but not additionally reviewed in this regression. No runtime
change or new product RED is claimed. The later multi-Process gate above adds
selected-only owner evidence; the broader new-candidate gates are recorded above.

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

Check In629 was subsequently verified on this candidate above. This
one-Process positive fixture does not establish multi-Process completion;
the later gate above adds that focused case. Packaged smoke/layout and live-role/
full-chain subsequently passed above. Interruption behavior, observations and
full human/NAS Release1 acceptance remain open.
