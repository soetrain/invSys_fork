# Slice 4be.3 Events publication

Last verified 2026-09-13. Release 1 and Slice 4be remain active. This checkpoint
records tests and a working candidate for Core-owned Events publication under
Architecture v4.11 D18. The runtime implementation remains uncommitted pending
its required regression gates; this test/evidence checkpoint does
not accept Viewer projection reads/paging, comprehensive control coverage,
recording, guides, How-To/Diagnostic/Compare both, NAS or human acceptance.

## Contract and implementation

The ordinary Inventory snapshot owner captures its existing read-only source and
delegates Events publication to `cEventsPublication`. Inventory's Boolean/path
result remains independent; explicit Admin Generate Inventory Snapshot includes
the separate Events notice. Bootstrap without a matching allowed target defers
Events without retargeting. Core and Domain gain no UI.

The publisher reads Inventory, the existing Designs PUBLICATION_EVENTS query,
the Activity owner store, and the declared fixed Operations Shipping owner read.
It retains every permitted line and exact identity, groups durable records by
source/ID, and selects the newest 5,000 complete groups. Activity outcomes preserve
original result records; current-state Shipping evidence cannot invent historical
IDs or completed outcomes. Coverage separates unavailable counts from zero and
distinguishes warehouse BOM from station-profile Holds.

`modEventsPublicationStore` validates and hashes the complete candidate, checks
the written temporary bytes, then uses same-directory atomic replacement. Failure
retains the prior artifact and removes only its own pending file. The full Events
parser entry preserves the Activity/guide 1 MiB limits. Array serialization and
private parser calls avoid repeated whole-buffer copies; ordinary JSON validation
and string escaping remain in use. The internal store reader is not yet the
authenticated, policy-aware Viewer read path required for final acceptance.

Shipping resolves required headers by normalized name and excludes unknown user
columns. It borrows clean sources, rejects dirty sources, and closes only its
own transient read-only handle. Its field codec reuses the existing Shipping
escape routine through a direct typed call. The fixed Core/Operations bridge is
the single declared additional literal Application.Run site. Sixteen source
harnesses import the new Core class and its dependency closure.

## Focused evidence

The test uses the public Admin snapshot command. Generated/seeded fixtures use
actual Settings, Box Designer, Box Maker, Add and Hold handlers. Only the volume
source table receives the existing controlled 5,001-group fixture substitution;
the owner acquisition and publication logic execute normally. Every prior check
is retained, including source-copy and canonical/local/activity byte preservation.

- Calibrated pre-implementation `validation-designs-publication-source`:
  **27 PASS / 39 FAIL**, 66 checks, no harness failure. The required Events
  artifact is absent; the new failure/recovery cases cannot pass without it.
- `validation-events-publication-atomic`: **66/66 focused checks GREEN**,
  no harness failure. All twenty BOM fields and eleven Hold fields are compared.
- The actual Box Maker handler publishes the exact durable BOX_BUILD owner event
  before Add/Hold, the explicit Admin command or volume substitution. The earlier
  65-check set is retained with no missing IDs, duplicates or lost GREEN results.
- Locked-destination tests preserve Inventory success, report Events failure,
  retain the exact prior bytes, leave no pending file, and replace successfully
  after unlocking with a new publication identity.
- Missing and dirty BOM tests publish unavailable counts and no invented state,
  preserve source files/borrowed handles, and recover original available evidence
  after restoration. The missing source is not recreated; the dirty source is
  neither saved nor closed by publication.

The diagnostic-only switch deliberately requires `-Phase RED`; its zero failures
prove these focused assertions, not combined Viewer or Release 1 acceptance.

## Fixture and unsuccessful diagnostic history

Earlier 44/47 and 51/54 diagnostics isolated Shipping BOM acquisition failures.
A fresh-instance comparison established that owner-saved BOM copies open normally,
whereas the PowerShell COM unknown-column edit produced an unreadable saved copy.
The equivalent VBA edit saved and reopened successfully. The maintained fixture
now performs that calibrated edit in VBA and requires a read-only reopen before
publication. No Shipping reader or owner repair was inferred from those failures.
Several immediate-reopen setup attempts failed before a complete publication run;
they are harness evidence, not meaningful product RED.

The corrected older candidate passed 55/55 including the first seven atomic
checks. Temporary stage/error and borrowed/direct-query probes were then removed;
the final 65-check run uses the real publisher without those observers.

An earlier candidate build contained an invalid reserved VBA identifier and was
corrected before compile acceptance. One slow diagnostic was intentionally
cancelled after 1,347 CPU seconds, producing an RPC harness failure. It is not
GREEN or behavioral RED. The serialization/parser changes address observed
whole-buffer work; that cancelled run does not supply an isolated benchmark.

## Package, regression and preservation gates

All five candidate XLAMs build and compile; Operations cold-start references
resolve within the candidate. Packaged source inspection compares 196 previous
components with 199 current components: three additions, six intended edits,
and 34 changes limited to capitalization with identical string literals. There
are no unexpected source differences, and all inspected package bytes remain
unchanged.

Static evidence is **206 components / 5,700 procedures / 127,033 lines**, nine
literal and 45 unresolved Application.Run calls, 189 duplicate-body groups and
28 oversized-module ratchets. Plan 022 records the four-line snapshot-orchestrator
exception (1,753 to 1,757); the implementation remains in the new bounded class.
No other oversized-module growth or duplicate-body increase is authorized.

The atomic candidate passes Detail **34/34**, Refresh **16/16**, Settings
**187/187**, populated Viewer and packaged smoke **86/86**. Two unchanged-package
full-chain attempts stop at projection rebuild with a native Excel failure:
top-level **4 PASS / 1 harness failure**, ordered live child **32 PASS / 1 harness
failure**. Neither attempt is behavioral RED or acceptance GREEN. Receiving has
not yet run against this candidate.

The interrupted CompileOnly diagnostic completed **48/48** live-role checks after
an unsaved, unused Core procedure forced recompilation. This is diagnostic evidence
only: it changes the test condition and does not prove unchanged-package acceptance
or establish a native-crash root cause. A separate `validation-events-publication-finalized`
candidate persists compilation state. All five packages compile, Operations resolves
its cold-start dependency within the candidate, and all **199 extracted component
hashes match** the atomic candidate. Its unchanged-package full-chain attempt also
failed natively: top-level **4 PASS / 1 harness failure**, live child **27 PASS /
1 harness failure**, during Production Complete Run. Saving compilation state
therefore does not resolve the gate. No runtime source or accepted deployment was
changed for this investigation; the finalized candidate is diagnostic, not accepted.

Running the unchanged atomic ordered child by itself also fails at projection
rebuild (**32 PASS / 1 harness failure**), so the preceding full-chain phases are
not required to reproduce that failure. A second unsaved diagnostic bypasses only
the Events Publish call and passes **48/48**. Because compile-only already passed
with publication enabled, this does not isolate publication as the crash cause.
The bypass is never saved, shipped or treated as acceptance. Its first attempt
stopped before workflow execution because an exact-case source anchor did not
match VBE-normalized identifiers; the corrected lookup requires exactly one match.

Preservation verification matches all **175 historical package pins**, **15 prior
publication-candidate pins** and **15 protected source files**. The remaining
protected Shipping source has exactly the reviewed EscapeHoldField visibility
change after line-ending normalization; it is not claimed unchanged. Verification
ran with Excel closed. The latest held-line capture was inspected: it shows the
held item and pending sync, not a completed shipment or Action Path conclusion.

## Reproduction and evidence

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-designs-publication-source -Phase RED -CheckViewerPublication -ViewerPublicationOnly
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-events-publication-atomic -Phase RED -CheckViewerPublication -ViewerPublicationOnly -CaptureEvidence
```

Ignored reports under `reports/runtime/`:

- `events-publication-calibrated-red.log` and
  `slice4be-viewer-groups/8e829a3db38d43e289a1a8e129a8612b/`: 27/38 RED.
- `events-publication-boxmaker-red.log` and
  `slice4be-viewer-groups/3e99ba87c54149d5a0696d1451579f95/`: expanded 27/39 RED.
- `events-publication-boxmaker-focused.log` and
  `slice4be-viewer-groups/e4240a4139864e759edcf6aca451b314/`: 66/66 and inspected
  held-line capture; `events-publication-boxmaker-check-comparison.json` preserves
  all prior check identities and GREEN results.
- `events-publication-atomic-focused.log` and
  `slice4be-viewer-groups/ab2373ef1362475295290428cb857c7e/`: 65/65, source
  preservation, coverage, command timing and held-line capture.
- `events-bom-open-matrix.json`, `events-bom-column-com-probe.json`, and
  `events-bom-column-vba-probe.json`: isolated fixture calibration.
- `events-publication-atomic-build.log`, `events-publication-atomic-compile.log`,
  `events-publication-atomic-compiled.json`, `events-publication-atomic-code-review.log`,
  and `events-publication-atomic-case-review.json`: package evidence.
- `events-publication-atomic-static.log`: generated static/schema checks.
- `events-publication-performance-cancellation.json`: intentional slow-run stop.
- `events-publication-atomic-regressions.json`, `events-publication-atomic-chain.log`
  and `events-publication-atomic-chain-retry.log`: completed regressions and the
  two native full-chain failures.
- `events-live-compileonly-825aeefb57154b0299bd7ca6a4e8aef1/result.json`: 48/48
  diagnostic, explicitly not acceptance.
- `events-publication-finalized/`: saved compilation, source comparison and pins.
- `events-live-unchanged-checks.json` and
  `events-live-bypasspublication-c46564f602fd4b37b91d4d2c977ea738/result.json`:
  isolated live failure and non-acceptance bypass diagnostic.
- `events-publication-preservation.json`: historical package/source preservation.

Accepted deployment and operational/NAS workbooks remain outside this candidate.
Unrelated handoff 067 and critique 023 must remain unstaged and unchanged.
