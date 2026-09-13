# Slice 4be.3 Designs publication source

Last verified 2026-09-13. Release 1 and Slice 4be remain active. This implements
the Designs owner read required by D18 publication; the Admin publisher does not
yet consume it or create the Events artifact. Viewer gains no canonical read.

## Contract and implementation

Architecture v4.11 D18 declares `PUBLICATION_EVENTS` on the existing Designs
query dispatcher, with captured WarehouseId/runtime-root strings and an escaped
`EVTSRC1` transport. The refinement implements approved owner/read separation,
exact identity, coverage and non-mutation rules under semantic inheritance.
Plan 022 and the controls catalog carry the same scope.

`cDesignsEventPublicationSource` reads only the existing captured Designs file.
It borrows a clean open source, rejects an unsaved source, or opens a transient
source read-only with events/macros disabled during open. It restores application
settings and closes only its transient workbook without saving. It never invokes
the create/ensure/save resolver. Required headers are resolved by normalized name;
missing/ambiguous headers and other-warehouse rows return unavailable coverage.
The thirteen declared fields retain every source line and exact EventID, escape
text losslessly, exclude PayloadJson/unknown columns, and do not claim verified UTC.

The runtime change is one 156-line, nine-procedure Designs class and four lines
in `modDesignsBridgeApi`. Core's existing generic query bridge carries the result;
there is no new Application.Run site. Both explicit source harnesses importing
that dispatcher also import the class. The class export uses CRLF for direct VBA
imports, as required by the previously discovered source-harness constraint.

## D13 evidence

`Slice4beDesignsPublicationSource.ps1` extends the actual Admin publication test.
Fixtures begin at Admin Generate Warehouse; Designs Domain applies a create and
release event, with distinct exact IDs, then the fixture adds an unknown leading
column and varies managed-header case. The actual Core/Domain dispatcher is tested
for full rows/identities/text, permitted fields, transient read-only mode/release,
source bytes, clean borrowing, dirty-source rejection without save, missing header,
missing file without recreation, root mismatch, changed target and valid empty source.

- Pre-implementation candidate `deploy/validation-publication-source-read`:
  **9 PASS / 27 FAIL**, including Designs **3 PASS / 12 FAIL**; no harness failure.
- New candidate `deploy/validation-designs-publication-source`:
  **21 PASS / 15 FAIL**, including Designs **15/15 GREEN**; no harness failure.
- Both runs retain the actual Admin Settings attempt/result fixture, successful
  public Generate Inventory Snapshot command, all four prior inventory source-read
  checks, and source-copy/canonical/activity byte preservation.
- The remaining fifteen failures are persisted Events-artifact requirements.
  Publication-only diagnosis cannot claim acceptance GREEN. Combined Viewer
  grouping/paging, actual all-source publication and policy-aware reads remain required.

The mixture now includes two owner-applied Designs events in addition to the
5,001 Inventory groups and one Activity group. The future global 5,000-group
publication must retain 4,997 Inventory groups, both Designs groups and the Activity
group, with source counts reconciling; direct owner-query success is supplemental
to that actual Admin publication requirement.

Four earlier setup runs failed while renaming/restoring a malformed-header fixture:
three recorded 4 PASS / 9 FAIL, one 4 PASS / 8 FAIL, each including a harness failure.
An explicit calibration showed the PowerShell COM rename had not persisted.
The final fixture helper performs and verifies that setup/restoration in VBA,
keeping the malformed source open to test clean borrowing independently of reopen.
These setup failures are not product RED. No runtime repair was based on them.

## Package and regression gates

All five isolated packages build and compile; Operations cold-start references
resolve inside the candidate. The compiled inventory has 196 components versus
195 previously. The new class and dispatcher account for the behavioral change.
Four other Designs modules differ only in VBA identifier capitalization, verified
against their actual packaged source; all other compiled component hashes match.
The first ad-hoc source comparison did not load its second package and is invalid;
the replacement review checks package acquisition and uses separate Excel instances.

Event Detail remains **34/34 GREEN**, including inspected default/maximize/restore
captures; Refresh failure remains **16/16 GREEN**. Settings is **187/187 GREEN**
with inspected Admin/Operations captures, populated Viewer checks pass, and
packaged smoke is **86/86 GREEN**. The full Release 1 chain finishes **31/31 GREEN**
(report 2026-09-13 12:37:59), including its actual ordered live child **48/48** and
Create Warehouse source integration **15/15**.

The first full Receiving run stopped at **606 PASS / 1 harness failure**:
`Intended navigation control focus unavailable` for `DISPOSITION_SELECT_AGGREGATE`.
The guard stopped before the intended native action. No product regression or
runtime cause is inferred. Excel closed; the complete unchanged-candidate retry
passes **854/854**, including the previously interrupted action. All 854 prior
GREEN identities remain, with no duplicate check. The original focus failure
remains recorded; the retry neither changes code nor weakens an assertion.

Static evidence records **203 components / 5,675 procedures / 126,352 lines**, **8 literal / 45
unresolved Application.Run calls**, **189 duplicate-body groups**, and **28
oversized-module ratchets**. No existing oversized module grows. The nine new
procedures implement the scoped owner read; no scanner candidate was deleted.
All three generated JSON reports validate and all five changed PowerShell files
parse. Document links and diffs are checked.

Final preservation verifies **170 prior package pins plus five candidate pins**
and **16 protected source pins**. Excel is closed and every test process is
terminal. The accepted deployment/NAS and unrelated handoff 067 (3 additions /
3 deletions) and untracked critique 023 remain untouched. Automated captures and
regression GREEN do not substitute for physical NAS or human acceptance.

## Reproduction and exact evidence

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-publication-source-read -ViewerPublicationOnly -Phase RED
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-designs-publication-source -ViewerPublicationOnly -Phase RED
```

Ignored evidence under `reports/runtime/`:

- `designs-publication-source-red.log`, `designs-publication-source-retry-red.log`,
  `designs-publication-source-borrowed-schema-red.log`, and
  `designs-publication-source-schema-calibration-red.log`: setup failures.
- `designs-publication-source-vba-schema-red.log` and
  `slice4be-viewer-groups/4998f3da36be45119c83e8ab901cc549/diagnostic-publication-red.json`:
  complete pre-implementation 9/27 RED.
- `designs-publication-source-focused.log`, `designs-publication-source-focused.json`,
  and `designs-publication-source-focused-index.json`: post-implementation 21/15;
  all nine prior GREEN identities retained, all 36 check identities unique.
- `designs-publication-source-build.log`, `designs-publication-source-compile.log`,
  `designs-publication-source-compiled.json`, `designs-publication-source-case-review.json`.
- `designs-publication-source-static.log`, `designs-publication-source-package-pins.json`.
- `slice4be-viewer-detail/b5ee2f0bb48a4f47a6dc5c1ad6ec845d/`: 34/34 and inspected captures.
- `slice4be-tracking-settings/da54477ee17247fb8b2dd97fdeecbabb/`: 187/187 and captures.
- `designs-publication-source-detail.log`, `designs-publication-source-refresh.log`,
  `designs-publication-source-settings.log`, `designs-publication-source-viewer.log`,
  `designs-publication-source-smoke.log`, `designs-publication-source-chain.log`.
- `designs-publication-source-receiving.log` and
  `designs-publication-source-receiving-focus-failure.json`: retained 606/1 setup failure;
  `designs-publication-source-receiving-retry.log`: complete retry.
- `designs-publication-source-receiving-green.json`,
  `designs-publication-source-receiving-preservation.json`, and
  `designs-publication-source-preservation.json`: 854 retained GREEN identities
  and final 175-package/16-source pin verification.

The accepted deployment, NAS workbooks and unrelated user changes remain outside
this work. Full Events publication/coverage, remaining Operations/Admin activity,
recording and conclusions, How-To/Diagnostic/Compare both, guide library,
physical NAS/multi-station and human acceptance remain open.
