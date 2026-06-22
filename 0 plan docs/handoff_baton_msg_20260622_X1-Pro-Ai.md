# Baton Message - X1-Pro-Ai Move, Shipping Add Fix, NAS Deploy Access

Date: 2026-06-22

This baton is for the next Codex chat running on `X1-Pro-Ai`. The current development machine can build and test the repo, but this Codex process cannot reach the NAS SMB shares even though the user's normal Windows PowerShell can. The next chat should first restore deployment capability, then continue user-side Shipping testing. Do not start aggregator work until deployment is solved unless the user explicitly redirects.

## Current Goal

Get the latest Shipping/Box Maker fixes deployed where the user can actually load them in Excel, then continue user-side testing of the Shipments form.

The immediate bug being chased:

> Shipping form shows `T28 v1` with `NAS Inv=20`, `Projected Inv=20`, `Locked=0`, but clicking `Add` for Qty `2` returns: `T28 v1 requires 2. but only 0. is available for that version.`

Current code has a fix and tests for this, but the user still saw the popup before reload/deploy was verified. Treat deployment/add-in reload as the first suspect, not necessarily a failed code fix.

## Important Environment Facts

Repo path on current machine:

```text
/mnt/c/Users/justu/source/repos/invSys_fork
```

Expected repo path on `X1-Pro-Ai` may differ. Start by locating/opening the same repo and checking the latest changes.

NAS shares from user Windows PowerShell:

```powershell
Test-Path '\\100.84.136.19\invSysWH1'
# True

Test-Path '\\100.84.136.19\invSysWH1\Addins'
# False
```

Current Codex process on the laptop tested:

```text
\\100.84.136.19\invSysWH1      -> False
\\100.84.136.19\invSys-deploy  -> False
```

Interpretation:

- The NAS itself is reachable from the user's normal Windows session.
- This Codex/WSL process does not have the same SMB network/auth context.
- `invSysWH1` is the warehouse runtime root, not an add-in distribution root.
- `invSys-deploy` is the documented staging area where Codex should normally push built `.xlam` artifacts.

Relevant doc:

```text
0 plan docs/xlam_invSys/NAS.md
```

Key contract from that doc:

- `invSysWH1`: warehouse runtime root. Codex service account should be read-only here.
- `invSys-deploy`: staging area where Codex pushes built `.xlam` artifacts. Codex service account should have read/write here.
- `invSys-backups`: locked down.

The user says the screenshot shows the actual `invSysWH1` NAS runtime root. There is no `Addins` folder there. The left-side `invSys-deploy` share is the normal deploy target, but because invSys is not operational yet, the user may ask for direct deployment to `invSysWH1`. Do not assume a direct runtime-root copy is safe; make the target explicit and back it up.

## Current Code Changes To Preserve

The current working tree is dirty and contains many older changes from this Shipping/debugging thread. Do not revert broad files. Work with the existing dirty state.

Most recent relevant Shipping Add fix:

### `src/Shipping/Forms/frmShipmentsTally.frm`

`CommitCurrentLine` now computes:

```vba
displayedAvailableQty = SelectedShippableProjectedInventoryText()
```

and passes it to:

```vba
modTS_Shipments.ShipmentsFormCommitLine(..., report, displayedAvailableQty)
```

`SelectedShippableProjectedInventoryText()` was added. It first reads the selected shippables listbox row, then falls back to cached `mShippables` by `ROW`, box, and version. This was added because the live form can have text fields populated even when the listbox selected row is missing/stale.

### `src/Shipping/Modules/modTS_Shipments.bas`

`ShipmentsFormCommitLine` signature now includes:

```vba
Optional ByVal displayedAvailableQty As Variant
```

For `ADD`/reserve, it builds:

```vba
Set versionAvailabilityOverrides = ShippingVersionAvailabilityOverride(rowValue, descriptionValue, displayedAvailableQty)
```

and passes it into:

```vba
BuildSelectedShipmentRowsDeltas(..., "Warehouse", errNotes, versionAvailabilityOverrides)
```

`BuildSelectedShipmentRowsDeltas` now accepts optional `versionAvailabilityOverrides`.

`SelectedVersionInventoryAvailable` now accepts that override and treats override quantities as already projected, so it does not subtract local staged or NAS reserved quantities again.

Additional fallback added:

- `VersionInventoryHasAnyPositiveQty(versionInv)`
- `CurrentRowInventoryAvailableQty(invLo, rowVal, itemName)`

If the version ledger returns no usable quantity and the picker/read-model availability is zero, it falls back to current row `TOTAL INV` instead of treating the version as unavailable. This is specifically to avoid `requires 2 but only 0` when the current form/display has valid inventory but the reconstructed version ledger is empty.

### `tests/unit/TestPhase6CoreSurfaces.bas`

Added/updated:

```vba
TestShippingAdd_UsesDisplayedProjectedInventoryWhenVersionLedgerIsEmpty
```

Important: the final version of this test no longer passes an explicit displayed availability override. It covers the no-override path:

- `T28`
- `ROW=89`
- `TOTAL INV=20`
- two active BOM versions in `ShippingBOMView`
- empty/no useful version ledger
- `Add` Qty `2` should not produce `requires 2 but only 0`

The test is registered in:

```text
tools/run_phase6_excel_validation.ps1
```

near the Shipping Add tests.

## Tests Already Passed On Current Machine

Build:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools/build-xlam.ps1 -Apply
```

Focused tests:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools/run_phase6_excel_validation.ps1 -StartAt 121 -EndAt 126
```

Result:

```text
PHASE6_VALIDATION_OK
PASSED=6 FAILED=0 TOTAL=6
RANGE=121-126 AVAILABLE=166
```

Broad Shipping/Box Maker regression slice:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools/run_phase6_excel_validation.ps1 -StartAt 119 -EndAt 144
```

Result:

```text
PHASE6_VALIDATION_OK
PASSED=26 FAILED=0 TOTAL=26
RANGE=119-144 AVAILABLE=166
```

These tests passed after the no-override fallback was added.

## First Steps On X1-Pro-Ai

1. Confirm the repo and dirty state.

```bash
pwd
git status --short
```

Do not revert user/generated changes. The repo has many modified fixture/build files from Excel validation.

2. Confirm X1-Pro-Ai can reach NAS shares from the Codex command environment, not just from a normal user PowerShell.

```powershell
Test-Path '\\100.84.136.19\invSysWH1'
Test-Path '\\100.84.136.19\invSys-deploy'
Get-ChildItem '\\100.84.136.19\invSys-deploy'
```

If these fail in Codex but pass in normal PowerShell, stop and report that Codex still lacks SMB/auth context. Do not pretend deployment is done.

3. Rebuild from X1-Pro-Ai.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools/build-xlam.ps1 -Apply
```

4. Re-run the focused test first.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools/run_phase6_excel_validation.ps1 -StartAt 121 -EndAt 126
```

5. If focused passes, run the broader slice.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools/run_phase6_excel_validation.ps1 -StartAt 119 -EndAt 144
```

## Deployment Work Needed

Current repo does not have a simple operator-facing deploy command that matches the present NAS reality. Build creates `.xlam` outputs in:

```text
deploy/current/
```

Expected files:

```text
invSys.Core.xlam
invSys.Inventory.Domain.xlam
invSys.Designs.Domain.xlam
invSys.Receiving.xlam
invSys.Shipping.xlam
invSys.Production.xlam
invSys.Admin.xlam
```

The next session should add a small deploy script rather than manually copying by memory.

Suggested script:

```text
tools/deploy_current_xlams_to_nas.ps1
```

Suggested behavior:

- Parameters:
  - `-TargetRoot` required, e.g. `\\100.84.136.19\invSys-deploy`
  - `-SourceRoot` default `deploy/current`
  - `-Backup` switch default true
  - `-WhatIf` support if easy
- Verifies all required `.xlam` files exist and are non-zero.
- Creates target root if allowed and appropriate. If target is `invSysWH1`, do not create new folder structure without user approval.
- Backs up overwritten `.xlam` files to a timestamped folder under target, e.g. `_backup_xlam_YYYYMMDD_HHMMSS`.
- Copies files.
- Verifies target file sizes match source sizes.
- Writes a small deployment manifest text/json with timestamp, source path, target path, file names, sizes.

Normal target should be:

```text
\\100.84.136.19\invSys-deploy
```

If the user explicitly wants direct testing from `invSysWH1`, clarify whether the files should be copied to:

```text
\\100.84.136.19\invSysWH1
```

or to a new subfolder under it. The user said direct deployment may be acceptable because invSys is not operational yet, but the architecture doc marks `invSysWH1` as warehouse runtime, not add-in staging.

## Do Not Confuse These Roots

`invSysWH1`:

- Runtime warehouse root.
- Contains config/inbox/outbox/snapshots/data workbooks.
- Not normally the add-in update root.

`invSys-deploy`:

- Deployment staging share.
- Codex should have read/write through service account.
- Safe place to push built `.xlam` artifacts.

`PathSharePointRoot`:

- WAN publish/distribution root.
- Later expected to contain `Addins`, `Events`, `Snapshots`, etc.
- Users eventually download/update from SharePoint distribution.

## Aggregator Decision

The user agrees aggregator is needed, but not as a workaround for this Shipping deploy issue.

Current agreed order:

1. Restore deployment capability and user-loadable add-in updates.
2. Continue Shipping/Box Maker user-side testing.
3. Then develop/verify aggregator for two-warehouse WAN sync.

Aggregator should not be in the hot path for Shipping locks. Shipping locks/reservations must remain local to the authoritative warehouse processor. Aggregator/HQ is advisory/global visibility from published artifacts.

Relevant docs:

```text
0 plan docs/xlam_invSys/WAN dev slices.md
0 plan docs/xlam_invSys/WAN proving path.md
```

Key invariant from WAN docs:

```text
PathDataRoot workbooks are warehouse-authoritative.
PathSharePointRoot is publish/distribution only.
WH1 and WH2 are independent.
HQ reads only published artifacts from SharePoint.
Two-warehouse HQ snapshot is advisory only and has no write-back path.
```

## User-Side Test To Repeat After Deployment

After the updated Shipping add-in is loaded by Excel:

1. Open Shipping Shipments form.
2. Confirm display shows:

```text
T28 v1  NAS Inv 20  Projected Inv 20  Locked 0
T29 v1  NAS Inv 20  Projected Inv 20  Locked 0
```

3. Enter:

```text
Ref: 31 or new ref
Box: T28
Version: v1
Qty: 2
Carrier: UPS
```

4. Click `Add`.

Expected:

- No `requires 2 but only 0` popup.
- Row appears in Shipments list.
- Area remains `Warehouse`.
- Inventory is locked/reserved according to current design.
- Projected Inv should update immediately.

If the same popup still appears after confirmed deployment/reload, inspect the live loaded add-in path first. The most likely remaining cause is Excel still using an old `invSys.Shipping.xlam`.

## How To Check Loaded Add-In Path In Excel

Use Immediate Window or a small macro:

```vba
? Workbooks("invSys.Shipping.xlam").FullName
```

Compare it against the deployed target and file modified time.

If Excel has an older add-in open:

- Close Excel completely.
- Reopen from the newly deployed `.xlam`.
- Re-test.

## Communication Style For Next Chat

The user is testing fast and expects bugs to become tests. Keep implementing and verifying. For every user-side regression:

- Reproduce with a focused test where possible.
- Fix the smallest path that explains the user-visible behavior.
- Register the test in `tools/run_phase6_excel_validation.ps1`.
- Run focused slice and a broader adjacent slice.
- Do not leave dead experimental hooks.

## Known Risk

There are many accumulated dirty files and generated Excel fixtures in the working tree. Avoid broad cleanup unless explicitly requested. Only touch code/tests/scripts needed for the current deploy and Shipping fix.

