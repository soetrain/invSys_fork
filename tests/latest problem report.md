**Problem Report: Shipping Form-Path Tests Wedging Excel**

While working on the Shipments form inventory-refresh issue, I attempted to add two additional regression tests that directly exercised the user-facing form paths. Both tests caused the Excel validation harness to hang indefinitely and leave hidden Excel/VBE processes holding the deployed XLAM files open.

**Context**
The user-facing bug was:

- Shipping an order did not visibly change `Projected Inv`.
- `NAS Inv` did not update after processor catch-up.
- Clicking `Refresh` on the Shipments form appeared to do nothing.

A lower-level regression test already passes:

- `TestShippingReserve_RunBatchRefreshUpdatesNasInvFromProjected`

That test exercises the processor/read-model path and verifies that after a reserve event is processed, the operator workbook and shippables data can show the updated NAS inventory.

The missing test coverage was closer to the actual userform/button experience.

**Attempted Tests**
I attempted to add two tests:

1. `TestShippingAdd_PostImmediatelyUpdatesProjectedOverlay`

Purpose:
Verify that using the same path as the Shipments form `Add` button immediately updates the projected inventory overlay.

Approach:
- Set up a shipping runtime session.
- Create a shipping inbox workbook.
- Create an operator workbook with `invSys`, `ShipmentsTally`, and `ShippingBOMView`.
- Call the Shipping add-in macro:
  - `ShipmentsFormCommitLine`
- Assert that `PendingBoxVersionInventoryOverlayText` returns the lowered projected quantity.

Observed result:
- The test hung before returning from the macro call.
- No VBA failure was surfaced to the harness.
- Excel stayed open as a hidden/mostly VBE process.
- Subsequent `build-xlam.ps1` runs failed because deployed XLAM files were locked.

2. `TestShippingRefresh_ProcessesReserveAndUpdatesVisibleNasInv`

Purpose:
Verify that the Shipments form `Refresh` action processes queued reserve events, publishes the inventory snapshot, refreshes the operator workbook, and reloads visible shippable NAS inventory.

Approach:
- Seed canonical inventory at qty 10.
- Seed a stale snapshot also at qty 10.
- Queue a `SHIP_RESERVE` event for qty 1.
- Create an operator workbook still showing qty 10.
- Call the new Shipping add-in macro:
  - `ShipmentsFormRefreshRuntimeInventory`
- Assert operator `invSys.TOTAL INV` and visible shippables NAS inventory become 9.

Observed result:
- The test also hung indefinitely when invoked through the Shipping add-in macro boundary.
- Again, no normal VBA error returned to the PowerShell harness.
- Excel/VBE process remained open and locked `deploy/current/*.xlam`.

**Symptoms**
When the tests wedged:

- PowerShell validation output stopped at:
  - `Starting test ...`
- No `PASS`, `FAIL`, or VBA error was emitted.
- `Ctrl+C` stopped the PowerShell command, but Excel remained alive.
- `build-xlam.ps1` then failed with file lock errors like:

```text
Cannot remove item ... deploy\current\invSys.Core.xlam:
The process cannot access the file because it is being used by another process.
```

- `Get-Process EXCEL` showed processes with `MainWindowTitle` like:
  - `Microsoft Visual Basic`

I had to forcibly stop those stale Excel processes before rebuilding.

**Likely Failure Area**
The hangs appear specific to invoking certain Shipping add-in public procedures from the validation harness via `Application.Run`, especially procedures that cross into form-adjacent or runtime-refresh paths.

Suspect areas:

- `Application.Run` into `invSys.Shipping.xlam` while the test harness workbook is active.
- Shipping add-in procedures that call:
  - `ActiveWorkbook`
  - `SheetExists`
  - `GetInvSysTable`
  - `BoxMakerFormLoadShippableVersionInventory`
  - `RunBatchAndRefreshOperatorWorkbook`
  - `RefreshShippingBomViewForWorkbook`
- Hidden modal VBA/runtime dialogs not being surfaced to PowerShell.
- Workbook/add-in locking from the validation harness opening deployed XLAMs.
- Re-entrancy or active workbook confusion between:
  - harness workbook
  - operator workbook under test
  - deployed `invSys.Shipping.xlam`
  - deployed `invSys.Core.xlam`
- Possibly a silent compile/runtime dialog in VBE, since `MainWindowTitle` showed `Microsoft Visual Basic`.

**What Was Kept**
The production fixes were kept:

- Shipments form `Refresh` now calls a module helper that runs `RunBatchAndRefreshOperatorWorkbook` before reloading form data.
- Add/lock path now writes projected inventory overlay immediately after local reserve/deduction.
- Overlay calculation now prefers current operator `invSys` values before slower/staler lookup paths.
- Snapshot publication after processor events remains in place.

The unstable experimental tests and temporary validation hooks were removed to avoid dead or hanging test clutter.

**Current Passing Coverage**
This shipping block passes:

```text
PHASE6_VALIDATION_OK
PASSED=15 FAILED=0 TOTAL=15
RANGE=119-133 AVAILABLE=151
```

Important passing tests include:

- `TestShippingAdd_DefaultsOrderToWarehouseArea`
- `TestShippingAdd_BlankCarrierRequiresCarrier`
- `TestShippingRemove_LockedRowReleasesInventory`
- `TestShippingSentRows_DoesNotIncreaseProjectedInventoryOverlay`
- `TestShippingSentRows_FullRunNeverIncreasesProjectedInventory`
- `TestShippingProjectedOverlay_PersistsAcrossRestartUntilNasCatchesUp`
- `TestShippingReserve_RunBatchRefreshUpdatesNasInvFromProjected`

**Need Help With**
We need a stable way to test the actual Shipments form/button-level behavior without wedging Excel.

Useful help would be:

1. Determine why `Application.Run` into Shipping add-in public procedures hangs instead of returning or failing.
2. Find whether a hidden VBA modal/compile dialog is being opened.
3. Identify whether active workbook/add-in context is wrong during test execution.
4. Build a safer test seam for userform-backed behavior, maybe by moving form actions into pure module functions that:
   - accept explicit workbook arguments
   - avoid `ActiveWorkbook`
   - avoid showing forms/msgboxes
   - avoid `Application.Run` across add-in boundaries where possible
5. Add a timeout/cleanup strategy to the validation harness so a hung Excel test does not leave locked XLAMs behind.

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

**Problem Report: Shipments Sent Dirty Transaction Increased Inventory**

**Summary**
The latest user-side Shipments test shows a serious dirty transaction after clicking `Shipments Sent`: visible inventory values increased instead of staying deducted or catching up to the already-deducted projected value.

Observed screen state after `Shipments Sent`:

| Box | Version | NAS Inv | Projected Inv | Locked |
| --- | --- | ---: | ---: | ---: |
| T24 | v1 | 5 | 5 | 0 |
| T26 | v1 | 8 | 8 | 0 |
| T27 | v1 | 10 | 10 | 1 |
| T27 | v2 | 20 | 19 | 1 |

The user reports T26 and T27 increased after clicking `Shipments Sent`. That must never happen. A shipment is an output from inventory. It is always a deduction from inventory, never an add-back.

**Expected Behavior**
For shipment workflow:

1. `Add` posts an order row to the Shipments listbox and locks inventory from use by others.
2. `Projected Inv` should immediately reflect the user-side deduction.
3. `To Shipments` moves the row from Warehouse to Shipments/dock area. It must not add inventory.
4. `Shipments Sent` completes the shipment. It must not add inventory.
5. After processor catch-up, `NAS Inv` should move toward the deducted authoritative value, not back upward.
6. `Locked` should clear only for completed/released rows and should not leave stale locks.

**Actual Behavior**
After `Shipments Sent`, at least these visible values moved upward:

- T26 v1 increased back to `NAS Inv 8`, `Projected Inv 8`.
- T27 v1 increased back to `NAS Inv 10`, `Projected Inv 10`, while still showing `Locked 1`.
- T27 v2 still shows `Projected Inv 19` and `Locked 1`.

This indicates the form or refresh path is treating an old NAS snapshot or stale version inventory source as authoritative and overwriting the user-side deducted state.

**Why Existing Tests Did Not Catch It**
Existing tests were too narrow. They proved selected helper paths did not increase projected inventory, but they did not reproduce the full user-side state transition visible in the Shipments form.

Likely gaps:

- Tests validated lower-level inventory functions instead of the same form-backed path used by `Shipments Sent`.
- Tests used direct table state and did not reload the shippables list after completion the way the form does.
- Tests checked one row/version at a time, not a mixed visible list with multiple boxes and versions.
- Tests did not assert that every visible shippable row is monotonic after `Shipments Sent`.
- Tests did not verify that stale NAS/version-log values cannot overwrite `Projected Inv`.
- Tests did not verify that `Locked` cannot stay positive after all corresponding shipment rows are completed or absent.
- Tests did not run the exact sequence the user ran: Add/order posted, lock visible, Shipments Sent, form reload/refresh, then compare visible `NAS Inv`, `Projected Inv`, and `Locked`.

**Required Regression Tests**
The test suite needs new or redeveloped tests that model the user-visible form contract, not just internal helper behavior.

**Test 1: Shipments Sent Never Increases Visible Inventory**
Create a test with multiple shippables, including at least:

| Box | Version | Starting NAS | Starting Projected | Starting Locked |
| --- | --- | ---: | ---: | ---: |
| T26 | v1 | 8 | 7 | 1 |
| T27 | v1 | 10 | 9 | 1 |
| T27 | v2 | 20 | 19 | 1 |

Run the same completion path used by the Shipments form for `Shipments Sent`.

Then reload the shippables view and assert:

- `NAS Inv` must not increase for any row touched by the shipment.
- `Projected Inv` must not increase for any row touched by the shipment.
- `Projected Inv <= NAS Inv` unless there is a clearly documented exceptional state.
- Completed shipment rows must not still contribute to `Locked`.

Failure condition:

- Any visible row has `NAS Inv after > NAS Inv before`.
- Any visible row has `Projected Inv after > Projected Inv before`.
- Any completed row still contributes to `Locked`.

**Test 2: Stale NAS Snapshot Cannot Overwrite Projected Deduction**
Set up:

- Canonical/current user-side projected value is already deducted.
- NAS snapshot still contains the older higher quantity.
- Shipment row is completed with `Shipments Sent`.

Assert:

- Reloading the form must preserve the lower projected value.
- The stale higher snapshot must not overwrite `Projected Inv`.
- If NAS has not caught up, the difference should remain visible as `NAS Inv > Projected Inv`, not collapse back upward.

**Test 3: Completed Reservations Do Not Appear Locked**
Set up:

- A shipment row has a reservation/lock.
- `Shipments Sent` completes it.
- The row is removed from active Shipments.

Assert:

- Reservation ledger marks the row completed.
- `ShipmentsFormLoadNasReservationTotals` does not count the completed reservation.
- The visible `Locked` column is `0` for that box/version unless another active shipment row exists.

**Test 4: Multi-Version Rows Are Checked Independently**
Use one box with multiple versions, for example:

- T27 v1
- T27 v2

Ship one version only.

Assert:

- Only the shipped version changes.
- The other version does not get an add-back.
- The other version does not inherit or retain the shipped version's lock.

**Suspected Code Areas**
Likely areas to inspect:

- `ShipmentsFormRunShipmentsSentRows`
- `ApplyShipmentsSentRowsInventory`
- `ApplyShipmentsSentVersionInventoryOverlay`
- `PendingBoxVersionInventoryOverlayText`
- `BoxMakerFormLoadShippableVersionInventory`
- `BoxMakerFormLoadBoxVersionInventory`
- `ShipmentsFormLoadNasReservationTotals`
- `BuildActiveShippingReservationTotalsFromTable`
- form reload path in `frmShipmentsTally.RefreshProjectedShippableInventory`

The key risk is that one path correctly deducts local inventory, but a later reload path reads stale version-log or stale NAS snapshot data and writes the visible quantities back upward.

**Test Design Requirement**
The new tests should compare visible form data before and after the action, not only raw helper outputs.

At minimum, capture before/after rows shaped like:

```text
Box, Version, NAS Inv, Projected Inv, Locked
```

Then assert monotonic non-increase for shipped rows:

```text
NAS Inv after <= NAS Inv before
Projected Inv after <= Projected Inv before
Locked after = 0 when no active shipment row/reservation remains
```

These tests need to fail loudly if the UI ever shows inventory increasing from a shipment action again.

**Current Status**
Existing tests are not sufficient. They passed while the user-facing form still showed a dirty transaction. The test suite must be redeveloped around user-visible workflow state so these inventory disasters are caught before user testing.
