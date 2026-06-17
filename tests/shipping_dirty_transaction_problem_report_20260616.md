# Shipping Dirty Transaction / Resurrected Shipment Rows Problem Report

Date: 2026-06-16

## Summary

The new `frmShipmentsTally` UserForm can add shippable box orders, move them `To Shipments`, and run `Shipments Sent`. Basic paths sometimes appear to work, but the system is not transactionally consistent. Orders that were already deducted from inventory later reappear in the `Shipments` listbox, and inventory quantities drift away from expected values.

The user calls this a dirty transaction. That is accurate: the same logical shipment currently has multiple unsynchronized representations:

- UserForm `Shipments` listbox / hidden `ShipmentsTally` table.
- Local persistent active shipment cache file.
- Local persistent sent-row tombstone file.
- Local writable `InventoryManagement!invSys` staging columns.
- Backend queued `SHIP` event / local staging JSONL.
- Version inventory overlay derived from local staging and/or inventory log.
- Backend read model after processor sync.

The system is allowing these stores to disagree, then later treating a stale store as authoritative.

## Current User-Observed Failure

Latest test sequence:

```text
Shipments sent: 2. package(s).
Boxes sent:
- 2. T24 v1
Inbox EventID: FA932671-9493-4209-9DE5-BCA1DF75468B

Completed in 2,121 ms.
```

Then:

```text
Moved 2. package(s) to Shipments.

Completed in 3,953 ms.
```

Then:

```text
Shipments sent: 2. package(s).
Boxes sent:
- 2. T25 v1
Inbox EventID: B1F54946-2D97-48E5-AEBD-E08BEEE606EE

Completed in 2,855 ms.
```

After that, the form showed inventory values that looked partially right:

- `T24 v1` current inventory showed `5`.
- `T25 v1` current inventory showed `7`.
- `T25 v2` current inventory showed `10`.
- `T26 v1` current inventory showed `10`.

But the `Shipments` listbox still contained a staged row:

```text
Ref 6486 | T25 | Qty 2 | Area Shipments | ROW 87 | Version v1 | Carrier USPS
```

Clicking `Shipments Sent` failed:

```text
ROW 87 only has 0. staged but needs 2..

Completed in 2,203 ms.
```

The user then reported a still worse symptom:

> still doesn't work, dirty transaction, a order i deducted from inventory has come back into the Shipments listbox.

This means a row that had already been processed by `Shipments Sent` was later reloaded into active Shipments.

## Expected Behavior

The shipment workflow needs exactly one coherent state transition:

1. Order starts in active `Shipments` list with `Area = Warehouse`.
2. User clicks `To Shipments`.
3. Row remains in active `Shipments` list, but `Area = Shipments`.
4. Local staged quantity is available for `Shipments Sent` validation.
5. User clicks `Shipments Sent`.
6. Inventory is deducted / event is queued.
7. The row is removed from active `Shipments`.
8. The row must never reappear unless the user explicitly creates a new order.
9. Inventory display must continue to show the deducted quantity across Excel reopen, unsaved `Book1`, and backend processor delay.

## Actual Behavior

Observed across multiple rounds:

- `Shipments Sent` reports success and queues an Inbox event.
- The row is sometimes deleted from the list for the current form session.
- After closing/reopening Excel, a previously sent row can reappear in the `Shipments` listbox.
- Inventory may show values that imply deduction happened, while the row still exists as active/staged.
- `Shipments Sent` may then fail because the form row says `Area = Shipments`, but local `invSys.SHIPMENTS` is `0`.
- The performance is also above target, typically `2,000-6,000 ms`, but the correctness bug is the blocker.

## Relevant Files

- `src/Shipping/Forms/frmShipmentsTally.frm`
- `src/Shipping/Modules/modTS_Shipments.bas`
- `src/Core/Modules/modRoleEventWriter.bas`
- `src/Core/Modules/modCarrierSettings.bas`
- `src/Admin/Forms/frmAdminSettings.frm`
- `src/Admin/Modules/modAdmin.bas`

## Recent Changes That Matter

### Active and Hold Local Cache

To support unsaved `Book1` workflows, active shipment rows and hold rows were persisted locally:

- Active rows: `%LOCALAPPDATA%\invSys\shipping_active_<warehouse>.tsv`
- Hold rows: `%LOCALAPPDATA%\invSys\shipping_hold_<warehouse>.tsv`

The active cache is loaded in:

```vb
Public Function ShipmentsFormLoadLines(Optional ByVal holdRows As Boolean = False) As Variant
    ...
    If holdRows Then
        LoadPersistentHoldRowsLocal lo
    Else
        LoadPersistentActiveShipmentRowsLocal lo
    End If
    ...
End Function
```

This is likely part of the problem: loading active persistent rows clears/replaces the workbook table before the system has established whether those cached rows have already been sent.

### Sent-Row Tombstone Attempt

A sent-row tombstone file was added:

- `%LOCALAPPDATA%\invSys\shipping_sent_<warehouse>.tsv`

On send, `ShipmentsFormRunShipmentsSentRows` now does:

```vb
SyncSingleVersionInventoryOverlayFromInvSysRows invLo, loShip, rowIndexes
AppendSentShipmentRowsLocal loShip, rowIndexes
DeleteShipmentRows loShip, rowIndexes
InvalidateAggregates True
PersistActiveShipmentRowsLocal loShip
ClearInstructionStaging ws
```

Persistent active rows are filtered during load:

```vb
If StrComp(defaultArea, "Warehouse", vbTextCompare) = 0 Then
    If PersistentSentShipmentRowExists(parts) Then GoTo NextPersistedLine
End If
```

Problems with this approach:

- The filter only runs when `defaultArea = "Warehouse"`, even though cached rows can contain `Area = Shipments`.
- The tombstone key uses row details and quantity; if quantity, carrier, version, normalization, or table content differs, stale rows may not match.
- A sent row can still be reintroduced if the active cache write after deletion fails silently.
- There is no transaction ID shared by active row, send event, tombstone, and inventory overlay.
- There is no atomic "mark sent + remove active + persist cache + queue event" boundary.

### SHIP Event Version Overlay Attempt

`SHIP` events were added to local staging and version inventory overlay.

In `modRoleEventWriter`, `GetLocalStagedBoxVersionInventoryDeltas` now includes `SHIP`:

```vb
If eventType = ROLE_EVENT_TYPE_SHIP _
   Or eventType = ROLE_EVENT_TYPE_BOX_BUILD _
   Or eventType = ROLE_EVENT_TYPE_BOX_UNBOX Then
    payloadJson = CStr(rowDict("PayloadJson"))
    AccumulateBoxPayloadVersionInventoryDeltasRole payloadJson, eventType, packageRow, result
End If
```

And `AccumulateBoxPayloadVersionInventoryDeltasRole` now treats `SHIP` as negative inventory:

```vb
If eventType = ROLE_EVENT_TYPE_SHIP Then
    deltaValue = -qtyValue
ElseIf eventType = ROLE_EVENT_TYPE_BOX_UNBOX Or ioType = "UNMADE" Then
    deltaValue = -qtyValue
Else
    deltaValue = qtyValue
End If
```

This helped displayed version inventory in some cases, but did not solve active row resurrection.

### SHIPMENTS Staging Rebuild Attempt

When a persisted row has `Area = Shipments`, `invSys.SHIPMENTS` must be rebuilt or `Shipments Sent` validation fails.

Added reconciliation:

```vb
Private Sub ReconcileShipmentStagingFromShipmentLines(ByVal invLo As ListObject, ByVal loShip As ListObject)
    ...
    For r = 1 To invLo.ListRows.Count
        invLo.DataBodyRange.Cells(r, cInvShip).Value = 0
    Next r

    For r = 1 To loShip.ListRows.Count
        If cShipArea > 0 Then
            If StrComp(NormalizeShipmentArea(...), "Shipments", vbTextCompare) <> 0 Then GoTo NextLine
        End If
        ...
        stagedByRow(CStr(rowVal)) = stagedByRow(CStr(rowVal)) + qtyVal
    Next r

    For Each key In stagedByRow.Keys
        invIdx = FindInvRowIndexByRow(invLo, CLng(key))
        If invIdx > 0 Then invLo.DataBodyRange.Cells(invIdx, cInvShip).Value = stagedByRow(CStr(key))
    Next key
End Sub
```

This may make validation pass, but it can also make a stale resurrected active row become valid again, causing a duplicate deduction risk.

## Why This Is Dangerous

The current design can double-ship/double-deduct:

1. User sends row.
2. Inventory/event layer records deduction.
3. Stale active row reappears.
4. Reconciliation rebuilds `invSys.SHIPMENTS` from that stale row.
5. `Shipments Sent` can become valid again.
6. User may send the same order a second time.

This is worse than a display bug. It is a data integrity bug.

## Suspected Root Cause

The active shipment row cache is being used as both:

- A UI draft/order persistence mechanism.
- A source of truth for operational state.

But sent shipment completion is represented elsewhere:

- Tombstone file.
- Queued `SHIP` event.
- Local staging JSONL.
- Inventory log after processor run.
- Version inventory overlay.

There is no durable per-row transaction ID that ties all these together. The code currently tries to match sent rows by values:

```vb
ref | item | qty | row | uom | location | version
```

That is not reliable enough for inventory transactions. It also does not distinguish:

- A legitimately new order with the same ref/box/qty.
- A stale cached old order.
- A partially sent order.
- A row that was moved to Hold and later returned.
- A row whose carrier/version/normalization changed.

## Recommended Design Fix

### 1. Add a Stable ShipmentLineId

Every row created in the Shipping form should get a GUID-like `ShipmentLineId`.

Columns needed in active and hold tables/cache:

- `LINE_ID`
- `REF_NUMBER`
- `ITEMS`
- `QUANTITY`
- `ROW`
- `UOM`
- `LOCATION`
- `DESCRIPTION` / version
- `AREA`
- `CARRIER`
- `STATUS`

Suggested statuses:

- `Warehouse`
- `Shipments`
- `Hold`
- `Sent`
- `Removed`

Do not infer state solely from table membership or `Area` text.

### 2. Replace Tombstone Matching by Value With Tombstone Matching by LineId

The sent tombstone should contain:

- `LINE_ID`
- `SentAtUTC`
- `EventID`
- `Ref`
- `Box`
- `Version`
- `Qty`

On active cache load:

```text
If LINE_ID exists in sent tombstones, do not load it into active Shipments.
```

No fuzzy matching.

### 3. Make `Shipments Sent` Transactionally Ordered

Preferred local transaction order:

1. Validate selected line IDs.
2. Build event payload.
3. Queue `SHIP` event and get EventID.
4. Mark selected line IDs `Sent` in local state ledger.
5. Remove selected line IDs from active rows.
6. Persist local state.
7. Refresh UI from local state.

If any step fails after event queueing, the local state should still mark the line as `SentPendingSync`, not restore it as active.

### 4. Do Not Rebuild `SHIPMENTS` From Stale Active Rows Without Status/LineId Guard

`ReconcileShipmentStagingFromShipmentLines` is useful only if rows are known to be currently active and staged. It should ignore:

- `Sent`
- `Removed`
- tombstoned line IDs
- rows without `LINE_ID` after migration unless explicitly treated as unsent legacy drafts

### 5. Move From Multiple Flat Files to One Local Shipping State Ledger

Current files:

- `shipping_active_<warehouse>.tsv`
- `shipping_hold_<warehouse>.tsv`
- `shipping_sent_<warehouse>.tsv`

Recommended:

- `shipping_state_<warehouse>.jsonl` or `.tsv`

One row per line item with explicit status. This prevents the same row from living in both active and sent stores.

### 6. Make UI Lists Projections Only

The listboxes should be filtered projections:

- Active `Shipments` listbox: `Status in ("Warehouse", "Shipments")`
- Not Shipped listbox: `Status = "Hold"`
- Hidden worksheet tables: UI/cache projection only, not source of truth

The source of truth for unsaved `Book1` user workflows should be the local state ledger until backend sync completes.

## Minimum Short-Term Patch If Full Redesign Is Too Large

If a smaller patch is needed first:

1. Add `LINE_ID` to active/hold persisted rows.
2. Generate a `LINE_ID` on add if missing.
3. Persist `LINE_ID` in all local files.
4. Tombstone by `LINE_ID`.
5. On load, never load a line ID that appears in sent tombstones.
6. On `Shipments Sent`, write tombstone before queue or immediately after queue, but before any possible UI refresh.
7. If `LINE_ID` is missing for legacy rows, assign one once and rewrite active cache before allowing send.

This should stop resurrection even before the full state-ledger redesign.

## Performance Note

Performance remains above target:

- `Shipments Sent`: roughly `2,100-2,900 ms`
- `To Shipments`: roughly `3,900-6,300 ms`

The likely performance sinks are:

- `GetWritableShippingInvSysTable`
- `ReconcileShippableTotalsFromVersionInventory`
- repeated list/table hydration
- local staging scans
- workbook/table surface checks

Do not optimize this until correctness is fixed. The current correctness bug can duplicate inventory deductions.

## Validation Run So Far

After latest changes, these automated validations pass:

```text
tools/build-xlam.ps1 -Apply
tools/run_phase6_excel_validation.ps1 -StartAt 109

PHASE6_VALIDATION_OK
PASSED=21 FAILED=0 TOTAL=21
```

These tests do not currently cover the dirty transaction case:

- unsaved `Book1`
- active shipment persisted across Excel close/reopen
- selected row sent
- row removed from active cache
- Excel closed/reopened again
- row must not reappear
- version inventory must remain deducted
- duplicate send must be impossible

## Proposed Regression Test

Add an automated/manual test named something like:

```text
TestShippingForm_SentRowsDoNotResurrectAcrossUnsavedWorkbookReopen
```

Steps:

1. Open a fresh unsaved workbook.
2. Open Shipping form.
3. Add order:
   - Ref: `TXN-001`
   - Box: `T25`
   - Version: `v1`
   - Qty: `2`
4. Move it `To Shipments`.
5. Verify active list has row with `Area=Shipments`.
6. Click `Shipments Sent`.
7. Verify active list no longer has `TXN-001`.
8. Close form and workbook without saving.
9. Fully close Excel.
10. Reopen Excel and Shipping form.
11. Verify `TXN-001` is not in active Shipments.
12. Verify `TXN-001` is not in Not Shipped.
13. Verify T25 v1 inventory is reduced by 2.
14. Verify no active cache file can reload the sent row.

Expected result:

```text
PASS: sent row does not resurrect, inventory remains deducted, duplicate send impossible.
```

## Key Ask for Help

Please review the Shipping state model and propose/implement a transaction-safe local state mechanism. The current patch series is trying to reconcile multiple local/projection stores after the fact, and that is producing dirty transactions. The likely correct fix is to introduce a stable `ShipmentLineId` and a single local state ledger with explicit row statuses, then treat the worksheet/listbox tables as projections only.
