# Box Maker Shippable Version Inventory Sync Problem Report

Date: 2026-06-30

## Executive summary

Box Maker now calculates component requirements correctly and the component inventory rows are syncing back from the server/NAS read model. The remaining defect is narrower: after clicking `Make Boxes`, the finished shippable box/version row in the Box Picker stays on the local projected overlay and does not resolve to a visible NAS quantity.

The same autosync/overlay pattern is working in the Shipments form. Shipments can post local work, refresh the server read model, clear completed overlays, and move the sync gauge back to complete. Box Maker was recently wired to use the same refresh path, but its finished-box version inventory still remains pending.

Current live symptom:

- Before `Make Boxes`, T31 v1 with Qty 10 shows shippable inventory as `NAS Inv: unknown`, `Projected Inv: unknown`.
- Component rows show correct math: `Projected Inv = NAS Inv - Required`.
- More than one minute after `Make Boxes`, component rows have caught up: their `NAS Inv` equals the prior projected values and `Required` is back to zero.
- The finished shippable row still shows T31 v1 as `NAS Inv: unknown`, `Projected Inv: 10`.
- The Box Maker sync gauge remains red: `Sync: pending (1 inventory row(s))`.

This indicates the runtime queue and component inventory refresh are functioning. The unresolved row is specifically the finished-good box/version inventory overlay.

## Recent relevant commits

- `e3896cc` - Fix Box Maker projected inventory math
- `79fe507` - Fix Box Maker NAS inventory auto update
- `f94f53b` - Add Box Maker inventory auto sync
- `f0854a2` - Refresh operator inventory log during shipping autosync
- `b4d7a9f` - Fix Box Maker version inventory display refresh

These commits repaired the component-side projected math and added an autosync path for Box Maker. They did not fully solve finished-box version inventory catch-up.

## What works

### Components To Deduct

Before `Make Boxes`:

- `Required` equals `Qty * Per Box`.
- `Projected Inv` equals `NAS Inv - Required`.

After `Make Boxes` and after autosync has had time to run:

- `Required` resets to zero.
- Component `NAS Inv` catches up to the prior projected value.
- Component projected display matches NAS.

This proves the following are at least partly working:

- Box Maker can enqueue/process the make-box event.
- Component inventory deltas reach the server/NAS-owned inventory surface.
- The operator workbook/read-model refresh can bring component NAS quantities back into the form.
- Component pending overlays can clear.

### Shipments form

The equivalent local-overlay pattern is working in the Shipments form:

- Local actions create a projected overlay.
- Autosync calls `ShipmentsFormAutoSyncRefresh`.
- The form reloads shippables from the read model.
- Completed overlays are evicted.
- The sync gauge returns to complete when NAS and projected values match.

Important procedures:

- `frmShipmentsTally.AutoSyncIfPending`
- `frmShipmentsTally.PendingShipmentSyncCount`
- `modTS_Shipments.ShipmentsFormAutoSyncRefresh`
- `modTS_Shipments.EvictCompletedShipmentInventoryOverlaysForShippables`

## What does not work

### Box Maker finished shippable version inventory

After `Make Boxes`, the selected box/version remains in a pending projected state:

- Box Picker row: T31 v1, row 92
- `NAS Inv`: `unknown`
- `Projected Inv`: `10`
- Sync gauge: `pending (1 inventory row(s))`

The pending count strongly suggests this is the single remaining entry in `mPendingVersionInv`, not a component reservation.

Important procedures:

- `frmShippingBoxMaker.PostBoxMakerAction`
- `frmShippingBoxMaker.RecordPendingVersionInventory`
- `frmShippingBoxMaker.AutoSyncIfPending`
- `frmShippingBoxMaker.PendingShippableInventoryCount`
- `frmShippingBoxMaker.DisplayBoxVersionInventoryText`
- `modTS_Shipments.BoxMakerFormLoadShippableVersionInventory`
- `modTS_Shipments.BuildBoxVersionInventoryCache`
- `modTS_Shipments.PendingBoxVersionInventoryOverlayValue`
- `modTS_Shipments.RefreshOperatorInventoryLogForWorkbook`

## Current data flow for Box Maker

1. User selects saved box T31 v1 and enters Qty 10.
2. Box Maker renders component requirements in memory.
3. User clicks `Make Boxes`.
4. `PostBoxMakerAction` calls `modTS_Shipments.CommitBoxMakerFormAction`.
5. The payload includes component deductions and a finished-good package line.
6. Box Maker records local pending overlays:
   - `RecordPendingComponentInventory` for component rows.
   - `RecordPendingVersionInventory` for the finished box/version.
7. Box Maker arms autosync.
8. `AutoSyncIfPending` calls `ShipmentsFormAutoSyncRefresh`.
9. `ShipmentsFormAutoSyncRefresh` runs the shipping runtime queue, refreshes the read model, refreshes the BOM view when needed, and attempts to refresh the operator workbook's inventory log.
10. Box Maker reloads shippable inventory through `BoxMakerFormLoadShippableVersionInventory`.
11. Box Maker rerenders package and shippable inventory.

The failure appears between steps 9 and 11 for the finished-good version quantity only.

## Latest popup evidence

The latest live test produced this Excel status message after clicking `Make Boxes` for T31 v1 Qty 10:

```text
Box build event queued for 10 T31 v1. Uses 120 component units and adds 10 shippable units after processor sync.

Sync complete.
Processed=1; StagingReport=LocalStagingMerged=1; LocalStagingFailed=0; BatchReport=Applied=1; SkipDup=0; Poison=0; RunId=RUN-invsys_Zenbook_WH-INVENTORY-20260630144232-701883; SnapshotError=Snapshot workbook not resolved.; PublishWarning=Snapshot workbook not resolved.; TimingMs=Total:9043;Batch:8285;Refresh:0
Inbox EventID: 0FA5A076-07BE-46DD-AC62-77B230CB0650

Completed in 9,676 ms.
```

This is important because it narrows the defect:

- The event was not stuck in staging: `LocalStagingMerged=1`.
- The batch processor applied it: `Processed=1`, `Applied=1`, `Poison=0`.
- The UI-visible problem persists after an applied event.
- The status reports `SnapshotError=Snapshot workbook not resolved.` and `PublishWarning=Snapshot workbook not resolved.`
- `Refresh=0`, so no separate refresh time is visible in this status payload.

The snapshot failure may be central. It can explain why Box Maker keeps displaying unknown NAS quantity for the finished box/version even though the processor applied the event. The expert should verify whether the component rows are updating from a different surface than the Box Picker version row, or whether the Box Picker version row specifically requires the snapshot/read-model workbook that failed to resolve.

## Why this likely differs from Shipments

Shipments shippable inventory is row-oriented. Its pending state compares visible NAS inventory against visible projected inventory for each shippable row.

Box Maker finished-good inventory is version-oriented. When a box has multiple active versions, `BoxMakerFormLoadShippableVersionInventory` intentionally avoids treating row-level `TOTAL INV` as the version quantity. Instead it depends on `BuildBoxVersionInventoryCache`, which reconstructs version quantities from:

- local staged box-version inventory deltas, and
- operator workbook `tblInventoryLog` rows containing matching SKU and version evidence.

That distinction matters for T31 because both v1 and v2 are active. A row-level `TOTAL INV` value on invSys row 92 is not enough to tell the form how much belongs to v1 versus v2. The Box Maker display needs version-specific evidence from the inventory log or staged version deltas.

## Leading hypotheses

### 1. Finished-good BOX_BUILD evidence is not present in the operator inventory log

`BuildBoxVersionInventoryCache` reads `tblInventoryLog` from the operator workbook. If `RefreshOperatorInventoryLogForWorkbook` does not copy the canonical server log into the operator workbook in the live environment, Box Maker will not see applied version evidence.

Things to inspect:

- Does operator `tblInventoryLog` contain the T31 row 92 BOX_BUILD after autosync?
- Does the server/NAS inventory workbook `tblInventoryLog` contain it?
- Does `RefreshOperatorInventoryLogForWorkbook` report `OK`, or does it silently skip because the inventory workbook or target table is not resolved?
- How does `SnapshotError=Snapshot workbook not resolved.` relate to `RefreshOperatorInventoryLogForWorkbook` and to the version inventory cache?

### 2. SKU identity mismatch prevents log rows from matching T31

`BuildBoxVersionInventoryCache` builds candidate keys from saved box name, invSys `ITEM_CODE`, and invSys `ITEM`. Then it only reads log rows whose `SKU` matches one of those candidates.

Earlier exported history showed a suspicious identity issue where rows for T28 appeared under `ServerSKU` value `T27`. If similar SKU drift exists for T31, the log may contain the right quantity evidence but under a SKU value not considered by the cache.

Things to inspect:

- T31 row 92 in operator `invSys`: `ITEM`, `ITEM_CODE`, `ROW`, `TOTAL INV`.
- The corresponding server inventory log row: `SKU`, `QtyDelta`, `EventType`, `Note`.
- Whether the log `SKU` equals `T31`, the row's `ITEM_CODE`, or some stale/incorrect value.

### 3. Version token is missing or not parsed

The version cache only counts log rows when `ExtractBoxVersionLabelFromNoteShipping` can extract a version label from `Note`.

Things to inspect:

- Does the BOX_BUILD log note contain an exact version token such as `VERSION=v1`?
- Is the token casing or field name different?
- Is the version stored in another column instead of `Note`?

### 4. The applied event updates row TOTAL INV but not version evidence

For multiple active versions, Box Maker cannot rely on row-level `TOTAL INV`, because row 92 may represent all T31 versions combined. It needs a per-version ledger total.

If the server applies the BOX_BUILD to row-level inventory but does not emit or preserve version-specific log evidence, the UI can be correct to keep NAS version quantity as `unknown`.

Things to inspect:

- Did server/NAS row 92 `TOTAL INV` increase by 10?
- Did a version-specific inventory log row exist for the same event?
- Does the version ledger include both row 92 and `VERSION=v1`?

### 5. Persistent overlay clearing is too strict for a version row with unknown backend

`DisplayBoxVersionInventoryText` only clears `mPendingVersionInv` when the backend text is numeric and equals the pending quantity. If `BoxMakerFormLoadShippableVersionInventory` keeps returning backend text blank/unknown for row 92 v1, the overlay cannot clear.

This is probably a symptom rather than the root cause. The real question is why the version read model is still unknown after the server has processed the make event.

## Recommended diagnostics for the second expert

### Inspect live workbook/server evidence after one Make Boxes action

Use a fresh T31 v1 make action with a small quantity and capture these values.

Operator workbook:

- `InventoryManagement!invSys` row 92
- `ITEM`
- `ITEM_CODE`
- `ROW`
- `TOTAL INV`
- `tblInventoryLog` rows for row 92, T31, and recent `BOX_BUILD`

Server/NAS inventory workbook:

- `tblInventoryLog` rows for the same event
- `SKU`
- `QtyDelta`
- `EventType`
- `Note`
- any EventID/RunId that can tie the log entry to the queued event

Local staging/runtime state:

- pending queue item for the make action, if still present
- local staged box-version inventory deltas for package row 92
- persistent pending overlay for key `92|v1`

Do not include credentials in diagnostics output. The relevant NAS share path is known to the project, but the report should avoid storing passwords.

### Add temporary debug output

Add one temporary report procedure or guarded debug print that can be called after autosync:

```text
BoxMakerDebugSelectedInventoryReport(packageRow:=92, version:="v1")
```

It should print:

- selected package row and selected version
- Box Picker rowData(4) NAS value
- Box Picker rowData(8) projected/overlay value
- `mPendingVersionInv("92|v1")`
- persistent overlay value/baseline for `92|v1`
- staged version deltas from `modRoleEventWriter.GetLocalStagedBoxVersionInventoryDeltas(92)`
- operator `tblInventoryLog` matches by SKU candidates
- server `tblInventoryLog` matches by SKU candidates
- extracted version labels from those matched log rows

### Make autosync log refresh visible

`ShipmentsFormAutoSyncRefresh` includes `logReport` only when it is non-OK. For this issue, a more detailed temporary report would help:

- inventory workbook path resolved
- whether it was opened transiently
- source `tblInventoryLog` row count
- target `tblInventoryLog` row count before/after copy
- number of copied log rows
- recent BOX_BUILD row count

This will separate "server evidence missing" from "operator workbook did not receive the evidence" from "Box Maker cache did not parse the evidence."

### Add a focused regression test

Create a test with a saved box that has two active versions, because that forces the version-ledger path:

1. Create or fixture package row 92 / T31 with v1 and v2 active.
2. Queue/apply `BOX_BUILD` Qty 10 for T31 v1.
3. Refresh the operator read model and inventory log.
4. Call `BoxMakerFormLoadShippableVersionInventory`.
5. Assert:
   - row 92 v1 NAS inventory is `10`
   - row 92 v1 projected inventory is `10`
   - row 92 v2 remains unchanged/unknown unless it has its own ledger evidence
   - pending Box Maker version overlay count is zero after render

This test should fail on the current live behavior and pass once the version evidence path is repaired.

## Expected correct behavior

After `Make Boxes` for T31 v1 Qty 10:

- Component rows should deduct from NAS inventory after sync. This already appears to work.
- Box Picker T31 v1 should show `NAS Inv: 10` and `Projected Inv: 10` once the server/NAS read model has applied the event.
- The top selected shippable summary should show `NAS Inv: 10`, `Projected Inv: 10`.
- Sync gauge should return to green/complete.
- T31 v2 should not be inflated by the v1 build.

## Key caution

Do not "fix" this by treating row-level `TOTAL INV` as the NAS quantity for every active version. That would hide the bug but would make multi-version box inventory mathematically wrong. If multiple active versions exist, Box Maker needs a version-specific read model or version-specific ledger evidence.

The correct boundary is:

- server/NAS row inventory remains authoritative for the whole shippable row,
- Box Maker version inventory must be derived from explicit version evidence,
- local projected overlays are temporary and must clear only when the server/read model reflects the same version-specific quantity.
