# BoxMaker UserForm: Make Boxes Performance and Runtime Lock Problem

Date: 2026-06-13

## Summary

The new Shipping BoxMaker UserForm can load saved box designs, display component requirements/current inventory, and queue a `BOX_BUILD` event. The user tested:

1. Open `Box Maker`.
2. Select a saved box.
3. Set `Qty Made = 5`.
4. Click `Make Boxes`.

The form reported a successful logical post, but the operation was slow and the runtime batch processor did not handle the queued event immediately because the inventory runtime workbook was locked/read-only.

Observed message:

```text
Box created. Used 30. component units; added 5. shippable units to TOTAL INV.

RunBatch did not handle the queued event after local post/write. BatchReport=Inventory workbook is read-only or locked by another Excel session. RefreshReport=OK
Inbox EventID: 50EB093B-4BBC-4B1A-9413-EE223E942B9Ak
```

The important distinction: the event was queued, but processing did not apply it to the runtime inventory workbook during the same click.

## Current UserForm Flow

Relevant files:

- `src/Shipping/Forms/frmShippingBoxMaker.frm`
- `src/Shipping/Modules/modTS_Shipments.bas`
- `src/Core/Modules/modOperatorReadModel.bas`
- `src/Core/Modules/modProcessor.bas`

Current call chain:

```text
frmShippingBoxMaker.mBtnMake_Click
  -> PostBoxMakerAction "MAKE"
  -> modTS_Shipments.CommitBoxMakerFormAction(...)
      -> writes selected design into BoxBuilder/BoxBOM projection tables
      -> ApplyBoxCreatedFromBuilder(...)
          -> QueueBoxBuildEventFromBuilder(...)
              -> modRoleEventWriter.QueuePayloadEventCurrent(EVENT_TYPE_BOX_BUILD, ...)
          -> modOperatorReadModel.RunBatchAndRefreshOperatorWorkbook(...)
              -> modProcessor.RunBatch(...)
              -> EnsureInventoryManagementSurface(...)
              -> RefreshInventoryReadModelForWorkbook(...)
```

Current implementation intentionally reuses the existing `BOX_BUILD` event path rather than mutating inventory directly from the form.

## What Is Working

- BoxMaker form opens from its own ribbon button.
- Saved box dropdown loads active saved ShippingBOM designs.
- Version dropdown loads active versions.
- Component list shows headers and current inventory.
- `Refresh` now preserves list state if a reload returns no designs.
- `Make Boxes` now bypasses the earlier local `invSys has no inventory rows` blocker.
- Event queueing succeeds enough to return an `Inbox EventID`.
- Read-model refresh can return `RefreshReport=OK` even when batch processing fails.

## Primary Failure

`RunBatchAndRefreshOperatorWorkbook` calls `modProcessor.RunBatch`, but the processor reports:

```text
Inventory workbook is read-only or locked by another Excel session.
```

Then `RunBatchAndRefreshOperatorWorkbook` returns:

```text
RunBatch did not handle the queued event after local post/write.
```

This means the queued `BOX_BUILD` event remains pending or unprocessed because the runtime inventory workbook could not be opened for write.

Likely target workbook:

```text
<RuntimeRoot>\<WarehouseId>.invSys.Data.Inventory.xlsb
```

## Performance Problem

The user reports `Make Boxes` is slow. Even when the event is eventually queued, the click currently does too much synchronous work:

1. Writes the form state back into sheet projection tables (`BoxBuilder`, `BoxBOM`).
2. Queues the `BOX_BUILD` event.
3. Immediately runs the batch processor.
4. Ensures/repairs the `InventoryManagement` surface.
5. Refreshes local read models/projections.
6. Returns a long status message.

The slowest likely parts are:

- `modProcessor.RunBatch(...)` trying to acquire/write the runtime inventory workbook.
- Workbook open/save/lock retry behavior inside runtime processing.
- `modRoleWorkbookSurfaces.EnsureInventoryManagementSurface(...)`.
- `RefreshInventoryReadModelForWorkbook(...)`.
- Any hidden Excel workbook open/close operations against runtime files.

The current UX is synchronous: the form waits for all of this inside the button click.

## Recent Changes Already Made

Recent BoxMaker changes:

- Added `frmShippingBoxMaker.frm`.
- Added `modTS_Shipments.BtnOpenBoxMaker`.
- Added `BoxMakerFormLoadSavedBoxes`, `BoxMakerFormLoadVersions`, `BoxMakerFormLoadVersionComponents`.
- Added runtime ShippingBOM fallback when `ShippingBOMView` is empty.
- Added `CommitBoxMakerFormAction`.
- Removed the form path's dependency on a populated local `invSys` table.
- Removed full `SetBoxMakerMode ws, True` from the form commit path.
- Removed one pre-post `RefreshBoxMakerCurrentInventory ws` call before event posting.

Validation after these changes:

```text
tools/build-xlam.ps1 -Apply                         PASS
tools/validate_phase6_packaged_ribbon.ps1           227 passed, 0 failed
tools/run_phase6_excel_validation.ps1 -StartAt 109  20 passed, 0 failed
```

## Reproduction Notes

Manual repro from user:

```text
Open new Excel workbook
Open Shipping Box Maker form
Select saved box
Set Qty Made = 5
Click Make Boxes
```

Observed output:

```text
Box created. Used 30. component units; added 5. shippable units to TOTAL INV.

RunBatch did not handle the queued event after local post/write.
BatchReport=Inventory workbook is read-only or locked by another Excel session.
RefreshReport=OK
Inbox EventID: ...
```

Prior related behavior:

- Before the latest fix, `Make Boxes` stopped with `invSys has no inventory rows`.
- That guard was removed for the form-backed event path.
- Now the post reaches batch processing and exposes the runtime workbook lock.

## Hypotheses

### 1. Runtime inventory workbook is actually open/locked

An Excel session may still have `<WarehouseId>.invSys.Data.Inventory.xlsb` open, visible or hidden. Because Excel COM can leave hidden workbooks alive, a workbook can be locked even if the user does not see it.

Things to check:

- Open Excel instances in Task Manager.
- Hidden `*.invSys.Data.Inventory.xlsb` workbooks in any Excel session.
- Whether a previous add-in validation/process left Excel running.
- Whether OneDrive/AV/indexer/file sync is holding the workbook briefly.

### 2. Processor lock handling is too binary

The batch processor may treat transient locks as hard failures. If so, the UX should not imply the inventory update happened. Better behavior could be:

- Queue event quickly.
- If runtime is locked, show `Queued, pending processor lock`.
- Retry processing later or provide a `Process Pending` button.

### 3. BoxMaker form should not synchronously run batch processing

For a UserForm-first workflow, a faster model may be:

```text
Make Boxes click:
  validate form state
  queue BOX_BUILD event
  return immediately with EventID

Background or explicit follow-up:
  process pending events
  refresh read model/current inventory
```

This would avoid making the user wait on runtime locks and heavy workbook refresh work.

### 4. Message text is misleading

The current message starts with:

```text
Box created. Used 30. component units; added 5. shippable units to TOTAL INV.
```

But the batch report says inventory was not processed because the runtime inventory workbook was locked. The event is queued, not actually applied. The message should be changed to avoid claiming inventory was updated until processing confirms it.

Recommended wording when batch does not process:

```text
Box build event queued but not processed.
Runtime inventory workbook is locked/read-only.
EventID: ...
No inventory quantities were updated yet.
```

## Suggested Investigation Tasks

1. Add lock diagnostics around the runtime inventory workbook open/write path.
   - Log full workbook path.
   - Log whether an open workbook with the same `FullName` exists.
   - Log `ReadOnly` status.
   - Log owner/lock file if detectable.

2. Inspect `modProcessor.RunBatch` workbook open mode.
   - Find the exact branch producing `Inventory workbook is read-only or locked by another Excel session.`
   - Determine whether it retries, and how many times.
   - Determine whether a hidden workbook in the same Excel instance is being treated as read-only.

3. Separate "queued" from "processed" result in `ApplyBoxCreatedFromBuilder`.
   - Return structured status or separate booleans:
     - `eventQueued`
     - `batchProcessed`
     - `readModelRefreshed`
   - Avoid success language when only queueing succeeded.

4. Consider an async/deferred BoxMaker UX.
   - `Make Boxes` queues event and returns fast.
   - `Refresh` or a separate button processes pending queue and refreshes inventory.
   - This aligns with the phasing-out of ribbon-to-table manipulation and reduces form latency.

5. Reduce projection table writes.
   - Current form still writes `BoxBuilder` and `BoxBOM` projection tables before calling the shared queue builder.
   - Long term, create a direct `QueueBoxBuildEventFromFormState(...)` that builds payload directly from the form arrays.
   - This avoids sheet writes and makes the form independent of table mode.

## Relevant Functions

Shipping:

- `modTS_Shipments.BtnOpenBoxMaker`
- `modTS_Shipments.CommitBoxMakerFormAction`
- `modTS_Shipments.ApplyBoxCreatedFromBuilder`
- `modTS_Shipments.QueueBoxBuildEventFromBuilder`
- `modTS_Shipments.AddBoxBuildComponentPayloadItems`
- `modTS_Shipments.BoxMakerFormLoadSavedBoxes`
- `modTS_Shipments.BoxMakerFormLoadVersionComponents`

Core:

- `modOperatorReadModel.RunBatchAndRefreshOperatorWorkbook`
- `modProcessor.RunBatch`
- `modRoleEventWriter.QueuePayloadEventCurrent`

## Current Working Theory

The BoxMaker form is now reaching the correct event-backed path. The remaining problem is not that the form cannot build the payload; it is that immediate synchronous runtime processing is blocked by a locked/read-only runtime inventory workbook. This also explains the bad perceived performance: the click waits for a processor attempt and read-model refresh even when the processor cannot write.

The next design decision is whether BoxMaker should:

1. Keep synchronous processing and improve lock handling/retries, or
2. Queue quickly and process/refresh separately.

Given the UserForm direction, option 2 is likely the better UX.
