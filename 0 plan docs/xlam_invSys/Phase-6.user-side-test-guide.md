# Phase 6 User-Side Test Guide

## Purpose

Use this guide to run real Excel-side proving of the current invSys build without VBE work or test harnesses.

The proving order is:

1. one-account local use
2. multi-PC LAN use
3. LAN + WAN use
4. central aggregation

Do not skip ahead. Each stage depends on the earlier one being stable.

## What this guide is proving

- local warehouse processing stays authoritative
- operators can keep working when WAN or SharePoint is unavailable
- operator inventory refresh can become stale without becoming unusable
- published warehouse artifacts can catch up later without data loss
- HQ aggregation reads published snapshots and stays advisory only

## Roles in this guide

- `Host PC`: the warehouse PC/session that owns the local runtime root and runs processor work
- `Station PC`: another LAN PC/session for the same warehouse
- `Remote/WAN PC`: a PC using published snapshots over SharePoint or remote access
- `HQ/Admin`: the PC/session that runs global aggregation

## Files you will inspect

- local authoritative inventory:
  - `WHx.invSys.Data.Inventory.xlsb`
- local outbox:
  - `WHx.Outbox.Events.xlsb`
- local snapshot:
  - `WHx.invSys.Snapshot.Inventory.xlsb`
- local publish log:
  - `invSys.Publish.log`
- SharePoint-published snapshots:
  - `Snapshots\WHx.invSys.Snapshot.Inventory.xlsb`
- SharePoint-published outboxes:
  - `Events\WHx.Outbox.Events.xlsb`
- global snapshot:
  - `Global\invSys.Global.InventorySnapshot.xlsb`

## Before you start

1. Close all old Excel sessions on every test machine.
2. Confirm each machine is using the current add-ins from `deploy/current`.
3. Pick one or two throwaway SKUs for testing, not live production SKUs.
4. Write down:
   - warehouse id
   - station ids
   - machine names
   - local runtime root
   - SharePoint root
   - start time
5. Make sure SharePoint has these folders:
   - `Events`
   - `Snapshots`
   - `Global`
6. If using LAN inbox shares, confirm each `PathInboxRoot` is reachable from the host PC.

## Pass/fail rule

- `PASS`: local warehouse posting and processing complete correctly, even if WAN publish is delayed
- `PASS`: operator refresh can show stale state and still allow posting
- `PASS`: HQ snapshot shows per-warehouse quantities from published snapshots only
- `FAIL`: local posting is blocked because SharePoint is down
- `FAIL`: stale refresh clears local staging or workbook-local logs
- `FAIL`: HQ/global snapshot overrides or is treated as authoritative warehouse truth

## Stage 1: One-Account Local Use

### Goal

Prove the saved operator workbook works on one PC/account with local authoritative processing.

### Steps

1. Open a saved receiving workbook for one warehouse/station.
2. Enter one small receive transaction for the test SKU.
3. Click the normal post action such as `Confirm Writes`.
4. Let processor work run, or run the warehouse batch from the Admin XLAM if needed.
5. Refresh the operator workbook inventory view.

### Expected result

- the inbox row posts and becomes processed
- `WHx.invSys.Data.Inventory.xlsb` reflects the new quantity
- `WHx.Outbox.Events.xlsb` exists or updates
- `WHx.invSys.Snapshot.Inventory.xlsb` exists or updates
- the operator inventory view shows current data
- freshness metadata is populated:
  - `LastRefreshUTC`
  - `SnapshotId`
  - `SourceType`
  - `IsStale`

### Evidence to capture

- screenshot of the operator workbook before post
- screenshot after refresh
- file timestamp of local snapshot
- file timestamp of local outbox

## Stage 2: Multi-PC LAN Use

### Goal

Prove two stations in the same warehouse can operate without cross-contaminating workbooks or corrupting the single-writer processor model.

### Setup

- `Host PC`: same warehouse runtime root
- `Station PC`: same warehouse, different station id
- both stations point to the same warehouse authority
- processor remains one lane per warehouse

### Steps

1. On the Station PC, open a saved receiving workbook and post one receive transaction.
2. On the Host PC, run processor work if it is not automatic.
3. On both PCs, refresh the operator inventory view.
4. Repeat once with the host inventory workbook intentionally open on the Host PC while the second station tries to process.

### Expected result

- the station can queue work without corrupting the warehouse runtime
- the processor handles one inventory writer lane only
- both operator workbooks refresh to the same warehouse quantity after successful processing
- if a lock boundary is hit, the second process attempt is denied cleanly rather than corrupting the inventory workbook

### Evidence to capture

- screenshot from each operator workbook after refresh
- note whether the denied lock case was clean and recoverable
- timestamp of local snapshot after successful retry

## Stage 3: WAN Publish and Recovery

### Goal

Prove the warehouse can keep operating on its local authoritative files when WAN is unavailable, then later publish the missed artifacts to SharePoint without data loss.

### Scope note

This stage is about WAN behavior only.

- do not break the warehouse LAN/runtime root
- only break internet access or SharePoint reachability
- local posting and local processing must still succeed
- SharePoint is a publish target, not the warehouse authority

### Machines in this stage

- `Warehouse PC`: the machine that owns or runs the warehouse processor path
- `Remote/WAN observer`: optional second machine that checks SharePoint or opens a remote operator workbook

### Test 3A: WAN available, normal publish

1. On the `Warehouse PC`, confirm SharePoint is reachable.
2. Post one small test transaction.
3. Run processor work if it is not automatic.
4. Immediately verify the local files updated:
   - `WHx.invSys.Data.Inventory.xlsb`
   - `WHx.Outbox.Events.xlsb`
   - `WHx.invSys.Snapshot.Inventory.xlsb`
5. Then verify the SharePoint copies updated:
   - `Events\WHx.Outbox.Events.xlsb`
   - `Snapshots\WHx.invSys.Snapshot.Inventory.xlsb`
6. Check `invSys.Publish.log`.

Expected:

- local authority updates first
- SharePoint publish happens after local write completes
- no posting delay is introduced by publish
- publish log shows success, or at minimum no failure warning for this run

What would be a real failure:

- SharePoint files update but local inventory did not
- posting waits on SharePoint before local processing finishes
- local snapshot/outbox are missing even though the publish copy exists

### Test 3B: WAN unavailable, offline-first processing

1. Make SharePoint unreachable from the `Warehouse PC`.
   - disconnect internet, or
   - disable the SharePoint sync/client path, or
   - temporarily point the machine at a disconnected network state
2. Leave the local warehouse runtime root available.
3. Post one small test transaction.
4. Run processor work.
5. Verify local results first:
   - operator post completed
   - local inventory changed
   - local outbox changed
   - local snapshot changed
6. Verify SharePoint did not update.
7. Open `invSys.Publish.log`.

Expected:

- warehouse posting still works
- local inventory remains authoritative
- local outbox and snapshot are preserved
- SharePoint publish failure is reported, not hidden
- no rollback of local processing occurs

What would be a real failure:

- post is blocked because SharePoint is down
- local outbox/snapshot are skipped because publish failed
- the transaction disappears after the failed publish

### Test 3C: delayed publish catch-up

1. Restore SharePoint/internet access.
2. Rerun the warehouse publish path.
   - use the normal scheduled/manual warehouse publish path if that is how the site will operate
3. Recheck SharePoint:
   - `Events\WHx.Outbox.Events.xlsb`
   - `Snapshots\WHx.invSys.Snapshot.Inventory.xlsb`
4. Compare the published snapshot against the local warehouse state.
5. Review the publish log again.

Expected:

- the previously missed publish catches up
- rerun is safe and deterministic
- published artifacts now match the local authoritative warehouse state
- old failed publish does not create duplicate warehouse effects

What would be a real failure:

- rerun creates duplicate inventory effects locally
- SharePoint receives partial or obviously stale copies after rerun
- leftover `.uploading` or broken temp files accumulate and prevent recovery

### WAN evidence to capture

- screenshot of the operator workbook after offline posting
- timestamp of local snapshot during the offline run
- timestamp of SharePoint snapshot before catch-up and after catch-up
- publish log lines showing:
  - one failed publish while offline
  - one later successful publish after recovery

## Stage 4: WAN Stale-State on Operator Workbooks

### Goal

Prove a remote or WAN-refreshed operator workbook can become stale without becoming unusable.

### Test 4A: missing SharePoint snapshot

1. Open a saved operator workbook that refreshes inventory from published snapshots.
2. Make the expected published snapshot unavailable.
3. Refresh the operator inventory view.
4. Try a normal operator posting action.

Expected:

- refresh completes without destroying local workflow tables
- workbook-local logs remain intact
- visible metadata shows stale/cached state
- posting still works because warehouse posting is not gated on snapshot freshness

Specifically inspect:

- `LastRefreshUTC`
- `SnapshotId`
- `SourceType`
- `IsStale`

### Test 4B: stale SharePoint snapshot

1. Refresh the same operator workbook from an older known published snapshot.
2. Inspect the visible metadata.
3. Confirm local staging rows are still present.

Expected:

- `SourceType` shows `SHAREPOINT` when a published snapshot was actually used
- `SourceType` shows `CACHED` when no current published snapshot was available and cached data was retained
- `IsStale=True`
- `SnapshotId` identifies the artifact actually used
- `LastRefreshUTC` changes on refresh
- local workflow tables such as `ReceivedTally` remain intact

What would be a real failure:

- stale refresh clears staging rows
- stale refresh clears workbook-local logs
- operator cannot post simply because published inventory is stale

### Operator stale-state evidence to capture

- screenshot of stale metadata in the operator workbook
- screenshot showing local staging rows still present after refresh

## Stage 5: WAN Central Aggregation

### Goal

Prove HQ aggregation rebuilds the global snapshot from published warehouse snapshots and remains advisory only.

### Setup

- use two warehouses, for example `WH1` and `WH2`
- each warehouse must have already published snapshots to SharePoint
- do not read directly from local warehouse authority for this stage; use published artifacts only

### Test 5A: initial aggregation from published artifacts

1. Publish current snapshots from both warehouses.
2. From Admin/HQ, run:
   - `Scheduler_RunHQAggregation`
   - or the equivalent scheduled job path
3. Open `Global\invSys.Global.InventorySnapshot.xlsb`.
4. Inspect:
   - the warehouse rows
   - the quantities per warehouse
   - the status sheet/table

Expected:

- the global snapshot is rebuilt from published snapshots
- per-warehouse quantities are preserved separately
- the snapshot is useful for visibility only
- warehouse-local `WHx.invSys.Data.Inventory.xlsb` remains the authority

### Test 5B: staggered catch-up

1. Update and publish only one warehouse.
2. Leave the other warehouse unchanged.
3. Run HQ aggregation again.
4. Reopen the global snapshot.

Expected:

- only the republished warehouse quantity moves
- the unchanged warehouse stays at its last published value
- the result is point-in-time advisory visibility, not live authority

### What would be a real failure

- HQ snapshot overwrites or is treated as warehouse authority
- global totals are shown without preserving per-warehouse source rows
- aggregation opens a live SharePoint file directly and fails on a partial sync instead of safely skipping/catching up later

### Central aggregation evidence to capture

- screenshot of the initial global snapshot rows for both warehouses
- screenshot of the catch-up global snapshot after one warehouse republishes
- screenshot of any status metadata showing warehouse count or skipped snapshot count

## Optional Scheduler Checks

If you want to prove routine operations instead of purely manual ones:

1. Run warehouse batch from the Admin XLAM scheduler-facing macro.
2. Run warehouse publish from the Admin XLAM scheduler-facing macro.
3. Run HQ aggregation from the Admin XLAM scheduler-facing macro.
4. If using Task Scheduler, confirm task history and any log file output.

Expected:

- rerunning warehouse batch is safe, even when it processes zero new rows
- rerunning warehouse publish is safe
- rerunning HQ aggregation is safe

## What counts as a real problem

- local operator post fails because WAN or SharePoint is down
- local quantity does not update even though the processor says the event was applied
- stale refresh clears `ReceivedTally`, shipping staging, production staging, or local logs
- repeated publish leaves broken `.uploading` leftovers and never recovers
- HQ aggregation opens a live SharePoint file directly and fails on partial sync
- users start relying on the global snapshot as warehouse authority

## What does not count as a failure

- SharePoint publish fails while local warehouse work still completes
- operator inventory shows stale metadata during WAN delay
- global snapshot lags behind current warehouse truth until the next successful publish/aggregation cycle

## Suggested real-world proving record

For each stage, record:

- date/time
- machine used
- warehouse id
- station id
- action taken
- expected result
- actual result
- screenshots taken
- files checked
- pass/fail

## Final judgment

You can call the system operational for Phase 6 only if all of these are true:

- one-account local use is stable
- LAN use is stable
- LAN + WAN use is stable
- stale-state behavior is visible and non-destructive
- central aggregation works from published artifacts
- everyone involved understands that the warehouse-local workbook is authoritative and the global snapshot is advisory
