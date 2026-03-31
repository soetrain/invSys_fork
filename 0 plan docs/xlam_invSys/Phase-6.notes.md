



Latest update:

- See `Expert Handoff - 2026-03-29 LAN Investigation` near the end of this file for the current state.
- The older `2026-03-28` notes below remain useful history, but they no longer describe the primary blocker.
- Current main blocker: S1 host-side lock/read-only contention on `WH1.invSys.Data.Inventory.xlsb`, not the original S2 UNC path issue.

This is a clean, well-structured baton. Here is a consolidated handoff document suitable for picking up tomorrow with full context.

***

## Phase 6 LAN Flow — Handoff State (2026-03-28)

### System Status

| Component | State |
|---|---|
| S1 `Confirm Writes` → S2 sync | ✅ Working |
| S2 → canonical runtime write | ⚠️ Config blocked |
| Bidirectional live-open subscriber | ✅ Working |
| `FRODECO.inventory_management.xlsb` subscribe | ✅ Working |
| Workbook open/close crash loops | ✅ Fixed |
| Shared perf/diagnostics logging | ✅ Landed |
| Processor malformed-path hardening | ✅ Landed (masks, not fixes) |

***

### Active Failure — Exact Evidence

The processor on S1 running `RunBatch` is logging this for every S2 inbox target:

```
SkipInboxTargetInvalidPath|Path=\\192.168.1.3\invSysStationS2\invSys.Inbox.Receiving.S2.xlsb|Error=Bad file name or number
```

Same pattern repeats for S2 Shipping and S2 Production inbox paths. After skipping all three, `InboxTargets=1` — only the local S1 target is processed. S2 work is queued but never picked up.

This is purely a **path configuration problem**. The sync engine, processor, and runtime propagation are all healthy.

***

### Root Cause Hypothesis

The UNC paths stored in `tblStationConfig.PathInboxRoot` for the S2 station row were written with a bad format or point to a share that is not reachable from the processor's execution context. The double-backslash pattern `\\\\192.168.1.3\\...` appearing in the log suggests either:

1. The path was written with escaped backslashes (`\\` → literal `\\`) into the config cell, producing a quadruple-backslash UNC that Windows cannot resolve
2. The share name `invSysStationS2` does not exist or requires credentials not available from the S1 session
3. The path was written as a mapped-drive path on S2 (e.g. `Z:\invSys\...`) and copied into config without converting to a UNC

***

### Tomorrow — Exact Steps

**Step 1 — Read the live S2 config row**

Open `WH1.invSys.Config.xlsb` on S1 (or read it via the Immediate Window) and print the raw value of `PathInboxRoot` for the S2 station row:

```vb
' Immediate Window on S1:
Dim wb As Workbook
Dim ws As Worksheet
Set wb = Workbooks("WH1.invSys.Config.xlsb")
Set ws = wb.Worksheets("tblStationConfig")
Dim r As Long
For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If InStr(LCase(ws.Cells(r, 1).Value), "s2") > 0 Then
        Debug.Print "Row " & r & ": StationId=" & ws.Cells(r, 1).Value
        Debug.Print "  PathInboxRoot=" & ws.Cells(r, 2).Value  ' adjust column index
        Debug.Print "  Len=" & Len(ws.Cells(r, 2).Value)
    End If
Next r
```

Adjust column indices to match the actual `tblStationConfig` schema. The `Len=` output will immediately reveal double-encoded backslashes — a valid UNC `\\192.168.1.3\share\file.xlsb` has `Len` of ~50; a quadruple-backslash version will be longer.

**Step 2 — Verify the share exists from S1**

From a PowerShell prompt on S1 (X1-Pro-Ai):

```powershell
# Test reachability:
Test-Path "\\192.168.1.3\invSysStationS2"

# List contents if reachable:
Get-ChildItem "\\192.168.1.3\invSysStationS2"

# Check specific inbox file:
Test-Path "\\192.168.1.3\invSysStationS2\invSys.Inbox.Receiving.S2.xlsb"
```

If `Test-Path` returns `False`, the share either doesn't exist, requires credentials, or the machine at `192.168.1.3` is not reachable. Confirm Arctic-Raptor's IP with `ipconfig` on S2.

**Step 3 — Fix the path producer**

There are two likely write locations:

**`tools/setup_lan_station.ps1`** — if this script writes `PathInboxRoot` into the config workbook during station provisioning, check how it constructs the UNC string. The most common bug is string interpolation producing `"\\\\$ip\\$share"` (four backslashes) instead of `"\\$ip\$share"` (two):

```powershell
# BAD — produces \\\\192.168.1.3\\share in the cell:
$path = "\\\\$StationIp\\$ShareName\\invSys.Inbox.Receiving.S2.xlsb"

# CORRECT — produces \\192.168.1.3\share in the cell:
$path = "\\$StationIp\$ShareName\invSys.Inbox.Receiving.S2.xlsb"
```

**`modRuntimeWorkbooks.bas`** — if inbox paths are written programmatically during runtime workbook creation, search for string concatenation building UNC paths and apply the same check.

**Step 4 — Write corrected paths back**

Once the correct path format is confirmed, either:
- Re-run `setup_lan_station.ps1` with corrected path logic, or
- Patch the config workbook directly from the Immediate Window for an immediate test:

```vb
' Direct patch — adjust row/column to match actual schema:
ws.Cells(r, 2).Value = "\\192.168.1.3\invSysStationS2\invSys.Inbox.Receiving.S2.xlsb"
wb.Save
```

**Step 5 — Retest and confirm in log**

After the path fix, run one S2 `Confirm Writes` and check `invSys.Inventory.Sync.log` for:

```
INBOX-STATUS | StationId=S2 | Role=Receiving | Path=\\192.168.1.3\... | Status=PROCESSED
[PERF-END] Total=...ms  Applied=1...
```

The absence of `SkipInboxTargetInvalidPath` and presence of `Status=PROCESSED` confirms the fix is complete. S2 cross-station writes should then propagate to canonical runtime and appear on S1.

***

### File Reference

| File | Purpose |
|---|---|
| `tools/setup_lan_station.ps1` | Primary suspect for path production bug |
| `src/Core/Modules/modRuntimeWorkbooks.bas` | Secondary suspect if paths are written at runtime |
| `src/Core/Modules/modConfig.bas` | Config load/save — check for path normalization on read |
| `src/Core/Modules/modProcessor.bas` | Already hardened — `SkipInboxTargetInvalidPath` diagnostic is here |
| `src/Receiving/Modules/modTS_Received.bas` | S2 entry point — not the issue, but useful for end-to-end retest |
| `C:\Users\Justin\AppData\Local\Temp\invSys.Inventory.Sync.log` | Ground truth — check after every retest |
***
a flip has happened, S2 was working, so we moved to optimize. i started testing the same workflow on S1, it did not work. Now S1 works but S2 does not. S1 & S2 are separate computers, this is real LAN testing.
This is a classic cross-machine config drift scenario. The most likely causes when one machine works and the other stops after a change are:

***

## Most Probable Causes

### 1 — Add-in paths diverged between S1 and S2
When you rebuilt and redeployed to fix S1, the `.xlam` files in `deploy/current` were updated. If S2's Excel still has the **old add-ins registered** from its previous install, it is running stale code while S1 runs the new code.

**Check on S2:**
```vb
' Immediate Window on S2:
Dim ai As AddIn
For Each ai In Application.AddIns
    If ai.Installed Then Debug.Print ai.Name & " | " & ai.FullName & " | Installed=" & ai.Installed
Next ai
```
Compare the `FullName` paths shown on S2 against S1. They must point to the same `deploy/current` files (or the S2-accessible equivalent — same share, same build).

***

### 2 — The `runId` / perf instrumentation changed a function signature that S2's stale workbook binding still calls by old name

If you added `runId` as a required parameter to `AppendToInbox`, `AppendToOutbox`, or `RunBatch`, any caller that was compiled against the **old signature** (including a stale `.xlsb` on S2) will fail silently or with a type mismatch — and the `Processed=1` result you were seeing before will now be `Processed=0` or no-op.

**Quick test:** In the Immediate Window on S2, manually call the new entry point:
```vb
modProcessor.RunBatch "WH1", report
Debug.Print report
```
If it errors with `Wrong number of arguments` or `Type mismatch`, a signature changed and S2's caller hasn't caught up.

***

### 3 — Screen/event suppression added to S1's path broke S2's event chain

If you wrapped S1's confirm path in `Application.EnableEvents = False` and that code is now also running on S2 (same `.xlam`), but S2's workbook relies on a `SheetChange` or `WorkbookActivate` event to trigger its read-model refresh, that event is now being suppressed and S2's UI never updates — even though the write succeeds.

***

### 4 — Shared file lock — S2 cannot open the inbox/outbox/snapshot workbooks

If S1 is now opening any of the shared workbooks (inbox, outbox, snapshot, canonical runtime) as part of its new sync path — even briefly — and S2 tries to open the same file simultaneously over the LAN, S2 will silently fail or get a read-only handle. The `ChangedRows=1` you see in the sync log on S1 means S1 is actively touching those files.

**Check:** Look at the timestamps in the sync log. If S1's `RuntimeStamp` entries are happening at the same wall-clock second as your S2 `Confirm Writes` attempt, there is a file contention window.

***

## Fastest Triage — One Question Per Machine

**On S2, Immediate Window:**
```vb
' 1. What build is loaded?
Debug.Print Application.AddIns("invSys.Receiving.xlam").FullName

' 2. Does the processor run at all?
Dim r As String
modProcessor.RunBatch "WH1", r
Debug.Print r
```

**On S1, check if it is holding a file open:**
```vb
Dim wb As Workbook
For Each wb In Workbooks
    Debug.Print wb.Name & " | " & wb.FullName
Next wb
```

If S1 has `WH1.invSys.Data.Inventory.xlsb` or `WH1.Outbox.Events.xlsb` in its open workbook list while S2 tries to write, that is your contention point.

***

## Net Assessment

The flip pattern — S1 works, S2 breaks — after a change that fixed S1 almost always means one of:
- S2 has stale add-ins not yet reflecting the new build
- A new required parameter broke S2's existing call chain
- S1's new sync path is holding shared files open long enough to block S2

Start with the add-in path check — it is the fastest to verify and the most common cause.
***
That changes the picture significantly. If Excel is holding a file lock on one of the transient workbooks — most likely `WH1.invSys.Snapshot.Inventory.xlsb`, `WH1.Outbox.Events.xlsb`, or `invSys.Inbox.Receiving.S2.xlsb` — the write pipeline either stalls silently or routes around the locked file entirely, which explains why the perf markers never appear: the instrumented code path is being skipped because a lock check earlier in the chain is short-circuiting to a fallback path or silent no-op.

***

## What Causes the Lock

The most likely sequence:

1. A previous `Confirm Writes` run opened one of the transient workbooks
2. The `EnableEvents = False` guard or a crash/interrupt caused `CleanExit` to be skipped
3. The workbook was never properly closed — it is still open in Excel's workbook collection, but invisible (hidden window or `Visible = False`)
4. The next run calls `Workbooks.Open` on the same path
5. Excel returns a "file locked by another user" dialog — which `DisplayAlerts = False` suppresses silently — and the open fails
6. The code receives `Nothing` from the open call, hits an `If wb Is Nothing Then Exit Function` guard, and returns without writing or logging

***

## Immediate Diagnosis — Check for Ghost Open Workbooks

Run this in the Immediate Window on S2 right now:

```vb
Dim wb As Workbook
For Each wb In Application.Workbooks
    Debug.Print wb.Name & " | Visible=" & wb.Windows(1).Visible & " | Path=" & wb.FullName
Next wb
```

If you see any of these in the output with `Visible=False`, those are your locks:
- `WH1.invSys.Snapshot.Inventory.xlsb`
- `invSys.Inbox.Receiving.S2.xlsb`
- `WH1.Outbox.Events.xlsb`

***

## Immediate Fix — Force Close the Ghost Workbooks

```vb
Dim names(2) As String
names(0) = "WH1.invSys.Snapshot.Inventory.xlsb"

***

## Expert Handoff - 2026-03-29 LAN Investigation

### Executive Summary

As of **March 29, 2026**, the Phase 6 LAN work is no longer blocked on the original S2 UNC path problem. That part was real and was fixed enough to prove cross-machine pickup. The current blocker is narrower and more serious:

- **S2 station inbox reachability over LAN was broken, then fixed**
- **Queued stale S2 receive rows were later processed successfully by S1 once share access was fixed**
- **Current main failure is on S1, the warehouse host**
- **S1 `Confirm Writes` and later S2 `Confirm Writes` runs are intermittently failing because the canonical inventory workbook is locked/read-only during runtime apply**
- **The operator workbook can still refresh from an older snapshot, which makes the UI look partially successful even when no new inventory apply happened**

This means the system is now in a mixed state:

- **S2 appears to be functionally closer to working**
- **S1 host-side runtime apply / canonical inventory lifecycle is not yet robust**
- **The remaining problem is not just "clogged inbox rows"; it is also host-side workbook lifetime / lock management**

### What Was Confirmed

#### 1. Original S2 LAN path failure was real

Shared config on S1 had S2 `PathInboxRoot` set to:

```text
\\192.168.1.3\invSysStationS2\
```

At first:

- `Test-Connection 192.168.1.3` succeeded from S1
- `Test-Path "\\192.168.1.3\invSysStationS2"` failed in normal S1 user context
- the processor logged:

```text
SkipInboxTargetInvalidPath|Path=\\192.168.1.3\invSysStationS2\invSys.Inbox.Receiving.S2.xlsb|Error=Bad file name or number
```

That was preventing S2 inbox rows from being picked up at all.

#### 2. S2 share/bootstrap issue was then fixed

On S2:

- local inbox folder existed at `C:\invSysStationS2`
- SMB share `invSysStationS2` was created
- after credential/access correction, from normal S1 user context:

```powershell
Test-Path "\\192.168.1.3\invSysStationS2"  -> True
Test-Path "\\192.168.1.3\invSysStationS2\invSys.Inbox.Receiving.S2.xlsb" -> True
```

That removed the original LAN reachability blocker.

#### 3. Old S2 backlog did eventually process

Once S1 could see the S2 share, previously queued S2 receive rows were processed by the S1 processor. This caused the unexpected quantity jump on `DEMO-RAW-BROWN-COLOR`.

Observed jump:

- expected: `+11`
- observed: `88 -> 143`

Confirmed reason:

- one current S1 event `+11`
- four older queued S2 receive events `+44`

That accounted exactly for the `+55`.

This proved that:

- S2 queue writes were reaching the S2 inbox workbook
- S1 processor could process S2 inbox rows once the share became accessible
- the canonical apply path itself can process S2-originated events

#### 4. A remaining inbox/backlog safety issue existed

`Receiving.ConfirmWrites` was clearing local staging before proving that the rows it just queued had actually left `NEW` state. That allowed silent backlog accumulation when runtime processing did not complete.

To reduce re-clogging, new guards were added so `ConfirmWrites` now:

- blocks if the receiving inbox already has pending `NEW` rows before posting more
- re-inspects the specific event IDs it just queued after runtime processing
- warns if those rows are still pending

This was intended to stop the system from compounding backlog while the host-side lock issue is being investigated.

### Current Primary Failure

The new main blocker is **host-side canonical inventory workbook locking**.

The critical evidence came from the S1 sync/runtime log at these exact times:

#### March 29, 2026 17:23:12

```text
RUNTIME | RunBatchAndRefresh|Workbook=FRODECO.inventory_management.xlsb|WarehouseId=WH1|Processed=0|BatchReport=Inventory workbook is read-only or locked by another Excel session.|RefreshReport=OK
RECEIVE-RUNTIME | Result=OK|Workbook=FRODECO.inventory_management.xlsb|WarehouseId=WH1|Report=Processed=0; BatchReport=Inventory workbook is read-only or locked by another Excel session.; RefreshReport=OK
```

#### March 29, 2026 17:36:20

```text
RUNTIME | RunBatchAndRefresh|Workbook=FRODECO.inventory_management.xlsb|WarehouseId=WH1|Processed=0|BatchReport=Inventory workbook is read-only or locked by another Excel session.|RefreshReport=OK
RECEIVE-RUNTIME | Result=OK|Workbook=FRODECO.inventory_management.xlsb|WarehouseId=WH1|Report=Processed=0; BatchReport=Inventory workbook is read-only or locked by another Excel session.; RefreshReport=OK
```

The user-visible symptom matched this exactly:

- `ReceivedLog` updated locally
- `InventoryManagement` showed a snapshot banner such as:

```text
INVENTORY SNAPSHOT CURRENT | Source=LOCAL | Refreshed=2026-03-29 17:23:11 | SnapshotId=WH1.invSys.Snapshot.Inventory.xlsb|20260329172043
```

- but `TOTAL INV` did not change

Interpretation:

- local role logging succeeded
- read-model refresh succeeded
- **no new canonical inventory apply occurred**
- the workbook refreshed from an **older snapshot**

### Why the UI Can Look "Partially Successful"

The current operator flow can show a refreshed snapshot surface even when `RunBatch` handled zero rows.

That created a misleading success condition:

- queue write succeeded
- local role log succeeded
- read-model refresh succeeded
- canonical inventory apply did **not** succeed

So the user sees:

- refreshed timestamp text
- intact UI
- maybe local log entries

but canonical quantity does not change.

This was patched in code so the runtime wrapper no longer treats `Processed=0` plus a lock message as success. However, at least one later observed run still showed the old `Result=OK` behavior in the log, which strongly suggests the currently loaded XLAM set was not yet fully aligned to the latest patched build during that test.

### Live Lock Evidence on S1

After the later failures, S1 still had an active Excel process and a live lock file:

- process: one `EXCEL.EXE`
- lock file:

```text
C:\invSys\WH1\~$WH1.invSys.Data.Inventory.xlsb
```

Direct lock test from PowerShell on S1:

```text
LOCKTEST=FAILED
The process cannot access the file 'C:\invSys\WH1\WH1.invSys.Data.Inventory.xlsb' because it is being used by another process.
```

So the canonical inventory workbook was genuinely locked at the file level at that moment.

### Working Theory for the Remaining Failure

The strongest current hypothesis is:

1. `RunBatch` or a closely related host-side path opens the canonical inventory workbook
2. the workbook remains open longer than intended inside the Excel host session
3. a later `Confirm Writes` run tries to reacquire/write the same canonical workbook
4. `ResolveInventoryWorkbookBridge` or lock acquisition sees it as read-only/locked
5. `RunBatch` returns `Processed=0`
6. the operator workbook refreshes from the most recent published snapshot instead of showing a hard failure

The key point is that the system is no longer failing first at S2 LAN inbox discovery. It is now failing at **warehouse-host inventory workbook lifecycle / contention**.

### Code Changes Already Made During Investigation

These changes were applied during the LAN/debugging pass.

#### Shared/Core compile-hardening

Direct compile-time dependencies on `modUiQuiet` and `modPerfLog` were removed from shared Core paths using safe `Application.Run` wrappers, so missing harness modules no longer break compile in shared code.

Files changed:

- `src/Core/Modules/modRoleEventWriter.bas`
- `src/Core/Modules/modProcessor.bas`
- `src/Core/Modules/modWarehouseSync.bas`
- `src/Core/Modules/modOperatorReadModel.bas`

#### Path normalization / LAN path handling

Path normalization was tightened to preserve UNC roots and reduce bad path serialization behavior.

Files changed:

- `src/Core/Modules/modConfig.bas`
- `src/Core/Modules/modRoleEventWriter.bas`
- `src/Core/Modules/modProcessor.bas`

#### Station inbox write/open behavior

Role-side queueing now fails cleanly on read-only inbox open instead of silently writing into the wrong state.

Additional folder/open handling was added for UNC paths.

#### Clipboard/save noise reduction

`Application.CutCopyMode = False` was added around save/close points to reduce Excel clipboard modal interruptions on transient inbox workbook handling.

#### Backlog / unclogging diagnostics

Processor-side backlog logging was added so inbox `NEW` row counts and age are visible in diagnostics.

Receive-side queueing now:

- blocks if pending receive rows already exist
- reports when just-queued rows remain pending after runtime processing

Files changed:

- `src/Core/Modules/modProcessor.bas`
- `src/Core/Modules/modRoleEventWriter.bas`
- `src/Receiving/Modules/modTS_Received.bas`

#### Runtime result hardening

`RunBatchAndRefreshOperatorWorkbook` was tightened so "no rows handled" no longer counts as success unless the batch report shows actual handling:

- `Applied > 0`, or
- `SkipDup > 0`

File changed:

- `src/Core/Modules/modOperatorReadModel.bas`

#### Attempted canonical workbook lifecycle fix

`RunBatch` was patched to track whether it opened the canonical inventory workbook transiently and close it on exit if it did.

File changed:

- `src/Core/Modules/modProcessor.bas`

This change has **not yet been conclusively validated in a clean rebuilt/reloaded Excel session**.

### Current State by Machine

#### S2

Status:

- no longer obviously blocked on share resolution
- can queue receive rows to its own inbox workbook
- has shown evidence of being processed by S1 once share access was fixed
- still not "proven robust"

Assessment:

- S2 is no longer the primary blocker
- S2 appears "closer to working" than S1 right now

#### S1 (warehouse host)

Status:

- still failing intermittently at runtime apply
- canonical inventory workbook lock/read-only condition is the active blocker
- can refresh from older snapshot even when apply fails

Assessment:

- **S1 warehouse-host runtime handling is the current main problem**

### Open Questions for Expert Review

1. **Is the canonical inventory workbook being intentionally kept open somewhere in the host session?**

   Candidate areas:
   - `modProcessor.RunBatch`
   - `modInventoryDomainBridge.ResolveInventoryWorkbookBridge`
   - inventory-domain open/create paths
   - snapshot publication path if it touches runtime workbook and leaves it open

2. **Is a hidden workbook / ghost workbook state present within the same Excel process?**

   There is strong reason to suspect hidden workbook retention or a workbook-open lifecycle mismatch.

3. **Is the lock being caused by this same Excel process, or by a second process / automation path?**

   The file-level lock is real, but the exact owner process / call path has not been definitively attributed yet.

4. **Should canonical runtime inventory be opened once per host session and managed centrally, rather than transiently per `RunBatch`?**

   Current behavior may be half-transient, half-session-scoped, which is a dangerous middle ground.

5. **Should `Confirm Writes` treat any `Processed=0` host-side result as a hard failure when queueing a new event?**

   Current patches move in that direction, but the loaded build at test time may not have been fully updated.

### What the Expert Should Probably Inspect First

#### 1. Canonical inventory workbook open/close ownership

Start with these files:

- `src/Core/Modules/modProcessor.bas`
- `src/Core/Modules/modInventoryDomainBridge.bas`
- `src/InventoryDomain/Modules/modInventoryApply.bas`
- `src/InventoryDomain/Modules/modInventoryPublisher.bas`
- `src/Core/Modules/modLockManager.bas`

The central question:

- when `RunBatch` resolves `WH1.invSys.Data.Inventory.xlsb`, who owns that workbook object, and who closes it?

#### 2. Same-session hidden/open workbook enumeration

At failure time on S1, enumerate `Application.Workbooks` and inspect:

- visible and hidden workbooks
- full paths
- `ReadOnly`
- whether `WH1.invSys.Data.Inventory.xlsb` is already open in-session

The current evidence strongly suggests that the workbook may remain open in the same Excel instance between runs.

#### 3. Rebuild/reload discipline

At least one later runtime log still showed the old false-success behavior after the code-side hardening. That means the live test environment may have been running a stale loaded XLAM set.

So before concluding any code patch failed, confirm:

- latest built `invSys.Core.xlam` is actually loaded
- latest built role/domain XLAMs are actually loaded
- stale `deploy/current/~$*.xlam` lock files are not confusing the load state

### Practical Reproduction State

Best current minimal reproduction:

1. Start with clean Excel close on S1
2. Open the host/operator workbook with the current XLAM set
3. Run one `Confirm Writes`
4. Run another `Confirm Writes`
5. Check whether:
   - `C:\invSys\WH1\~$WH1.invSys.Data.Inventory.xlsb` appears
   - canonical inventory workbook remains locked
   - second run reports `Processed=0`
   - operator read model refreshes from the previous snapshot anyway

That seems to be the most likely route to reproducing the current failure chain.

### Net Assessment

This has evolved from a station-LAN configuration problem into a **host-side runtime workbook lifecycle problem**.

The important progress is:

- S2 LAN inbox access was not imaginary; it was broken and was fixed
- S2 backlog processing later proved the cross-machine receive path can work
- the remaining blocker is now concentrated on S1 host-side canonical inventory access

So the current expert task is not "why can't S2 talk over LAN?" anymore.

It is:

**Why does the S1 host session intermittently retain or collide on `WH1.invSys.Data.Inventory.xlsb`, causing `RunBatch` to return `Processed=0` while the UI still refreshes from an older snapshot?**

That appears to be the highest-signal next debugging target.
names(1) = "invSys.Inbox.Receiving.S2.xlsb"
names(2) = "WH1.Outbox.Events.xlsb"

Dim i As Integer
Dim wb As Workbook
For i = 0 To 2
    On Error Resume Next
    Set wb = Workbooks(names(i))
    On Error GoTo 0
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
        Debug.Print "Force-closed: " & names(i)
        Set wb = Nothing
    End If
Next i
```

After running this, click `Confirm Writes` once. The perf markers should now appear.

***

## Permanent Fix — Defensive Open in Each Writer

In `modRoleEventWriter.bas` and `modWarehouseSync.bas`, wherever a transient workbook is opened, check if it is already in the collection **before** calling `Workbooks.Open`. If it is already open, reuse it rather than trying to open the path again:

```vb
Private Function GetOrOpenWorkbook(ByVal fullPath As String) As Workbook
    Dim wb As Workbook
    Dim nm As String
    nm = Dir(fullPath)                ' just the filename
    On Error Resume Next
    Set wb = Workbooks(nm)
    On Error GoTo 0
    If wb Is Nothing Then
        Set wb = Workbooks.Open(fullPath, ReadOnly:=False, UpdateLinks:=False)
    End If
    Set GetOrOpenWorkbook = wb
End Function
```

Replace every bare `Workbooks.Open(...)` call for transient files with `GetOrOpenWorkbook(...)`. This prevents the double-open that creates the lock and also eliminates the silent failure when `DisplayAlerts = False` swallows the lock dialog.
