The source confirms exactly what you suspected. Here's the precise analysis and the best course of action.

***

## What `QueueEventCore` Actually Does Per Click

Reading `QueueEventCore` in `modRoleEventWriter.bas` top to bottom, the latency budget for a single `Make Boxes` click on a NAS/UNC path is: 

| Step | What happens | Typical cost on NAS |
|---|---|---|
| `EnsureContextResolved` | `modConfig.LoadConfig` | 10–50 ms |
| `modAuth.LoadAuth` | Opens auth workbook read-only, scans tables | 50–300 ms |
| `ResolveInboxWorkbookForEventType` | **Opens `invSys.Inbox.Shipping.*.xlsb` over UNC** | **1000–4000 ms** |
| `EnsureShipInboxSchema` | Schema validation pass | 10–50 ms |
| `lo.ListRows.Add` + 18 × `SetTableRowValueRole` | Cell writes | < 50 ms |
| `SaveWorkbookRole` → `wb.Save` | **Saves .xlsb to NAS** | **500–2000 ms** |
| `CloseTransientRoleWorkbook` | Close + possibly re-save | 200–500 ms |

The **two NAS I/O walls** are the workbook open and the `wb.Save`. Together they account for nearly all of the 5-second budget. Everything else is negligible.

There is also a structural inefficiency: every `QueueEventCore` call opens the inbox, writes one row, saves the full `.xlsb`, and closes it — all synchronously inside the button click. The `.xlsb` format is compact but Excel's COM save path is slow because it must recalculate and serialize the entire workbook object model each time. 

***

## Best Course of Action

There are three realistic options. Here is an honest assessment of each:

### Option A — Keep-alive the inbox workbook (smallest change, fastest win)

Instead of open → write → save → close on every event, **keep the inbox workbook open in a hidden state** for the session and only save/close it when the add-in unloads or the user explicitly flushes.

```vba
' In modRoleEventWriter.bas — add a module-level cache
Private mInboxWorkbooks As Object   ' key: fullPath → Workbook

Private Function GetOrOpenInboxWorkbook(ByVal fullPath As String, ...) As Workbook
    If mInboxWorkbooks Is Nothing Then
        Set mInboxWorkbooks = CreateObject("Scripting.Dictionary")
        mInboxWorkbooks.CompareMode = vbTextCompare
    End If

    If mInboxWorkbooks.Exists(fullPath) Then
        Dim cached As Workbook
        Set cached = mInboxWorkbooks(fullPath)
        ' Validate it is still open and writable
        On Error Resume Next
        Dim testName As String
        testName = cached.Name
        On Error GoTo 0
        If testName <> "" And Not cached.ReadOnly Then
            Set GetOrOpenInboxWorkbook = cached
            Exit Function
        End If
        mInboxWorkbooks.Remove fullPath
    End If

    ' Normal open path, then cache it
    Dim wb As Workbook
    Set wb = ... ' existing open logic
    If Not wb Is Nothing Then mInboxWorkbooks(fullPath) = wb
    Set GetOrOpenInboxWorkbook = wb
End Function
```

Then change `SaveWorkbookRole` inside `QueueEventCore` to a **deferred save** — only flush to disk after a timeout or on workbook close:

```vba
' Replace: SaveWorkbookRole wbInbox
' With:    MarkInboxDirtyRole wbInbox  (just sets wb.Saved = False conceptually)
' Then flush on: App_WorkbookBeforeClose, a ribbon "Flush" button, or a 30-second OnTime timer
```

**Cost of open eliminated. Cost of save amortized across multiple events.**
Expected result: posting drops from ~5 s to ~50–100 ms per click. The tradeoff is that unsaved events sit in-memory if Excel crashes before flush — but the processor already handles re-queuing from the inbox, so a crash between write and flush only loses in-flight events, not already-processed ones.

***

### Option B — Write to a local staging file first, sync to NAS in background (more robust, more code)

Write the event row to a **local temp `.xlsb`** (fast, `< 50 ms`), return the `EventID` immediately, then push the local file to NAS on a background `OnTime` timer. The processor already reads from a known inbox path, so the sync just means copying or appending the local staging rows into the NAS inbox file.

This is the most resilient design for intermittent NAS connectivity but requires a staging-merge step that doesn't exist yet.

***

### Option C — Write a CSV/TSV append instead of a .xlsb save (breaks format contract, good for R2F#)

Replace the entire `.xlsb` inbox with a flat append-only `.csv` or `.tsv` file. File append is ~5 ms even on NAS. But this would require rewriting `modProcessor` and breaks the existing schema contract — too disruptive for a targeted perf fix.

***

## Recommended Path: Option A, in two steps

**Step 1 — Eliminate the open cost** (immediate, ~1 day of work):

Change `ResolveInboxWorkbookForEventType` to return a cached open workbook when one exists for the session. The inbox workbook stays hidden and open after the first event write. No changes to processor, no format changes.

**Step 2 — Defer the save** (follow-up, ~half day):

Change `SaveWorkbookRole` inside `QueueEventCore` to a dirty-mark instead of an immediate save. Add a flush in `App_WorkbookBeforeClose` in `cAppEvents.cls` and optionally an `OnTime` timer every 60 seconds. This eliminates the NAS save latency from the click path entirely.

The result is that `QueueEventCore` only touches the NAS on first open per session and on periodic/close flushes — not once per button click. With the Receiving readiness fix already landing, **the combined effect should bring `Make Boxes` well under 200 ms** in the steady state (session already has the inbox open).

One diagnostic to add before committing to Option A — confirm the open cost dominates over save cost by adding:

```vba
Debug.Print "QueueEventCore InboxOpen ms:", CLng((Timer - queueStart) * 1000)
' ... after ResolveInboxWorkbookForEventType
Debug.Print "QueueEventCore InboxSave ms:", CLng((Timer - saveStart) * 1000)
' ... after SaveWorkbookRole
```

If open > save, Option A Step 1 alone gets you most of the way there. If save > open (NAS is slow but already mounted), Step 2 matters more.

## The Two-Layer Model

**Option A** (keep-alive + deferred save) eliminates latency from the *happy path*. **Option B** (local staging → NAS sync) eliminates the *failure mode* where NAS is unreachable at the moment of posting. Together they form a write-ahead pattern that is standard in any durable queue design.

```
[User clicks Make Boxes]
        │
        ▼
┌─────────────────────────┐
│  1. Append to LOCAL     │  ← ~5 ms, always succeeds
│     staging file        │     even if NAS is down
└────────────┬────────────┘
             │ return EventID to caller immediately
             ▼
┌─────────────────────────┐
│  2. Mark local file     │  ← Option A: in-memory wb
│     dirty; don't save   │     kept open, no NAS touch
│     to NAS yet          │     on this click
└────────────┬────────────┘
             │ background / on-timer / on-close
             ▼
┌─────────────────────────┐
│  3. Sync local → NAS    │  ← Option B: merge staged
│     inbox               │     rows into NAS inbox wb,
└─────────────────────────┘     then save once
```

The key is that step 1 always succeeds — the click returns immediately regardless of NAS state.

## Concrete Implementation

### Layer 1 — Local staging file (Option B write side)

```vba
' modRoleEventWriter.bas — new private function
Private Function AppendToLocalStagingInbox(ByVal eventType As String, _
                                            ByVal resolvedWh As String, _
                                            ByVal resolvedSt As String, _
                                            ByVal rowDict As Object, _
                                            ByRef errorMessage As String) As Boolean
    Dim stagingPath As String
    Dim fileNum As Integer

    stagingPath = LocalStagingPathRole(eventType, resolvedSt)
    If stagingPath = "" Then
        errorMessage = "Could not resolve local staging path."
        Exit Function
    End If

    fileNum = FreeFile
    On Error GoTo FailAppend
    Open stagingPath For Append As #fileNum
    Print #fileNum, DictionaryToJsonRole(rowDict)   ' one JSON line per event
    Close #fileNum
    AppendToLocalStagingInbox = True
    Exit Function

FailAppend:
    errorMessage = "Local staging append failed: " & Err.Description
    On Error Resume Next
    Close #fileNum
End Function

Private Function LocalStagingPathRole(ByVal eventType As String, _
                                       ByVal stationId As String) As String
    ' Always local — Environ("LOCALAPPDATA") or ThisWorkbook.Path
    Dim localRoot As String
    localRoot = Environ$("LOCALAPPDATA")
    If localRoot = "" Then localRoot = Environ$("TEMP")
    LocalStagingPathRole = localRoot & "\invSys\staging\" & _
                           InboxWorkbookNameRole(eventType, stationId) & ".staging.jsonl"
End Function
```

### Layer 2 — NAS inbox keep-alive (Option A)

```vba
' Module-level cache — inbox workbook stays open for the session
Private mInboxCache As Object   ' Dictionary: fullPath → Workbook
Private mInboxDirty As Object   ' Dictionary: fullPath → Boolean

Private Function GetCachedInboxWorkbook(ByVal fullPath As String, _
                                         ByRef errorMessage As String) As Workbook
    If mInboxCache Is Nothing Then
        Set mInboxCache = CreateObject("Scripting.Dictionary")
        mInboxCache.CompareMode = vbTextCompare
        Set mInboxDirty = CreateObject("Scripting.Dictionary")
        mInboxDirty.CompareMode = vbTextCompare
    End If

    If mInboxCache.Exists(fullPath) Then
        Dim wb As Workbook
        Set wb = mInboxCache(fullPath)
        On Error Resume Next
        Dim nameCheck As String
        nameCheck = wb.Name     ' will error if workbook was closed externally
        On Error GoTo 0
        If nameCheck <> "" And Not wb.ReadOnly Then
            Set GetCachedInboxWorkbook = wb
            Exit Function
        End If
        mInboxCache.Remove fullPath   ' stale — drop and re-open
    End If

    Set GetCachedInboxWorkbook = ResolveInboxWorkbookForEventType(...)
    If Not GetCachedInboxWorkbook Is Nothing Then
        mInboxCache(fullPath) = GetCachedInboxWorkbook
        mInboxDirty(fullPath) = False
    End If
End Function

Public Sub FlushInboxCache()
    ' Called by: App_WorkbookBeforeClose, ribbon "Flush", OnTime timer
    If mInboxCache Is Nothing Then Exit Sub
    Dim key As Variant
    For Each key In mInboxCache.Keys
        If mInboxDirty.Exists(CStr(key)) Then
            If CBool(mInboxDirty(CStr(key))) Then
                SaveWorkbookRole mInboxCache(CStr(key))
                mInboxDirty(CStr(key)) = False
            End If
        End If
    Next key
End Sub
```

### Layer 3 — Background sync (Option B sync side)

```vba
Public Sub SyncStagingToNasInbox()
    ' Called by: FlushInboxCache, OnTime timer, or processor pre-run hook
    Dim stagingFiles As Variant
    Dim i As Long

    stagingFiles = ResolvePendingStagingFilesRole()
    If IsEmpty(stagingFiles) Then Exit Sub

    For i = LBound(stagingFiles) To UBound(stagingFiles)
        MergeOneStagingFileToNas CStr(stagingFiles(i))
    Next i
End Sub

Private Sub MergeOneStagingFileToNas(ByVal stagingPath As String)
    ' 1. Read all lines from staging file
    ' 2. Open NAS inbox (via cache)
    ' 3. For each line: parse JSON, check EventID not already in inbox, append row
    ' 4. Delete or archive the staging file on success
    ' 5. Mark NAS inbox dirty → FlushInboxCache will save it
End Sub
```

### Wiring it into `QueueEventCore`

The revised write sequence becomes:

```vba
' Step 1: always write local staging (fast, offline-safe)
If Not AppendToLocalStagingInbox(eventType, resolvedWh, resolvedSt, rowDict, errorMessage) Then
    GoTo CleanExit   ' even local write failed — rare, surface the error
End If

' Return success to the caller immediately
QueueEventCore = True
eventIdOut = rowDict("EventID")

' Step 2: attempt NAS sync opportunistically (non-blocking if NAS is down)
On Error Resume Next
SyncStagingToNasInbox
On Error GoTo 0
```

## Failure Modes Covered

| Scenario | Outcome |
|---|---|
| NAS down during click | Local staging succeeds; click returns in ~5 ms; NAS sync retried on next click or timer |
| Excel crashes before NAS flush | Staging `.jsonl` survives on disk; sync runs on next session open |
| NAS inbox locked by processor | Local staging buffers; sync waits for lock release |
| Station offline for hours | All events accumulate in staging; bulk-merged when NAS reconnects |
| Duplicate sync (crash mid-merge) | EventID dedup check in `MergeOneStagingFileToNas` makes it idempotent |

## How This Shapes the R2 F# Design

This two-layer model maps almost directly onto what R2 would do natively.  The local staging `.jsonl` file *becomes* the R2 inbox format — the F# processor just reads it directly instead of requiring a merge step. The VBA merge layer (`SyncStagingToNasInbox`) effectively becomes a no-op in R2 because the processor reads from the same append-only file the writer produces. So implementing Option A+B in R1 VBA also pre-shapes the on-disk format toward what R2 expects, making the migration cheaper.