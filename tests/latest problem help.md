Good — now I have the full picture. The problem is multi-layered and the sources confirm exactly what is happening. Here is a complete performance diagnosis and fix plan. 

***

## Performance Root Causes

There are four distinct problems causing the blinking and slowness, ordered by impact.

### 1. `Workbooks.Open` on Every Box Selection (The Biggest Offender)

Every time the user selects a box, the system is calling `Workbooks.Open` on inventory/inbox/auth workbooks that are **already open** in the session. In `modInventoryDomainBridge.bas`, `modRoleEventWriter.bas`, `modWarehouseSync.bas`, and `modInventoryApply.bas`, the pattern is: 

```vba
' Current bad pattern:
Set wb = Application.Workbooks.Open(targetPath)   ' ← opens EVEN IF already open
```

When Excel opens a workbook that is already open, it either: (a) flashes the workbook's window, causing visible blink, or (b) triggers a "file modified" dialog that `DisplayAlerts = False` eats silently, then re-opens it anyway. Either way = blink + overhead.

The `modRoleEventWriter.bas` already sets `Application.ScreenUpdating = False` before opening, but that only suppresses one layer of the flash. 

**Fix:** Add an "already open?" guard before every `Workbooks.Open` call:

```vba
' Shared helper — add to modRuntimeWorkbooks.bas
Public Function GetOrOpenWorkbook(ByVal filePath As String, _
                                  Optional ByVal readOnly As Boolean = False) As Workbook
    Dim normalPath As String
    normalPath = NormalizePath(filePath)         ' lowercase, canonical separators

    ' Check if already open — zero I/O cost
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(NormalizePath(wb.FullName), normalPath, vbTextCompare) = 0 Then
            Set GetOrOpenWorkbook = wb
            Exit Function
        End If
    Next wb

    ' Only opens if not already in the collection
    Dim prev As Long
    prev = Application.AutomationSecurity
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Set GetOrOpenWorkbook = Application.Workbooks.Open( _
        Filename:=filePath, UpdateLinks:=0, ReadOnly:=readOnly, _
        Notify:=False, AddToMru:=False)
    Application.AutomationSecurity = prev
End Function
```

Then replace every bare `Workbooks.Open` in `modInventoryDomainBridge`, `modRoleEventWriter`, `modWarehouseSync`, and `modInventoryApply` with `GetOrOpenWorkbook(path)`.

***

### 2. Missing `modUiQuiet` Wrapper Around BoxBOM Save Path

`modTS_Received.bas` uses `modUiQuiet.BeginQuietUi` / `EndQuietUi` properly.  The BoxBOM save path almost certainly does not, because it was written after those guards existed. Every cell write, table row addition, and `ListObject` operation during a box save triggers a recalculation + repaint unless wrapped.

**Fix:** Wrap the entire BoxBOM commit operation:

```vba
' In whatever module handles BoxBOM save (modTS_Shipments or similar):
Public Sub SaveBoxBOMSelection(...)
    Dim quietReport As String
    Dim wb As Workbook
    Set wb = callbackCell.Parent.Parent

    modUiQuiet.BeginQuietUi wb         ' ← ScreenUpdating=F, Calc=Manual, Events=F
    On Error GoTo CleanupQuiet

    '--- all BoxBOM writes here ---

CleanupQuiet:
    modUiQuiet.EndQuietUi wb           ' ← always restores, even on error
    If Err.Number <> 0 Then Err.Raise Err.Number
End Sub
```

If `modUiQuiet` is not yet used in the Shipping module, add a direct guard at minimum:

```vba
Dim prevScreen As Boolean
Dim prevCalc As XlCalculation
Dim prevEvents As Boolean
prevScreen = Application.ScreenUpdating
prevCalc   = Application.Calculation
prevEvents = Application.EnableEvents

Application.ScreenUpdating = False
Application.Calculation    = xlCalculationManual
Application.EnableEvents   = False

On Error GoTo RestoreAppState
    '--- writes ---
RestoreAppState:
Application.ScreenUpdating = prevScreen
Application.Calculation    = prevCalc
Application.EnableEvents   = prevEvents
```

***

### 3. `TryRefreshSearchInventoryReadModel` Triggers a Full Workbook Surface Rebuild on Picker Open

In `cDynItemSearch.LoadManagedInventoryItems`, if the table is empty it calls `TryRefreshSearchInventoryReadModel()`, which calls: 

```vba
modRoleWorkbookSurfaces.EnsureInventoryManagementSurface(wb, surfaceReport)
modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wb, ...)
```

That surface rebuild writes cells, possibly adds sheets, and fires recalculations — all during picker open. This explains the "a lot of time between box selections."

**Fix:** Decouple the refresh from the picker. The picker should **never trigger a background refresh** in the hot path. Refresh should be explicit (a ribbon button or a timed background event):

```vba
' In cDynItemSearch.LoadManagedInventoryItems — remove the retry/refresh loop:
Private Function LoadManagedInventoryItems(...) As Variant
    Dim loInv As ListObject
    Set loInv = ResolveManagedInventoryTable()
    If loInv Is Nothing Then Exit Function
    If loInv.DataBodyRange Is Nothing Then Exit Function
    ' No refresh attempt — show empty state instead, user refreshes explicitly
    LoadManagedInventoryItems = BuildInventoryPickerItemsFromTable(loInv, includeCategory)
End Function
```

If an auto-refresh is truly needed, move it to a non-blocking `Application.OnTime` call that fires after the picker is already visible:

```vba
' Fire refresh 500ms after picker opens, non-blocking:
Application.OnTime Now + TimeValue("00:00:01"), "mBackground.RefreshInventoryReadModel"
```

***

### 4. BoxBOM Version Autofill is Writing on Every `Change` Event Without Debounce

The "versions are auto filling" confirmation means something is wired to `Worksheet_Change` or `SelectionChange`. If that handler calls `Workbooks.Open` or triggers a surface refresh, every keystroke in BoxBOM costs a full open/close cycle.

**Diagnostic — add to the Shipping worksheet's `Worksheet_Change`:**
```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    Debug.Print "Worksheet_Change fired: " & Target.Address & " @ " & Now
    ' ... existing logic
End Sub
```

If the Immediate Window floods with Change events on a single box select, the handler needs a guard:

```vba
Private mSuppressChange As Boolean

Private Sub Worksheet_Change(ByVal Target As Range)
    If mSuppressChange Then Exit Sub
    mSuppressChange = True
    On Error GoTo Done
        ' ... writes happen here, not recursively
Done:
    mSuppressChange = False
End Sub
```

***

## Priority Order

| # | Fix | Effort | Impact |
|---|-----|--------|--------|
| 1 | Add `GetOrOpenWorkbook` guard in `modRuntimeWorkbooks` and replace bare `Workbooks.Open` in `modInventoryDomainBridge`, `modRoleEventWriter`, `modWarehouseSync` | Medium | **Eliminates all blink from redundant opens** |
| 2 | Wrap BoxBOM commit in `modUiQuiet.BeginQuietUi/EndQuietUi` | Small | **Eliminates repaint during save** |
| 3 | Remove `TryRefreshSearchInventoryReadModel` from picker hot path, move to `OnTime` | Small | **Eliminates the multi-second delay between selections** |
| 4 | Add `mSuppressChange` guard to Shipping `Worksheet_Change` | Small | **Stops recursive/cascading Change events** |

Fix #1 alone will eliminate the visible open/close flash. Fixes #2 and #3 together will make box selections feel near-instant. Fix #4 is a safety net to prevent regressions as BoxBOM grows.