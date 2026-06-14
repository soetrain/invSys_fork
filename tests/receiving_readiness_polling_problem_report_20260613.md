# Repeated RECEIVING-READINESS Polling While Book1 Is Open

Date: 2026-06-13

## Summary

While testing the Shipping BoxMaker UserForm from a new unsaved workbook (`Book1`), the Immediate Window is flooded with repeated `RECEIVING-READINESS` diagnostics for that workbook. This is unrelated to the BoxMaker inventory lock bug that was fixed earlier. The BoxMaker action now processes successfully, but the Excel session still feels sluggish because the Receiving add-in appears to be repeatedly evaluating and rendering readiness status against `Book1`.

Observed after a successful BoxMaker `Make Boxes` action:

```text
Box created. Used 60. component units; added 10. shippable units to TOTAL INV.

Processed=8; BatchReport=Applied=8; SkipDup=0; Poison=0; RunId=RUN-invsys_Zenbook_WH-INVENTORY-20260613181201-812420; SnapshotError=Snapshot workbook not resolved.; RefreshReport=OK
Inbox EventID: 5D37ED1B-79F4-456A-B020-574AB99F129Dk
```

Even though this is a Shipping workflow, `tests/immediate_window_dump.txt` contains 114 `RECEIVING-READINESS` entries for `Workbook=Book1`.

## Evidence

Representative repeated diagnostic:

```text
2026-06-13 18:10:25 DIAG RECEIVING-READINESS Workbook=Book1|SnapshotStatus=MISSING|AuthStatus=NO_USER|RuntimeStatus=MISSING_TABLES|Messages=Inventory snapshot was not found for the selected warehouse. Publish or refresh the warehouse snapshot, then run Setup UI.|Signed-in user was not found in Users & Roles for this warehouse. Ask an admin to add the account.|This workbook is missing Receiving tables. Click Setup UI to repair the operator workbook.
```

The same status repeats from approximately `18:10:25` through `18:12:59`, including bursts where multiple entries occur in the same second.

The repeated state is always effectively:

```text
Workbook=Book1
SnapshotStatus=MISSING
AuthStatus=NO_USER
RuntimeStatus=MISSING_TABLES
```

That means the Receiving readiness layer is not just logging once. It is repeatedly doing readiness work on a workbook that does not appear to be a Receiving operator workbook.

## Why This Matters

This adds UI and runtime overhead during unrelated Shipping workflows:

- It runs readiness checks for the wrong role while BoxMaker is open.
- It repeatedly probes workbook/runtime/auth/snapshot state.
- It may render or clear readiness shapes on activation/open paths.
- It adds noise to diagnostics and obscures real BoxMaker/Shipping events.
- It likely contributes to the user's report that `Make Boxes` still feels slow even after the runtime inventory lock was fixed.

This is especially visible because the user is intentionally testing by opening a new Excel workbook, using the Shipping form, then closing without saving. That workflow makes `Book1` the active workbook, and the Receiving event hooks appear to treat it as a possible Receiving workbook.

## Relevant Code

Primary files:

- `src/Receiving/ClassModules/cAppEvents.cls`
- `src/Receiving/Modules/modReceivingInit.bas`
- `tests/immediate_window_dump.txt`

The Receiving app event sink calls readiness checks on broad application events:

```vba
Private Sub App_NewWorkbook(ByVal Wb As Workbook)
    Application.EnableEvents = True
    modReceivingInit.EnsureReceivingSurfaceForWorkbook Wb
End Sub

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    Application.EnableEvents = True
    modReceivingInit.EnsureReceivingSurfaceForWorkbook Wb
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    Application.EnableEvents = True
    modReceivingInit.EnsureReceivingSurfaceForWorkbook Wb
End Sub
```

`EnsureReceivingSurfaceForWorkbook` is gated, but the role detection is too broad:

```vba
Public Sub EnsureReceivingSurfaceForWorkbook(ByVal wb As Workbook)
    Dim prevEvents As Boolean

    If wb Is Nothing Then Exit Sub
    If Not modRoleWorkbookSurfaces.ShouldBootstrapRoleWorkbookSurface(wb) Then Exit Sub
    If Not IsLikelyReceivingWorkbookReadiness(wb) Then Exit Sub

    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    ApplyReceivingReadinessForWorkbook wb, True
    Application.EnableEvents = prevEvents
End Sub
```

The likely over-broad condition is in `IsLikelyReceivingWorkbookReadiness`:

```vba
Private Function IsLikelyReceivingWorkbookReadiness(ByVal wb As Workbook) As Boolean
    Dim wbName As String

    If wb Is Nothing Or wb.IsAddin Then Exit Function
    wbName = LCase$(Trim$(wb.Name))

    If WorkbookHasReceivingSurfacesReadiness(wb) Then
        IsLikelyReceivingWorkbookReadiness = True
        Exit Function
    End If

    If wbName Like "*.receiving.operator.xls*" Then
        IsLikelyReceivingWorkbookReadiness = True
        Exit Function
    End If

    If modConfig.IsLoaded() Then
        IsLikelyReceivingWorkbookReadiness = (Trim$(modConfig.GetWarehouseId()) <> "")
    End If
End Function
```

That final fallback means: once config is loaded and has a warehouse id, almost any non-addin workbook can become "likely Receiving." A blank `Book1` in a Shipping workflow can pass this test, then get full Receiving readiness evaluation.

## Current Call Chain

Observed likely path:

```text
Excel activates or opens Book1
  -> Receiving cAppEvents.App_WorkbookActivate/App_WorkbookOpen/App_NewWorkbook
      -> modReceivingInit.EnsureReceivingSurfaceForWorkbook(Book1)
          -> IsLikelyReceivingWorkbookReadiness(Book1)
              -> returns True because modConfig is loaded and WarehouseId is nonblank
          -> ApplyReceivingReadinessForWorkbook(Book1, True)
              -> CheckReceivingReadinessForWorkbook(Book1)
                  -> ResolveRuntimeStatusReadiness
                  -> ResolveSnapshotStatusReadiness
                  -> ResolveAuthStatusReadiness
              -> RenderReceivingReadinessPanel or ClearReceivingReadinessPanel
              -> LogDiagnosticEvent "RECEIVING-READINESS", "Workbook=Book1|..."
```

## Suspected Root Cause

Receiving readiness role detection is using global runtime/config state as a proxy for workbook role.

The line:

```vba
IsLikelyReceivingWorkbookReadiness = (Trim$(modConfig.GetWarehouseId()) <> "")
```

is probably a legacy convenience from earlier ribbon/table workflows, but it is unsafe now that:

- The system has multiple role add-ins loaded at once.
- Users can operate role-specific UserForms from a blank unsaved workbook.
- Shipping BoxMaker and BoxBuilder no longer require setup of the active workbook first.
- Role surfaces are being phased away from "every active workbook is a candidate" behavior.

## Expected Behavior

When the user opens `Book1` and works in the Shipping BoxMaker UserForm:

- Receiving should not run readiness checks against `Book1`.
- Receiving should not render readiness panels into `Book1`.
- Receiving should not log repeated `RECEIVING-READINESS` diagnostics for `Book1`.
- A blank workbook should remain role-neutral unless the user explicitly runs Receiving Setup UI or opens a saved Receiving operator workbook.

## Actual Behavior

Receiving readiness repeatedly evaluates `Book1` and logs:

```text
Workbook=Book1|SnapshotStatus=MISSING|AuthStatus=NO_USER|RuntimeStatus=MISSING_TABLES
```

This happens during a Shipping form workflow and continues across many activation/interaction events.

## Recommended Fix

### Fix 1 - Tighten Receiving workbook detection

Change `IsLikelyReceivingWorkbookReadiness` so it only returns true for:

- Workbooks that already have Receiving surfaces/tables.
- Workbooks named like `*.receiving.operator.xls*`.
- Workbooks with explicit Receiving role metadata, if such metadata exists or can be added.
- Possibly workbooks opened via Receiving Setup UI during that explicit setup call.

Remove or heavily guard the global config fallback:

```vba
If modConfig.IsLoaded() Then
    IsLikelyReceivingWorkbookReadiness = (Trim$(modConfig.GetWarehouseId()) <> "")
End If
```

A safer replacement:

```vba
' Do not infer Receiving role from global config alone.
IsLikelyReceivingWorkbookReadiness = False
```

If backwards compatibility requires auto-bootstrapping fresh Receiving workbooks, make it explicit through a setup flag or call path, not passive workbook activation.

### Fix 2 - Add debounce/caching to readiness logging

Even for real Receiving workbooks, repeated identical diagnostics should be suppressed.

Possible module-level cache:

```vba
Private mLastReadinessLog As Object

' key: workbook.FullName or workbook.Name for unsaved
' value: packed readiness signature + timestamp
```

Only log when:

- status changes,
- messages change,
- a minimum interval has elapsed, or
- the caller explicitly requests a refresh.

This would prevent 100+ identical log rows from one unchanged workbook state.

### Fix 3 - Avoid readiness work on every WorkbookActivate

`WorkbookActivate` can fire often during UserForm work, hidden workbook opens, add-in interactions, and Excel focus changes. It should not run full readiness checks every time.

Options:

- Do only a cheap role check on activation.
- Defer full readiness to explicit Receiving ribbon actions.
- Run full readiness on `WorkbookOpen`/`NewWorkbook`, but debounce activation.
- Skip if the same workbook was checked recently and no relevant workbook/runtime state changed.

### Fix 4 - Do not force `Application.EnableEvents = True` in app events

Each event handler currently starts with:

```vba
Application.EnableEvents = True
```

This can undo intentional event suppression from other role modules or quiet UI sections. It may also make cross-role event recursion harder to control.

Prefer preserving caller/application state:

```vba
If Not Application.EnableEvents Then Exit Sub
```

or only re-enable events in controlled cleanup paths owned by the same module.

## Acceptance Criteria

After the fix:

1. Open a new blank Excel workbook (`Book1`).
2. Open Shipping BoxMaker.
3. Select a box.
4. Set `Qty = 10`.
5. Click `Make Boxes`.

Expected:

- BoxMaker still processes the `BOX_BUILD` event successfully.
- `tests/immediate_window_dump.txt` should not show repeated `RECEIVING-READINESS Workbook=Book1` entries during the Shipping workflow.
- If any Receiving readiness log appears for `Book1`, it should be at most one initial diagnostic and only if explicitly justified.
- Opening a real Receiving operator workbook should still render readiness status when required.
- Running Receiving Setup UI should still prepare/repair the Receiving workbook.

## Suggested Tests

Add or update unit/integration coverage:

1. `IsLikelyReceivingWorkbookReadiness_BlankBookWithConfigLoaded_ReturnsFalse`
   - Arrange `modConfig.IsLoaded = True` and a valid warehouse id.
   - Use a blank workbook with no Receiving surfaces.
   - Assert `IsLikelyReceivingWorkbookReadiness` is false, or assert `EnsureReceivingSurfaceForWorkbook` does not apply readiness.

2. `EnsureReceivingSurface_BlankWorkbook_DoesNotLogReadiness`
   - Call `EnsureReceivingSurfaceForWorkbook` on a blank workbook.
   - Assert no `RECEIVING-READINESS` diagnostic is emitted.

3. `EnsureReceivingSurface_ReceivingWorkbook_StillAppliesReadiness`
   - Use a workbook with Receiving surfaces or a receiving-operator filename.
   - Assert readiness still runs.

4. `ReceivingReadiness_DebouncesRepeatedIdenticalStatus`
   - Call `ApplyReceivingReadinessForWorkbook` repeatedly with unchanged status.
   - Assert diagnostic output is not duplicated excessively.

## Current Workaround

No reliable user-level workaround besides avoiding blank active workbooks or unloading the Receiving add-in. That is not acceptable for normal multi-role usage, because Shipping should be usable from a fresh workbook without Receiving repeatedly inspecting it.

## Priority

Medium-high.

The bug does not corrupt inventory now that BoxMaker processing succeeds, but it creates significant UI noise and likely contributes to perceived slowness. It also reflects a broader architectural issue: role add-ins are still treating unrelated workbooks as role candidates because global config is loaded.
