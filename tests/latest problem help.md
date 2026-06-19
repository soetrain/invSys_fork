Now I have full context. The file is very large — I have enough to give a complete, precise analysis. Here's the diagnosis:

***

## Why the Tests Don't Catch the Problems

There are two distinct problems in the report, and both share the same fundamental testing gap: **the tests call module-level helpers directly and bypass the actual form-backed code paths** that produce the observed behavior.

***

### Problem 1: The Hanging Tests (`ShipmentsFormCommitLine` / `ShipmentsFormRefreshRuntimeInventory`)

The hang is **not a test logic error** — it's an architectural barrier. The tests call public procedures in `invSys.Shipping.xlam` via `Application.Run` while the validation harness workbook is the active workbook. This triggers at least one of three deadlock scenarios:

**Why it hangs instead of failing:**

| Scenario | What Happens |
|---|---|
| `ActiveWorkbook` called inside the Shipping add-in | Returns the *harness* workbook, not the operator workbook. The procedure tries to find `invSys`/`ShipmentsTally` in the harness, can't, and either loops or blocks on a MsgBox |
| A silent VBE compile dialog fires | The deployed XLAM has a missing reference or compile error that only manifests when invoked cross-process. The dialog blocks forever, PowerShell never gets a return |
| Re-entrancy via `Application.OnTime` or a modal form | If the Shipping add-in internally schedules work or opens a form, the `Application.Run` call never returns — Excel's message pump is waiting on user input |

The evidence: `Get-Process EXCEL` showing `MainWindowTitle = "Microsoft Visual Basic"` means a **VBE compile/error dialog is open and waiting for a click** that PowerShell can never deliver.

**The tests cannot work as written because `Application.Run` is synchronous and blocks if the called procedure blocks.** There is no timeout in the harness.

***

### Problem 2: The Dirty-Transaction Tests (Inventory Increasing After `Shipments Sent`)

The existing passing tests **do not reproduce this bug** for these specific reasons:

**What the passing tests actually cover:**

Looking at [TestPhase6CoreSurfaces.bas](https://github.com/justinwj/invSys_fork/blob/codex/fix-tester-station-nas-setup/tests/unit/TestPhase6CoreSurfaces.bas), the shipping tests (`TestSavedShippingWorkbook_*`) call:
1. `modRoleEventWriter.QueuePayloadEvent` → `modProcessor.RunBatch` → `modOperatorReadModel.RefreshInventoryReadModelForWorkbook`

That path exercises the **processor-level deduction** and the **read-model refresh from snapshot**. What it does NOT exercise:

```
TestSavedShippingWorkbook_ReopenQueueProcessRefreshPreservesStagingAndLogs
→ manually seeds invSys at qty 1, ships qty 6, asserts TOTAL INV = 4 (10−6)
→ verifies RefreshInventoryReadModelForWorkbook gives the right number
```

**The gap:** None of the tests run the code path that `Shipments Sent` actually uses in the form: `ShipmentsFormRunShipmentsSentRows` → `ApplyShipmentsSentRowsInventory` → the subsequent `frmShipmentsTally.RefreshProjectedShippableInventory` reload. The form reload path reads from a version-log or version-inventory table — **not** from the snapshot path tested in `RefreshInventoryReadModelForWorkbook`. That secondary reload is what overwrites `Projected Inv` with the stale NAS value.

Specifically: the existing tests assert `TOTAL INV` from the `invSys` table after `RefreshInventoryReadModelForWorkbook`. The user-visible bug is in `PendingBoxVersionInventoryOverlayText` and `BoxMakerFormLoadShippableVersionInventory` — a completely different code path that runs **after** `Shipments Sent` reloads the shippables list.

***

## What the Tests Need to Do Instead

### For Problem 1 (Hanging Tests) — Three viable approaches:

**Option A: Extract pure module functions (the right fix long-term)**

Refactor the Shipping add-in so `ShipmentsFormCommitLine` and `ShipmentsFormRefreshRuntimeInventory` are thin wrappers that call explicit, testable module functions:

```vba
' In modShipmentsTallyActions.bas (new pure module, no ActiveWorkbook)
Public Function CommitShipmentLine(ByVal wbOperator As Workbook, _
                                   ByVal wbInbox As Workbook, _
                                   ByVal userId As String, _
                                   ... ) As Boolean
```

Tests then call `modShipmentsTallyActions.CommitShipmentLine(wbOps, wbInbox, ...)` directly — **no `Application.Run`, no add-in boundary crossing, no ActiveWorkbook dependency**. This is the same pattern every other passing test in Phase 6 uses.

**Option B: Add a harness timeout + cleanup guard (immediate mitigation)**

Before the hung-test can be removed from the harness entirely, add a PowerShell watchdog around `Application.Run` calls into Shipping:

```powershell
$job = Start-Job { Application.Run "invSys.Shipping.xlam!ShipmentsFormCommitLine" }
if (-not ($job | Wait-Job -Timeout 15)) {
    $job | Stop-Job
    Get-Process EXCEL | Stop-Process -Force
    Write-Error "TIMEOUT: ShipmentsFormCommitLine hung"
}
```

This at least prevents the XLAM file-lock cascade. But it does not make the test pass — it just fails cleanly.

**Option C: Replace `Application.Run` with direct module calls in a test-mode XLAM build**

Add a compile flag (`#Const TESTMODE = True`) to the Shipping add-in that disables form-show calls and MsgBoxes. Tests invoke the XLAM with the test flag set. This is the standard VBA integration-test pattern when pure extraction isn't feasible yet.

***

### For Problem 2 (Dirty Transaction) — Exact test structure needed:

The new tests must call the **same reload path the form uses**, not `RefreshInventoryReadModelForWorkbook`. Based on the suspected code areas in the report:

```vba
Public Function TestShipmentsSent_NeverIncreasesVisibleProjectedInventory() As Long
    ' 1. Create operator workbook with ShippingBOMView, ShipmentsTally, invSys
    '    Seed multi-row state:
    '      T26 v1: NAS=8, Projected=7, Locked=1
    '      T27 v1: NAS=10, Projected=9, Locked=1
    '      T27 v2: NAS=20, Projected=19, Locked=1

    ' 2. Snapshot the BEFORE state from the visible shippables list:
    '    Call BoxMakerFormLoadShippableVersionInventory(wbOps)
    '    or BuildActiveShippingReservationTotalsFromTable(...)
    '    Capture: NasInvBefore, ProjectedInvBefore per (Box, Version)

    ' 3. Run the same path as the form's Shipments Sent button:
    '    modShipmentsForm.RunShipmentsSentRows wbOps, wbInbox, ...
    '    (not Application.Run — this is the pure module call)

    ' 4. Reload the visible shippables list using the EXACT form reload path:
    '    Call ShipmentsFormLoadNasReservationTotals(wbOps)
    '    or frmShipmentsTally.RefreshProjectedShippableInventory equivalent

    ' 5. Assert monotonic non-increase for every visible row:
    For Each row In visibleShippables
        If row.NasInvAfter > row.NasInvBefore Then FAIL "NAS increased for " & row.BoxVersion
        If row.ProjectedInvAfter > row.ProjectedInvBefore Then FAIL "Projected increased for " & row.BoxVersion
    Next row

    TestShipmentsSent_NeverIncreasesVisibleProjectedInventory = 1
End Function
```

The critical detail: **step 4 must call `ShipmentsFormLoadNasReservationTotals` / `ApplyShipmentsSentVersionInventoryOverlay` or whatever the form's actual reload function is** — not `RefreshInventoryReadModelForWorkbook`. The existing tests use the wrong reload path, so they pass even when the real form reload produces the dirty result.

The stale-snapshot overwrite test follows the same pattern but seeds a stale `NAS Inv` that is higher than the projected deduction, then asserts that after `Shipments Sent` + form reload, `Projected Inv` does not move back up to the stale NAS value.

***

## Summary

The tests don't catch the bugs because:

1. **Hanging tests** cross the `Application.Run` / add-in boundary into form-backed procedures that depend on `ActiveWorkbook` and can open modal dialogs — behaviors that block the harness forever. The fix is to extract pure module functions that accept explicit workbook arguments, matching the pattern every passing test already uses.

2. **Dirty transaction tests** call `RefreshInventoryReadModelForWorkbook` for their reload assertion, but the real bug lives in the form's shippables-list reload path (`ShipmentsFormLoadNasReservationTotals`, `ApplyShipmentsSentVersionInventoryOverlay`, `PendingBoxVersionInventoryOverlayText`). Tests that skip those functions will always pass even when the user-visible form shows inventory increasing.