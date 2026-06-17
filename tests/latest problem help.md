Now I have everything I need. Here is the complete analysis and test specification.

***

## What's Actually In the Repo Right Now

The `frmShipmentsTally.frm` on main is a **stub** — it has `btnSend_Click → ProcessShipmentsBatch` and `PopulateShipmentsForm`, but none of the local-cache, tombstone, or `LINE_ID` machinery described in the problem report exists in the committed code yet.  The `modTS_Shipments.bas` contains only the worksheet-surface flow (`BtnToShipments`, `BtnShipmentsSent`, `AggregatePackages`, etc.) and has zero references to `LINE_ID`, `shipping_active_*.tsv`, `PersistentSentShipmentRowExists`, or `ReconcileShipmentStagingFromShipmentLines`. 

That means the dirty-transaction machinery described in the problem report is either in uncommitted local code or was described as the intended design that needs to be built. Either way, the tests need to be written **against the intended behavior**, so they catch violations before or as the code is added.

***

## The Core Failure Modes to Guard Against

From the problem report, there are four distinct failure classes:

| # | Class | What goes wrong |
|---|---|---|
| **A** | **Row resurrection** | A sent row reappears in active Shipments after close/reopen |
| **B** | **SHIPMENTS staging drift** | `invSys.SHIPMENTS` is 0, but the form row claims `Area=Shipments`, so `Shipments Sent` fails |
| **C** | **Double-deduction** | Stale active row is reloaded → reconcile rebuilds SHIPMENTS from it → second send succeeds |
| **D** | **Illogical workflow sequence** | Buttons pressed out of order (e.g. `Shipments Sent` before `To Shipments`, `Boxes Made` before `Confirm Inventory`) produce wrong inventory math |

***

## Test Specifications

### Test Group 1 — Workflow Sequence Guard (Class D)

These are the "illogical workflow" tests the problem report asks for. Each test calls the validation-only surface (`ValidateQueueShipmentsSentEventFromCurrentWorkbook`) or the delta builders in isolation, with a controlled invSys state, and asserts the correct error or early-exit.

**Test 1.1 — `Shipments Sent` with zero SHIPMENTS staged must fail cleanly**

```text
Setup:
  invSys ROW 87 has:  TOTAL INV=10, MADE=0, SHIPMENTS=0
  AggregatePackages has ROW=87, QUANTITY=2

Action:
  Call ValidateQueueShipmentsSentEventFromCurrentWorkbook()

Expected:
  Returns error string containing "No staged shipments found in invSys.SHIPMENTS"
  Does NOT return "OK"
  invSys.SHIPMENTS unchanged (still 0)
```

**Test 1.2 — `To Shipments` with TOTAL INV < requested quantity must fail cleanly**

```text
Setup:
  invSys ROW 87 has:  TOTAL INV=1, SHIPMENTS=0
  AggregatePackages has ROW=87, QUANTITY=2

Action:
  Call BtnToShipments (or BuildShipmentDeltaPacket directly)

Expected:
  Returns/shows error "ROW 87 requires 2 but only 1 in TOTAL INV"
  invSys.SHIPMENTS unchanged (still 0)
```

**Test 1.3 — `Boxes Made` with insufficient component inventory must fail cleanly**

```text
Setup:
  invSys ROW 10 (kraft paper) has: TOTAL INV=3, USED=0
  AggregateBoxBOM requires ROW=10, QUANTITY=5

Action:
  Call ValidateComponentInventory (or BtnBoxesMade path)

Expected:
  Returns/shows "ROW 10 requires 5 but only 3 available"
  invSys.USED unchanged (still 0)
```

**Test 1.4 — `Confirm Inventory` with Use Existing Inventory checked must warn and exit**

```text
Setup:
  CHK_USE_EXISTING checkbox is checked (Value=1)

Action:
  Call BtnConfirmInventory

Expected:
  Shows "Use existing inventory is enabled. Skip Confirm inventory..."
  No staging changes applied
```

**Test 1.5 — Math: `Boxes Made` component deduction must equal BOM qty × package qty**

```vba
' This is the critical math guard.
' Box T25 BOM: kraft paper ROW=10 qty=2, tape ROW=11 qty=1
' ShipmentsTally: T25 qty=3
' Expected AggregateBoxBOM: ROW=10 qty=6, ROW=11 qty=3
'
' Assert:
'   AggregateBoxBOM.ROW(10).QUANTITY = 6   (2 per box × 3 boxes)
'   AggregateBoxBOM.ROW(11).QUANTITY = 3   (1 per box × 3 boxes)
'
' If these are wrong, BOM expansion (BuildBomSummary) is broken.
```

This test can be run as a worksheet validation: hydrate `ShipmentsTally` with one T25 row qty=3, run `RebuildShippingAggregates`, read `AggregateBoxBOM`, assert.

***

### Test Group 2 — Transaction State Guards (Classes A, B, C)

These require the `LINE_ID` and local state ledger that the problem report recommends building. Write these as the spec before or alongside the implementation.

**Test 2.1 — `TestShippingForm_SentRowsDoNotResurrectAcrossUnsavedWorkbookReopen`** *(exact name from problem report)*

```text
Steps:
  1. Open fresh unsaved workbook
  2. Open Shipping form
  3. Add order: Ref=TXN-001, Box=T25, Version=v1, Qty=2
  4. Move To Shipments → verify active list row has Area=Shipments
  5. Click Shipments Sent → capture EventID
  6. Assert active list does NOT contain TXN-001
  7. Assert LINE_ID of TXN-001 is in sent tombstone file
  8. Close form and workbook without saving
  9. Fully close Excel
 10. Reopen Excel, open Shipping form
 11. Assert TXN-001 is NOT in active Shipments listbox
 12. Assert TXN-001 is NOT in Not Shipped listbox
 13. Assert T25 v1 TOTAL INV or SHIPMENTS column is deducted by 2

Guards against:
  - Row resurrection (Class A)
  - Tombstone-by-value false negative (Class A)
  - Missing LINE_ID allowing match bypass (Class A)
```

**Test 2.2 — Reconcile must not reactivate a tombstoned row**

```text
Setup:
  Local active cache contains one row with LINE_ID=abc123, Area=Shipments
  Tombstone file contains LINE_ID=abc123

Action:
  Call LoadPersistentActiveShipmentRowsLocal (or ShipmentsFormLoadLines)

Expected:
  Zero rows loaded into active list
  LINE_ID abc123 was filtered by tombstone check

Guards against: Class A and Class C
```

**Test 2.3 — Double-send of same order must be impossible**

```text
Setup:
  Order TXN-002, LINE_ID=def456 is sent successfully
  Tombstone file has def456
  invSys.SHIPMENTS for ROW=87 was decremented to 0

Action:
  Force-load active cache back with same LINE_ID=def456
  Attempt to call Shipments Sent

Expected:
  Either:
    (a) Row is filtered before send attempt → "No staged shipments found"
    OR
    (b) Validation fails → "ROW 87 only has 0 staged but needs 2"
  Either outcome is acceptable. Both prevent double-deduction.
  invSys.SHIPMENTS must not go negative.

Guards against: Class C
```

**Test 2.4 — SHIPMENTS staging and active cache must agree**

```text
Setup:
  Active cache has ROW=87, Area=Shipments, Qty=2
  invSys.SHIPMENTS for ROW=87 = 0 (drift / stale)

Action:
  Call ReconcileShipmentStagingFromShipmentLines (if it exists)
  OR call BuildQueueableShipmentsSentDeltas

Expected (strict contract):
  If the row has no tombstone (i.e., it is a legitimate pending order):
    invSys.SHIPMENTS for ROW=87 must be rebuilt to 2 before send proceeds
  If the row IS tombstoned:
    Row must be excluded from reconciliation
    invSys.SHIPMENTS must remain 0

Guards against: Class B and Class C (the dangerous reconcile path)
```

***

### Test Group 3 — Column/Schema Guards

These can run as part of the existing Phase 6 validation.

**Test 3.1 — `BtnShipmentsSent` must fail if invSys lacks SHIPMENTS column**

```text
Setup: invSys has no SHIPMENTS column

Expected: Returns "invSys table missing SHIPMENTS/ROW columns."
```

**Test 3.2 — `BtnBoxesMade` must fail if AggregateBoxBOM has no ROW column**

```text
Expected: "AggregateBoxBOM missing QUANTITY/ROW columns." or similar early exit
```

**Test 3.3 — `BuildShipmentDeltaPacket` returns Nothing if AggPack is empty**

```text
Setup: AggregatePackages has no DataBodyRange

Expected: Returns Nothing, errNotes = ""  (the "no shipments required" path)
```

***

## Implementation Order

Given that `LINE_ID` doesn't exist in the repo yet, the suggested build-then-test order is:

1. **Add `LINE_ID` column to the active cache TSV schema** — generate it on row creation, persist on every write.
2. **Write Test 2.1 as a manual checklist** (automated later) — run it against current code to confirm resurrection.
3. **Implement tombstone by `LINE_ID`** — filter in `LoadPersistentActiveShipmentRowsLocal`.
4. **Re-run Test 2.1** — should now pass.
5. **Add Tests 1.1–1.5 as automated Phase 6 entries** — these test the already-committed `modTS_Shipments.bas` math and validation guards and can be added to `run_phase6_excel_validation.ps1` immediately. 
6. **Add Tests 2.2–2.4** once the state ledger is implemented.

The Phase 6 tests that already pass (PASSED=21) don't cover any of the seven failure scenarios listed in the problem report, so all six of the above groups would be net-new coverage. 