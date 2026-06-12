The architecture is now clear. `modTS_Shipments.bas` at 117 KB is the central file — it's huge, and the combination of a giant module driving ribbon callbacks + `Worksheet_Change` handlers + Table writes is the stability trap.  Here is a direct, honest assessment.

***

## Short Answer

**Yes, a UserForm-centered architecture will be dramatically more stable and performant than the current ribbon-buttons-on-Tables approach** — for this specific use case. But the reason is not "UserForms are faster." The reason is that a UserForm forces a clean separation of concerns that the ribbon-on-sheet model inherently resists, and that separation is what eliminates the cascading event/refresh/open cycles that are killing you.

***

## Why the Current Model Keeps Regressing

The ribbon-on-Tables architecture creates a structural trap that no amount of patching fully escapes. Every fix introduces a new interaction surface. 

**The cascade chain that fires on every box selection:**

```
User clicks cell in BoxBOM table
  → Worksheet_SelectionChange fires
    → cDynItemSearch.ShowForShippingComponentCell
      → LoadCanonicalManagedInventoryItems
        → modInventoryDomainBridge.ResolveInventoryWorkbookBridge
          → Workbooks.Open (if not cached)
            → Excel window flash
  → User picks item → CommitSelection
    → callbackCell.Value = x        → Worksheet_Change fires
    → callbackCell.ListObject write → Worksheet_Change fires again
      → version autofill logic       → more cell writes
        → Worksheet_Change fires again (re-entrant)
          → Application.EnableEvents = False needed
            → but someone forgot it somewhere
              → regression
```

`modTS_Shipments.bas` being 117 KB means this logic has grown organically inside one module, and the event handler / quiet-ui guard boundaries are now inconsistent.  Codex keeps "gassing out" because it is patching a system where every write can trigger another write — the state machine is implicit and lives in Excel's event system.

***

## The UserForm Model — What Changes Structurally

A UserForm does not replace ribbon buttons — you keep the ribbon. What changes is **where state lives and where writes happen**.

### Current model (broken loop):
```
Sheet ←→ Worksheet_Change ←→ modTS_Shipments ←→ Workbooks.Open
  ↑_____________writes_____________________________|
```

### UserForm model (clean pipeline):
```
Ribbon button → opens frmBoxBuilder (UserForm, modeless)
  frmBoxBuilder loads ALL data into memory at open time (one I/O burst)
  User makes all selections inside the form (zero sheet writes)
  User clicks "Save Box" → ONE batch write to the sheet
  Form closes → Worksheet_Change fires exactly once, intentionally
```

The UserForm is the **state container**. The sheet becomes a **display/storage layer**, not an interaction layer. `Worksheet_Change` fires once per commit, not on every field.

***

## Concrete Architecture

### The UserForm: `frmBoxBuilder`

```
┌─────────────────────────────────────────────┐
│  Box Builder                           [X]  │
├─────────────────────────────────────────────┤
│  Box Name: [_______________]  Ver: [auto]   │
│  Box Code: [_______________]                │
├─────────────────────────────────────────────┤
│  BOM Components                             │
│  ┌─────────────────────────────────────┐   │
│  │ ROW  │ ITEM        │ QTY │ UOM      │   │
│  │  12  │ Kraft Paper │  2  │ sheets   │   │
│  │  34  │ Tape        │  1  │ roll     │   │
│  └─────────────────────────────────────┘   │
│  [+ Add Component]  [- Remove]             │
├─────────────────────────────────────────────┤
│           [Cancel]        [Save Box]        │
└─────────────────────────────────────────────┘
```

The ListBox inside `frmBoxBuilder` holds the BOM in memory. No sheet writes until "Save Box."

### Module structure

```
modShipping_BoxUI.bas          ← thin, <200 lines
  ShowBoxBuilder(targetRow)   ← opens frmBoxBuilder, passes context
  CommitBoxBuilderResult(...)  ← called by form on Save, does ONE quiet batch write

frmBoxBuilder.frm              ← all interaction lives here
  Private mInventoryArr()     ← inventory loaded once at Initialize
  Private mBomRows()          ← current BOM state
  LoadInventory               ← one call to LoadCanonicalManagedInventoryItems
  cmdAddComponent_Click       ← opens cDynItemSearch picker, adds to mBomRows
  cmdSave_Click               ← calls modShipping_BoxUI.CommitBoxBuilderResult
```

### The single batch write

```vba
' modShipping_BoxUI.CommitBoxBuilderResult
Public Sub CommitBoxBuilderResult(ByVal bomRows() As Variant, ByVal boxName As String)
    Dim wb As Workbook
    Dim loBoxBOM As ListObject
    ' ... resolve table

    Application.ScreenUpdating = False
    Application.Calculation    = xlCalculationManual
    Application.EnableEvents   = False          ' ← fires Worksheet_Change ZERO times during write

    On Error GoTo Restore
        ClearExistingBoxBOMRows loBoxBOM, boxName
        WriteBoxBOMRows loBoxBOM, bomRows, boxName
        ' version autofill happens HERE, inside the quiet block, not via Worksheet_Change

Restore:
    Application.EnableEvents   = True
    Application.Calculation    = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
```

`EnableEvents = False` during the entire batch means `Worksheet_Change` fires **zero times** during the write. It fires once when the function returns and Excel resumes events, and by then all writes are done.

***

## Performance Comparison

| Factor | Ribbon + Tables (current) | UserForm model |
|---|---|---|
| `Workbooks.Open` calls per box select | 1–3 (inventory, auth, inbox) | 0 (loaded at form open, cached in `mInventoryArr`) |
| `Worksheet_Change` fires per box commit | 5–20+ (each cell write) | 0 during write, 1 after |
| `ScreenUpdating` discipline | Patched per-function, inconsistent | 1 wrapper in `CommitBoxBuilderResult`, always fires |
| Re-entrancy risk | High (`Worksheet_Change` can re-trigger) | None (form is the only writer) |
| Codex patch surface | 117 KB monolith, any edit can break any other path | Form ↔ one commit function, isolated |
| Perceived speed | Blink + delay between each selection | Form stays open; all selections instant inside UI |

***

## Migration Path — Incremental, Not a Rewrite

You do not need to gut `modTS_Shipments.bas` in one shot. The safe migration is:

1. **Extract `CommitBoxBuilderResult`** from inside `modTS_Shipments` into `modShipping_BoxUI`. This is a copy-and-fix, not a delete. (~1 day)

2. **Build `frmBoxBuilder`** as a new UserForm. Wire the ribbon "Box Builder" button to `modShipping_BoxUI.ShowBoxBuilder`. The old sheet-click path still works in parallel. (~2 days)

3. **Move `LoadCanonicalManagedInventoryItems`** call into `frmBoxBuilder.UserForm_Initialize`. Remove it from the `SelectionChange` path. (~1 hour)

4. **Disable the `SelectionChange`→picker trigger for BoxBOM** once the form is the canonical entry point. The picker still exists — it now opens *from the form's Add Component button*, not from a cell click. (~1 hour)

5. **Delete dead code in `modTS_Shipments`** after the form path is confirmed stable. This shrinks the 117 KB monolith over time rather than all at once.

The result: no more workbook open/close on cell selection, no re-entrant `Worksheet_Change`, one quiet batch write per box save, and a UI that is faster because all filtering and selection happens in memory inside the form before any data ever touches the sheet.