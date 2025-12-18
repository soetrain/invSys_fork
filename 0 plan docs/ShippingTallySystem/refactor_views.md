Refactor views – Shipping system
================================

View 1 – Module consolidation (Before + After)
----------------------------------------------
```mermaid
C4Component
title BEFORE: scattered shipping logic

Boundary(ShipForm_b, "frmShipPackages") {
  Component(Form, "Legacy form buttons", "UI")
}
Boundary(PkgMod_b, "modTS_Packaging") {
  Component(PkgMix, "Mixed helpers (package + component math)", "mixed")
}
Boundary(ShipMod_b, "modTS_ShipOut") {
  Component(ShipMacros, "Ship macros (confirm/undo)", "mixed")
}
Boundary(LogMod_b, "modTS_Log") {
  Component(LogShip, "Logging helpers", "logging")
}
Boundary(InvMod_b, "modInvSync") {
  Component(InvSync, "invSys synchronizers", "data")
}

Rel(Form, ShipMacros, "calls (confirm/undo)")
Rel(ShipMacros, PkgMix, "uses BOM math")
Rel(ShipMacros, LogShip, "writes log")
Rel(ShipMacros, InvSync, "updates inventory")
Rel(Form, PkgMix, "calls (package creation)")
```

```mermaid
C4Component
title AFTER: consolidated shipping workflow

Boundary(ShipNew_b, "modTS_Shipping") {
  Component(AddPkg, "AddOrMergeFromSearch()", "proc")
  Component(BuildPkg, "CreateOrUpdatePackage()", "proc")
  Component(RebuildAssy, "RebuildShippingAssembly()", "proc")
  Component(RebuildShip, "RebuildShipments()", "proc")
  Component(ConfirmShip, "ConfirmShipments()", "proc")
  Component(AppendLog, "AppendShippingLog()", "proc")
}
Boundary(Undo_b, "modUndoRedo") {
  Component(UndoShip, "MacroUndo()", "proc")
  Component(RedoShip, "MacroRedo()", "proc")
}
Component(Picker, "Shipping picker form", "UI")
Component(BuilderPanels, "Package Builder sheet lists", "UI")

Rel(BuilderPanels, BuildPkg, "writes ShippingBOM + invSys row")
Rel(Picker, AddPkg, "adds/merges package rows")
Rel(AddPkg, RebuildAssy, "rebuild component view")
Rel(AddPkg, RebuildShip, "rebuild package aggregation")
Rel(RebuildAssy, ConfirmShip, "supplies component totals")
Rel(RebuildShip, ConfirmShip, "supplies package totals")
Rel(ConfirmShip, AppendLog, "calls")
Rel(UndoShip, RebuildAssy, "restore staging/posted/log")
Rel(RedoShip, RebuildAssy, "reapply staging/posted/log")
```

View 2 – Runtime sequence (builder + shipping confirm)
------------------------------------------------------
```mermaid
sequenceDiagram
    participant Builder as Package Builder lists
    participant BOM as ShippingBOM
    participant invPkg as invSys (package row)
    participant Picker as Shipping picker form
    participant Tally as ShippingTally
    participant Assembly as ShippingAssembly
    participant ShipAgg as Shipments
    participant Confirm as ConfirmShipments
    participant InvUsed as invSys.USED
    participant InvMade as invSys.MADE
    participant Log as ShippingLog
    participant Undo as MacroUndo
    participant Redo as MacroRedo

    Builder->>BOM: save header + component rows
    Builder->>invPkg: ensure managed package row exists

    Picker->>Tally: add/merge package refs
    Tally->>Assembly: rebuild component requirements
    Tally->>ShipAgg: rebuild package aggregation

    ShipAgg->>Confirm: user clicks Confirm shipments
    Confirm->>Assembly: validate ROW/UOM/INV_CHECK
    Confirm->>InvUsed: add component qty (USED)
    Confirm->>InvMade: add package qty (MADE)
    Confirm->>Log: append package + component rows
    Confirm->>Tally: clear staging (Tally/Assembly/ShipAgg)

    Undo-->>Tally: restore staging
    Undo-->>Assembly: restore components
    Undo-->>ShipAgg: restore packages
    Undo-->>InvUsed: revert USED deltas
    Undo-->>InvMade: revert MADE deltas
    Undo-->>Log: delete inserted log rows

    Redo-->>Tally: reapply staging
    Redo-->>Assembly: reapply components
    Redo-->>ShipAgg: reapply packages
    Redo-->>InvUsed: reapply USED deltas
    Redo-->>InvMade: reapply MADE deltas
    Redo-->>Log: reinsert log rows
```

View 3 – Migration plan (flowchart)
-----------------------------------
```mermaid
flowchart TD
    classDef step fill:#dde7ff,stroke:#2f4e9c,color:#000,stroke-width:1.1px;
    classDef done fill:#dff7df,stroke:#2f6f2f,color:#000;
    classDef warn fill:#fff8d7,stroke:#b5a542,color:#000;

    M1["Catalogue existing shipping macros/forms"]:::step
    M2["Design Package Builder lists (header + BOM)"]:::step
    M3["Implement ShippingTally, ShippingAssembly, Shipments lists"]:::step
    M4["Implement modTS_Shipping (add/merge, rebuild, confirm)"]:::step
    M5["Wire generated buttons (Confirm/Undo/Redo) on ShippingTally"]:::step
    M6["Deprecate legacy modTS_ShipOut + forms"]:::warn
    M7["Remove old modules/forms once parity reached"]:::done

    M1 --> M2 --> M3 --> M4 --> M5 --> M6 --> M7
```
