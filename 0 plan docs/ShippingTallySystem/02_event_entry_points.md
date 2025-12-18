Event Entry Points – Shipping
============================

Two primary entry streams exist: (1) package creation (builder) and (2) package tally/confirm. The diagram keeps layout noise out and focuses on who triggers what.

```mermaid
sequenceDiagram
    participant Builder as Package Builder sheet
    participant BOM as ShippingBOM lists
    participant invPkg as invSys row (managed package)
    participant Picker as Shipping picker form
    participant Tally as ShippingTally (list)
    participant Assembly as ShippingAssembly (component view)
    participant ShipAgg as Shipments (package aggregation)
    participant Confirm as Confirm shipments macro
    participant Used as invSys.USED
    participant Made as invSys.MADE
    participant Log as ShippingLog
    participant Undo as Undo shipping macro
    participant Redo as Redo shipping macro

    Builder->>BOM: capture package header + recipe
    Builder->>invPkg: create/update managed package row (derives from invSys schema)

    Picker->>Tally: add/merge package (sum qty, concat refs)
    Tally->>Assembly: rebuild component requirements
    Tally->>ShipAgg: rebuild package aggregation

    ShipAgg->>Confirm: click Confirm shipments
    alt validation OK + inventory available
        Confirm-->>Used: add needed quantities (per component ROW) to USED
        Confirm-->>Made: add package quantities (per package ROW) to MADE
        Confirm-->>Log: append ShippingLog rows per REF_NUMBER
        Confirm-->>Tally: clear staging (ShippingTally/Assembly/Shipments)
    else validation fails or inventory short
        Confirm-->>Confirm: show error; no writes
    end
    Undo-->>Tally: restore ShippingTally rows
    Undo-->>Assembly: restore ShippingAssembly
    Undo-->>ShipAgg: restore Shipments
    Undo-->>Used: revert USED deltas
    Undo-->>Made: revert MADE deltas
    Undo-->>Log: delete ShippingLog rows created by last confirm

    Redo-->>Tally: reapply last undone tally rows
    Redo-->>Assembly: rebuild component view
    Redo-->>ShipAgg: rebuild package aggregation
    Redo-->>Used: reapply USED deltas
    Redo-->>Made: reapply MADE deltas
    Redo-->>Log: reinsert ShippingLog rows
```
