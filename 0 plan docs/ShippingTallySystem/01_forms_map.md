Forms Map / Controls Interaction – Shipping
==========================================

```mermaid
flowchart LR
    classDef form fill:#e7e7ff,stroke:#4d4d8c,color:#000,stroke-width:1.2px;
    classDef genbtn fill:#dde7ff,stroke:#2f4e9c,color:#000,stroke-width:1.2px;
    classDef sheet fill:#e8f9ff,stroke:#2c7a9b,color:#000,stroke-width:1.2px;
    classDef action fill:#fff8d7,stroke:#b5a542,color:#000,stroke-dasharray:3 3;

    BUILDER["Package Builder sheet\n(ListObjects: PackageHeaderEntry, PackageComponentsEntry)"]:::sheet
    BTN_CREATE["Button: Add package to ShippingBOM + invSys row"]:::genbtn
    BOM["ShippingBOM sheet\n(ListObjects: ShippingPackages, PackageRecipes)"]:::sheet

    PICKER["Shipping picker form\n(Managed packages selector)"]:::form
    TALLY["ShippingTally list object\n(REF_NUMBER, ITEMS, QUANTITY)"]:::sheet
    ASSY["ShippingAssembly list object\n(component explosion + inventory check)"]:::sheet
    SHIP["Shipments list object\n(package aggregation)"]:::sheet

    BTN_CONFIRM["Generated button: Confirm shipments"]:::genbtn
    BTN_UNDO["Generated button: Undo shipping macro"]:::genbtn
    BTN_REDO["Generated button: Redo shipping macro"]:::genbtn

    ACT_BUILD["Capture package header + recipe\n(derives item structure from invSys)"]:::action
    ACT_TALLY["Fast entry of packages to ship\n(tab/enter launches picker)"]:::action
    ACT_CONFIRM["Validate inventory, add USED/MADE, log"]:::action
    ACT_UNDO["Undo macro (tally, assembly, invSys, log)"]:::action
    ACT_REDO["Redo macro (same scope)"]:::action
    NOTE_BTN["Generated buttons on ShippingTally: Confirm shipments, Undo, Redo.\nOnly create when missing; never duplicate."]:::action

    BUILDER -->|user enters package + components| ACT_BUILD --> BTN_CREATE --> BOM
    BTN_CREATE -->|writes managed package row| BOM

    PICKER -->|choose package| ACT_TALLY --> TALLY
    TALLY -->|auto rebuild| SHIP
    TALLY -->|explode BOM| ASSY
    SHIP --> BTN_CONFIRM
    BTN_CONFIRM --> ACT_CONFIRM
    BTN_UNDO --> ACT_UNDO
    BTN_REDO --> ACT_REDO

    SHIP --- NOTE_BTN
```
