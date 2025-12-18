List-object-based shipping flow – Mermaid views
===============================================

This mirrors the received system diagrams but emphasizes package creation plus shipping tally.

View A – Package Builder + Shipping staging
-------------------------------------------
```mermaid
flowchart LR
    classDef button fill:#dde7ff,stroke:#2f4e9c,color:#000,stroke-width:1.3px;
    classDef list fill:#e8f9ff,stroke:#2c7a9b,color:#000,stroke-width:1.3px;
    classDef note fill:#fff8d7,stroke:#b5a542,color:#000,stroke-dasharray:3 3;

    BUILDER_HDR["PackageHeaderEntry<br/>(future package row)"]:::list
    BUILDER_BOM["PackageComponentsEntry<br/>(component list)"]:::list
    BTN_CREATE["Create/Update package<br/>(writes ShippingBOM + invSys row)"]:::button
    SHIPPING_PACKAGES["ShippingPackages<br/>(managed packages)"]:::list
    SHIPPING_RECIPES["PackageRecipes<br/>(component BOM)"]:::list
    PICKER["Shipping picker form"]:::button
    TALLY["ShippingTally<br/>(REF_NUMBER, ITEMS, QUANTITY)"]:::list
    ASSY["ShippingAssembly<br/>(component explosion + INV_CHECK)"]:::list
    SHIP["Shipments<br/>(package aggregate)"]:::list

    BUILDER_HDR --> BTN_CREATE
    BUILDER_BOM --> BTN_CREATE
    BTN_CREATE --> SHIPPING_PACKAGES
    BTN_CREATE --> SHIPPING_RECIPES
    SHIPPING_PACKAGES -->|picker source| PICKER -->|add/merge package| TALLY
    TALLY -->|"explode BOM"| ASSY
    TALLY -->|"group packages"| SHIP

    NOTE1["Package Builder enforces invSys-aligned schema (ITEM_CODE/ITEM/UOM/LOCATION + metadata).\nEach package is a managed item with its own ROW."]:::note
    NOTE2["ShippingAssembly calculates component demand = Quantity × BOM; INV_CHECK = TOTAL_INV - demand."]:::note

    SHIPPING_PACKAGES --- NOTE1
    ASSY --- NOTE2
```

View B – Confirm / Undo / Redo / Logging
----------------------------------------
```mermaid
flowchart LR
    classDef button fill:#dde7ff,stroke:#2f4e9c,color:#000,stroke-width:1.3px;
    classDef list fill:#e8f9ff,stroke:#2c7a9b,color:#000,stroke-width:1.3px;
    classDef data fill:#dff7df,stroke:#2f6f2f,color:#000;
    classDef log fill:#f7eadb,stroke:#8c6239,color:#000,stroke-dasharray:4 3;
    classDef note fill:#fff8d7,stroke:#b5a542,color:#000,stroke-dasharray:3 3;

    SHIP["Shipments"]:::list
    ASSY["ShippingAssembly"]:::list
    CNF["Confirm shipments"]:::button
    INV_USED["invSys.USED"]:::data
    INV_MADE["invSys.MADE"]:::data
    SLOG["ShippingLog"]:::log
    UNDO["MacroUndo"]:::button
    REDO["MacroRedo"]:::button

    ASSY -->|"component totals"| CNF
    SHIP -->|"package totals"| CNF

    CNF -->|add component qty| INV_USED
    CNF -->|add package qty| INV_MADE
    CNF -->|append per REF rows| SLOG
    CNF -->|clear staging| SHIP
    CNF -->|clear components| ASSY

    UNDO -.->|restore rows| SHIP
    UNDO -.->|restore components| ASSY
    UNDO -.->|revert USED| INV_USED
    UNDO -.->|revert MADE| INV_MADE
    UNDO -.->|delete log| SLOG

    REDO -.->|reapply rows| SHIP
    REDO -.->|reapply components| ASSY
    REDO -.->|reapply USED| INV_USED
    REDO -.->|reapply MADE| INV_MADE
    REDO -.->|reinsert log| SLOG

    NOTE3["Confirm fails fast if INV_CHECK < 0 or package/BOM missing; nothing is written until all checks pass."]:::note
    NOTE4["Undo/Redo operate on the macro snapshot (ShippingTally + derived lists + invSys deltas + ShippingLog rows)."]:::note

    CNF --- NOTE3
    UNDO --- NOTE4
```
