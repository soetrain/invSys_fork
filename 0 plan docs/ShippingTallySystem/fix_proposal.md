Shipping system proposal (Mermaid)
==================================

Goal: replace ad-hoc shipping macros with a structured, list-object-driven workflow that mirrors ReceivedTally but supports package BOMs and invSys managed items.

```mermaid
flowchart LR
    classDef table fill:#e8f9ff,stroke:#2c7a9b,color:#000,stroke-width:1.2px;
    classDef proc fill:#dde7ff,stroke:#2f4e9c,color:#000,stroke-width:1.2px;
    classDef note fill:#fff8d7,stroke:#b5a542,color:#000,stroke-dasharray:3 3;
    classDef data fill:#dff7df,stroke:#2f6f2f,color:#000;
    classDef log fill:#f7eadb,stroke:#8c6239,color:#000,stroke-dasharray:4 3;

    BUILDER["Package Builder lists\n(PackageHeaderEntry + PackageComponentsEntry)"]:::table
    BOMHDR["ShippingPackages\n(managed items)"]:::table
    BOMDET["PackageRecipes\n(component BOM)"]:::table
    PICKER["Shipping picker form"]:::proc
    TALLY["ShippingTally\n(REF_NUMBER, ITEMS, QUANTITY)"]:::table
    ASSY["ShippingAssembly\n(component view + INV_CHECK)"]:::table
    SHIP["Shipments\n(package aggregation)"]:::table
    CNF["Confirm shipments macro"]:::proc
    INV_USED["invSys.USED (component consumption)"]:::data
    INV_MADE["invSys.MADE (package output)"]:::data
    SLOG["ShippingLog"]:::log

    BUILDER -->|create| BOMHDR
    BUILDER -->|create| BOMDET
    BOMHDR -->|picker source| PICKER -->|add/merge| TALLY
    TALLY -->|explode via BOM| ASSY
    TALLY -->|aggregate packages| SHIP
    SHIP -->|click| CNF
    CNF -->|add qty| INV_USED
    CNF -->|add qty| INV_MADE
    CNF -->|append rows| SLOG

    NOTE1["Package Builder ensures every package has:\n• invSys row (managed item)\n• Component BOM (PackageRecipes)\n• Optional metadata (UOM, default location, carton dims).\nUsers cannot ship undefined packages."]:::note
    NOTE2["ShippingAssembly enforces inventory sufficiency by showing INV_CHECK (TOTAL_INV - needed).\nConfirm refuses to run if any INV_CHECK < 0."]:::note
    NOTE3["Confirm writes capture per-REF log rows for both package + component (traceability for audits)."]:::note

    BUILDER --- NOTE1
    ASSY --- NOTE2
    SLOG --- NOTE3
```

Key deltas vs. legacy
---------------------
1. **Managed packages** – packages exist as real invSys rows, making them searchable, countable, and ship-ready.
2. **BOM-aware tally** – ShippingAssembly automatically explodes BOMs using ShippingBOM so the operator never hand-calculates component usage.
3. **Inventory check before confirm** – INV_CHECK column highlights shortages immediately; confirm enforces sufficiency.
4. **Symmetric logging** – ShippingLog mirrors ReceivedLog but tracks both package (output) and component (input) deltas for reconciliation.
5. **Single-source macros** – All confirm/undo/redo logic lives in `modTS_Shipping`, giving parity with `modTS_Received`.
