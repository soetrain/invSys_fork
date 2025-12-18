Data Contracts (Shipping tables/lists) – Mermaid
===============================================

Flowchart
---------
```mermaid
flowchart TB
    classDef list fill:#e8f9ff,stroke:#2c7a9b,color:#000,stroke-width:1.2px;
    classDef data fill:#dff7df,stroke:#2f6f2f,color:#000;
    classDef log fill:#f7eadb,stroke:#8c6239,color:#000,stroke-dasharray:4 3;
    classDef note fill:#fff8d7,stroke:#b5a542,color:#000,stroke-dasharray:3 3;

    INV_SRC["invSys (catalog source)\nROW, ITEM_CODE, ITEM, UOM, LOCATION, TOTAL_INV, USED, MADE"]:::data
    PKG_ENTRY["Package Builder Entry lists\nPackageHeaderEntry (ITEM meta)\nPackageComponentsEntry (component rows)"]:::list
    BOM_HDR["ShippingPackages (persisted packages)\nPACKAGE_ID, ITEM_CODE, ITEM, UOM, LOCATION, ROW, NOTES"]:::list
    BOM_DET["PackageRecipes (component BOM)\nPACKAGE_ID, COMPONENT_ROW, COMPONENT_ITEM, COMPONENT_QTY"]:::list
    SHIP_TALLY["ShippingTally\nREF_NUMBER, ITEMS (package), QUANTITY"]:::list
    SHIP_ASSY["ShippingAssembly (component explosion)\nROW, ITEM, UOM, TOTAL_INV, COMPONENT_QTY, INV_CHECK"]:::list
    SHIP_AGG["Shipments (package aggregation)\nREF_NUMBER display, PACKAGE_ROW, ITEM, ITEM_CODE, QUANTITY, LOCATION"]:::list
    SHIPPING_LOG["ShippingLog\nREF_NUMBER, PACKAGE_ITEM, PACKAGE_QTY, COMPONENT_ITEM, COMPONENT_QTY, ROW, UOM, LOCATION, SNAPSHOT_ID, ENTRY_DATE"]:::log
    INV_TARGET["invSys (inventory target)\nROWS for components (USED) + packages (MADE)"]:::data

    INV_SRC -->|"prefill builder lists"| PKG_ENTRY
    PKG_ENTRY -->|"Create package"| BOM_HDR
    PKG_ENTRY -->|"Create BOM"| BOM_DET
    BOM_HDR -->|"Package picker source"| SHIP_TALLY
    BOM_DET -->|"Component explosion"| SHIP_ASSY
    SHIP_TALLY -->|"auto aggregate packages"| SHIP_AGG
    SHIP_TALLY -->|"fan-out per component"| SHIP_ASSY
    SHIP_ASSY -->|"Confirm shipments (consume components)"| INV_TARGET
    SHIP_AGG -->|"Confirm shipments (produce packages)"| INV_TARGET
    SHIP_AGG -->|"Confirm shipments"| SHIPPING_LOG

    NOTE1["All package definitions are managed items; they exist as rows in invSys so they can be stocked, searched, and shipped."]:::note
    NOTE2["ShippingAssembly compares needed quantity vs invSys.TOTAL_INV to populate INV_CHECK = TOTAL_INV - COMPONENT_QTY. Validation blocks negatives."]:::note
    NOTE3["ShippingLog keeps per-REF detail: both the package and its underlying component movements for traceability."]:::note

    BOM_HDR --- NOTE1
    SHIP_ASSY --- NOTE2
    SHIPPING_LOG --- NOTE3
```

Block Diagram (relationships & merge rules)
-------------------------------------------
```mermaid
block-beta
columns 4
  INV_SRC["invSys catalog"]
  PKG_HDR["ShippingPackages\n(package headers)"]
  PKG_REC["PackageRecipes\n(component rows)"]
  SHIP_TALLY["ShippingTally\n(REF_NUMBER, ITEMS, QUANTITY)"]
  blockArrowINVPKG<["Lookup invSys rows to prefill header + components"]>(right)
  blockArrowPKGREC<["Package creation writes header + BOM to ShippingBOM"]>(right)
  blockArrowPKGSHIP<["Picker pulls only managed packages (those with PACKAGE_ID + ROW)"]>(right)
  blockArrowSHIPASSY<["Explode BOM × quantity; group by component ROW"]>(right)
  blockArrowSHIPAGG<["Group packages by ROW/ITEM"]>(right)
  blockArrowASSYINV<["Confirm: add COMPONENT_QTY to invSys.USED"]>(down)
  blockArrowAGGINV<["Confirm: add package QUANTITY to invSys.MADE"]>(down)
  blockArrowAGGLOG<["Confirm: append ShippingLog per REF_NUMBER"]>(down)
  SHIP_ASSY["ShippingAssembly\n(ROW, ITEM, COMPONENT_QTY, INV_CHECK)"]
  SHIP_AGG["Shipments\n(PACKAGE_ROW, ITEM, QUANTITY, LOCATION)"]
  SHIPPING_LOG["ShippingLog\n(package + component trace)"]
  INV_TARGET["invSys target rows\ncomponents (USED) + packages (MADE)"]
  style INV_SRC fill:#dff7df,stroke:#2f6f2f,stroke-width:2px
  style PKG_HDR fill:#e8f9ff,stroke:#2c7a9b,stroke-width:2px
  style PKG_REC fill:#e8f9ff,stroke:#2c7a9b,stroke-width:2px
  style SHIP_TALLY fill:#e8f9ff,stroke:#2c7a9b,stroke-width:2px
  style SHIP_ASSY fill:#e8f9ff,stroke:#2c7a9b,stroke-width:2px
  style SHIP_AGG fill:#e8f9ff,stroke:#2c7a9b,stroke-width:2px
  style INV_TARGET fill:#dff7df,stroke:#2f6f2f,stroke-width:2px
  style SHIPPING_LOG fill:#f7eadb,stroke:#8c6239,stroke-width:2px,stroke-dasharray:4 3
```

Entity Relationship View
------------------------
```mermaid
erDiagram
    INV_SYS ||--o{ PACKAGE_BUILDER : "prefill header/components"
    PACKAGE_BUILDER ||--|{ SHIPPING_PACKAGES : "persist managed package"
    PACKAGE_BUILDER ||--|{ PACKAGE_RECIPES : "persist BOM rows"
    SHIPPING_PACKAGES ||--o{ SHIPPING_TALLY : "user selects package to ship"
    PACKAGE_RECIPES ||--o{ SHIPPING_ASSEMBLY : "component explosion (per package)"
    SHIPPING_TALLY ||--o{ SHIPMENTS : "aggregate packages"
    SHIPPING_ASSEMBLY ||--|{ INV_SYS_USED : "writes USED delta"
    SHIPMENTS ||--|{ INV_SYS_MADE : "writes MADE delta"
    SHIPMENTS ||--|{ SHIPPING_LOG : "log per REF_NUMBER"

    INV_SYS {
        string ROW
        string ITEM_CODE
        string ITEM
        string UOM
        string LOCATION
        decimal TOTAL_INV
        decimal USED
        decimal MADE
    }
    SHIPPING_PACKAGES {
        string PACKAGE_ID
        string ITEM_CODE
        string ITEM
        string UOM
        string LOCATION
        string ROW
        string NOTES
    }
    PACKAGE_RECIPES {
        string PACKAGE_ID
        string COMPONENT_ROW
        string COMPONENT_ITEM
        decimal COMPONENT_QTY
        string COMPONENT_UOM
    }
    SHIPPING_TALLY {
        string REF_NUMBER
        string ITEMS
        decimal QUANTITY
    }
    SHIPPING_ASSEMBLY {
        string ROW
        string ITEM
        decimal COMPONENT_QTY
        decimal TOTAL_INV
        decimal INV_CHECK
    }
    SHIPMENTS {
        string PACKAGE_ROW
        string ITEM
        string LOCATION
        decimal QUANTITY
        string REF_LIST
    }
    SHIPPING_LOG {
        string REF_NUMBER
        string PACKAGE_ITEM
        decimal PACKAGE_QTY
        string COMPONENT_ITEM
        decimal COMPONENT_QTY
        string ROW
        string UOM
        string LOCATION
        string SNAPSHOT_ID
        datetime ENTRY_DATE
    }
```
