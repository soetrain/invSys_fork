# invSys v2 Architecture Diagrams (Mermaid)

## 1) Directory structure (repo)

```mermaid
flowchart TB
  subgraph DS["directory-structure (repo)"]
    R["invSys/"]

    R --> SRC["src/"]
    R --> DATA["data/"]
    R --> ASSETS["assets/"]
    R --> DEPLOY["deploy/"]
    R --> TOOLS["tools/"]
    R --> DOCS["docs/"]

    %% src targets
    SRC --> CORE["Core/"]
    SRC --> INVDOM["InventoryDomain/"]
    SRC --> DESDOM["DesignsDomain/"]
    SRC --> RECV["Receiving/"]
    SRC --> SHIP["Shipping/"]
    SRC --> PROD["Production/"]
    SRC --> ADMIN["Admin/"]

    CORE --> CORE_M["Modules/"]
    CORE --> CORE_C["ClassModules/"]
    CORE --> CORE_F["Forms/ (shared only)"]
    CORE --> CORE_R["Ribbon/"]
    CORE --> CORE_CFG["Config/"]

    INVDOM --> INVDOM_M["Modules/"]
    INVDOM --> INVDOM_C["ClassModules/"]

    DESDOM --> DESDOM_M["Modules/"]
    DESDOM --> DESDOM_C["ClassModules/"]

    RECV --> RECV_M["Modules/"]
    RECV --> RECV_F["Forms/"]
    RECV --> RECV_R["Ribbon/"]

    SHIP --> SHIP_M["Modules/"]
    SHIP --> SHIP_F["Forms/"]
    SHIP --> SHIP_R["Ribbon/"]

    PROD --> PROD_M["Modules/"]
    PROD --> PROD_F["Forms/"]
    PROD --> PROD_R["Ribbon/"]

    ADMIN --> ADMIN_M["Modules/"]
    ADMIN --> ADMIN_F["Forms/"]
    ADMIN --> ADMIN_R["Ribbon/"]

    %% data domain schemas + samples
    DATA --> INV_D["InventoryDomain/"]
    DATA --> DES_D["DesignsDomain/"]

    INV_D --> INV_SCHEMA["schema/"]
    INV_D --> INV_SAMPLES["samples/"]

    DES_D --> DES_SCHEMA["schema/"]
    DES_D --> DES_SAMPLES["samples/"]

    %% assets
    ASSETS --> AS_RIB["ribbon/"]
    AS_RIB --> AS_IMG["images/"]
    AS_RIB --> AS_XML["xml/"]

    %% deploy
    DEPLOY --> DEP_SP["sharepoint/"]
    DEPLOY --> DEP_LAN["lan/"]

    %% tools
    TOOLS --> T_BUILD["build/"]
    TOOLS --> T_EXPORT["export-import/"]
    TOOLS --> T_RIB["ribbon-editing/"]

    %% docs
    DOCS --> D_ARCH["architecture/"]
    DOCS --> D_WORK["workflows/"]
    DOCS --> D_ROLES["roles-permissions/"]
    DOCS --> D_REL["release-notes/"]
  end
```

---

## 2) File diagram (SharePoint-synced, supports out-of-state users)

```mermaid
flowchart TB
  subgraph SP["SharePoint Document Library (single source of distribution)"]
    ROOT["invSys/"]
    ROOT --> ADDINS["Addins/"]
    ROOT --> DATASTORE["Data/"]
    ROOT --> USERWBS["UserWorkbooks/"]

    %% add-ins (synced to every station)
    ADDINS --> ACORE["invSys.Core.xlam"]
    ADDINS --> AINVD["invSys.Inventory.Domain.xlam"]
    ADDINS --> ADESD["invSys.Designs.Domain.xlam"]
    ADDINS --> ARECV["invSys.Receiving.xlam"]
    ADDINS --> ASHIP["invSys.Shipping.xlam"]
    ADDINS --> APROD["invSys.Production.xlam"]
    ADDINS --> AADM["invSys.Admin.xlam"]

    %% authoritative data stores (ideally single-writer)
    DATASTORE --> INVF["Data/Inventory/"]
    DATASTORE --> DESF["Data/Designs/"]
    INVF --> INVX["invSys.Data.Inventory.xlsb"]
    DESF --> DESX["invSys.Data.Designs.xlsb"]

    %% inbox queue (multi-writer friendly because each station has its own file)
    INVF --> INBOX["Data/Inventory/Inbox/"]
    INBOX --> IB1["invSys.Inbox.ShippingA.xlsb"]
    INBOX --> IB2["invSys.Inbox.ShippingB.xlsb"]
    INBOX --> IB3["invSys.Inbox.Receiving1.xlsb"]
    INBOX --> IB4["invSys.Inbox.Production1.xlsb"]

    %% user workbooks (optional)
    USERWBS --> URECV["Receiving/invSys.Receiving.Job.xlsm"]
    USERWBS --> USHIP["Shipping/invSys.Shipping.Job.xlsm"]
    USERWBS --> UPROD["Production/invSys.Production.Job.xlsm"]
    USERWBS --> UADM["Admin/invSys.Admin.Console.xlsm"]
  end

  %% stations (can be on LAN in WA or remote out-of-state)
  subgraph WA["Washington HQ station(s)"]
    WA_PC["Excel workstation(s)"]
  end

  subgraph REM["Remote station (e.g., Colorado)"]
    REM_PC["Excel workstation"]
  end

  SP --- WA_PC
  SP --- REM_PC

  note1["Key idea: Remote users do NOT write directly into invSys.Data.Inventory.xlsb.
They append events into their own invSys.Inbox.<Station>.xlsb, which syncs via SharePoint."]
```

---

## 3) Inventory workflow combinations (explicit off-local-network flow, role inboxes, lock order, logs)

```mermaid
flowchart TB
  %% =========================
  %% Places
  %% =========================
  subgraph REM["Remote user (out-of-state)"]
    RWB["User workbook (.xlsm)
(optional UI/staging)"]
    RROLE["Role XLAM
(Receiving/Shipping/Production)"]
    RCORE["invSys.Core.xlam"]
  end

  subgraph SP["SharePoint sync (convenience layer)"]
    INB_RECV["Inbox: Receiving (xlsb)
invSys.Inbox.Receiving.<Station>.xlsb"]
    INB_SHIP["Inbox: Shipping (xlsb)
invSys.Inbox.Shipping.<Station>.xlsb"]
    INB_PROD["Inbox: Production (xlsb)
invSys.Inbox.Production.<Station>.xlsb"]
  end

  subgraph WA["Washington HQ (authoritative processing)"]
    PROC["Processor station
(Admin XLAM job or scheduled macro)"]

    ACORE["invSys.Core.xlam
(authority: roles + orchestration)"]
    GATE{{"Core-owned capability check
Core.CanPerform(CAPABILITY)"}}
    REJECT["REJECT + LOG
(no write)"]

    INVDOM["invSys.Inventory.Domain.xlam
(domain rules + writes)"]
    DESDOM["invSys.Designs.Domain.xlam
(domain rules + writes)"]

    INVDB["invSys.Data.Inventory.xlsb
(authoritative)"]
    DESDB["invSys.Data.Designs.xlsb
(authoritative)"]

    INVLOG["InventoryLog (table)
(in Inventory.xlsb)"]
    PRUNS["ProductionRuns (table)
(in Designs.xlsb)"]
  end

  %% =========================
  %% Remote: create an event (never touches authoritative XLSB)
  %% =========================
  RWB --> RROLE --> RCORE --> GATE
  GATE -- "NO" --> REJECT

  %% When allowed, remote writes ONLY to role-specific inbox file
  GATE -- "CAP: RECEIVE_POST" --> INB_RECV
  GATE -- "CAP: SHIP_POST" --> INB_SHIP
  GATE -- "CAP: PROD_POST" --> INB_PROD

  %% =========================
  %% SharePoint sync carries inbox files to HQ
  %% =========================
  INB_RECV --> PROC
  INB_SHIP --> PROC
  INB_PROD --> PROC

  %% =========================
  %% HQ: processor applies events with explicit lock order
  %% =========================
  PROC --> ACORE --> GATE

  %% Processor capability (separate from posting)
  GATE -- "CAP: INBOX_PROCESS" --> L1["Lock order:
1) Inventory domain
2) Designs domain (only if needed)"]

  %% Apply Receiving + Shipping (Inventory only)
  L1 --> INVDOM
  INVDOM -->|"Open + Lock Inventory.xlsb"| INVDB
  INVDOM -->|"Apply RECEIVE events"| INVDB
  INVDOM -->|"Apply SHIP events"| INVDB
  INVDOM -->|"Write InventoryLog"| INVLOG

  %% Apply Production (Inventory + Designs)
  INB_PROD -. "needs BOM/version" .-> DESDOM
  L1 --> DESDOM

  %% Designs lock happens after Inventory lock (only when processing production)
  DESDOM -->|"Open + Lock Designs.xlsb"| DESDB
  DESDOM -->|"Read BOM/version"| DESDB

  %% Production consumption/outputs go to Inventory
  DESDOM -->|"Resolved BOM lines"| INVDOM
  INVDOM -->|"Consume parts + add outputs"| INVDB
  INVDOM -->|"Write InventoryLog"| INVLOG

  %% Production run history goes to Designs
  DESDOM -->|"Write ProductionRuns"| PRUNS

  %% =========================
  %% Notes
  %% =========================
  N1["Why this avoids file-lock conflicts:
- Remote stations only append to their own role Inbox file
- Only the HQ Processor writes the authoritative .xlsb files
- Lock order prevents deadlocks (Inventory then Designs)
- Logs land in authoritative stores (InventoryLog in Inventory, ProductionRuns in Designs)"]

  %% tidy (non-executing link for clarity)
  INVLOG --> INVDB
  PRUNS --> DESDB
```

### Core-owned capability check (important point)

* Role XLAMs may show UI and collect inputs.
* **Only Core** authorizes actions (`Core.CanPerform`).
* Remote users create **events** in their role-specific inbox file.
* The HQ processor applies events to authoritative `.xlsb` with domain XLAM rules.
* **Lock order** during processing: **Inventory first**, then **Designs only if needed**.
* **InventoryLog** is written in `invSys.Data.Inventory.xlsb`.
* **ProductionRuns** is written in `invSys.Data.Designs.xlsb`.

---

## 4) Multi-warehouse topology (LAN-first per warehouse + cloud convenience + global visibility)

```mermaid
flowchart LR
  %% =========================
  %% Cloud convenience layer
  %% =========================
  subgraph SP["SharePoint (convenience sync + distribution)"]
    SP_Addins["/Addins (xlam distribution)"]
    SP_Events["/Events (warehouse outbox uploads)"]
    SP_Global["/Global (read models)"]
    GlobalSnap["invSys.Global.InventorySnapshot.xlsb"]
    SP_Global --> GlobalSnap
  end

  %% =========================
  %% HQ
  %% =========================
  subgraph HQ["HQ (WA)"]
    HQ_Admin["Admin/Processor station"]
    HQ_Core["invSys.Core.xlam"]
    HQ_Agg["Aggregator job
(build global snapshot)"]
  end

  %% =========================
  %% Warehouse 1
  %% =========================
  subgraph WH1["Warehouse WH1 (LAN-first)"]
    WH1_Core["invSys.Core.xlam"]
    WH1_InvDom["invSys.Inventory.Domain.xlam"]
    WH1_DesDom["invSys.Designs.Domain.xlam (optional)"]

    WH1_InvDB["WH1.invSys.Data.Inventory.xlsb"]
    WH1_DesDB["WH1.invSys.Data.Designs.xlsb (optional)"]

    WH1_Inbox["WH1 Inbox files (per station)
(append-only)"]
    WH1_Proc["WH1 Processor
(applies inbox locally)"]

    WH1_Outbox["WH1 Outbox
InventoryEvents + RunEvents
(append-only export)"]
  end

  %% =========================
  %% Warehouse 2
  %% =========================
  subgraph WH2["Warehouse WH2 (LAN-first)"]
    WH2_Core["invSys.Core.xlam"]
    WH2_InvDom["invSys.Inventory.Domain.xlam"]
    WH2_DesDom["invSys.Designs.Domain.xlam (optional)"]

    WH2_InvDB["WH2.invSys.Data.Inventory.xlsb"]
    WH2_DesDB["WH2.invSys.Data.Designs.xlsb (optional)"]

    WH2_Inbox["WH2 Inbox files (per station)
(append-only)"]
    WH2_Proc["WH2 Processor
(applies inbox locally)"]

    WH2_Outbox["WH2 Outbox
InventoryEvents + RunEvents
(append-only export)"]
  end

  %% =========================
  %% Remote sales UI
  %% =========================
  subgraph SALES["Remote Sales (small screens)"]
    Streamlit["Streamlit UI
(read-mostly)"]
  end

  %% =========================
  %% Local warehouse flows (offline-capable)
  %% =========================
  WH1_Inbox --> WH1_Proc --> WH1_Core --> WH1_InvDom --> WH1_InvDB
  WH1_Core --> WH1_DesDom --> WH1_DesDB
  WH1_Proc --> WH1_Outbox

  WH2_Inbox --> WH2_Proc --> WH2_Core --> WH2_InvDom --> WH2_InvDB
  WH2_Core --> WH2_DesDom --> WH2_DesDB
  WH2_Proc --> WH2_Outbox

  %% =========================
  %% Sync when internet is available
  %% =========================
  WH1_Outbox --> SP_Events
  WH2_Outbox --> SP_Events

  %% =========================
  %% HQ aggregation and distribution
  %% =========================
  SP_Addins --- HQ_Admin
  SP_Events --> HQ_Agg --> GlobalSnap

  %% Sales reads global snapshot (no direct writes)
  GlobalSnap --> Streamlit

  %% Add-ins distributed to warehouses via SharePoint
  SP_Addins --- WH1_Core
  SP_Addins --- WH2_Core

  %% Optional note: transfers are events between warehouses
  Xfer["Transfers: request/fulfill/receive
(events across WH1/WH2)"]
  Streamlit -. "creates request (optional)" .-> Xfer
  Xfer -. "synced events" .-> SP_Events
```

**Intent of this topology**

* Each warehouse is **LAN-first**: it can receive/ship/produce even if the internet is down.
* Each warehouse writes only to its **local authoritative** `.xlsb` stores.
* Warehouses periodically publish **outbox events** to SharePoint when internet returns.
* HQ builds a **global snapshot** for cross-warehouse visibility.
* Remote Sales uses **Streamlit read-mostly**, fed from the global snapshot (no direct stock writes).

---

## 5) Test topology (unit + integration + load simulation)

```mermaid
flowchart TB
  %% =========================
  %% Code under test
  %% =========================
  subgraph CUT["Code under test"]
    Core["invSys.Core.xlam"]
    InvDom["invSys.Inventory.Domain.xlam"]
    DesDom["invSys.Designs.Domain.xlam"]
    Roles["Role XLAMs
(Receiving/Shipping/Production/Admin)"]
  end

  %% =========================
  %% Test types
  %% =========================
  subgraph UT["Unit tests (fast)"]
    UT_Core["Core tests
capability matrix
EventID idempotency
lock-order rules"]
    UT_Inv["Inventory domain tests
math + invariants
log row formation"]
    UT_Des["Designs domain tests
BOM versioning
production run formatting"]
  end

  subgraph IT["Integration tests (Excel automation)"]
    Harness["Test harness
(Python xlwings or VBA runner)"]
    TestData["Synthetic test data
items + BOMs + warehouses"]
  end

  subgraph LS["Load simulation"]
    Gen["Generate N inbox files
(1..100 stations/users)"]
    Proc["Run processor
INBOX_PROCESS"]
    Check["Assertions
counts, sums, logs
no negatives
EventID applied once"]
    Perf["Perf metrics
rows/sec, peak time
lock durations"]
  end

  %% =========================
  %% Data fixtures
  %% =========================
  subgraph FIX["Test fixtures (.xlsb)"]
    InvDB["Test invSys.Data.Inventory.xlsb"]
    DesDB["Test invSys.Data.Designs.xlsb"]
    Inboxes["Test Inbox folder
invSys.Inbox.*.xlsb"]
  end

  %% Wiring
  UT_Core --> Core
  UT_Inv --> InvDom
  UT_Des --> DesDom

  Harness --> TestData
  Harness --> InvDB
  Harness --> DesDB
  Harness --> Inboxes

  Gen --> Inboxes
  Proc --> Core
  Proc --> InvDom
  Proc --> DesDom

  Proc --> InvDB
  Proc --> DesDB
  Check --> InvDB
  Check --> DesDB
  Perf --> Proc

  %% Role XLAM paths (optional UI tests)
  Harness -. "optional UI smoke" .-> Roles
```

### Release 1 test priorities

* **Correctness invariants**: no negative on-hand, EventID applied once, logs match deltas.
* **Capability enforcement**: cannot post/processing without Core permission.
* **Processor determinism**: same inbox set => same authoritative results.
* **Performance**: simulate 100 inboxes with realistic row counts and measure processing throughput.

---

## 6) XLAM build spec (tables, controls, forms, and pseudocode)

### 6.1 What the XLAMs must contain vs. generate

**Contained in XLAMs (code assets)**

* RibbonX tabs/groups/buttons + callbacks.
* Shared modules/classes (Core) and domain modules/classes (InventoryDomain, DesignsDomain).
* UserForms for role workflows (Receiving/Shipping/Production/Admin).
* A small configuration layer:

  * default SharePoint sync root
  * default warehouse ID
  * station ID
  * feature flags (offline-only, remote-only, etc.)

**Generated/ensured by XLAM (data assets)**

* If a required data workbook/table is missing, Admin XLAM can generate it from schema defaults.
* If an inbox/outbox workbook is missing for a station, Role XLAM can generate it.

---

### 6.2 Required ListObjects (tables)

#### Inventory store (per warehouse)

**Workbook:** `WHx.invSys.Data.Inventory.xlsb`

* `tblItems`

  * Item master (SKU, description, UOM, category, active)
* `tblOnHand`

  * Quantities by `WarehouseID + Location + SKU`
* `tblLocations`

  * Warehouse locations (bin/rack/zone)
* `tblInventoryLog`

  * Authoritative transaction ledger (append-only)
* `tblWarehouseConfig`

  * Warehouse metadata (WarehouseID, timezone, default location)
* `tblProcessingState`

  * Processor state (last processed EventID per inbox, last run times)

#### Designs store (optional per warehouse; always at HQ if used)

**Workbook:** `WHx.invSys.Data.Designs.xlsb`

* `tblDesignHeader`

  * Design/recipe metadata (DesignID, Name, Version, Status, UOM)
* `tblBOMLines`

  * BOM lines (DesignID, Version, ComponentSKU, QtyPerUnit, ScrapFactor)
* `tblProductionRuns`

  * Run header (RunID, DesignID, Version, QtyPlanned, QtyActual, Status, WarehouseID)
* `tblProductionRunLines`

  * Consumption/output lines (RunID, SKU, Qty, LineType: Consume/Produce)

#### Inbox workbooks (per station; SharePoint synced)

**Workbook:** `invSys.Inbox.<Role>.<Station>.xlsb`

* Receiving inbox: `tblInboxReceive`
* Shipping inbox: `tblInboxShip`
* Production inbox: `tblInboxProd`

Each inbox table is append-only from the station’s perspective.

#### Outbox workbooks (per warehouse; SharePoint synced)

**Workbook:** `WHx.invSys.Outbox.Events.xlsb`

* `tblOutboxEvents`

  * Warehouse publishes applied events for HQ aggregation.

---

### 6.3 Required Controls

#### RibbonX (per XLAM)

**Core Ribbon (optional shared tab)**

* Status group

  * `btnOpenConsole`
  * `btnShowConfig`
  * `btnHealthCheck`

**Receiving Ribbon**

* `btnReceiveOpenForm`
* `btnReceivePostToInbox`
* `btnReceiveViewInbox`

**Shipping Ribbon**

* `btnShipOpenForm`
* `btnShipPostToInbox`
* `btnShipViewInbox`

**Production Ribbon**

* `btnProdOpenForm`
* `btnProdPostToInbox`
* `btnProdViewRuns`

**Admin Ribbon**

* `btnAdminRunProcessor`
* `btnAdminReprocessErrors`
* `btnAdminBuildWarehouse`
* `btnAdminBuildStationInbox`
* `btnAdminBuildGlobalSnapshot`

#### Worksheet controls (optional; for job workbooks)

* “Open Form” buttons near the top of each job sheet.
* A status label cell (e.g., named range `rngStatus`) for non-modal messages.

---

### 6.4 Required UserForms

#### frmReceive (Receiving)

* `txtSearch` (TextBox)
* `lstItems` (ListBox)
* `txtQty` (TextBox)
* `cmbLocation` (ComboBox)
* `btnSubmit`, `btnCancel`

#### frmShip (Shipping)

* `txtSearch`
* `lstItems`
* `txtQty`
* `cmbLocation`
* `txtDestination` (TextBox)
* `btnSubmit`, `btnCancel`

#### frmProduction (Production)

* `txtDesignSearch`
* `lstDesigns`
* `txtQtyToMake`
* `btnSubmit`, `btnCancel`

#### frmAdminConsole (Admin)

* `btnRunProcessor`
* `btnBuildWarehouse`
* `btnBuildInbox`
* `btnBuildGlobalSnapshot`
* `lstErrors` (ListBox)

---

## 6.5 Mermaid: Table relationship sketch

```mermaid
erDiagram
  tblItems ||--o{ tblOnHand : "SKU"
  tblWarehouseConfig ||--o{ tblLocations : "WarehouseID"
  tblOnHand ||--o{ tblInventoryLog : "affects"

  tblDesignHeader ||--o{ tblBOMLines : "DesignID+Version"
  tblDesignHeader ||--o{ tblProductionRuns : "DesignID+Version"
  tblProductionRuns ||--o{ tblProductionRunLines : "RunID"

  tblInboxReceive ||--o{ tblInventoryLog : "processed_into"
  tblInboxShip ||--o{ tblInventoryLog : "processed_into"
  tblInboxProd ||--o{ tblInventoryLog : "processed_into"
  tblInboxProd ||--o{ tblProductionRuns : "processed_into"
```

---

## 6.6 Mermaid: Runtime call graph (capability + domain ownership)

```mermaid
flowchart TB
  UI["RibbonX / Forms"] --> RoleXLAM["Role XLAM"] --> Core["Core"] --> Gate{{"Core.CanPerform"}}
  Gate -- "Denied" --> Deny["Reject + Log Attempt"]

  Gate -- "Allowed" --> InvDom["Inventory.Domain"]
  Gate -- "Allowed (if needed)" --> DesDom["Designs.Domain"]

  InvDom --> Inbox["Role Inbox (.xlsb)"]
  Inbox --> Proc["Processor"] --> Core

  Core --> Lock1["Lock Inventory"] --> InvDom --> InvDB["Inventory.xlsb"]
  Core --> Lock2["Lock Designs (optional)" ] --> DesDom --> DesDB["Designs.xlsb"]

  InvDom --> InvLog["InventoryLog"]
  DesDom --> Runs["ProductionRuns"]
```

---

## 6.7 Pseudocode (functional style, VB-ish)

> Naming note: these are *contracts*; each procedure should live in the owning XLAM.

### Core: capabilities and orchestration

```vb
'========== Core ==========
Enum InvSysCapability
    RECEIVE_POST
    SHIP_POST
    PROD_POST
    INBOX_PROCESS
    DESIGN_RELEASE
    ADMIN_MAINT
End Enum

Function Core_CanPerform(cap As InvSysCapability, userId As String, warehouseId As String) As Boolean
    ' Look up user role in credentials store (or cached session).
    ' Return True/False.
End Function

Sub Core_Require(cap As InvSysCapability, userId As String, warehouseId As String)
    If Not Core_CanPerform(cap, userId, warehouseId) Then
        Core_LogSecurityEvent userId, cap, warehouseId, "DENY"
        Err.Raise vbObjectError + 1001, , "Not authorized"
    End If
End Sub

Sub Core_LogSecurityEvent(userId As String, cap As InvSysCapability, warehouseId As String, result As String)
    ' Append to a small security/audit log (can be in Inventory store or separate).
End Sub

Function Core_NowUTC() As Date
    ' Return UTC (store both UTC and local if desired)
End Function

Function Core_NewEventId() As String
    ' Return GUID string
End Function

Function Core_GetWarehouseId() As String
    ' From config
End Function

Function Core_GetStationId() As String
    ' From config
End Function
```

### Inventory Domain: append events to inbox

```vb
'========== Inventory.Domain ==========

Sub InvDom_AppendReceiveEvent(inboxWbPath As String, sku As String, qty As Double, location As String, note As String)
    Dim eId As String: eId = Core_NewEventId()
    Dim wh As String: wh = Core_GetWarehouseId()
    Dim st As String: st = Core_GetStationId()

    ' Open/create inbox workbook and table tblInboxReceive
    ' Append row:
    '   EventID, EventUTC, WarehouseID, StationID, SKU, Qty, Location, Note, Status="NEW"
End Sub

Sub InvDom_AppendShipEvent(inboxWbPath As String, sku As String, qty As Double, location As String, destination As String, note As String)
    Dim eId As String: eId = Core_NewEventId()
    ' Append to tblInboxShip
End Sub
```

### Designs Domain: read BOM and (optionally) append production event

```vb
'========== Designs.Domain ==========

Function DesDom_GetActiveDesignVersion(designsDbPath As String, designId As String) As String
    ' Return latest RELEASED version (or active version)
End Function

Function DesDom_GetBOMLines(designsDbPath As String, designId As String, version As String) As Variant
    ' Return array/list of (ComponentSKU, QtyPerUnit, ScrapFactor)
End Function

Sub DesDom_AppendProdEvent(inboxWbPath As String, designId As String, version As String, qtyToMake As Double, note As String)
    Dim eId As String: eId = Core_NewEventId()
    ' Append to tblInboxProd with design/version and qty
End Sub
```

### Processor: apply inbox events to authoritative stores (lock order)

```vb
'========== Admin/Processor ==========

Sub Processor_RunOnce()
    Dim userId As String: userId = Core_GetCurrentUserId()
    Dim wh As String: wh = Core_GetWarehouseId()

    Core_Require INBOX_PROCESS, userId, wh

    ' 1) Lock Inventory first (always)
    InvDom_OpenAndLockInventoryDb

    ' 2) Determine if we need Designs lock (only if there are PROD events)
    Dim hasProd As Boolean: hasProd = Processor_InboxHasNewProdEvents()
    If hasProd Then
        DesDom_OpenAndLockDesignsDb
    End If

    ' 3) Apply receive events
    Processor_ApplyReceiveInbox

    ' 4) Apply ship events
    Processor_ApplyShipInbox

    ' 5) Apply production events (needs BOM)
    If hasProd Then
        Processor_ApplyProdInbox
    End If

    ' 6) Unlock (reverse order)
    If hasProd Then DesDom_UnlockDesignsDb
    InvDom_UnlockInventoryDb
End Sub

Sub Processor_ApplyReceiveInbox()
    For Each row In tblInboxReceive.Rows Where Status="NEW"
        If Processor_EventAlreadyApplied(row.EventID) Then
            row.Status = "SKIP_DUP"
        Else
            InvDom_ApplyReceiveToOnHand row
            InvDom_AppendInventoryLog row.EventID, "RECEIVE", row.SKU, row.Qty, row.Location
            row.Status = "PROCESSED"
        End If
    Next
End Sub

Sub Processor_ApplyShipInbox()
    For Each row In tblInboxShip.Rows Where Status="NEW"
        If Processor_EventAlreadyApplied(row.EventID) Then
            row.Status = "SKIP_DUP"
        Else
            InvDom_ApplyShipToOnHand row
            InvDom_AppendInventoryLog row.EventID, "SHIP", row.SKU, -row.Qty, row.Location
            row.Status = "PROCESSED"
        End If
    Next
End Sub

Sub Processor_ApplyProdInbox()
    For Each row In tblInboxProd.Rows Where Status="NEW"
        If Processor_EventAlreadyApplied(row.EventID) Then
            row.Status = "SKIP_DUP"
        Else
            Dim bomLines As Variant
            bomLines = DesDom_GetBOMLines(DesignsDbPath, row.DesignID, row.Version)

            ' Consume components
            For Each line In bomLines
                Dim need As Double
                need = row.QtyToMake * line.QtyPerUnit * (1 + line.ScrapFactor)
                InvDom_ConsumeComponent line.ComponentSKU, need, row.Location
                InvDom_AppendInventoryLog row.EventID, "PROD_CONSUME", line.ComponentSKU, -need, row.Location
            Next

            ' Produce finished good into Inventory
            InvDom_AddFinishedGood row.DesignID, row.QtyToMake, row.Location
            InvDom_AppendInventoryLog row.EventID, "PROD_OUTPUT", row.DesignID, row.QtyToMake, row.Location

            ' Record ProductionRuns in Designs
            DesDom_AppendProductionRun row.EventID, row.DesignID, row.Version, row.QtyToMake, wh

            row.Status = "PROCESSED"
        End If
    Next
End Sub

Function Processor_EventAlreadyApplied(eventId As String) As Boolean
    ' Look up eventId in tblInventoryLog.EventID (and/or tblProductionRuns.EventID)
End Function
```

---

## 6.8 Notes on generation (Admin)

```vb
Sub Admin_BuildWarehouseDataStores(warehouseId As String)
    ' Create Inventory.xlsb with required sheets + ListObjects if missing
    ' Create Designs.xlsb similarly (optional)
End Sub

Sub Admin_BuildStationInboxFiles(stationId As String)
    ' Create invSys.Inbox.Receiving.<Station>.xlsb with tblInboxReceive
    ' Create invSys.Inbox.Shipping.<Station>.xlsb with tblInboxShip
    ' Create invSys.Inbox.Production.<Station>.xlsb with tblInboxProd
End Sub
```
