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