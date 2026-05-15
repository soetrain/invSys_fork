# invSys R1 Controls Status

**Date:** 2026-05-14  
**Scope:** Release 1 Excel/VBA controls, forms, ribbon actions, and generated workbook buttons.

## Completeness Levels

| Level | Meaning |
| --- | --- |
| Complete | Usable for R1 with current architecture and enough test evidence to rely on it. |
| Partial | Exists and does useful work, but has known gaps, rough edges, or incomplete user flow. |
| Broken | Exists but currently fails or gives an unusable result. |
| Planned | Needed for R1 workflow but not implemented yet. |
| Needs Definition | The control name or concept exists, but its real R1 responsibility is not settled. |

## R1 Control Priorities

| Priority | Control / Area | Current Level | R1 Need |
| --- | --- | --- | --- |
| P0 | Setup Tester Station | Partial | Disposable smoke-test warehouse that proves NAS, SharePoint path, auth, seed data, processor, and operator workbook flow. |
| P0 | Delete Tester Station Generated | Partial | Remove the test warehouse/runtime artifacts created by Setup Tester Station so the user can reset and move on. Initial control exists; needs live validation on NAS. |
| P0 | Create New Warehouse | Partial | Real warehouse onboarding path after Tester Station passes. Creates runtime, config, auth, receiving inbox, initial admin/receiving user capabilities, seeded demo inventory, published discovery artifacts, and a receiving operator workbook for Confirm Writes. COM smoke passed; needs live NAS Confirm Writes validation. |
| P0 | Add Users and Roles | Partial / Needs Definition | Must provision real warehouse users, PINs, roles/capabilities, active/inactive state, and station scoping. |
| P0 | Admin Console | Needs Definition | Needs a clear R1 job description: operator-facing admin launcher, processor dashboard, warehouse maintenance console, or all of these. |
| P0 | Verify Add-ins Published | Broken | User reports it does not work. Must be debugged before R1 deployment flow can be trusted. |
| P1 | Publish Add-ins | Partial | Publishes required XLAMs and manifest to SharePoint add-ins location, but needs user-level validation with real SharePoint sync path. |
| P1 | Open Workbook after Tester Setup | Partial | Opens generated receiving workbook, but generated workbook location and readiness/auth behavior need cleanup. |
| P1 | Publish Initial Warehouse Artifacts | Partial | Used by Create New Warehouse; needs NAS + SharePoint validation. |
| P1 | Run Processor from Console | Partial | Existing admin call exists; needs clear UI surface and expected status/reporting. |
| P1 | Generate Inventory Snapshot | Partial | Existing admin call exists; needs confirmation it is wired into the final admin workflow. |
| P1 | Publish Warehouse Artifacts | Partial | Existing admin call exists; needed for SharePoint/WAN sync. Needs visible operator report. |
| P1 | Break Inventory Lock | Partial | Existing admin call exists; should require re-auth / admin capability and clear audit trail. |
| P1 | Poison Queue Review / Reissue | Partial | Existing functions exist; UI and expected operator flow need hardening. |
| P2 | Retire / Migrate Warehouse | Partial | Form and console functions exist, but not on the immediate Tester Station -> Create Warehouse path. |
| P2 | Admin Email | Needs Definition | Form exists; R1 value and required behavior need definition. |
| P2 | HQ Aggregation | Partial | Scheduled/admin functions exist; probably later in R1 after local warehouse operations are stable. |

## Admin Add-in Controls

| Control | Source / Surface | Current Level | Notes |
| --- | --- | --- | --- |
| Admin Controls | `frmAdminControls` | Partial | Launcher currently opens create/delete user and edit user forms. Needs to become coherent R1 admin entry point or be replaced by Admin Console. |
| Open Admin Console | `modAdminConsole.OpenAdminConsole` | Needs Definition | Schema and refresh helpers exist, but the user-facing purpose needs to be defined before polishing. |
| Open User Management | `modAdminConsole.OpenUserManagement` | Partial | Existing path opens user management surfaces, but current user forms still look older than the role/capability model. |
| Setup Tester Station | `frmSetupTesterStation`, `modTesterSetup.SetupTesterStation` | Partial | Now creates runtime, auth/config, seed data, processor run, and receiving workbook. Known gap: generated operator workbook currently lands on the warehouse hub path. |
| Setup Tester Station - Find Hub Path | `frmSetupTesterStation` dynamic button | Partial | Helps choose warehouse hub path. Needs NAS path expectation documented in UI or defaulting. |
| Setup Tester Station - Find SharePoint Root | `frmSetupTesterStation` dynamic button | Partial | Detects/browses SharePoint sync root. Works enough for current testing, but X1-Pro-Ai vs Zenbook path discovery may vary. |
| Setup Tester Station - Open Workbook | `frmSetupTesterStation` dynamic button | Partial | Opens generated workbook after setup. The workbook can still be unusable if generated in the wrong location or readiness/auth context is wrong. |
| Delete Tester Station Generated | `frmSetupTesterStation`, `modTesterSetup.DeleteTesterStationGenerated` | Partial | Removes known tester-generated runtime, inbox, snapshot, outbox, operator workbook, and tester package artifacts with confirmation and a report. Guarded to tester-looking warehouse ids. Needs live NAS validation. |
| Create New Warehouse | `frmCreateWarehouse`, `modAdminConsole.BootstrapWarehouseLocalAdmin` | Partial | Creates hub artifacts plus the first station receiving inbox. The generated receiving operator workbook is station-local under the current Windows user's Documents tree, with local config/auth copies beside it that still point `PathDataRoot` at the hub. Existing hub folders are allowed when warehouse artifacts do not already exist. Seeds demo inventory, lays out receiving tables horizontally, and adds Confirm Writes/Undo/Redo buttons. Needs live NAS Confirm Writes validation. |
| Create New Warehouse - Publish Initial | `frmCreateWarehouse.chkPublishInitial`, `PublishInitialArtifactsAdmin` | Partial | Can retry publish. Needs verification against real SharePoint sync path. |
| Create / Delete User | `frmCreateDeleteUser` | Partial | Exists, but appears table-bound to `UserCredentials`; needs alignment with current `Auth.xlsb` users/capabilities model. |
| Edit User | `frmEditUser` | Partial | Exists, but same concern as Create/Delete User: likely older role model assumptions. |
| Add Role / Edit Role / Assign Capability | New or revised user management controls | Planned | R1 needs capability-scoped roles/users, not just a simple role string. |
| Re-Authenticate | `frmReAuthGate` | Partial | Exists for admin-sensitive actions. Needs to be consistently applied to destructive/admin controls. |
| Verify Add-ins Published | `modAddinsPublish.VerifyAddinsPublished` | Broken | User reports it does not work. It currently checks required add-in files under resolved SharePoint add-ins root; likely path/root/manifest expectations need debugging. |
| Publish Add-ins | `modAddinsPublish.PublishAddins` | Partial | Publishes add-ins and manifest with staging/rollback behavior. Needs tested control surface and clear reports. |
| Loaded Package Diagnostics | `modPackageDiagnostics` | Partial | Useful support diagnostic. Needs UI placement if kept for R1. |
| Run Processor | `modAdminConsole.RunProcessorFromConsole` | Partial | Exists as function. Needs R1 UI button/report and expected schedule/manual behavior. |
| Break Inventory Lock | `modAdminConsole.BreakInventoryLock` | Partial | Exists as function. Needs confirmation UX, re-auth, and audit/report clarity. |
| Refresh Poison Queue | `modAdminConsole.RefreshPoisonQueue` | Partial | Exists. Needs UI table and operator action flow. |
| Reissue Poison Event | `modAdminConsole.ReissuePoisonEvent` | Partial | Exists. Needs guardrails and report path. |
| Generate Inventory Snapshot | `modAdminConsole.GenerateInventorySnapshot` | Partial | Exists. Needs place in Admin Console and expected success/failure report. |
| Publish Warehouse Artifacts | `modAdminConsole.PublishWarehouseArtifacts` | Partial | Exists. Needed for SharePoint/WAN relay. Needs user-visible publish report. |
| Scheduled Warehouse Batch | `modAdmin.Scheduler_RunWarehouseBatch` | Partial | Automation hook exists. R1 UI/admin scheduling story still needs definition. |
| Scheduled Warehouse Publish | `modAdmin.Scheduler_RunWarehousePublish` | Partial | Automation hook exists. Needs deployment/scheduling decision. |
| Scheduled HQ Aggregation | `modAdmin.Scheduler_RunHQAggregation` | Partial | Automation hook exists. Later R1 concern after warehouse path is stable. |
| Retire / Migrate Warehouse | `frmRetireMigrateWarehouse`, retire/migrate admin functions | Partial | Form and workflow exist. Needs R1 priority decision and live testing before relying on it. |

## Role Workbook Controls

| Control | Surface | Current Level | Notes |
| --- | --- | --- | --- |
| Receiving - Confirm Writes | `ReceivedTally` worksheet button, `modTS_Received.ConfirmWrites` | Partial | Core receive-post path exists and is capability-gated. Create Warehouse now generates the button and demo inventory rows; needs live NAS Confirm Writes validation. |
| Receiving - Undo | `ReceivedTally` worksheet button | Partial | Generated by receiving surface. Needs current behavior confirmed in R1 workflow. |
| Receiving - Redo | `ReceivedTally` worksheet button | Partial | Generated by receiving surface. Needs current behavior confirmed in R1 workflow. |
| Receiving - Setup UI | Receiving ribbon/control | Partial | Useful diagnostic. Has exposed config/readiness problems; should become less prominent or admin-only after R1 hardening. |
| Shipping - Post Shipment / Confirm | Shipping worksheet/ribbon controls | Partial | Existing role UI/event creator path exists. Needs same real-workbook validation as Receiving. |
| Shipping - Undo / Redo | Shipping worksheet controls | Partial | Needs R1 validation. |
| Production - Post Production / Confirm | Production worksheet/ribbon controls | Partial | Existing role UI/event creator path exists. Recent fixes prevent Production surfaces from contaminating Receiving workbooks. Needs real workbook validation. |
| Production - Undo / Redo | Production worksheet controls | Partial | Needs R1 validation. |
| Item Search | Role-specific search forms | Complete / Partial | Shared search architecture is implemented. Needs final user workflow validation per role workbook. |

## Domain / Inventory Controls

| Control | Surface | Current Level | Notes |
| --- | --- | --- | --- |
| Inventory Snapshot Generation | Admin/domain function | Partial | Needed for read models and SharePoint publishing. Must be reliable before R1. |
| Inventory Read Model Refresh | Operator workbook/readiness path | Partial | Works in tests, but generated workbook readiness/auth needs continued live validation. |
| Inventory Adjustments | Inventory domain/ribbon concept | Partial | R1 requirement likely exists, but final admin/operator ownership needs definition. |
| Inventory Logs | Inventory domain/ribbon concept | Partial | Useful for audit/debug. Needs final placement and operator expectations. |
| Locations | Inventory domain/ribbon concept | Partial | Needs R1 workflow definition if location management is user-facing. |
| Designs Domain Controls | Designs domain/ribbon concept | Partial / Needs Definition | Present in source tree but not on current warehouse setup critical path. |

## Tester Station Required Cleanup Control

Add a destructive-but-contained control named something like **Delete Tester Station Generated**.

Minimum behavior:

- Confirm the target warehouse id is exactly `TestStation` or another explicit tester id.
- Show the hub path and SharePoint path it will clean before deleting.
- Delete only files/folders known to be generated by Tester Station.
- Never delete the parent NAS share, deploy folder, or real warehouse folders.
- Include a dry-run/report mode internally so the UI can show what will be removed.
- Clear any generated receiving operator workbook path captured by Tester Station.
- Return a plain success/failure report.

Likely generated artifacts to remove:

- `TestStation.invSys.Config.xlsb`
- `TestStation.invSys.Auth.xlsb`
- `TestStation.invSys.Data.Inventory.xlsb`
- `TestStation.invSys.Snapshot.Inventory.xlsb`
- `TestStation.Outbox.Events.xlsb`
- `TestStation.Receiving.Operator.xlsm`
- `inbox\invSys.Inbox.Receiving.TS1.xlsb`
- Any Tester Station publish/test bundle artifacts created under SharePoint sync root.

Open decision:

- Whether Tester Station should create the operator workbook on the local station path, the SharePoint sync path, or another operator-safe path. It should not make the NAS hub workbook the normal operator workbook if that workbook is expected to be opened interactively.

## Admin Console Definition Needed

Before polishing Admin Console, define whether it is:

1. A warehouse setup launcher only.
2. A live warehouse maintenance console.
3. A processor/publish/diagnostics dashboard.
4. The single R1 admin home that includes all of the above.

Recommended R1 direction:

- Admin Console should be the single admin home.
- It should expose setup, users/roles, processor, publish/sync, diagnostics, and destructive maintenance.
- Destructive actions should require re-auth and produce an audit/report entry.
- Tester Station should remain a separate smoke-test section, clearly marked disposable.

## Immediate Next Work

1. Run **Create New Warehouse** against the NAS path and perform a Confirm Writes validation from the generated station-local receiving operator workbook.
2. Fix **Verify Add-ins Published**.
3. Decide the correct generated operator workbook location for Tester Station.
4. Replace or upgrade older Create/Delete/Edit User forms so they manage the real `Auth.xlsb` users/capabilities model.
5. Define the final Admin Console shape for R1.
