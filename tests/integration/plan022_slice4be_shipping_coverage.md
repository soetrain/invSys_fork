# Plan 022 Slice 4be.1 Shipping/Boxing activity discovery

Last verified: 2026-09-12, including the Shipping context/timer candidates described
in [context evidence](plan022_slice4be_shipping_context_results.md). This source map
advances D18's comprehensive coverage inventory. It neither registers new
ControlIds nor claims packaged activity coverage or changes a business contract.
Core catalog7 currently has no Shipping/Boxing activity entries. D18's existing
owner-fact, captured-context, optional-tracking and programmatic-call rules apply.

## Existing form routes

All handlers below belong to `src/Shipping/Forms/frmShipmentsTally.frm` in
`invSys.Operations.xlam`. Each row is **pending shared activity implementation
and packaged D13 evidence**. Captions are current source wording.

| Control/caption | Actual handler -> existing action boundary | Evidence to obtain from owner |
|---|---|---|
| `btnAdd` / Add | `mBtnAdd_Click` -> `CommitCurrentLine("ADD")` -> `modTS_Shipments.ShipmentsFormCommitLine` | Pre-validation rejection, exact-key reservation/staging, actual owner outcome and every emitted event identity. |
| `btnUpdate` / Update Row | `mBtnUpdate_Click` -> `CommitCurrentLine("UPDATE")` -> same owner | Exact selected line; rejection versus completed/partial mutation. |
| `btnRemove` / Remove | `mBtnRemove_Click` -> `RemoveSelectedShipmentRows` -> same owner with DELETE for each selected row | Every selected-line result; a partial multi-row failure cannot claim rollback. |
| `btnHold` / Send Hold; `btnReturn` / Return | `mBtnHold_Click` / `mBtnReturn_Click` -> `MoveSelectedShipmentHold` -> `ShipmentsFormMoveHoldRows` | Actual Hold/Return outcome; never relabel an internal reservation as Hold. |
| `btnStage` / To Shipments | `mBtnStage_Click` -> `RunShippingAction(True)` | Existing selected-row staging outcome, separate from shipment application. |
| `btnSend` / Shipments Sent | `mBtnSend_Click` -> `modShippingPostingService.ExecuteShipmentsSent` -> `ShipmentsFormRunShipmentsSentRows` | Exact submitted events, pending/uncertain state and later Domain application remain distinct. |
| `btnRefresh` / Refresh | `mBtnRefresh_Click` -> `ShipmentsFormRefreshRuntimeInventoryForWorkbook` | Owner freshness/failure; subsequent form rendering is not another click. |
| `btnHistory` / History | `mBtnHistory_Click` -> `ShipmentsFormRecentHistoryText` | Read/display outcome; never copy returned history text into activity. |
| `btnHistorySheet` / Export | `mBtnHistorySheet_Click` -> `ShipmentsFormExportHistoryToSheet` | Local export outcome and captured destination; no arbitrary report text in activity. |
| `btnClose` / Close | `mBtnClose_Click` -> `Me.Hide` | Observed dismissal only; do not silently adopt Receiving's unload lifecycle. |
| `btnBoxBuilderNewPage` / New Box | `mBtnBoxBuilderNew_Click` -> `ClearBoxBuilderSelection` | Local new-design staging, not a saved design. |
| `btnBoxBuilderRefreshPage` / Refresh Box Designs | `mBtnBoxBuilderRefresh_Click` -> `RefreshBoxBuilderPage` | Explicit refresh distinguished from initialization/page rendering. |
| `btnBoxBuilderAddComponentPage` / Add | `mBtnBoxBuilderAddComponent_Click` | Local component staging or validation rejection. |
| `btnBoxBuilderRemoveComponentPage` / Remove | `mBtnBoxBuilderRemoveComponent_Click` | Local component removal or missing-selection rejection. |
| `btnBoxBuilderSavePage` / Save Box | `mBtnBoxBuilderSave_Click` -> `modBoxingService.SaveBoxDesign` | Existing owner result; no inferred durable event ID. |
| `btnBoxBuilderUpdateVersionPage` / Update Alternative; `btnBoxBuilderNewVersionPage` / New Alternative | `mBtnBoxBuilderUpdateVersion_Click` / `mBtnBoxBuilderNewVersion_Click` -> `RunBoxBuilderVersionAction` -> `SaveBoxDesign` | Existing version action and result, preserving its authorization. |
| `btnBoxBuilderDeleteVersionPage` / Delete Alternative | `mBtnBoxBuilderDeleteVersion_Click` -> `modBoxingService.DeleteBoxDesignVersion` | Missing selection, user cancellation, authorization and owner completion remain distinct. |
| `btnBoxBuilderArchivePage` / Archive Box | `mBtnBoxBuilderArchive_Click` -> `modBoxingService.ArchiveBoxDesign` | Same distinctions; observation never authorizes archive. |
| `btnBoxBuilderDeletePage` / Delete Box | `mBtnBoxBuilderDelete_Click` -> `modBoxingService.DeleteBoxDesign` | Same distinctions; preserve existing explicit confirmation. |
| `btnBoxMakerRefreshPage` / Refresh Box Maker | `mBtnBoxMakerRefresh_Click` -> `RefreshBoxMakerPage` | Explicit refresh distinguished from internal refresh. |
| `btnBoxMakerMakePage` / Make Boxes; `btnBoxMakerUnmakePage` / Unbox | `mBtnBoxMakerMake_Click` / `mBtnBoxMakerUnmake_Click` -> `modBoxingService.PostBoxMakerAction` | Exact package/component identities and owner results; retain every emitted event. |

The normal Shipping launcher (`modTS_Shipments.BtnOpenShipmentsForm`) checks
`SHIP_POST`. Box Designer/Maker entry wrappers call that launcher and select a
page; their internal page-selection calls must not fabricate a second user click.
`SaveBoxDesign` and `PostBoxMakerAction` require `SHIP_POST`; the separate
Delete Alternative/Archive Box/Delete Box service paths use
`ValidateBoxMaintenanceRequest`, which requires `ADMIN_MAINT`. Discovery does
not broaden or consolidate these distinct permissions.

## Navigation, exclusions and unresolved reachability

- `mPages_Change` calls `ApplyShippingPage` for Shipping, Box Designer and Box
  Maker. Explicit page selection is eligible navigation under D18; initialization
  and programmatic selection are not user input.
- Shippable/Shipment/Hold list clicks, Use Existing, Box Designer/Maker design
  selection and alternative selection have existing handlers. Inventory/component
  lists and the Box Designer Status selector also need deliberate-selection
  coverage assessment even where no dedicated event handler currently exists.
- Text search/change, entered quantity/reference/carrier/description values,
  labels, automatic synchronization, layout and rendering are excluded from
  control-usage collection under D18. Exclusion does not suppress the surrounding
  command's attempt/result or required business/audit events.
- Native window close, `frmBoxVersionSaveChoice`, generated Ribbon dispatch and
  legacy worksheet-button reachability remain to be proven through real packaged
  entry points. Source existence or a test-only callable wrapper is insufficient
  to declare them operator-reachable, retired or safely deletable.

## Packaged findings and next D13 action

The [packaged activity test](plan022_slice4be_shipping_activity_results.md) now
proves the eight-action normal sequence, exact source identities and shipment
application, with missing shared activity RED. Its negative cases show that
zero quantity reaches existing validation without staging or canonical mutation,
while the unchanged `8e64268` source allowed same-user reauthentication to leave
the old form able to change local staging. Its `CommitCurrentLine` lacked a
captured-session check before `ShipmentsFormCommitLine`; that contradicted D18.
The test does not establish an unauthorized Domain write. The isolated context
candidate now guards all seven mutation handlers and explicit stale-form relaunch:
46 former guard failures pass, while46 missing-activity failures remain. Broader
candidate gates and the remaining context/failure cases are still required.

Normal workbook closure and pre-entry capability loss now have the focused evidence
below. Extend the real-handler matrix to permission loss during a yield and unavailable
Auth/Config, retaining the passing pending-yield/timer checks and all prior GREENs. Protect pending,
uncertain-submission, storage and policy outcomes separately. A later owner
rejection or rollback cannot substitute for the required pre-owner stale-form guard.
Before implementation, record discovered ControlIds and precise owner outcomes
under D18 and synchronize Architecture, Plan022 and the controls catalog. Do not
use a Boolean success, report-text parsing, direct owner call or test-only auth
bypass as a substitute for the required user-action and application evidence.

## Owner branches that constrain activity implementation

Source review on2026-09-12 adds these observations; they are not new runtime or
outcome contracts. D18's owner-fact rule governs the next protecting tests.

- `ShipmentsFormCommitLine` can release an earlier reservation and create a new
  one in the same edit. Its delta-only and preserved-reservation branches may
  complete without a new submission. Record every newly emitted identity, never
  copy an existing row's reservation ID as a new source event. DELETE may remove
  a row after a release-queue warning; Boolean success alone cannot prove that
  release was accepted. Failure may follow prior local or submitted changes.
- `ShipmentsFormRunToShipmentsRows` has two successful already-locked branches
  that update local staging without queueing a new reservation. Its other branch
  submits a reserve event before subsequent local work can fail. Empty references
  and retained references must each follow actual owner evidence.
- `ShipmentsFormRunShipmentsSentRows` explicitly checks SHIP_POST, queues SHIP,
  finalizes local staging, then runs processing/refresh. Its Boolean result can
  be True with `runtimeProcessed=False`. A local completion must not become an
  all-events-applied conclusion. Exceptions after queue entry retain uncertainty.
- `QueueShipmentsReserveEvent`, `QueueShipmentsReleaseEvent` and
  `QueueShipmentsSentEvent` delegate to `QueueShippingPayloadEventServerFirst`.
  Capture evidence at the owning submission boundary, with acceptance state per
  emitted identity; neither report parsing nor reading the latest current-state
  row can substitute for that evidence. Automatic catch-up is not another click.
- The server-first helper falls back to `QueuePayloadEventCurrent` after a
  negative or exceptional server result. Both calls receive the same ByRef
  event ID; Core `QueueEventCore` generates an ID only when that value is empty.
  Do not manufacture a second ID or duplicate reference merely because both
  submission paths execute. Core's current-target path may confirm local staging
  rather than server application. The pending/lost-acknowledgement tests must
  prove the actual accepted or uncertain boundary and preserve that exact ID.
  Both Core paths retain their capability checks; source inspection alone does
  not prove that revoking capability preserves every preceding local mutation.

The pending-yield/timer tests now separately protect stale owner dispatch; their
stopped-owner probes do not cover the failure/outcome branches above.

The subsequent [native validation/lifecycle checkpoint](plan022_slice4be_native_validation_results.md)
adds11/11 actual workbook-close checks, producing333 PASS /46 missing-activity
FAIL with all prior368 checks retained. Normal Shipping startup installs its event
hook; actual workbook close releases form/callback bindings, and late timer dispatch
does not reopen, submit, fabricate a user click or redirect to another workbook.
Read-only reopen preserves saved active/held staging, exact keys and unknown values.
This closes ordinary workbook-shutdown coverage only. A separately retained stale
form and the pending/uncertain owner branches remain unproven.

The [capability checkpoint](plan022_slice4be_shipping_capability_results.md) proves
same-session SHIP_POST revocation before all seven mutation handlers. Seven healthy
owner probes calibrate dispatch; the denied probes originally enter owners, and a
real denied Hold moves staging. RED374/61 becomes GREEN389/46 after normative
clarification0cc61c9 and the bounded form/context-helper repair. All435 checks and
prior GREENs remain. Owner checks and Core security authority are retained; the
timer keeps its existing context entry contract. Permission loss during a UI yield
and unavailable Auth/Config are separate cases, not implied by pre-entry revocation.

The subsequent [access-interruption evidence](plan022_slice4be_shipping_access_results.md)
completes558 checks:505 PASS/53 FAIL, retaining every prior435 check/389 GREEN.
All actual-yield revocations and missing-Config cases pass. Every missing-Auth
action also rejects mutation, but the ordinary read recreates Auth. Those seven
creation assertions expose the pending D8-A contract decision; they are separate
from the46 missing-activity failures. This does not approve an Auth behavior change
or substitute for pending/uncertain submission and comprehensive activity evidence.

[Submission discovery](plan022_slice4be_shipping_submission_results.md) now proves
pending Add references and exact ID survival through negative/exceptional server
acknowledgment, including owner failure after real server acceptance. Initial
605-check discovery retains all558 prior checks; isolated focused validation then
passes136/136, including66 submission/persistence/path checks. Two revised full
runs stop early with native faults and remain unresolved. The normal activity
test's applied-log-only expectation is subsequently corrected in the
[owner-reference RED](plan022_slice4be_shipping_owner_reference_results.md):650 checks,
597 PASS/53 FAIL, all preceding checks/GREENs retained and all66 source-submission
checks passing within the full route. References include pending owner submissions
and exclude earlier actions merely applied during the current click. Those
source facts do not yet prove Update/Remove mixed outcomes or Send processing.
