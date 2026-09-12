# Plan 022 Slice 4be.1 Shipping/Boxing activity discovery

Last verified: 2026-09-12 against source checkpoint `01891bb`. This source map
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
while reauthentication of the same user leaves the old form able to change local
staging. Source `CommitCurrentLine` has no captured-session check before
`ShipmentsFormCommitLine`; this contradicts existing D18 behavior, not the
approved architecture. The test does not establish an unauthorized Domain write.

Extend the real-handler context matrix to other mutation controls, target change,
sign-out and lost capability, retaining all prior GREEN checks. Protect pending,
uncertain-submission, storage and policy outcomes separately. A later owner
rejection or rollback cannot substitute for the required pre-owner stale-form guard.
Before implementation, record discovered ControlIds and precise owner outcomes
under D18 and synchronize Architecture, Plan022 and the controls catalog. Do not
use a Boolean success, report-text parsing, direct owner call or test-only auth
bypass as a substitute for the required user-action and application evidence.
