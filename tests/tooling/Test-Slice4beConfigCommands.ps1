[CmdletBinding()]
param(
    [string]$RepoRoot = '',
    [string]$DeployRoot = 'deploy/current',
    [ValidateSet('RED','GREEN')][string]$Phase = 'GREEN',
    [switch]$CaptureEvidence,
    [switch]$CheckAuthReadOnly,
    [switch]$CheckActivityEvidence,
    [switch]$CheckActivityFoundation,
    [switch]$CheckAdminUomActivity,
    [switch]$CheckProductionDesignerActivity,
    [switch]$CheckProductionClose,
    [switch]$CheckProcessWorksheetHeaders,
    [switch]$CheckProcessWorksheetPicker,
    [switch]$CheckProcessWorksheetActivity,
    [switch]$CheckProcessWorksheetPaths,
    [switch]$ProcessWorksheetClosedDiagnostic,
    [switch]$ProcessWorksheetCatalogOnly,
    [switch]$CheckProductionClosePaths,
    [switch]$CheckInventoryQueryReadOnly,
    [switch]$InventoryQueriesOnly,
    [switch]$CheckProductionUomStaging,
    [switch]$CheckProductionUomActivity,
    [switch]$UomAdapterDiagnostic,
    [switch]$CheckProductionUomPublicClose,
    [switch]$CheckProductionUomPaths,
    [switch]$CheckProductionInstructions,
    [switch]$CheckProductionComponents,
    [switch]$CheckProductionRecipeOrder,
    [switch]$CheckProductionRecipeStructure,
    [switch]$CheckProductionDesignReads,
    [switch]$CheckProductionAssignment,
    [switch]$CheckProductionRunPresentation,
    [switch]$RunPresentationPathsOnly,
    [switch]$CheckProductionRunLocal,
    [switch]$RunLocalClosedDiagnostic,
    [switch]$RunLocalRefillDiagnostic,
    [switch]$RunCompleteActivityOnly,
    [switch]$RunCompleteSubmissionFaultOnly,
    [switch]$RunCompleteStatusOnly,
    [switch]$RunCompletePathsOnly,
    [ValidateSet('Reusable','Worksheet')][string]$RunCompletePathMode='Reusable',
    [switch]$RunCompleteBaselineOnly,
    [switch]$RunCompletePreparationDiagnostic,
    [ValidateSet('All','CompletePending')][string]$RunCompleteClosedBoundary='All',
    [switch]$RunCompleteVisibleHostForTest,
    [switch]$RunCompleteSavedDecoyForTest,
    [ValidateSet('None','Submission','OutputReturn','Entry','Interruptions','Initial','PriorSequence')][string]$RunCompletePreparationPrelude='None',
    [switch]$RunPrintBaselineOnly,
    [switch]$RunPrintRecordedOnly,
    [switch]$RunPrintPathsDiagnostic,
    [ValidateSet('None','BareSeed','FreshSeed','SeedOnly')][string]$PrintSeedBoundaryForTest='None',
    [switch]$RunNextBaselineOnly,
    [switch]$RunNextActivityOnly,
    [switch]$RunNextTerminalOnly,
    [switch]$RunNextYieldOnly,
    [switch]$RunNextClosedOnly,
    [switch]$RunNextClosedEscapeForTest,
    [switch]$RunNextPathsOnly,
    [ValidateSet('Reusable','Worksheet')][string]$RunNextPathMode='Reusable',
    [switch]$RunCheckInBaselineOnly,
    [switch]$RunCheckInRoutedOnly,
    [switch]$RunCheckInClosedOnly,
    [switch]$RunCheckInActivityOnly,
    [switch]$RunCheckInPathsOnly,
    [ValidateSet('Reusable','Worksheet')][string]$RunCheckInPathMode='Reusable',
    [switch]$RunLocalContractOnly,
    [switch]$RunLocalPolicyOnly,
    [switch]$RunLocalFaultOnly,
    [switch]$RunLocalYieldOnly,
    [switch]$RunLocalStockOnly,
    [switch]$RunLocalWorksheetOnly,
    [switch]$RunLocalWorksheetOwnerOnly,
    [switch]$RunLocalWorksheetScaleOnly,
    [switch]$RunClearPathsOnly,
    [switch]$RunLoadPathsOnly,
    [switch]$RunRefreshPathsOnly,
    [switch]$RunAllocatePathsOnly,
    [switch]$CheckProductionAssignmentSafety,
    [switch]$CheckProductionAssignmentPaths,
    [switch]$CheckProductionAssignmentSourcePaths,
    [switch]$CheckProductionRegulation,
    [switch]$CheckProductionRegulationPaths,
    [switch]$CheckProductionDesignReadPaths,
    [switch]$CheckProductionRecipeStructureReleasedData,
    [switch]$CheckProductionRecipeStructurePaths,
    [switch]$CheckProductionRecipeOrderPaths,
    [switch]$CheckProductionComponentPaths,
    [switch]$CheckProductionInstructionPaths,
    [switch]$CheckProductionLifecycle,
    [switch]$CheckProductionLifecyclePaths,
    [switch]$CheckProductionLifecycleNative,
    [switch]$CheckProductionLifecyclePresentation,
    [switch]$CheckProductionDesignerPaths,
    [switch]$CaptureProductionDesignerPaths,
    [switch]$CheckSettingsEditorActivity,
    [switch]$GeneralSettingsOnly,
    [switch]$GeneralSettingsPolicyOnly,
    [switch]$CheckSettingsDiagnostics,
    [switch]$SettingsSafetyOnly,
    [switch]$CheckAdminUomExpectationChoice,
    [switch]$CheckTrackingSettings,
    [switch]$CheckTrackingPolicy,
    [switch]$CheckUserTrackingPolicy,
    [switch]$CheckDetailProfile,
    [switch]$CheckDetailColumns,
    [switch]$CheckActionPathPreference,
    [switch]$CheckOperationsTrackingSettings,
    [switch]$CheckAdminSettingsClose,
    [switch]$AdminSettingsCloseOnly,
    [switch]$CheckViewerRefreshFailure,
    [switch]$CheckViewerEventDetail,
    [switch]$DetailScrollLockDiagnostic,
    [switch]$CheckDetailScrollMovement,
    [switch]$CheckViewerEventGroups,
    [switch]$CheckViewerPublishedRead,
    [switch]$CheckViewerShippingState,
    [switch]$CheckViewerFilters,
    [switch]$CheckActionRecording,
    [switch]$CheckRecordingLimits,
    [switch]$CheckRecordingStorageBounds,
    [switch]$CheckRecordingReader,
    [switch]$CheckGuideDraft,
    [switch]$CheckGuideSave,
    [switch]$CheckGuideLibrary,
    [switch]$CheckGuideExpectation,
    [switch]$CheckGuideEvaluation,
    [switch]$CheckGuidePresentation,
    [switch]$CheckGuidePresentationAvailability,
    [switch]$CheckGuidePresentationRestart,
    [switch]$GuidePresentationRestartOnly,
    [switch]$TraceGuideResourcesForTest,
    [switch]$TraceGuideViewCallsForTest,
    [switch]$CheckGuideLayoutStabilityForTest,
    [switch]$GuideResourceSavedWorkbookForTest,
    [switch]$CheckPublishedGuideEdit,
    [switch]$PublishedGuideEditOnly,
    [switch]$GuideActionCurationOnly,
    [switch]$CheckGuideTransfer,
    [switch]$CheckGuideTransferRoundTrip,
    [switch]$CheckGuideTransferActivity,
    [switch]$CheckGuideTransferPolicy,
    [switch]$CheckGuideTransferRecorded,
    [switch]$RetryActionPathViewCountForTest,
    [switch]$RetryGuideObservationForTest,
    [switch]$WaitForExcelReadyForTest,
    [ValidateRange(1,120)][int]$ExcelReadyReadLimitForTest = 8,
    [switch]$CaptureGuideEvidence,
    [switch]$GuideCaptureVisibleExcelForTest,
    [switch]$GuideCaptureSavedWorkbookForTest,
    [switch]$TraceViewerStartupForTest,
    [switch]$ViewerStartupSavedWorkbookForTest,
    [ValidateSet('None','Caption','Status','Config','ShownConfig','ConfigSteps','ConfigCloseVisible')]
    [string]$ViewerStartupCalibrationForTest = 'None',
    [ValidateSet('OriginalReadOnly','WritableCopies','SavedCopies')]
    [string]$ViewerStartupPackageStateForTest = 'OriginalReadOnly',
    [switch]$GuideDraftOnly,
    [switch]$CheckRecordingIsolation,
    [switch]$CheckRecordingRestart,
    [switch]$CheckRecordingOperations,
    [switch]$CheckReceivingReplay,
    [switch]$CheckRecordingEvaluation,
    [switch]$CheckEvaluationContracts,
    [switch]$CheckExpectationCompatibility,
    [switch]$CheckEvaluationVisualEvidence,
    [switch]$CheckOperationsGuidePresentation,
    [switch]$RecordingEvaluationDiagnostic,
    [Alias('CompileViewerProbesForTest')][switch]$CompileEvaluationProbesForTest,
    [switch]$CheckExpectationEditor,
    [string]$RecordingContinuationPipeName = '',
    [switch]$CheckViewerPublication,
    [switch]$ViewerPublicationOnly,
    [switch]$CheckShippingActivity,
    [switch]$CheckShippingRecording,
    [switch]$CheckBoxingActivity,
    [switch]$CheckOwnerCommandCompletion,
    [switch]$OwnerCommandCompletionOnly,
    [switch]$TraceBootstrapForTest,
    [switch]$TraceSettingsOpenForTest,
    [switch]$TraceReceivingRunCloseForTest,
    [switch]$ShippingBeforeSharedFormsForTest,
    [switch]$PrepareShippingFixturesBeforeProbesForTest,
    [switch]$ShippingSubmissionOnly,
    [switch]$CheckReceivingActivity,
    [switch]$CheckReceivingStagingActivity,
    [switch]$CheckReceivingLocalActivity,
    [switch]$CheckReceivingLifecycleActivity,
    [switch]$CheckReceivingNavigationActivity,
    [switch]$ReceivingNavigationOnly,
    [switch]$CheckReceivingSurfaceCoverage,
    [switch]$ReceivingSurfaceOnly,
    [switch]$CheckReceivingNativeSurface,
    [switch]$CheckReceivingWorksheetActivity,
    [switch]$CheckReceivingWorksheetScenarios,
    [switch]$CheckReceivingWorksheetGuards,
    [switch]$CheckReceivingLauncherDenial,
    [switch]$ReceivingLauncherDenialOnly,
    [switch]$CaptureDenialDialogs,
    [switch]$ReceivingLifecycleOnly,
    [ValidateSet('None','SkipTerminationEvidence','KeepLauncherReference')]
    [string]$LifecycleDiagnostic = 'None'
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if($RunPrintBaselineOnly -and (-not $CheckProductionRunLocal -or $RunCompleteBaselineOnly -or $RunNextBaselineOnly)){throw 'Print baseline requires its isolated Run fixture.'}
if($PrintSeedBoundaryForTest -ne 'None' -and -not $RunPrintBaselineOnly){throw 'Seed boundary diagnosis requires the Print fixture.'}
if($RunPrintRecordedOnly -and (-not $RunPrintBaselineOnly -or $PrintSeedBoundaryForTest -ne 'None')){throw 'Print recording proof requires its isolated Print fixture.'}
if($RunPrintPathsDiagnostic -and -not $RunPrintRecordedOnly){throw 'Print view diagnosis requires recorded Print mode.'}
if($RunNextBaselineOnly -and -not $RunCompleteBaselineOnly){throw 'Next Batch baseline requires Complete baseline fixture setup.'}
if($RunCompletePreparationDiagnostic -and (-not $RunCompleteBaselineOnly -or $RunNextBaselineOnly)){throw 'Preparation diagnosis requires the isolated Complete Run fixture.'}
if($RunCompleteClosedBoundary -ne 'All' -and -not $RunCompletePreparationDiagnostic){throw 'A narrowed closure boundary requires explicit diagnostic mode.'}
if($RunCompletePreparationPrelude -ne 'None' -and -not $RunCompletePreparationDiagnostic){throw 'Preparation prelude requires explicit diagnostic mode.'}
if($RunCompleteVisibleHostForTest -and (-not $RunCompleteBaselineOnly -or $RunNextBaselineOnly)){throw 'Visible-host comparison requires the isolated Complete Run fixture.'}
if($RunCompleteSavedDecoyForTest -and (-not $RunCompleteBaselineOnly -or $RunNextBaselineOnly)){throw 'Saved decoy requires the isolated completion gate.'}
if($RunCompleteActivityOnly -and (-not $RunCompleteBaselineOnly -or $RunNextBaselineOnly -or $RunCompletePreparationDiagnostic)){throw 'Complete activity requires its isolated baseline fixture.'}
if($RunCompleteSubmissionFaultOnly -and -not $RunCompleteActivityOnly){throw 'Complete submission faults require the activity fixture.'}
if($RunCompleteStatusOnly -and (-not $RunCompleteActivityOnly -or $RunCompleteSubmissionFaultOnly)){throw 'Status-only evidence requires the completion activity fixture without submission faults.'}
if($RunCompletePathsOnly -and (-not $RunCompleteBaselineOnly -or $RunCompleteActivityOnly -or $RunCompletePreparationDiagnostic -or $RunNextBaselineOnly)){throw 'Complete paths require their isolated baseline fixture.'}
if($RunNextActivityOnly -and -not $RunNextBaselineOnly){throw 'Next Batch activity requires its binding baseline.'}
if($RunNextYieldOnly -and (-not $RunNextBaselineOnly -or $RunNextActivityOnly -or $RunNextPathsOnly)){throw 'Next Batch interruptions require the isolated Next baseline fixture.'}
if($RunNextTerminalOnly -and (-not $RunNextBaselineOnly -or $RunNextActivityOnly -or $RunNextYieldOnly -or $RunNextPathsOnly)){throw 'Next terminal proof requires the isolated Next baseline.'}
if($RunNextClosedOnly -and -not $RunNextYieldOnly){throw 'Next Batch native closure requires its isolated interruption fixture.'}
if($RunNextClosedEscapeForTest -and -not $RunNextClosedOnly){throw 'Next Batch diagnostic escape requires native closure testing.'}
if($RunNextPathsOnly -and (-not $RunNextBaselineOnly -or $RunNextActivityOnly)){throw 'Next Batch paths require the isolated Next baseline fixture.'}
if($TraceGuideResourcesForTest){
    if(-not $GuidePresentationRestartOnly){throw 'Resource tracing requires the isolated restart diagnostic.'}
    . (Join-Path $PSScriptRoot 'Slice4beGuideResourceTrace.ps1')
}
if($GuideResourceSavedWorkbookForTest -and -not $TraceGuideResourcesForTest){throw 'Saved-workbook resource control requires tracing.'}
if($TraceGuideViewCallsForTest -and -not $GuideResourceSavedWorkbookForTest){throw 'View-call tracing requires the saved-workbook resource control.'}
if($CheckGuideLayoutStabilityForTest -and -not $TraceGuideViewCallsForTest){throw 'Layout stability requires the instrumented handler trace.'}
if($CheckDetailScrollMovement -and (-not $CheckViewerEventDetail -or -not $CaptureEvidence -or $DetailScrollLockDiagnostic)){
    throw 'Native scrolling checks require the isolated visible detail gate without temporary unlocking.'
}
if($CheckGuideTransferRoundTrip -and -not $CheckGuideTransfer){throw 'Round trip requires the isolated guide transfer fixture.'}
if($CheckGuideTransferActivity -and (-not $CheckGuideTransfer -or $CheckGuideTransferRoundTrip)){throw 'Transfer activity requires its separate compiled transfer fixture.'}
if($CheckGuideTransferPolicy -and -not $CheckGuideTransferActivity){throw 'Transfer policy requires the activity fixture.'}
if($CheckGuideTransferRecorded -and (-not $CheckGuideTransferActivity -or $CheckGuideTransferPolicy)){throw 'Recorded transfers require their separate activity fixture.'}
if($CheckGuideTransfer){
    if(-not $GuideDraftOnly -or -not $CheckViewerPublishedRead -or -not $CompileEvaluationProbesForTest -or $ViewerStartupPackageStateForTest -ne 'SavedCopies'){throw 'Guide transfer requires its compiled saved-copy guide fixture.'}
    if($CheckReceivingReplay -or $GuideActionCurationOnly -or $PublishedGuideEditOnly -or $GuidePresentationRestartOnly){throw 'Guide transfer requires an isolated fixture.'}
    $CheckGuideExpectation=$true
}
if($GuideActionCurationOnly){
    if($PublishedGuideEditOnly -or $CheckPublishedGuideEdit -or $CheckGuidePresentation -or $GuidePresentationRestartOnly -or $CheckGuidePresentationRestart){throw 'Direct curation uses its own focused packaged gate.'}
    if(-not $GuideDraftOnly -or -not $CheckViewerPublishedRead -or -not $CompileEvaluationProbesForTest -or $ViewerStartupPackageStateForTest -ne 'SavedCopies'){throw 'Direct curation requires the isolated compiled saved-copy guide fixture.'}
    $CheckGuideExpectation=$true
}
if($PublishedGuideEditOnly){$CheckPublishedGuideEdit=$true}
if($CheckPublishedGuideEdit){
    if($GuidePresentationRestartOnly -or $CheckGuidePresentationRestart){throw 'Published guide editing and preference restart use separate gates.'}
    $CheckGuidePresentation=$true
}
if($GuidePresentationRestartOnly){
    if($CheckGuidePresentationAvailability){throw 'The focused restart gate is separate from the availability regression gate.'}
    $CheckGuidePresentationRestart=$true
}
if($CheckGuidePresentationRestart){
    if(-not $GuideDraftOnly -or $ViewerStartupPackageStateForTest -ne 'SavedCopies' -or $GuideCaptureSavedWorkbookForTest){
        throw 'Paired preference restart requires the isolated saved-copy guide gate without a retained capture workbook.'
    }
    $CheckGuidePresentation=$true
}
if($CheckGuidePresentationAvailability){
    if(-not $CaptureGuideEvidence){throw 'Presentation availability requires the complete visible presentation gate.'}
    $CheckGuidePresentation=$true
}
if($CheckGuidePresentation){$CheckGuideEvaluation=$true}
if($CheckGuideEvaluation){$CheckGuideExpectation=$true}
if($GuideCaptureVisibleExcelForTest -and -not $TraceViewerStartupForTest -and (-not $GuideDraftOnly -or (-not $CaptureGuideEvidence -and -not $TraceGuideResourcesForTest))){
    throw 'Visible-Excel diagnosis requires Viewer startup tracing, or GuideDraftOnly with capture/resource tracing.'
}
if($TraceViewerStartupForTest -and ($Phase -ne 'RED' -or -not $CheckViewerPublishedRead -or -not $CompileEvaluationProbesForTest)){
    throw 'Viewer startup tracing requires RED and the published-Viewer fixture with instrumented compilation.'
}
if($ViewerStartupSavedWorkbookForTest -and -not $TraceViewerStartupForTest){
    throw 'The saved startup workbook is restricted to Viewer startup diagnosis.'
}
if($GuideCaptureSavedWorkbookForTest -and (-not $GuideCaptureVisibleExcelForTest -or -not $GuideDraftOnly -or -not $CaptureGuideEvidence -or $ViewerStartupPackageStateForTest -ne 'SavedCopies')){
    throw 'Saved guide-capture workbook requires the visible saved-copy guide gate.'
}
if($ViewerStartupCalibrationForTest -ne 'None' -and -not $TraceViewerStartupForTest){
    throw 'Viewer startup calibration requires explicit startup tracing.'
}
if($ViewerStartupPackageStateForTest -ne 'OriginalReadOnly' -and -not $TraceViewerStartupForTest){
    $savedGuide=$GuideDraftOnly -and $CheckGuideExpectation -and $CheckViewerPublishedRead
    $savedPrint=$RunPrintBaselineOnly -and $CheckProductionDesignerActivity -and $CheckProductionRunLocal -and $PrintSeedBoundaryForTest -cne 'BareSeed'
    $savedCheckIn=$RunCheckInActivityOnly -and $CheckProductionDesignerActivity -and $CheckProductionRunLocal -and -not $CheckTrackingSettings -and -not $CheckSettingsEditorActivity -and -not $CheckViewerPublishedRead -and -not $CheckShippingRecording
    $savedComplete=$RunCompleteBaselineOnly -and $CheckProductionDesignerActivity -and $CheckProductionRunLocal -and -not $RunNextBaselineOnly -and -not $CheckTrackingSettings -and -not $CheckSettingsEditorActivity -and -not $CheckViewerPublishedRead -and -not $CheckShippingRecording
    if($ViewerStartupPackageStateForTest -ne 'SavedCopies' -or -not $CompileEvaluationProbesForTest -or -not ($savedGuide -or $savedComplete -or $savedPrint -or $savedCheckIn)){
        throw 'Package-state comparison requires startup tracing or an isolated compiled saved-copy guide/completion/Print/Check In gate.'
    }
}
if($CheckBoxingActivity){$CheckShippingRecording=$true}
if($CheckOwnerCommandCompletion -and (-not $CheckBoxingActivity -or -not $CaptureEvidence -or $CheckExpectationEditor -or $CheckExpectationCompatibility)){
    throw 'Owner completion requires the isolated visible Boxing activity gate and its instrumented compilation.'
}
if($OwnerCommandCompletionOnly -and -not $CheckOwnerCommandCompletion){throw 'Focused owner completion requires its protecting gate.'}
if($GuideDraftOnly){$CheckGuideDraft=$true}
if($CaptureGuideEvidence){$CheckGuideLibrary=$true}
if($CheckGuideExpectation){$CheckGuideLibrary=$true}
if($CheckGuideLibrary){$CheckGuideSave=$true}
if($CheckGuideSave){$CheckGuideDraft=$true}
if($CheckOperationsGuidePresentation){
    if(-not $CompileEvaluationProbesForTest -or $GuideDraftOnly -or $CheckGuideDraft -or $CaptureGuideEvidence){throw 'Operations guide comparison requires the separate compiled evaluation gate.'}
    $CheckEvaluationVisualEvidence=$true
}
if($CheckEvaluationVisualEvidence){$CheckExpectationCompatibility=$true}
if($RecordingEvaluationDiagnostic){
    if($Phase -ne 'RED' -or -not $CheckExpectationCompatibility){throw 'Evaluation isolation requires RED and the full evaluation contract probes; it is not regression acceptance.'}
    if($CheckRecordingLimits -or $CheckRecordingStorageBounds -or $CheckRecordingRestart -or $CheckExpectationEditor){throw 'Run limits, storage, restart and editor-only gates separately from evaluation diagnosis.'}
    Write-Output 'DIAGNOSTIC: evaluation fixture first; full regression gate remains required.'
}
if($CheckAdminUomExpectationChoice -and (-not $CompileEvaluationProbesForTest -or -not ($CheckExpectationEditor -or $CheckExpectationCompatibility))){
    throw 'Admin UOM expectation choice requires a compiled expectation editor gate.'
}
if($CheckAdminUomActivity){
    if(-not $CompileEvaluationProbesForTest -or $CheckViewerPublishedRead -or $CheckTrackingSettings -or $CheckShippingActivity -or $CheckReceivingActivity -or $CheckActivityFoundation){throw 'Admin UOM commands require their separate compiled activity gate.'}
    $CheckActivityEvidence=$true
}
if($CheckSettingsDiagnostics){$CheckSettingsEditorActivity=$true}
if($GeneralSettingsOnly -and (-not $CheckSettingsEditorActivity -or $CheckSettingsDiagnostics)){throw 'General Settings requires its separate compiled Settings callback gate.'}
if($GeneralSettingsPolicyOnly -and (-not $GeneralSettingsOnly -or $Phase -cne 'RED')){throw 'Policy-only fixture diagnosis requires the General gate and RED; it is not full acceptance GREEN.'}
if($CheckSettingsEditorActivity){
    if(-not $CompileEvaluationProbesForTest -or -not $CaptureEvidence -or $CheckAdminUomActivity -or $CheckViewerPublishedRead -or $CheckTrackingSettings -or $CheckShippingActivity -or $CheckReceivingActivity -or $CheckActivityFoundation){throw 'Settings editor observations require their separate visible compiled gate.'}
    $CheckActivityEvidence=$true
}
if($SettingsSafetyOnly -and (-not $CheckSettingsEditorActivity -or $Phase -ne 'RED')){throw 'Focused Settings safety diagnosis requires the Settings callback gate and RED; it is not full acceptance GREEN.'}
if($InventoryQueriesOnly -and -not $CheckInventoryQueryReadOnly){throw 'Query-only diagnosis requires the supplemental query gate.'}
if($CheckInventoryQueryReadOnly -and -not $CheckProductionClose){throw 'Inventory query preservation supplements the compiled public Production Close gate.'}
if($RunLocalClosedDiagnostic -and (-not $CheckProductionRunLocal -or $Phase -cne 'RED')){throw 'Run closed-boundary diagnosis requires its separate RED fixture gate.'}
if($RunLocalContractOnly -and (-not $CheckProductionRunLocal -or $RunLocalClosedDiagnostic)){throw 'Run Core contract checks supplement a separate actual-handler gate.'}
if($RunLocalPolicyOnly -and (-not $CheckProductionRunLocal -or $RunLocalClosedDiagnostic -or $RunLocalContractOnly)){throw 'Run policy checks require their separate actual-handler gate.'}
if($RunLocalFaultOnly -and (-not $CheckProductionRunLocal -or $RunLocalClosedDiagnostic -or $RunLocalContractOnly -or $RunLocalPolicyOnly)){throw 'Run fault checks require their separate actual-handler gate.'}
if($RunLocalYieldOnly -and (-not $CheckProductionRunLocal -or $RunLocalClosedDiagnostic -or $RunLocalContractOnly -or $RunLocalPolicyOnly -or $RunLocalFaultOnly)){throw 'Run read-yield checks require their separate actual-handler gate.'}
if($RunLocalStockOnly -and (-not $CheckProductionRunLocal -or $RunLocalClosedDiagnostic -or $RunLocalContractOnly -or $RunLocalPolicyOnly -or $RunLocalFaultOnly -or $RunLocalYieldOnly)){throw 'Run stock-bucket checks require their separate actual-handler gate.'}
if($RunLocalRefillDiagnostic -and (-not $CheckProductionRunLocal -or @($RunLocalClosedDiagnostic,$RunLocalContractOnly,$RunLocalPolicyOnly,$RunLocalFaultOnly,$RunLocalYieldOnly,$RunLocalStockOnly,$RunLocalWorksheetOnly,$RunLocalWorksheetOwnerOnly,$RunLocalWorksheetScaleOnly,$RunClearPathsOnly,$RunLoadPathsOnly,$RunRefreshPathsOnly|Where-Object{$_}).Count)){throw 'Run refill diagnostic requires its separate actual-handler gate.'}
if($RunLocalWorksheetOnly -and (-not $CheckProductionRunLocal -or $RunLocalClosedDiagnostic -or $RunLocalContractOnly -or $RunLocalPolicyOnly -or $RunLocalFaultOnly -or $RunLocalYieldOnly -or $RunLocalStockOnly)){throw 'Run worksheet checks require their separate actual-handler gate.'}
if($RunLocalWorksheetOwnerOnly -and (-not $CheckProductionRunLocal -or $RunLocalClosedDiagnostic -or $RunLocalContractOnly -or $RunLocalPolicyOnly -or $RunLocalFaultOnly -or $RunLocalYieldOnly -or $RunLocalStockOnly -or $RunLocalWorksheetOnly)){throw 'Run worksheet owner checks require their separate actual-handler gate.'}
if($RunLocalWorksheetScaleOnly -and (-not $CheckProductionRunLocal -or $RunLocalClosedDiagnostic -or $RunLocalContractOnly -or $RunLocalPolicyOnly -or $RunLocalFaultOnly -or $RunLocalYieldOnly -or $RunLocalStockOnly -or $RunLocalWorksheetOnly -or $RunLocalWorksheetOwnerOnly)){throw 'Worksheet Scale requires its separate actual-handler gate.'}
if($RunClearPathsOnly -and (-not $CheckProductionRunLocal -or -not $CaptureEvidence -or $RunLocalClosedDiagnostic -or $RunLocalContractOnly -or $RunLocalPolicyOnly -or $RunLocalFaultOnly -or $RunLocalYieldOnly -or $RunLocalStockOnly -or $RunLocalWorksheetOnly -or $RunLocalWorksheetOwnerOnly -or $RunLocalWorksheetScaleOnly)){throw 'Clear paths require a separate actual-handler gate and visible evidence.'}
if($RunLoadPathsOnly -and (-not $CheckProductionRunLocal -or -not $CaptureEvidence -or $RunClearPathsOnly -or $RunLocalClosedDiagnostic -or $RunLocalContractOnly -or $RunLocalPolicyOnly -or $RunLocalFaultOnly -or $RunLocalYieldOnly -or $RunLocalStockOnly -or $RunLocalWorksheetOnly -or $RunLocalWorksheetOwnerOnly -or $RunLocalWorksheetScaleOnly)){throw 'Load paths require a separate actual-handler gate and visible evidence.'}
if($RunAllocatePathsOnly -and (-not $CheckProductionRunLocal -or -not $CaptureEvidence -or @($RunRefreshPathsOnly,$RunClearPathsOnly,$RunLoadPathsOnly,$RunLocalRefillDiagnostic,$RunLocalClosedDiagnostic,$RunLocalContractOnly,$RunLocalPolicyOnly,$RunLocalFaultOnly,$RunLocalYieldOnly,$RunLocalStockOnly,$RunLocalWorksheetOnly,$RunLocalWorksheetOwnerOnly,$RunLocalWorksheetScaleOnly|Where-Object{$_}).Count)){throw 'Apply paths require a separate actual-handler gate and visible evidence.'}
if($RunRefreshPathsOnly -and (-not $CheckProductionRunLocal -or -not $CaptureEvidence -or $RunLoadPathsOnly -or $RunClearPathsOnly -or $RunLocalClosedDiagnostic -or $RunLocalContractOnly -or $RunLocalPolicyOnly -or $RunLocalFaultOnly -or $RunLocalYieldOnly -or $RunLocalStockOnly -or $RunLocalWorksheetOnly -or $RunLocalWorksheetOwnerOnly -or $RunLocalWorksheetScaleOnly)){throw 'Refresh paths require a separate actual-handler gate and visible evidence.'}
if(($RunCheckInBaselineOnly -or $RunCheckInRoutedOnly -or $RunCheckInClosedOnly) -and (-not $CheckProductionRunLocal -or (@($RunCheckInBaselineOnly,$RunCheckInRoutedOnly,$RunCheckInClosedOnly|Where-Object{$_}).Count -gt 1) -or @($RunAllocatePathsOnly,$RunRefreshPathsOnly,$RunClearPathsOnly,$RunLoadPathsOnly,$RunLocalRefillDiagnostic,$RunLocalClosedDiagnostic,$RunLocalContractOnly,$RunLocalPolicyOnly,$RunLocalFaultOnly,$RunLocalYieldOnly,$RunLocalStockOnly,$RunLocalWorksheetOnly,$RunLocalWorksheetOwnerOnly,$RunLocalWorksheetScaleOnly|Where-Object{$_}).Count)){throw 'Check In baseline requires a separate actual-handler gate.'}
if($RunCompleteBaselineOnly -and (-not $CheckProductionRunLocal -or -not $CaptureEvidence -or @($RunCheckInPathsOnly,$RunCheckInActivityOnly,$RunCheckInBaselineOnly,$RunCheckInClosedOnly,$RunCheckInRoutedOnly,$RunAllocatePathsOnly,$RunRefreshPathsOnly,$RunClearPathsOnly,$RunLoadPathsOnly,$RunLocalRefillDiagnostic,$RunLocalClosedDiagnostic,$RunLocalContractOnly,$RunLocalPolicyOnly,$RunLocalFaultOnly,$RunLocalYieldOnly,$RunLocalStockOnly,$RunLocalWorksheetOnly,$RunLocalWorksheetOwnerOnly,$RunLocalWorksheetScaleOnly|Where-Object{$_}).Count)){throw 'Complete baseline requires a separate actual-handler gate and visible evidence.'}
if($RunCheckInPathsOnly -and (-not $CheckProductionRunLocal -or -not $CaptureEvidence -or @($RunCheckInActivityOnly,$RunCheckInBaselineOnly,$RunCheckInClosedOnly,$RunCheckInRoutedOnly,$RunAllocatePathsOnly,$RunRefreshPathsOnly,$RunClearPathsOnly,$RunLoadPathsOnly,$RunLocalRefillDiagnostic,$RunLocalClosedDiagnostic,$RunLocalContractOnly,$RunLocalPolicyOnly,$RunLocalFaultOnly,$RunLocalYieldOnly,$RunLocalStockOnly,$RunLocalWorksheetOnly,$RunLocalWorksheetOwnerOnly,$RunLocalWorksheetScaleOnly|Where-Object{$_}).Count)){throw 'Check In paths require a separate actual-handler gate and visible evidence.'}
if($CheckProductionRunLocal){
    $otherRunGates=@($PSBoundParameters.Keys|Where-Object{($_ -like 'CheckProduction*' -or $_ -like 'CheckProcessWorksheet*') -and $_ -notin @('CheckProductionRunLocal','CheckProductionDesignerActivity') -and [bool]$PSBoundParameters[$_]})
    if(-not $CheckProductionDesignerActivity -or $otherRunGates.Count -gt 0){throw 'Run local observations require their separate compiled actual-handler gate.'}
}
if($RunPresentationPathsOnly -and (-not $CheckProductionRunPresentation -or -not $CaptureEvidence)){throw 'Run presentation paths require their compiled actual-handler gate and visible evidence.'}
if($CheckProductionRunPresentation){
    $otherProductionGates=@($PSBoundParameters.Keys|Where-Object{($_ -like 'CheckProduction*' -or $_ -like 'CheckProcessWorksheet*') -and $_ -notin @('CheckProductionRunPresentation','CheckProductionDesignerActivity') -and [bool]$PSBoundParameters[$_]})
    if(-not $CheckProductionDesignerActivity -or $otherProductionGates.Count -gt 0){throw 'Run presentation requires its separate compiled actual-handler gate.'}
}
if($CheckProcessWorksheetPicker -and -not $CheckProcessWorksheetHeaders){throw 'Process picker checks require the compiled worksheet fixture adapters.'}
if($CheckProcessWorksheetActivity -and (-not $CheckProcessWorksheetHeaders -or $CheckProcessWorksheetPicker)){throw 'Worksheet observations require their separate compiled worksheet fixture gate.'}
if($CheckProcessWorksheetPaths -and (-not $CheckProcessWorksheetActivity -or $ProcessWorksheetClosedDiagnostic -or $ProcessWorksheetCatalogOnly)){throw 'Worksheet paths require their separate compiled activity adapters.'}
if($ProcessWorksheetClosedDiagnostic -and -not $CheckProcessWorksheetActivity){throw 'Closed diagnostic requires worksheet activity fixtures.'}
if($ProcessWorksheetCatalogOnly -and (-not $CheckProcessWorksheetActivity -or $ProcessWorksheetClosedDiagnostic)){throw 'Catalog-only coverage requires separate worksheet activity fixtures.'}
if($CheckProductionClosePaths -and (-not $CheckProductionClose -or $CheckInventoryQueryReadOnly -or $InventoryQueriesOnly)){throw 'Close paths require a separate compiled Close gate.'}
if($CheckProcessWorksheetHeaders -and (-not $CheckProductionDesignerActivity -or -not $CaptureEvidence -or $CheckProductionClose -or $CheckProductionUomStaging -or $CheckProductionInstructions -or $CheckProductionComponents -or $CheckProductionRecipeOrder -or $CheckProductionRecipeStructure -or $CheckProductionDesignReads -or $CheckProductionRegulation -or $CheckProductionLifecycle -or $CheckProductionDesignerPaths)){throw 'Process worksheet headers require their separate visible compiled handler gate.'}
if($CheckProductionClose -and (-not $CheckProductionDesignerActivity -or -not $CaptureEvidence -or $CheckProductionUomStaging -or $CheckProductionInstructions -or $CheckProductionComponents -or $CheckProductionRecipeOrder -or $CheckProductionRecipeStructure -or $CheckProductionDesignReads -or $CheckProductionRegulation -or $CheckProductionLifecycle -or $CheckProductionDesignerPaths)){throw 'Production Close requires its separate visible compiled handler gate.'}
if($CheckProductionUomStaging -and (-not $CheckProductionDesignerActivity -or $CheckProductionInstructions -or $CheckProductionLifecycle -or $CheckProductionDesignerPaths)){throw 'UOM staging preservation requires its separate compiled Production handler gate.'}
if($CheckProductionUomActivity -and -not $CheckProductionUomStaging){throw 'UOM activity requires its packaged staging gate.'}
if($UomAdapterDiagnostic -and (-not $CheckProductionUomActivity -or -not $CaptureEvidence -or $CheckProductionUomPaths -or $Phase -cne 'RED')){throw 'Adapter diagnosis requires the combined visible UOM activity gate in diagnostic RED.'}
if($CheckProductionUomPublicClose -and (-not $CheckProductionUomActivity -or -not $CaptureEvidence -or $CheckProductionUomPaths -or $UomAdapterDiagnostic)){throw 'Public close requires the separate visible UOM activity fixture.'}
if($CheckProductionUomPaths -and (-not $CheckProductionUomStaging -or -not $CaptureEvidence -or $CheckProductionUomActivity)){throw 'UOM paths require the separate staging adapters and visible evidence.'}
if($CheckProductionInstructions -and (-not $CheckProductionDesignerActivity -or $CheckProductionLifecycle -or $CheckProductionDesignerPaths)){throw 'Instruction checks require their separate compiled Production designer gate.'}
if($CheckProductionComponents -and (-not $CheckProductionDesignerActivity -or $CheckProductionInstructions -or $CheckProductionLifecycle -or $CheckProductionDesignerPaths -or $CheckProductionUomStaging)){throw 'Component checks require their separate compiled Production designer gate.'}
if($CheckProductionRecipeOrder -and (-not $CheckProductionDesignerActivity -or $CheckProductionComponents -or $CheckProductionInstructions -or $CheckProductionLifecycle -or $CheckProductionDesignerPaths -or $CheckProductionUomStaging)){throw 'Recipe ordering requires its separate compiled Production designer gate.'}
if($CheckProductionRecipeOrderPaths -and (-not $CheckProductionRecipeOrder -or -not $CaptureEvidence)){throw 'Recipe ordering paths require compiled ordering adapters and visible evidence.'}
if($CheckProductionRecipeStructure -and (-not $CheckProductionDesignerActivity -or $CheckProductionRecipeOrder -or $CheckProductionComponents -or $CheckProductionInstructions -or $CheckProductionLifecycle -or $CheckProductionDesignerPaths -or $CheckProductionUomStaging)){throw 'Recipe structure requires its separate compiled Production designer gate.'}
if($CheckProductionRecipeStructureReleasedData -and -not $CheckProductionRecipeStructure){throw 'Released-data structure checks require the compiled structure adapters.'}
if($CheckProductionRecipeStructurePaths -and (-not $CheckProductionRecipeStructure -or $CheckProductionRecipeStructureReleasedData -or -not $CaptureEvidence)){throw 'Structure paths require compiled adapters and visible evidence in a separate gate.'}
if($CheckProductionDesignReads -and (-not $CheckProductionDesignerActivity -or $CheckProductionRecipeStructure -or $CheckProductionRecipeOrder -or $CheckProductionComponents -or $CheckProductionInstructions -or $CheckProductionLifecycle -or $CheckProductionDesignerPaths -or $CheckProductionUomStaging)){throw 'Designer read checks require a separate compiled designer gate.'}
if($CheckProductionDesignReadPaths -and (-not $CheckProductionDesignReads -or -not $CaptureEvidence)){throw 'Designer read paths require compiled adapters and visible evidence in a separate gate.'}
if($CheckProductionAssignment -and (-not $CheckProductionDesignerActivity -or $CheckProductionDesignReads -or $CheckProductionRegulation -or $CheckProductionClose -or $CheckProcessWorksheetHeaders -or $CheckProductionRecipeStructure -or $CheckProductionRecipeOrder -or $CheckProductionComponents -or $CheckProductionInstructions -or $CheckProductionLifecycle -or $CheckProductionDesignerPaths -or $CheckProductionUomStaging)){throw 'Assignment observations require their separate compiled actual-handler gate.'}
if($CheckProductionAssignmentPaths -and (-not $CheckProductionAssignment -or $CheckProductionAssignmentSafety -or -not $CaptureEvidence)){throw 'Assignment paths require separate compiled adapters and visible evidence.'}
if($CheckProductionAssignmentSourcePaths -and (-not $CheckProductionAssignment -or $CheckProductionAssignmentPaths -or $CheckProductionAssignmentSafety)){throw 'Assignment source paths require a separate compiled actual-handler gate.'}
if($CheckProductionAssignmentSafety -and -not $CheckProductionAssignment){throw 'Assignment safety requires the Assignment adapters.'}
if($CheckProductionRegulationPaths -and (-not $CheckProductionRegulation -or -not $CaptureEvidence)){throw 'Regulation paths require compiled adapters and visible evidence in a separate gate.'}
if($CheckProductionRegulation -and (-not $CheckProductionDesignerActivity -or $CheckProductionDesignReads -or $CheckProductionRecipeStructure -or $CheckProductionRecipeOrder -or $CheckProductionComponents -or $CheckProductionInstructions -or $CheckProductionLifecycle -or $CheckProductionDesignerPaths -or $CheckProductionUomStaging)){throw 'Output regulation requires its separate compiled designer gate.'}
if($CheckProductionInstructionPaths -and (-not $CheckProductionInstructions -or -not $CaptureEvidence)){throw 'Instruction paths require the compiled instruction adapters and visible evidence.'}
if($CheckProductionComponentPaths -and (-not $CheckProductionComponents -or -not $CaptureEvidence)){throw 'Component paths require the compiled component adapters and visible evidence.'}
if($CheckProductionLifecycle -and (-not $CheckProductionDesignerActivity -or $CheckProductionDesignerPaths)){throw 'Lifecycle checks require the separate compiled Production designer gate.'}
if($CheckProductionLifecyclePaths -and -not $CheckProductionLifecycle){throw 'Lifecycle paths require the lifecycle adapters.'}
if(($CheckProductionLifecycleNative -or $CheckProductionLifecyclePresentation) -and (-not $CheckProductionLifecycle -or -not $CaptureEvidence)){throw 'Native lifecycle and presentation checks require the compiled lifecycle adapters and visible evidence.'}
if(($CheckProductionLifecyclePaths -and ($CheckProductionLifecycleNative -or $CheckProductionLifecyclePresentation)) -or ($CheckProductionLifecycleNative -and $CheckProductionLifecyclePresentation)){throw 'Run lifecycle extensions separately.'}
if($CheckProductionDesignerActivity){
    if(-not $CompileEvaluationProbesForTest -or $CheckViewerPublishedRead -or $CheckSettingsEditorActivity -or $CheckAdminUomActivity -or $CheckShippingActivity -or $CheckReceivingActivity -or $CheckTrackingSettings){throw 'Production designer observations require their separate compiled gate.'}
    $CheckActivityEvidence=$true
}
if($CheckProductionDesignerPaths -and -not $CheckProductionDesignerActivity){throw 'Production paths require the designer gate.'}
if($CaptureProductionDesignerPaths -and -not $CheckProductionDesignerPaths){throw 'Production captures require the path gate.'}
if($CompileEvaluationProbesForTest -and -not ($CheckAuthReadOnly -or $RecordingEvaluationDiagnostic -or $CheckEvaluationVisualEvidence -or $CheckViewerPublishedRead -or $CheckAdminUomActivity -or $CheckSettingsEditorActivity -or $CheckTrackingSettings -or $CheckProductionDesignerActivity)){
    throw 'Instrumented project compilation requires a supported focused gate.'
}
if($TraceSettingsOpenForTest -and ($Phase -ne 'RED' -or -not $CheckTrackingSettings)) {
    throw 'Settings constructor tracing requires RED and the Settings checks; it is not acceptance GREEN.'
}
if($TraceReceivingRunCloseForTest -and ($Phase -ne 'RED' -or -not $CheckReceivingReplay)) {
    throw 'Receiving close tracing requires RED and the replay checks; it is not acceptance GREEN.'
}
if($AdminSettingsCloseOnly) { $CheckAdminSettingsClose = $true }
if($ViewerPublicationOnly) {
    if($Phase -ne 'RED'){throw 'Publication-only diagnosis is not acceptance GREEN.'}
    $CheckViewerPublication = $true
}
if($CheckViewerPublication) { $CheckViewerEventGroups = $true }
if($CheckViewerShippingState) { $CheckViewerPublishedRead = $true }
if($CheckViewerFilters) { $CheckViewerPublishedRead = $true }
if($CheckRecordingLimits) { $CheckActionRecording = $true }
if($CheckRecordingStorageBounds) { $CheckActionRecording = $true }
if($CheckExpectationCompatibility) { $CheckEvaluationContracts = $true }
if($CheckExpectationEditor) {
    if($CheckExpectationCompatibility -or $CheckRecordingOperations -or $CheckRecordingRestart){throw 'Use the editor-focused gate separately; the full Operations gate remains required.'}
    $CheckActionRecording = $true
}
if($CheckEvaluationContracts) { $CheckRecordingEvaluation = $true }
if($CheckRecordingEvaluation) { $CheckRecordingOperations = $true }
if($CheckRecordingOperations) {
    if($CheckRecordingRestart){throw 'Operations sequence and cold restart use separate runs.'}
    $CheckRecordingReader = $true
}
if($CheckRecordingRestart) {
    if($CheckRecordingLimits -or $CheckRecordingStorageBounds -or $CheckRecordingIsolation -or $CheckViewerFilters -or $CheckViewerShippingState) {
        throw 'Cold recording restart uses its separate reader gate; run the preserved limits/isolation/filter gates separately.'
    }
    $CheckRecordingReader = $true
}
if($CheckRecordingReader) { $CheckActionRecording = $true }
if($CheckReceivingReplay) {
    if(-not $CompileEvaluationProbesForTest){throw 'Receiving replay requires compiled packaged probes.'}
    $CheckGuideDraft = $true
}
if($CheckGuideDraft) { $CheckActionRecording = $true }
if($CheckRecordingIsolation) { $CheckActionRecording = $true }
if($CheckActionRecording) { $CheckViewerPublishedRead = $true }
if($CheckAdminSettingsClose -and -not $CheckActionPathPreference) { throw 'Admin close requires the complete preference probes.' }
if ([string]::IsNullOrWhiteSpace($RepoRoot)) { $RepoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot) }
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$deploy = (Resolve-Path -LiteralPath (Join-Path $repo $DeployRoot)).Path
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Close Excel before isolated packaged validation.' }
$recordingHandoff=$null; $recordingFixtureTransferred=$false
. (Join-Path $PSScriptRoot 'Slice4beRecordingTransfer.ps1')
if($RecordingContinuationPipeName -ne ''){
    if(-not $CheckRecordingRestart -or $RecordingContinuationPipeName -notmatch '^invsys-recording-[0-9a-f]{32}$'){
        throw 'Invalid recording continuation mode.'
    }
    $recordingHandoff=[IO.Pipes.NamedPipeClientStream]::new('.', $RecordingContinuationPipeName, [IO.Pipes.PipeDirection]::Out)
    $recordingHandoff.Connect(10000)
}
$runRoot = Join-Path ([IO.Path]::GetTempPath()) ('invsys-config-command-' + [guid]::NewGuid().ToString('N'))
$reportRoot = Join-Path $repo 'reports/runtime/config-commands'
if ($CheckActivityEvidence) {
    $reportRoot = Join-Path $repo 'reports/runtime/slice4be-activity'
    . (Join-Path $PSScriptRoot 'Slice4beActivityAssertions.ps1')
    if ($CheckActivityFoundation) { . (Join-Path $PSScriptRoot 'Slice4beActivityFoundation.ps1') }
}
if ($CheckActivityFoundation -and -not $CheckActivityEvidence) { throw 'Foundation checks require activity evidence mode.' }
if ($CheckShippingActivity -and (-not $CheckActivityFoundation -or $CheckReceivingActivity)) { throw 'Shipping activity requires the foundation and a separate run from Receiving.' }
# Boxing needs the recording probes installed before its forms, including the
# existing submission-only diagnostic. That route does not claim recording coverage.
if ($CheckShippingRecording -and (-not $CheckShippingActivity -or ($ShippingSubmissionOnly -and (-not $CheckBoxingActivity -or $CheckOwnerCommandCompletion)))) { throw 'Shipping recording requires the complete Shipping activity route.' }
if ($ShippingSubmissionOnly -and -not $CheckShippingActivity) { throw 'Shipping submission-only discovery requires Shipping activity mode.' }
if ($TraceBootstrapForTest -and -not $CheckShippingActivity -and -not $RecordingEvaluationDiagnostic) {
    throw 'Bootstrap tracing requires the isolated Shipping or evaluation diagnostic route.'
}
if ($ShippingBeforeSharedFormsForTest -and (-not $CheckShippingActivity -or $ShippingSubmissionOnly)) { throw 'Shipping-first ordering requires the complete Shipping route.' }
if ($PrepareShippingFixturesBeforeProbesForTest -and -not $ShippingBeforeSharedFormsForTest) { throw 'Prepared Shipping fixtures require the complete Shipping-first route.' }
$preparedShippingFixtures=@{}
$preparedShippingBoundaries=@{}
if ($CheckShippingActivity) {
    . (Join-Path $PSScriptRoot 'Slice4beShippingActivity.ps1')
    $reportRoot = Join-Path $repo ('reports/runtime/slice4be-shipping-activity/'+[guid]::NewGuid().ToString('N'))
}
if ($CheckReceivingStagingActivity -and -not $CheckReceivingActivity) { throw 'Staging coverage requires Receiving activity mode.' }
if ($CheckReceivingLocalActivity -and -not $CheckReceivingStagingActivity) { throw 'Local-action coverage requires staging activity mode.' }
if ($CheckReceivingLifecycleActivity -and -not $CheckReceivingLocalActivity) { throw 'Lifecycle coverage requires the preserved local-action baseline.' }
if ($CheckReceivingNavigationActivity -and (-not $CheckReceivingLifecycleActivity -or $ReceivingLifecycleOnly)) { throw 'Navigation coverage requires the full preserved lifecycle baseline.' }
if ($ReceivingNavigationOnly -and -not $CheckReceivingNavigationActivity) { throw 'Navigation-only diagnosis requires navigation coverage.' }
if ($CheckReceivingSurfaceCoverage -and -not $CheckReceivingNavigationActivity) { throw 'Surface coverage requires the preserved navigation baseline.' }
if ($ReceivingSurfaceOnly -and (-not $CheckReceivingSurfaceCoverage -or $ReceivingNavigationOnly)) { throw 'Surface-only diagnosis requires surface coverage without another diagnostic-only mode.' }
if ($CheckReceivingNativeSurface -and -not $ReceivingSurfaceOnly) { throw 'Native worksheet discovery requires the separate surface-only run.' }
if ($CheckReceivingWorksheetActivity -and -not $CheckReceivingNativeSurface) { throw 'Worksheet activity requires calibrated native surface coverage.' }
if ($CheckReceivingWorksheetScenarios -and -not $CheckReceivingWorksheetActivity) { throw 'Worksheet scenarios require the protected native activity baseline.' }
if ($CheckReceivingWorksheetGuards -and -not $CheckReceivingWorksheetActivity) { throw 'Worksheet guards require the protected native activity baseline.' }
if ($CheckReceivingLauncherDenial -and -not $CheckReceivingNavigationActivity) { throw 'Launcher denial coverage requires the preserved navigation baseline.' }
if ($ReceivingLauncherDenialOnly -and (-not $CheckReceivingLauncherDenial -or $ReceivingSurfaceOnly -or $ReceivingNavigationOnly)) { throw 'Launcher-denial-only diagnosis requires its coverage without another diagnostic-only mode.' }
if ($CaptureDenialDialogs -and -not $ReceivingLauncherDenialOnly) { throw 'Native denial dialog evidence uses the separate focused run.' }
if ($ReceivingLifecycleOnly -and -not $CheckReceivingLifecycleActivity) { throw 'Lifecycle-only diagnosis requires lifecycle coverage.' }
if ($LifecycleDiagnostic -ne 'None' -and -not $ReceivingLifecycleOnly) { throw 'Mutation diagnostics require the separate lifecycle-only report.' }
if ($LifecycleDiagnostic -ne 'None' -and $Phase -ne 'RED') { throw 'Diagnostic mutations cannot be run as acceptance GREEN.' }
if ($CheckReceivingActivity) {
    $reportRoot = Join-Path $repo 'reports/runtime/slice4be-receiving-activity'
    . (Join-Path $PSScriptRoot 'Slice4beActivityAssertions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beReceivingActivity.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beReceivingReferences.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beReceivingRetry.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beReceivingStaging.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beReceivingLocal.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beReceivingFreshness.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beReceivingLifecycle.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beReceivingNavigation.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beReceivingSurface.ps1')
    if ($CheckReceivingNativeSurface) {
        . (Join-Path $PSScriptRoot 'Slice4beReceivingNativeSurface.ps1')
        if ($CheckReceivingWorksheetScenarios) { . (Join-Path $PSScriptRoot 'Slice4beReceivingWorksheetScenarios.ps1') }
        if ($CheckReceivingWorksheetGuards) { . (Join-Path $PSScriptRoot 'Slice4beReceivingWorksheetGuards.ps1') }
        $reportRoot=Join-Path $reportRoot ('native-surface-'+[guid]::NewGuid().ToString('N'))
        if ($CheckReceivingWorksheetActivity) { $reportRoot=Join-Path $reportRoot 'worksheet-activity' }
    }
    . (Join-Path $PSScriptRoot 'Slice4beReceivingLauncherDenial.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beReceivingDenialDialogs.ps1')
    if (-not $CheckActivityFoundation) { . (Join-Path $PSScriptRoot 'Slice4beActivityFoundation.ps1') }
}
if ($CheckTrackingSettings) {
    if ($CheckShippingActivity -or $CheckReceivingActivity -or $CheckActivityEvidence) {
        throw 'Tracking Settings discovery uses the separate preserved D5 route.'
    }
    . (Join-Path $PSScriptRoot 'Slice4beTrackingSettings.ps1')
    $reportRoot = Join-Path $repo ('reports/runtime/slice4be-tracking-settings/'+[guid]::NewGuid().ToString('N'))
}
if ($CheckTrackingPolicy -and -not $CheckTrackingSettings) { throw 'Tracking policy checks require the Settings route.' }
if ($CheckUserTrackingPolicy -and (-not $CheckTrackingSettings -or -not $CompileEvaluationProbesForTest)) { throw 'User policy checks require the compiled Settings route.' }
if ($CheckDetailProfile -and -not $CheckTrackingSettings) { throw 'Detail profile checks require the Settings route.' }
if ($CheckDetailProfile -and -not $CheckTrackingPolicy) { throw 'Detail profile checks retain the tracking policy baseline and cancellation observer.' }
if ($CheckDetailColumns) {
    if (-not $CheckDetailProfile -or -not $CompileEvaluationProbesForTest) { throw 'Detail column geometry requires the compiled Settings/profile route.' }
    . (Join-Path $PSScriptRoot 'Slice4beDetailColumns.ps1')
    $reportRoot = Join-Path $repo ('reports/runtime/slice4be-detail-columns/'+[guid]::NewGuid().ToString('N'))
}
if ($CheckActionPathPreference -and -not $CheckDetailProfile) { throw 'Preference checks retain the detail and policy baseline.' }
if ($CheckOperationsTrackingSettings -and -not $CheckActionPathPreference) { throw 'Operations Settings checks retain the full personal preference baseline.' }
if ($CheckViewerRefreshFailure) {
    if ($CheckTrackingSettings -or $CheckActivityEvidence -or $CheckReceivingActivity -or $CheckShippingActivity) { throw 'Viewer refresh failure uses a separate focused run.' }
    $reportRoot = Join-Path $repo ('reports/runtime/slice4be-viewer-refresh/'+[guid]::NewGuid().ToString('N'))
    . (Join-Path $PSScriptRoot 'Slice4beViewerRefreshFailure.ps1')
}
if ($CheckViewerEventDetail) {
    if ($CheckTrackingSettings -or $CheckActivityEvidence -or $CheckReceivingActivity -or $CheckShippingActivity -or $CheckViewerRefreshFailure) { throw 'Viewer event detail uses a separate focused run.' }
    $reportRoot = Join-Path $repo ('reports/runtime/slice4be-viewer-detail/'+[guid]::NewGuid().ToString('N'))
    . (Join-Path $PSScriptRoot 'Slice4beViewerEventDetail.ps1')
}
if ($CheckViewerPublishedRead) {
    if ($CheckTrackingSettings -or $CheckActivityEvidence -or $CheckReceivingActivity -or $CheckShippingActivity -or $CheckViewerRefreshFailure -or $CheckViewerEventDetail -or $CheckViewerEventGroups -or $CheckViewerPublication) { throw 'Published Viewer reads use a separate focused run.' }
    $reportRoot = Join-Path $repo ('reports/runtime/slice4be-viewer-published-read/'+[guid]::NewGuid().ToString('N'))
    . (Join-Path $PSScriptRoot 'Slice4beViewerPublishedRead.ps1')
}
if ($CheckViewerEventGroups) {
    if ($CheckTrackingSettings -or $CheckActivityEvidence -or $CheckReceivingActivity -or $CheckShippingActivity -or $CheckViewerRefreshFailure -or $CheckViewerEventDetail) { throw 'Viewer event groups uses a separate focused run.' }
    $reportRoot = Join-Path $repo ('reports/runtime/slice4be-viewer-groups/'+[guid]::NewGuid().ToString('N'))
    . (Join-Path $PSScriptRoot 'Slice4beViewerEventGroups.ps1')
    if ($CheckViewerPublication) { . (Join-Path $PSScriptRoot 'Slice4beViewerPublication.ps1') }
}
if($CheckAdminUomActivity){
    $reportRoot=Join-Path $repo ('reports/runtime/slice4be-admin-uom/'+[guid]::NewGuid().ToString('N'))
    . (Join-Path $PSScriptRoot 'Slice4beAdminUomProbe.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beAdminUomActivity.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beAdminUomTracking.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beActivityFoundation.ps1')
}
if($CheckAdminUomExpectationChoice){
    . (Join-Path $PSScriptRoot 'Slice4beAdminUomProbe.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beAdminUomExpectation.ps1')
}
if($CheckSettingsEditorActivity){
    $reportRoot=Join-Path $repo ('reports/runtime/slice4be-settings-activity/'+[guid]::NewGuid().ToString('N'))
    . (Join-Path $PSScriptRoot 'Slice4beSettingsEditorProbe.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beSettingsEditorActivity.ps1')
    if($GeneralSettingsOnly){
        $reportRoot=Join-Path $repo ('reports/runtime/slice4be-general-settings/'+[guid]::NewGuid().ToString('N'))
        . (Join-Path $PSScriptRoot 'Slice4beGeneralSettingsProbe.ps1')
        . (Join-Path $PSScriptRoot 'Slice4beGeneralSettings.ps1')
        Write-Output ('General Settings evidence: '+$reportRoot)
    }
}
if($CheckProductionDesignerActivity){
    $reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-designer/'+[guid]::NewGuid().ToString('N'))
    if($CheckProcessWorksheetHeaders){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-process-worksheet-headers/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProcessWorksheetPicker){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-process-worksheet-picker/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProcessWorksheetActivity){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-process-worksheet-activity/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionClose){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-close/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProcessWorksheetPaths){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-process-worksheet-paths/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionClosePaths){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-close-paths/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionUomStaging){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-uom-staging/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionInstructions){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-instructions/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionComponents){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-components/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionRecipeOrder){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-recipe-order/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionRecipeStructure){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-recipe-structure/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionDesignReads){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-design-reads/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionAssignment){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-assignment/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionRunPresentation){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-run-presentation/'+[guid]::NewGuid().ToString('N'))}
    if($RunPresentationPathsOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-run-presentation-paths/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionRunLocal){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-run-local/'+[guid]::NewGuid().ToString('N'))}
    if($RunLocalClosedDiagnostic){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-run-local-closed/'+[guid]::NewGuid().ToString('N'))}
    if($RunCompleteBaselineOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-complete-baseline/'+[guid]::NewGuid().ToString('N'))}
    if($RunCompletePreparationDiagnostic){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-complete-preparation/'+[guid]::NewGuid().ToString('N'))}
    if($RunPrintBaselineOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-print-baseline/'+[guid]::NewGuid().ToString('N'))}
    if($RunPrintRecordedOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-print-recorded/'+[guid]::NewGuid().ToString('N'))}
    if($RunNextBaselineOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-next-baseline/'+[guid]::NewGuid().ToString('N'))}
    if($RunCompleteActivityOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-complete-activity/'+[guid]::NewGuid().ToString('N'))}
    if($RunCompleteSubmissionFaultOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-complete-submission-fault/'+[guid]::NewGuid().ToString('N'))}
    if($RunCompletePathsOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-complete-paths-'+$RunCompletePathMode.ToLowerInvariant()+'/'+[guid]::NewGuid().ToString('N'))}
    if($RunNextActivityOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-next-activity/'+[guid]::NewGuid().ToString('N'))}
    if($RunNextTerminalOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-next-terminal/'+[guid]::NewGuid().ToString('N'))}
    if($RunNextYieldOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-next-yield/'+[guid]::NewGuid().ToString('N'))}
    if($RunNextClosedOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-next-closed/'+[guid]::NewGuid().ToString('N'))}
    if($RunNextPathsOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-next-paths-'+$RunNextPathMode.ToLowerInvariant()+'/'+[guid]::NewGuid().ToString('N'))}
    if($RunCheckInActivityOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-check-in-activity/'+[guid]::NewGuid().ToString('N'))}
    if($RunCheckInClosedOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-check-in-closed/'+[guid]::NewGuid().ToString('N'))}
    if($RunCheckInRoutedOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-check-in-routed/'+[guid]::NewGuid().ToString('N'))}
    if($RunCheckInBaselineOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-check-in-baseline/'+[guid]::NewGuid().ToString('N'))}
    if($RunLocalRefillDiagnostic){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-run-refill/'+[guid]::NewGuid().ToString('N'))}
    if($RunLocalContractOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-run-contract/'+[guid]::NewGuid().ToString('N'))}
    if($RunLocalPolicyOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-run-policy/'+[guid]::NewGuid().ToString('N'))}
    if($RunLocalFaultOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-run-fault/'+[guid]::NewGuid().ToString('N'))}
    if($RunLocalYieldOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-run-yield/'+[guid]::NewGuid().ToString('N'))}
    if($RunLocalStockOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-run-stock/'+[guid]::NewGuid().ToString('N'))}
    if($RunLocalWorksheetScaleOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-run-worksheet-scale/'+[guid]::NewGuid().ToString('N'))}
    if($RunLocalWorksheetOwnerOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-run-worksheet-owner/'+[guid]::NewGuid().ToString('N'))}
    if($RunLocalWorksheetOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-run-worksheet/'+[guid]::NewGuid().ToString('N'))}
    if($RunClearPathsOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-run-clear-paths/'+[guid]::NewGuid().ToString('N'))}
    if($RunLoadPathsOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-run-load-paths/'+[guid]::NewGuid().ToString('N'))}
    if($RunCheckInPathsOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-check-in-paths-'+$RunCheckInPathMode.ToLowerInvariant()+'/'+[guid]::NewGuid().ToString('N'))}
    if($RunAllocatePathsOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-run-allocate-paths/'+[guid]::NewGuid().ToString('N'))}
    if($RunRefreshPathsOnly){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-run-refresh-paths/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionAssignmentSafety){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-assignment-safety/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionAssignmentPaths){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-assignment-paths/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionAssignmentSourcePaths){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-assignment-source-paths/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionRegulation){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-regulation/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionRegulationPaths){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-regulation-paths/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionDesignReadPaths){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-design-read-paths/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionRecipeStructureReleasedData){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-recipe-structure-released/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionRecipeStructurePaths){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-recipe-structure-paths/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionRecipeOrderPaths){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-recipe-order-paths/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionInstructionPaths){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-instruction-paths/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionUomPaths){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-uom-paths/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionUomPublicClose){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-uom-public-close/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionLifecyclePaths){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-lifecycle-paths/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionLifecycleNative){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-lifecycle-native/'+[guid]::NewGuid().ToString('N'))}
    if($CheckProductionLifecyclePresentation){$reportRoot=Join-Path $repo ('reports/runtime/slice4be-production-lifecycle-presentation/'+[guid]::NewGuid().ToString('N'))}
    Write-Output ('Production designer report: '+$reportRoot)
    . (Join-Path $PSScriptRoot 'Slice4beProductionDesignerProbe.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beProductionDesignerActivity.ps1')
}
if($CheckAuthReadOnly){
    if(-not $CompileEvaluationProbesForTest -or $CheckActivityEvidence -or $CheckTrackingSettings -or $CheckViewerPublishedRead -or $CheckViewerEventDetail -or $CheckViewerEventGroups -or $CheckViewerRefreshFailure -or $CheckAdminSettingsClose){throw 'Auth reads require the separate compiled Core/caller gate.'}
    $reportRoot=Join-Path $repo ('reports/runtime/slice4be-auth-read/'+[guid]::NewGuid().ToString('N'))
    Write-Output ('Auth read report: '+$reportRoot)
    . (Join-Path $PSScriptRoot 'Slice4beAuthReadOnly.ps1')
}
if($CheckReceivingReplay){
    $reportRoot=Join-Path $repo ('reports/runtime/slice4be-receiving-replay/'+[guid]::NewGuid().ToString('N'))
    Write-Output ('Receiving replay evidence: '+$reportRoot)
}
if($CheckGuideTransfer){
    $reportRoot=Join-Path $repo ('reports/runtime/slice4be-guide-transfer/'+[guid]::NewGuid().ToString('N'))
    Write-Output ('Guide transfer evidence: '+$reportRoot)
}
New-Item -ItemType Directory -Path $runRoot,$reportRoot -Force | Out-Null
$inputDeploy=$deploy
if($ViewerStartupPackageStateForTest -ne 'OriginalReadOnly'){
    $probeDeploy=Join-Path $runRoot 'startup-probe-packages'
    New-Item -ItemType Directory -Path $probeDeploy|Out-Null
    foreach($name in @('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam','invSys.Admin.xlam')){
        Copy-Item -LiteralPath (Join-Path $deploy $name) -Destination (Join-Path $probeDeploy $name)
    }
    $deploy=$probeDeploy
}
if ($CheckOperationsTrackingSettings) {
    $operationsSettingsDeploy = Join-Path $reportRoot 'operations-only'
    New-Item -ItemType Directory -Path $operationsSettingsDeploy | Out-Null
    foreach($name in @('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam')) {
        Copy-Item -LiteralPath (Join-Path $deploy $name) -Destination (Join-Path $operationsSettingsDeploy $name)
    }
    . (Join-Path $PSScriptRoot 'Slice4beOperationsTrackingSettings.ps1')
}
$results = [Collections.Generic.List[object]]::new()
$excel = $null
$guidePresentationRestartFixture = $null
$step = 'startup'
$settingsRoot = 'HKCU:\Software\VB and VBA Program Settings\invSys'
$registryBefore = @{}
if (Test-Path -LiteralPath $settingsRoot) {
    foreach ($key in @(Get-Item -LiteralPath $settingsRoot) + @(Get-ChildItem -LiteralPath $settingsRoot -Recurse)) {
        $values = @{}
        foreach ($name in $key.GetValueNames()) { $values[$name] = @($key.GetValue($name),$key.GetValueKind($name)) }
        $registryBefore[$key.Name] = $values
    }
}
function Check([string]$Name,[bool]$Passed) {
    $results.Add([pscustomobject]@{Check=$Name;Passed=$Passed})
    Write-Output ("{0}: {1}" -f $Name, $(if($Passed){'PASS'}else{'FAIL'}))
    if($CheckGuideTransfer -and $Name -clike 'Guide*' -and $Name -cne 'GuideCapture.SavedWorkbookBytesPreservedAfterClose'){Write-GuideTransferHostState $Name}
}
function Complete-ResultEvidence([string]$ReportPath,[scriptblock]$Cleanup) {
    # Preserve typed checks before fixture cleanup can encounter an external lock.
    $results | ConvertTo-Json | Set-Content -Encoding UTF8 -LiteralPath $ReportPath
    try { & $Cleanup }
    catch {
        Check 'Harness.Exception.disposable fixture cleanup' $false
        Write-Output 'Disposable fixture cleanup incomplete. Typed results retained; inspect the owned fixture after Excel closes.'
    }
    finally { $results | ConvertTo-Json | Set-Content -Encoding UTF8 -LiteralPath $ReportPath }
}
function Test-LoadedPackage([string]$Name) {
    try { return ($excel.Workbooks.Item($Name).Name -ieq $Name) }
    catch { return $false }
}
function Initialize-SettingsCapture {
    if (-not ('InvSysSettingsCapture' -as [type])) {
    Add-Type -ReferencedAssemblies System.Drawing @'
using System; using System.Drawing; using System.Runtime.InteropServices;
public static class InvSysSettingsCapture {
    static readonly System.Collections.Generic.List<string> captionFacts=new System.Collections.Generic.List<string>();
    public static string[] CaptionFacts() { return captionFacts.ToArray(); }
    public delegate bool WindowCallback(IntPtr hwnd, IntPtr state);
    [StructLayout(LayoutKind.Sequential)] public struct Rect { public int Left, Top, Right, Bottom; }
    [StructLayout(LayoutKind.Sequential)] public struct Point { public int X,Y; }
    [StructLayout(LayoutKind.Sequential)] public struct Mouse { public int X,Y; public uint Data,Flags,Time; public UIntPtr Extra; }
    [StructLayout(LayoutKind.Sequential)] public struct Input { public uint Type; public Mouse Mouse; }
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr FindWindow(string cls, string title);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hwnd, out Rect rect);
    [DllImport("user32.dll")] static extern bool GetClientRect(IntPtr hwnd,out Rect rect);
    [DllImport("user32.dll")] static extern bool ClientToScreen(IntPtr hwnd,ref Point point);
    [DllImport("user32.dll")] static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdc, uint flags);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);
    [DllImport("user32.dll")] static extern bool EnumWindows(WindowCallback callback, IntPtr state);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern bool IsWindowEnabled(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hwnd);
    [DllImport("user32.dll")] static extern IntPtr WindowFromPoint(Point point);
    [DllImport("user32.dll")] static extern IntPtr MonitorFromPoint(Point point,uint flags);
    [DllImport("user32.dll")] static extern bool GetCursorPos(out Point point);
    [DllImport("user32.dll")] static extern bool SetCursorPos(int x,int y);
    [DllImport("user32.dll")] static extern uint SendInput(uint count,Input[] inputs,int size);
    [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr window,IntPtr after,int x,int y,int width,int height,uint flags);
    [DllImport("user32.dll",EntryPoint="GetWindowLongPtrW")] static extern IntPtr GetWindowLong64(IntPtr window,int index);
    [DllImport("user32.dll",EntryPoint="GetWindowLongW")] static extern int GetWindowLong32(IntPtr window,int index);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd, System.Text.StringBuilder text, int length);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr hwnd, System.Text.StringBuilder text, int length);
    [DllImport("dwmapi.dll")] static extern int DwmGetWindowAttribute(IntPtr hwnd,int attribute,out int value,int size);
    [DllImport("user32.dll")] static extern uint GetGuiResources(IntPtr process,uint flags);
    static string WindowFact(IntPtr window) {
        uint process;GetWindowThreadProcessId(window,out process);
        var kind=new System.Text.StringBuilder(128);GetClassName(window,kind,kind.Capacity);
        Rect bounds;bool haveBounds=GetWindowRect(window,out bounds);int cloaked;
        int result=DwmGetWindowAttribute(window,14,out cloaked,4);
        return window.ToInt64()+"|"+process+"|"+kind+"|Visible="+IsWindowVisible(window)+"|Enabled="+IsWindowEnabled(window)+
            "|Minimized="+IsIconic(window)+"|Bounds="+haveBounds+","+bounds.Left+","+bounds.Top+","+bounds.Right+","+bounds.Bottom+
            "|Topmost="+IsTopmost(window)+"|Cloaked="+cloaked+"|DwmResult="+result;
    }
    static void RecordCaptionFailure(IntPtr form) {
        // Identifiers, classes, geometry and counts only: never titles or content.
        captionFacts.Add("FailureForm|"+WindowFact(form));
        captionFacts.Add("Foreground|"+WindowFact(GetForegroundWindow()));
        uint process;GetWindowThreadProcessId(form,out process);int owned=0,visible=0,order=0;
        EnumWindows((window,state)=>{
            uint current;GetWindowThreadProcessId(window,out current);
            if(current==process){owned++;if(IsWindowVisible(window))visible++;}
            if(IsWindowVisible(window) && order++<8)captionFacts.Add("ZOrder|"+(order-1)+"|"+WindowFact(window));
            return true;
        },IntPtr.Zero);
        captionFacts.Add("OwnedWindows|"+owned+"|Visible="+visible);
        if(process!=0)using(var owner=System.Diagnostics.Process.GetProcessById((int)process))
            captionFacts.Add("Resources|Handles="+owner.HandleCount+"|Gdi="+GetGuiResources(owner.Handle,0)+"|User="+GetGuiResources(owner.Handle,1));
    }
    // Window rectangles are DPI-virtualized; screen pixels and cursor targets
    // must use the same physical coordinate space. Never change process DPI.
    sealed class PhysicalPixels : IDisposable {
        readonly IntPtr previous;
        public PhysicalPixels() {
            previous=SetThreadDpiAwarenessContext(new IntPtr(-4));
            if(previous==IntPtr.Zero)throw new Exception("Physical capture coordinates unavailable.");
        }
        public void Dispose() {
            if(SetThreadDpiAwarenessContext(previous)==IntPtr.Zero)
                throw new Exception("Capture DPI context could not be restored.");
        }
    }
    public static string ForegroundIdentity(IntPtr intended) {
        IntPtr wanted=GetAncestor(intended,2), actual=GetAncestor(GetForegroundWindow(),2);
        uint wantedProcess,actualProcess;
        GetWindowThreadProcessId(wanted,out wantedProcess);GetWindowThreadProcessId(actual,out actualProcess);
        var wantedClass=new System.Text.StringBuilder(128);var actualClass=new System.Text.StringBuilder(128);
        GetClassName(wanted,wantedClass,wantedClass.Capacity);GetClassName(actual,actualClass,actualClass.Capacity);
        return wantedProcess+"|"+wanted.ToInt64()+"|"+wantedClass+"|"+actualProcess+"|"+actual.ToInt64()+"|"+actualClass;
    }
    public static IntPtr OwnedVisibleForm(string title, IntPtr ownerWindow) {
        uint ownerProcess; GetWindowThreadProcessId(ownerWindow,out ownerProcess);
        if(ownerProcess==0) return IntPtr.Zero;
        int count=0; IntPtr found=IntPtr.Zero;
        EnumWindows((hwnd,state)=>{
            uint process; GetWindowThreadProcessId(hwnd,out process);
            if(process==ownerProcess && IsWindowVisible(hwnd)) {
                var text=new System.Text.StringBuilder(512); GetWindowText(hwnd,text,text.Capacity);
                if(String.Equals(text.ToString(),title,StringComparison.Ordinal)){found=hwnd;count++;}
            }
            return true;
        },IntPtr.Zero);
        return count==1 ? found : IntPtr.Zero;
    }
    public static bool ActivateByCaptionClick(IntPtr form,IntPtr excel) {
        using(var pixels=new PhysicalPixels()) {
        captionFacts.Clear();captionFacts.Add("BeforeRaise|"+WindowFact(form));
        Rect rect;uint formProcess,excelProcess;
        GetWindowThreadProcessId(form,out formProcess);GetWindowThreadProcessId(excel,out excelProcess);
        if(formProcess==0 || formProcess!=excelProcess || !IsWindowVisible(form) || !IsWindowEnabled(form) || !GetWindowRect(form,out rect))
            throw new Exception("Owned form cannot receive capture focus.");
        if(GetAncestor(GetForegroundWindow(),2)==form)return false;
        bool wasTopmost=IsTopmost(form),raised=false;
        try {
            if(!wasTopmost) {
                if(!SetWindowPos(form,new IntPtr(-1),0,0,0,0,0x13))throw new Exception("Owned capture form could not be raised.");
                raised=true;
            }
            captionFacts.Add("AfterRaise|"+WindowFact(form));
            foreach(int offset in new int[] {36,(rect.Right-rect.Left)/3,(rect.Right-rect.Left)/2}) {
                var point=new Point {X=rect.Left+offset,Y=rect.Top+15};
                bool onMonitor=MonitorFromPoint(point,0)!=IntPtr.Zero;
                captionFacts.Add("Candidate|"+point.X+","+point.Y+"|OnMonitor="+onMonitor);
                if(!onMonitor)continue;
                var target=WindowFromPoint(point);uint targetProcess;GetWindowThreadProcessId(target,out targetProcess);
                captionFacts.Add("Hit|"+point.X+","+point.Y+"|"+WindowFact(target)+"|Root="+GetAncestor(target,2).ToInt64());
                if(targetProcess!=formProcess || GetAncestor(target,2)!=form)continue;
                if(!SetCursorPos(point.X,point.Y))throw new Exception("Capture focus cursor positioning failed.");
                Point actualCursor;
                if(!GetCursorPos(out actualCursor) || actualCursor.X!=point.X || actualCursor.Y!=point.Y)
                    throw new Exception("Capture focus cursor did not reach the verified owned point.");
                if(GetAncestor(WindowFromPoint(point),2)!=form)throw new Exception("Capture focus point became covered.");
                Input[] down={new Input {Mouse=new Mouse {Flags=2}}},up={new Input {Mouse=new Mouse {Flags=4}}};
                try {
                    if(SendInput(1,down,Marshal.SizeOf(typeof(Input)))!=1)throw new Exception("Capture focus press delivery failed.");
                } finally {
                    if(SendInput(1,up,Marshal.SizeOf(typeof(Input)))!=1)throw new Exception("Capture focus release delivery failed.");
                }
                for(int wait=0;wait<10 && GetAncestor(GetForegroundWindow(),2)!=form;wait++)System.Threading.Thread.Sleep(50);
                return true;
            }
            try {RecordCaptionFailure(form);} catch(Exception failure) {captionFacts.Add("DiagnosticUnavailable|"+failure.GetType().Name);}
            throw new Exception("No uncovered owned form caption point is available.");
        } finally {
            if(raised && (!SetWindowPos(form,new IntPtr(-2),0,0,0,0,0x13) || IsTopmost(form)!=wasTopmost))
                throw new Exception("Capture form topmost state could not be restored.");
        }
        }
    }
    static bool IsTopmost(IntPtr window) {
        long style=IntPtr.Size==8 ? GetWindowLong64(window,-20).ToInt64() : GetWindowLong32(window,-20);
        return (style & 8)!=0;
    }
    public static void Save(string title, string path) {
        SaveWindow(FindWindow(null,title),path);
    }
    public static void SaveWindow(IntPtr hwnd, string path) {
        using(var pixels=new PhysicalPixels()) {
        Rect r;
        if(hwnd==IntPtr.Zero || !GetWindowRect(hwnd,out r)) throw new Exception("Requested form window unavailable.");
        using(var bitmap=new Bitmap(r.Right-r.Left,r.Bottom-r.Top)) {
            using(var graphics=Graphics.FromImage(bitmap)) {
                var hdc=graphics.GetHdc(); bool captured;
                try { captured=PrintWindow(hwnd,hdc,2); } finally { graphics.ReleaseHdc(hdc); }
                if(!captured) graphics.CopyFromScreen(r.Left,r.Top,0,0,bitmap.Size);
            }
            bitmap.Save(path,System.Drawing.Imaging.ImageFormat.Png);
        }
        }
    }
    public static void SaveVisibleWindow(IntPtr hwnd, string path) {
        using(var pixels=new PhysicalPixels()) {
        Rect r;
        if(hwnd==IntPtr.Zero || !GetWindowRect(hwnd,out r)) throw new Exception("Requested form window unavailable.");
        SetForegroundWindow(hwnd);
        // A focused Excel form can remain behind a sibling. Raise within its
        // existing band; never turn an operator form into a topmost window.
        if(!SetWindowPos(hwnd,IntPtr.Zero,0,0,0,0,0x213))throw new Exception("Requested form could not be raised.");
        System.Threading.Thread.Sleep(200);
        if(GetAncestor(GetForegroundWindow(),2)!=GetAncestor(hwnd,2)) throw new Exception("Requested form is not in the foreground.");
        if(!ContentUncovered(hwnd))throw new Exception("Requested form content is obscured.");
        if(!GetWindowRect(hwnd,out r)) throw new Exception("Requested form window unavailable.");
        using(var bitmap=new Bitmap(r.Right-r.Left,r.Bottom-r.Top)) {
            using(var graphics=Graphics.FromImage(bitmap)) { graphics.CopyFromScreen(r.Left,r.Top,0,0,bitmap.Size); }
            bitmap.Save(path,System.Drawing.Imaging.ImageFormat.Png);
        }
        }
    }
    static bool ContentUncovered(IntPtr hwnd) {
        Rect client;
        if(!GetClientRect(hwnd,out client))return false;
        var origin=new Point {X=client.Left,Y=client.Top};
        if(!ClientToScreen(hwnd,ref origin))return false;
        int left=origin.X,top=origin.Y,right=left+client.Right-client.Left,bottom=top+client.Bottom-client.Top;
        if(right<=left || bottom<=top)return false;
        foreach(var point in new Point[] {new Point {X=left,Y=top},new Point {X=right-1,Y=top},new Point {X=left,Y=bottom-1},new Point {X=right-1,Y=bottom-1}})
            if(MonitorFromPoint(point,0)==IntPtr.Zero)return false;
        bool reached=false,clear=true;
        EnumWindows((window,state)=>{
            if(window==hwnd){reached=true;return false;}
            if(!IsWindowVisible(window) || IsIconic(window))return true;
            int cloaked;if(DwmGetWindowAttribute(window,14,out cloaked,4)==0 && cloaked!=0)return true;
            Rect cover;
            if(GetWindowRect(window,out cover) && cover.Left<right && cover.Right>left && cover.Top<bottom && cover.Bottom>top){clear=false;return false;}
            return true;
        },IntPtr.Zero);
        return reached && clear;
    }
    public static bool SaveOwnedForegroundForm(IntPtr owner, string[] titles, string path) {
        using(var pixels=new PhysicalPixels()) {
        IntPtr hwnd=GetAncestor(GetForegroundWindow(),2);
        uint ownerProcess,foregroundProcess;
        GetWindowThreadProcessId(owner,out ownerProcess);GetWindowThreadProcessId(hwnd,out foregroundProcess);
        if(ownerProcess==0 || ownerProcess!=foregroundProcess || !IsWindowVisible(hwnd))return false;
        var title=new System.Text.StringBuilder(512);GetWindowText(hwnd,title,title.Capacity);
        if(Array.IndexOf(titles,title.ToString())<0)return false;
        Rect r;if(!GetWindowRect(hwnd,out r))return false;
        using(var bitmap=new Bitmap(r.Right-r.Left,r.Bottom-r.Top)) {
            using(var graphics=Graphics.FromImage(bitmap)){graphics.CopyFromScreen(r.Left,r.Top,0,0,bitmap.Size);}
            bitmap.Save(path,System.Drawing.Imaging.ImageFormat.Png);
        }
        return true;
        }
    }
}
'@
    }
}
function CaptureFormEvidence([string]$Title,[string]$FileName,[long]$WindowHandle=0) {
    Initialize-SettingsCapture
    if($WindowHandle) { [InvSysSettingsCapture]::SaveVisibleWindow([IntPtr]$WindowHandle,(Join-Path $reportRoot $FileName)) }
    else { [InvSysSettingsCapture]::Save($Title,(Join-Path $reportRoot $FileName)) }
}
function CaptureOwnedFormByCaptionEvidence([string]$Title,[string]$FileName) {
    Initialize-SettingsCapture
    $owned=[InvSysSettingsCapture]::OwnedVisibleForm($Title,[IntPtr]$excel.Hwnd).ToInt64()
    CaptureOwnedFormEvidence $Title $FileName $owned
}
function CaptureOwnedFormEvidence([string]$Title,[string]$FileName,[long]$WindowHandle) {
    if($TraceGuideResourcesForTest){Write-GuideResourceMark ('CaptureBefore|'+$FileName)}
    try {
    Initialize-SettingsCapture
    for($attempt=1;$attempt -le 3;$attempt++){
        $owned=[InvSysSettingsCapture]::OwnedVisibleForm($Title,[IntPtr]$excel.Hwnd).ToInt64()
        if($WindowHandle -eq 0 -or $owned -ne $WindowHandle){throw 'Requested capture does not identify a unique owned visible form.'}
        if($GuideCaptureVisibleExcelForTest){
            $captureVisible=$excel.Visible
            if($captureVisible -isnot [bool]){throw 'Capture-time application visibility is unavailable.'}
            [pscustomobject]@{Image=$FileName;Attempt=$attempt;ExcelVisible=$captureVisible;FormEnabled=[InvSysSettingsCapture]::IsWindowEnabled([IntPtr]$WindowHandle);FormMinimized=[InvSysSettingsCapture]::IsIconic([IntPtr]$WindowHandle);WorkbookCount=$excel.Workbooks.Count;SavedWorkbookFixture=[bool]$GuideCaptureSavedWorkbookForTest}|
                ConvertTo-Json -Compress|Add-Content -LiteralPath (Join-Path $reportRoot 'capture-window-state.jsonl')
        }
        $activation=New-Object -ComObject WScript.Shell
        $activated=$null
        try{$activated=$activation.AppActivate($Title)}finally{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($activation)}
        try {$clicked=[InvSysSettingsCapture]::ActivateByCaptionClick([IntPtr]$WindowHandle,[IntPtr]$excel.Hwnd)}
        catch {
            # Read-only failure facts: no captions, workbook values or input fallback.
            $bounds=New-Object InvSysSettingsCapture+Rect
            $haveBounds=[InvSysSettingsCapture]::GetWindowRect([IntPtr]$WindowHandle,[ref]$bounds)
            [pscustomobject]@{Image=$FileName;Attempt=$attempt;Identity=[InvSysSettingsCapture]::ForegroundIdentity([IntPtr]$WindowHandle);BoundsAvailable=$haveBounds;CallerDpiBounds=$bounds;Enabled=[InvSysSettingsCapture]::IsWindowEnabled([IntPtr]$WindowHandle);Minimized=[InvSysSettingsCapture]::IsIconic([IntPtr]$WindowHandle);PhysicalCaptionFacts=[InvSysSettingsCapture]::CaptionFacts()}|
                ConvertTo-Json -Compress -Depth 3|Add-Content (Join-Path $reportRoot 'capture-caption-failure.jsonl')
            throw
        }
        [pscustomobject]@{Image=$FileName;Attempt=$attempt;OwnedCaptionClick=$clicked}|
            ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'capture-owned-caption.jsonl')
        if($clicked){Start-Sleep -Milliseconds 200}
        try {CaptureFormEvidence $Title $FileName $WindowHandle; return}
        catch {
            if($_.Exception.GetBaseException().Message -cne 'Requested form is not in the foreground.'){throw}
            # Window/process identifiers and classes only; never captions or workbook values.
            $activationResult=if($null -eq $activated){'Unavailable'}else{[string]$activated}
            $identity=[InvSysSettingsCapture]::ForegroundIdentity([IntPtr]$WindowHandle)
            ([DateTimeOffset]::UtcNow.ToString('o')+'|'+$FileName+'|'+$attempt+'|'+$activationResult+'|'+$identity) |
                Add-Content -LiteralPath (Join-Path $reportRoot 'capture-foreground-observations.tsv')
            if($attempt -eq 3){throw}
            Start-Sleep -Milliseconds 300
        }
    }
    } finally {
        if($TraceGuideResourcesForTest){Write-GuideResourceMark ('CaptureAfter|'+$FileName)}
    }
}
function Wait-ExcelReadyForTest([string]$Macro) {
    # Read-only readiness sampling precedes dispatch. No command is replayed.
    for($readAttempt=1;$readAttempt -le $ExcelReadyReadLimitForTest;$readAttempt++){
        $ready=$null
        try {$ready=$excel.Ready} catch {$ready=$null}
        $status=if($null -eq $ready){'Unavailable'}elseif($ready -isnot [bool]){'Invalid'}elseif($ready){'Ready'}else{'Busy'}
        [pscustomobject]@{Macro=$Macro;ReadAttempt=$readAttempt;Status=$status}|
            ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'readiness-before-dispatch.jsonl')
        if($status -ceq 'Ready'){return}
        if($status -ceq 'Invalid'){throw 'Excel readiness returned an unsupported type; macro was not dispatched.'}
        if($readAttempt -lt $ExcelReadyReadLimitForTest){Start-Sleep -Milliseconds 250}
    }
    throw 'Excel readiness remained unavailable or busy; macro was not dispatched.'
}
function Run([string]$Package,[string]$Macro,[object[]]$Values=@()) {
    $resourceBoundary=$Macro
    if($TraceGuideResourcesForTest){
        if($Macro -ceq 'modInventoryViewer.GuideDraftControlForTest' -and $Values.Count -eq 4){
            $resourceBoundary+='|'+$Values[0]+'|'+$Values[1]+'|'+$Values[2]
        }
        Write-GuideResourceMark ('Before|'+$resourceBoundary)
    }
    try {
    $name="'$Package'!$Macro"
    if($WaitForExcelReadyForTest){Wait-ExcelReadyForTest $Macro}
    if($TraceGuideViewCallsForTest){Write-GuideResourceMark ('Ready|'+$resourceBoundary)}
    $attemptLimit=1
    $retryLog='readonly-count-retries.jsonl'
    # Only the exact observational getter is eligible; commands retain one call.
    if($RetryActionPathViewCountForTest -and $Package -ceq 'invSys.Operations.xlam' -and
       $Macro -ceq 'modInventoryViewer.GuideDraftControlForTest' -and $Values.Count -eq 4 -and
       $Values[0] -ceq 'frmActionPathView' -and $Values[1] -ceq '' -and $Values[2] -ceq 'Count' -and $Values[3] -ceq ''){
        $attemptLimit=4
    }
    # These probe branches only inspect existing forms/list values. They do
    # not activate forms, invoke handlers, alter selection or read business data.
    if($RetryGuideObservationForTest -and $Package -ceq 'invSys.Operations.xlam' -and
       $Macro -ceq 'modInventoryViewer.GuideDraftControlForTest' -and $Values.Count -eq 4 -and $Values[3] -ceq '' -and
       (($Values[0] -ceq 'frmActionPathGuide' -and $Values[1] -ceq '' -and $Values[2] -ceq 'Count') -or
        ($Values[0] -ceq 'frmGuideActionPicker' -and $Values[1] -ceq 'lstGuideActions' -and $Values[2] -ceq 'Values') -or
        ($Values[0] -ceq 'frmActionPathView' -and $Values[1] -cin @('txtActionPathHowTo','txtActionPathDiagnostic') -and $Values[2] -ceq 'State'))){
        $attemptLimit=4
        $retryLog='readonly-observation-retries.jsonl'
    }
    for($attempt=1;$attempt -le $attemptLimit;$attempt++){
    try {
    switch($Values.Count) {
        0 { $excel.Run($name) }
        1 { $excel.Run($name,$Values[0]) }
        2 { $excel.Run($name,$Values[0],$Values[1]) }
        3 { $excel.Run($name,$Values[0],$Values[1],$Values[2]) }
        4 { $excel.Run($name,$Values[0],$Values[1],$Values[2],$Values[3]) }
        5 { $excel.Run($name,$Values[0],$Values[1],$Values[2],$Values[3],$Values[4]) }
        6 { $excel.Run($name,$Values[0],$Values[1],$Values[2],$Values[3],$Values[4],$Values[5]) }
        7 { $excel.Run($name,$Values[0],$Values[1],$Values[2],$Values[3],$Values[4],$Values[5],$Values[6]) }
        default { throw 'Unsupported macro argument count' }
    }
    if($attempt -gt 1){
        [pscustomobject]@{Attempt=$attempt;Recovered=$true;Macro=$Macro;Form=$Values[0];Action=$Values[2]}|
            ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot $retryLog)
    }
    return
    } catch {
        Write-Host ('Packaged call failed: '+$Macro+'; argument count='+$Values.Count)
        $codes=@()
        for($errorNode=$_.Exception;$null -ne $errorNode;$errorNode=$errorNode.InnerException){$codes+=('0x{0:X8}' -f $errorNode.HResult)}
        $failurePath = Join-Path $reportRoot 'first-call-failure.json'
        if(-not (Test-Path -LiteralPath $failurePath)) {
            $control=$null;$captured=$false;$identity='Unavailable'
            if($Macro -cin @('modInventoryViewer.GuideDraftControlForTest','modInventoryViewer.RecordingExpectationForTest') -and $Values.Count -eq 4){
                # Only fixed form/control/action names. Never field values or call arguments.
                $control=[pscustomobject]@{Form=[string]$Values[0];Control=[string]$Values[1];Action=[string]$Values[2]}
                Write-Host ('Failed form action: '+$control.Form+'/'+$control.Control+'/'+$control.Action)
            }
            if(($CheckGuidePresentation -or $CheckProductionDesignerPaths) -and ('InvSysSettingsCapture' -as [type])){
                try {
                    $identity=[InvSysSettingsCapture]::ForegroundIdentity([IntPtr]$initialExcelWindow)
                    $captured=[InvSysSettingsCapture]::SaveOwnedForegroundForm([IntPtr]$initialExcelWindow,@('Action Path view','Action Paths','Published guides','Event Tracking Settings'),(Join-Path $reportRoot 'first-control-failure.png'))
                } catch {$captured=$false}
            }
            [pscustomobject]@{
                Macro=$Macro; HResult=$_.Exception.HResult
                ExceptionHResults=$codes;FormControl=$control;OwnedForegroundCapture=$captured;ForegroundIdentity=$identity
                InitialExcelProcessIds=$initialExcelProcessIds
                LiveExcelProcesses=@(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object Id,StartTime)
            } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $failurePath
        }
        if($attempt -lt $attemptLimit -and ('0x800AC472' -cin $codes -or '0x80010001' -cin $codes)){
            $delay=250*$attempt
            [pscustomobject]@{Attempt=$attempt;Recovered=$false;Macro=$Macro;Form=$Values[0];Action=$Values[2];ExceptionHResults=$codes;DelayMs=$delay}|
                ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot $retryLog)
            Start-Sleep -Milliseconds $delay
            continue
        }
        throw
    }
    }
    } finally {
        if($TraceGuideResourcesForTest){Write-GuideResourceMark ('After|'+$resourceBoundary)}
    }
}
function Table($Workbook,[string]$Name) {
    foreach ($sheet in $Workbook.Worksheets) {
        foreach ($candidate in $sheet.ListObjects) { if($candidate.Name -eq $Name){return $candidate} }
    }
    throw "Fixture table missing: $Name"
}
function CredentialHash([string]$Secret) {
    [double]$acc=5381
    for($i=0;$i -lt $Secret.Length;$i++) { $acc=($acc*33 + [int][char]$Secret[$i] + $i+1) % 2147483647 }
    '{0:X8}' -f [int]$acc
}
function SelectTarget($Fixture,[string]$User='config-admin') {
    [void](Run 'invSys.Core.xlam' 'modRuntimeWorkbooks.SetCoreDataRootOverride' @($Fixture.Root))
    $selected = [string](Run 'invSys.Core.xlam' 'modNasConnection.SelectWarehouseTargetForAutomation' @($Fixture.Root,$Fixture.Root,'S1',$false))
    if(-not $selected.StartsWith('OK|')){
        $code=if($selected -match '^FAIL\|([0-9]+|ERROR)(?:\||$)'){$Matches[1]}else{'UNAVAILABLE'}
        throw ('Fixture target selection failed; status code='+$code+'.')
    }
    [void](Run 'invSys.Core.xlam' 'modNasConnection.SetCurrentTargetPathsForTest' @('\\fixture-host\config-command',$Fixture.Root))
    $signed = [string](Run 'invSys.Core.xlam' 'modAuth.SignInCurrentTargetForAutomation' @($User,$Fixture.Secret,''))
    if(-not $signed.StartsWith('OK|')){
        $code = if ($signed -match '^FAIL\|([0-9]+|NO_TARGET|ERROR)(?:\||$)') { $Matches[1] } else { 'UNAVAILABLE' }
        throw ('Fixture sign-in failed; status code='+$code+'.')
    }
}
function NewFixture([string]$Suffix) {
    if($preparedShippingFixtures.ContainsKey($Suffix)) {
        $prepared=$preparedShippingFixtures[$Suffix]
        $preparedShippingFixtures.Remove($Suffix)
        [void](Run 'invSys.Core.xlam' 'modRuntimeWorkbooks.SetCoreDataRootOverride' @($prepared.Root))
        return $prepared
    }
    if($preparedShippingBoundaries.ContainsKey($Suffix)){throw 'Prepared Shipping fixture was already consumed.'}
    $wh='WHD5'+[guid]::NewGuid().ToString('N').Substring(0,6).ToUpperInvariant()
    $root=Join-Path $runRoot $Suffix
    $share=Join-Path $runRoot ($Suffix+'-share')
    New-Item -ItemType Directory -Path $share -Force | Out-Null
    [void](Run 'invSys.Core.xlam' 'modRuntimeWorkbooks.SetCoreDataRootOverride' @($root))
    $creationArgs=@($wh,'Config command fixture','S1','config-admin',$root,$share)
    if($CheckReceivingReplay -and $Suffix -ceq 'a'){$creationArgs+='Training'}
    $created=[bool](Run 'invSys.Admin.xlam' 'modAdminConsole.BootstrapWarehouseLocalAdmin' $creationArgs)
    if(-not $created){throw 'Admin Generate Warehouse fixture failed.'}
    $secret=[guid]::NewGuid().ToString('N')
    $auth=$excel.Workbooks.Open((Join-Path $root ($wh+'.invSys.Auth.xlsb')),0,$false)
    $users=Table $auth 'tblUsers'
    $caps=Table $auth 'tblCapabilities'
    # Fixture credentials are text. Excel must not coerce a randomly generated
    # all-numeric value or drop its leading zero; never emit either value.
    $users.ListColumns.Item('PinHash').Range.NumberFormat='@'
    foreach($row in $users.ListRows) {
        if($row.Range.Cells.Item(1,$users.ListColumns.Item('UserId').Index).Value2 -eq 'config-admin') {
            $row.Range.Cells.Item(1,$users.ListColumns.Item('PinHash').Index).Value2=CredentialHash $secret
            if ([string]$row.Range.Cells.Item(1,$users.ListColumns.Item('PinHash').Index).Value2 -cne (CredentialHash $secret)) { throw 'Fixture credential text did not round-trip.' }
        }
    }
    foreach($identity in @('config-reader','config-producer')) {
        $row=$users.ListRows.Add()
        foreach($pair in @{UserId=$identity;DisplayName='Config fixture';PinHash=(CredentialHash $secret);Status='Active'}.GetEnumerator()) {
            $row.Range.Cells.Item(1,$users.ListColumns.Item($pair.Key).Index).Value2=$pair.Value
        }
        if ([string]$row.Range.Cells.Item(1,$users.ListColumns.Item('PinHash').Index).Value2 -cne (CredentialHash $secret)) { throw 'Fixture credential text did not round-trip.' }
        $row=$caps.ListRows.Add()
        $cap=if($identity -eq 'config-producer'){'PROD_POST'}else{'RECEIVE_POST'}
        foreach($pair in @{UserId=$identity;Capability=$cap;WarehouseId=$wh;StationId='S1';Status='Active'}.GetEnumerator()) {
            $row.Range.Cells.Item(1,$caps.ListColumns.Item($pair.Key).Index).Value2=$pair.Value
        }
    }
    if($RunPrintRecordedOnly -or $RunNextPathsOnly -or $RunCompletePathsOnly -or $RunCheckInPathsOnly -or $RunAllocatePathsOnly -or $RunRefreshPathsOnly -or $RunLoadPathsOnly -or $RunClearPathsOnly -or $RunPresentationPathsOnly -or $CheckGuideDraft -or $CheckOperationsGuidePresentation -or $CheckSettingsDiagnostics -or $CheckProductionLifecyclePresentation -or $CheckProductionInstructionPaths -or $CheckProductionUomPaths -or $CheckProductionComponentPaths -or $CheckProductionRecipeOrderPaths -or $CheckProductionRecipeStructurePaths -or $CheckProductionDesignReadPaths -or $CheckProductionRegulationPaths -or $CheckProductionClosePaths -or $CheckProcessWorksheetPaths -or $CheckProductionAssignmentPaths){
        # Explicit fixture grant: Admin bootstrap does not imply guide maintenance.
        $row=$caps.ListRows.Add()
        foreach($pair in @{UserId='config-admin';Capability='ACTION_PATH_MAINT';WarehouseId=$wh;StationId='S1';Status='Active'}.GetEnumerator()){
            $row.Range.Cells.Item(1,$caps.ListColumns.Item($pair.Key).Index).Value2=$pair.Value
        }
    }
    $auth.Save(); $auth.Close($false)
    [pscustomobject]@{Root=$root;Warehouse=$wh;Secret=$secret;Config=(Join-Path $root ($wh+'.invSys.Config.xlsb'))}
}
try {
    $excel=New-Object -ComObject Excel.Application
    $initialExcelWindow=[long]$excel.Hwnd
    $initialExcelProcessIds = @(Get-Process EXCEL -ErrorAction Stop | Select-Object -ExpandProperty Id)
    $excel.Visible=[bool]$GuideCaptureVisibleExcelForTest; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=1
    if($GuideCaptureVisibleExcelForTest){
        $visible=$excel.Visible
        Check 'Harness.GuideCaptureVisibleExcelVerified' ($visible -is [bool] -and $visible)
        if($visible -isnot [bool] -or -not $visible){throw 'Visible-Excel capture setup was not established.'}
    }
    $step='load packages'
    $packages=@{}
    foreach($name in @('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam','invSys.Admin.xlam')) {
        $packages[$name]=$excel.Workbooks.Open((Join-Path $deploy $name),0,($ViewerStartupPackageStateForTest -eq 'OriginalReadOnly'))
    }
    if ($CheckActivityEvidence) {
        Check 'Activity.ProducingPackageIdentity' (Test-Slice4bePackageIdentity $deploy)
    }
    # Test-only instrumentation in the unsaved Admin project: invokes the exact
    # existing form selection/save handlers without changing their implementation.
    $formCode=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('frmAdminSettings').CodeModule
    $formCode.AddFromString(@'
Public Function D5TestSave(ByVal keyName As String, ByVal valueText As String) As String
    Dim i As Long
    For i = 0 To mLstConfig.ListCount - 1
        If StrComp(CStr(mLstConfig.List(i, 0)), keyName, vbTextCompare) = 0 Then
            mLstConfig.ListIndex = i
            mLstConfig_Click
            mTxtConfigValue.Value = valueText
            mBtnSaveConfig_Click
            D5TestSave = mLblStatus.Caption
            Exit Function
        End If
    Next i
    Err.Raise 5, , "Fixture key missing"
End Function
'@)
    $testModule=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Add(1)
    $testModule.Name='TestD5Commands'
    $testModule.CodeModule.AddFromString(@'
Option Explicit
Private mForm As frmAdminSettings
Private mLastStatus As String
Public Sub OpenSettings()
    Set mForm = New frmAdminSettings
End Sub
Public Sub ShowSettings()
    mForm.Show vbModeless
    mForm.Repaint
End Sub
Public Sub CloseSettings()
    If Not mForm Is Nothing Then Unload mForm
    Set mForm = Nothing
End Sub
Public Function LoadedFormsForTest() As Long
    LoadedFormsForTest = VBA.UserForms.Count
End Function
Public Function SaveSettings(ByVal key As String, ByVal value As String) As Boolean
    mLastStatus = mForm.D5TestSave(key, value)
    SaveSettings = (InStr(1, mLastStatus, "saved", vbTextCompare) > 0)
End Function
Public Function LastStatus() As String
    LastStatus = mLastStatus
End Function
Public Function SaveDirect(ByVal key As String, ByVal value As String, ByVal wh As String, ByVal st As String) As Boolean
    Dim report As String
    SaveDirect = modConfig.UpdateConfigValue(key, value, report, wh, st)
End Function
Public Function PublishUom() As Boolean
    Dim rows As Variant, report As String
    rows = modUomSettings.GetUomCatalogRows()
    rows(3, 6) = Not CBool(rows(3, 6))
    PublishUom = modUomSettings.PublishUomCatalogRows(rows, report)
End Function
'@)
    $productionCode=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmProduction').CodeModule
    $productionCode.AddFromString(@'
Public Function D5UomRoundTrip(ByVal workbookName As String) As Boolean
    Dim prior As Workbook, lo As ListObject, version As Long
    On Error GoTo Done
    If Not mBuilt Then BuildLayout
    Set prior = mOperatorWorkbook
    Set mOperatorWorkbook = Application.Workbooks(workbookName)
    version = modConfig.GetLong("UomConversionCatalogVersion", 1)
    ' Prepare the Retrieve fixture independently; Edit has its own actual-handler gate.
    If Not modProductionUomCatalog.SendUomCatalogToWorksheet(mOperatorWorkbook) Then GoTo Done
    Set lo = mOperatorWorkbook.Worksheets("invSys UOM Catalog").ListObjects("tblInvSysUomCatalog")
    lo.DataBodyRange.Cells(3, 6).Value2 = Not CBool(lo.DataBodyRange.Cells(3, 6).Value2)
    lo.Parent.Activate
    lo.DataBodyRange.Cells(1, 1).Select
    mBtnUomCatalogRetrieve_Click
    D5UomRoundTrip = (mOperatorWorkbook.Worksheets("invSys UOM Catalog").ListObjects.Count = 0 _
        And modConfig.GetLong("UomConversionCatalogVersion", 1) = version + 1)
Done:
    Set mOperatorWorkbook = prior
End Function
'@)
    $productionTest=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Add(1)
    $productionTest.Name='TestD5Uom'
    $productionTest.CodeModule.AddFromString(@'
Private mHeld As frmProduction
Public Sub HoldForm()
    Set mHeld = New frmProduction
    Load mHeld
End Sub
Public Function HeldRoundTrip(ByVal workbookName As String) As Boolean
    HeldRoundTrip = mHeld.D5UomRoundTrip(workbookName)
    Unload mHeld
    Set mHeld = Nothing
End Function
Public Function RoundTrip(ByVal workbookName As String) As Boolean
    RoundTrip = frmProduction.D5UomRoundTrip(workbookName)
    Unload frmProduction
End Function
'@)
    if ($CheckGuideTransferActivity -or $GeneralSettingsOnly) { . (Join-Path $PSScriptRoot 'Slice4beTrackingSettings.ps1') }
    if ($CheckTrackingSettings -or $CheckGuideTransferActivity -or $GeneralSettingsOnly) { Install-Slice4beTrackingSettingsProbe $testModule }
    if ($CheckTrackingPolicy -or $CheckUserTrackingPolicy -or $CheckGuidePresentationAvailability -or $CheckGuideTransferActivity) {
        . (Join-Path $PSScriptRoot 'Slice4beTrackingPolicy.ps1')
        Install-Slice4beTrackingPolicyProbe $testModule $formCode
    }
    if ($CheckUserTrackingPolicy -or $CheckGuideTransferActivity -or $GeneralSettingsOnly) {
        . (Join-Path $PSScriptRoot 'Slice4beUserTrackingPolicy.ps1')
        Install-UserTrackingPolicyProbe $testModule
    }
    if ($CheckDetailProfile) {
        . (Join-Path $PSScriptRoot 'Slice4beDetailProfile.ps1')
        Install-Slice4beDetailProfileProbe $testModule $formCode
        if ($CheckDetailColumns) { Install-DetailColumnsProbe $testModule $formCode }
    }
    if ($CheckActionPathPreference) {
        . (Join-Path $PSScriptRoot 'Slice4beActionPathPreference.ps1')
        Install-Slice4beActionPathPreferenceProbe $testModule $formCode
    }
    if($CheckAdminSettingsClose){
        . (Join-Path $PSScriptRoot 'Slice4beAdminSettingsClose.ps1')
        Install-AdminSettingsDefaultCloseProbe $testModule $formCode
    }
    if($TraceSettingsOpenForTest) {
        . (Join-Path $PSScriptRoot 'Slice4beSettingsOpenTrace.ps1')
        Install-Slice4beSettingsOpenTrace $testModule $formCode (Join-Path $reportRoot 'settings-open-stages.csv')
    }
    $bootstrapCode=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modWarehouseBootstrap').CodeModule
    if($CheckShippingActivity){
        . (Join-Path $PSScriptRoot 'Slice4beShippingCatalog.ps1')
        Install-Slice4beShippingCatalogProbe
    }
    $bootstrapCode.AddFromString(@'
Public Function TestFixtureBootstrapRoots(ByVal templateRoot As String, ByVal operatorRoot As String) As String
    TestFixtureBootstrapRoots = CStr(StrComp(mBootstrapTemplateRootOverride, templateRoot, vbTextCompare) = 0) & "|" & _
        CStr(StrComp(mLocalOperatorRootOverride, operatorRoot, vbTextCompare) = 0)
End Function
'@)
    if($TraceBootstrapForTest){
        . (Join-Path $PSScriptRoot 'Slice4beBootstrapTrace.ps1')
        Install-Slice4beBootstrapTrace
    }
    if($CheckAdminUomActivity -or $CheckAdminUomExpectationChoice){Install-AdminUomActivityProbe}
    if($CheckAdminUomActivity){
        . (Join-Path $PSScriptRoot 'Slice4beShippingCatalog.ps1')
        Install-Slice4beShippingCatalogProbe
        . (Join-Path $PSScriptRoot 'Slice4beEvaluationNativeTrace.ps1')
        Compile-Slice4beEvaluationProbes
        $noForms=[long](Run 'invSys.Admin.xlam' 'TestD5Commands.LoadedFormsForTest') -eq 0
        Check 'Harness.AdminUomProbesInstalledBeforeForms' $noForms
        if(-not $noForms){throw 'Admin UOM probes must precede all forms; not product RED.'}
    }
    # Install the complete recording gate before any fixture/form activity.
    # Later helpers reuse these probes instead of editing active VBA projects.
    $script:PublishedReadProbeInstalled=$false
    $script:RecordingProbeInstalled=$false
    $script:RecordingReaderProbeInstalled=$false
    $script:RecordingOperationsProbeInstalled=$false
    $script:ShippingActivityProbeInstalled=$false
    $script:ShippingSubmissionProbeInstalled=$false
    if($CheckViewerPublishedRead -or $CheckShippingRecording){
        . (Join-Path $PSScriptRoot 'Slice4beViewerPublishedRead.ps1')
        Test-Slice4beViewerPublishedRead $null $null $true
        if($CheckActionRecording -or $CheckShippingRecording){
            . (Join-Path $PSScriptRoot 'Slice4beActionRecording.ps1')
            Test-Slice4beActionRecording $null $true
            if($CheckRecordingReader -or $CheckGuideDraft){
                . (Join-Path $PSScriptRoot 'Slice4beRecordingReader.ps1')
                Test-Slice4beRecordingReader $null $null $true
            }
            if($CheckGuideDraft -or $CheckOperationsGuidePresentation){
                . (Join-Path $PSScriptRoot 'Slice4beGuideDraft.ps1')
                Install-GuideDraftProbe
            }
            if($CheckGuideTransfer){
                . (Join-Path $PSScriptRoot 'Slice4beGuideTransfer.ps1')
                Install-GuideTransferDialogProbe
                if($CheckGuideTransferActivity){
                    . (Join-Path $PSScriptRoot 'Slice4beGuideTransferActivity.ps1')
                    Install-GuideTransferActivityProbe
                    if($CheckGuideTransferRecorded){
                        . (Join-Path $PSScriptRoot 'Slice4beGuideTransferRecordedProbe.ps1')
                        Install-GuideTransferRecordedProbe
                    }
                }
            }
            if($CheckReceivingReplay){
                . (Join-Path $PSScriptRoot 'Slice4beReceivingReplay.ps1')
                Install-ReceivingReplayProbe
            }
            if($CheckRecordingOperations){
                . (Join-Path $PSScriptRoot 'Slice4beRecordingOperations.ps1')
                Test-Slice4beRecordingOperations $true
            }
        }
        if($TraceBootstrapForTest -and $RecordingEvaluationDiagnostic){
            . (Join-Path $PSScriptRoot 'Slice4beEvaluationNativeTrace.ps1')
            Install-Slice4beEvaluationNativeTrace
        }
        if($CheckShippingRecording){Test-Slice4beShippingActivity $true}
        if($CheckBoxingActivity){
            if($CaptureEvidence){
                . (Join-Path $PSScriptRoot 'Slice4beRecordingReader.ps1')
                Test-Slice4beRecordingReader $null $null $true
            }
            . (Join-Path $PSScriptRoot 'Slice4beBoxingActivity.ps1')
            Install-Slice4beBoxingActivityProbe
            . (Join-Path $PSScriptRoot 'Slice4beBoxingPublishedRead.ps1')
            Install-Slice4beBoxingPublishedReadProbe
            . (Join-Path $PSScriptRoot 'Slice4beBoxingContext.ps1')
            Install-Slice4beBoxingContextProbe
            . (Join-Path $PSScriptRoot 'Slice4beShippingSubmission.ps1')
            Install-ShippingSubmissionProbe $packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modTS_Shipments').CodeModule
            . (Join-Path $PSScriptRoot 'Slice4beBoxingOutcomes.ps1')
            Install-Slice4beBoxingOutcomeProbes
        }
        if($CheckOwnerCommandCompletion){
            . (Join-Path $PSScriptRoot 'Slice4beAdminUomProbe.ps1')
            Install-AdminUomActivityProbe
            . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
            Install-RecordingEvaluationProbe
            . (Join-Path $PSScriptRoot 'Slice4beEvaluationContracts.ps1')
            . (Join-Path $PSScriptRoot 'Slice4beOwnerCommandCompletion.ps1')
        }
        if($TraceViewerStartupForTest){
            . (Join-Path $PSScriptRoot 'Slice4beViewerStartup.ps1')
            Install-Slice4beViewerStartupProbe
        }
        if($TraceGuideViewCallsForTest){
            . (Join-Path $PSScriptRoot 'Slice4beGuideViewTrace.ps1')
            Install-Slice4beGuideViewTrace
        }
        if($CompileEvaluationProbesForTest -or $CheckShippingRecording){
            . (Join-Path $PSScriptRoot 'Slice4beEvaluationNativeTrace.ps1')
            Compile-Slice4beEvaluationProbes
        }
        if($ViewerStartupPackageStateForTest -ne 'OriginalReadOnly'){
            . (Join-Path $PSScriptRoot 'Slice4beProbePackages.ps1')
            Save-DisposableStartupProbes $packages $probeDeploy $ViewerStartupPackageStateForTest
        }
        $noForms=[long](Run 'invSys.Admin.xlam' 'TestD5Commands.LoadedFormsForTest') -eq 0
        Check 'Harness.RecordingProbesInstalledBeforeForms' $noForms
        if(-not $noForms){throw 'A form was already loaded during initial probe setup; not product RED.'}
    }
    if($CheckProductionDesignerActivity -and $PrintSeedBoundaryForTest -cne 'BareSeed'){
        . (Join-Path $PSScriptRoot 'Slice4beShippingCatalog.ps1')
        Install-Slice4beShippingCatalogProbe
        Install-ProductionDesignerProbe
        if($CheckProductionRunLocal){
            . (Join-Path $PSScriptRoot 'Slice4beProductionDesignReadProbe.ps1')
            . (Join-Path $PSScriptRoot 'Slice4beProductionRunLocalProbe.ps1')
            . (Join-Path $PSScriptRoot 'Slice4beProductionRunLocalActivity.ps1')
            Install-ProductionDesignReadProbe
            Install-ProductionRunLocalProbe
            if($RunPrintBaselineOnly){
                . (Join-Path $PSScriptRoot 'Slice4beProductionPrintBaseline.ps1')
                Install-ProductionPrintBaselineProbe
                if($RunPrintRecordedOnly){
                    . (Join-Path $PSScriptRoot 'Slice4beProductionCheckInTerminal.ps1')
                    . (Join-Path $PSScriptRoot 'Slice4beProductionPrintRecorded.ps1')
                    Install-ProductionCheckInTerminalProbe -ControlId 'PRODUCTION_RUN_PRINT'
                }
            }
            if($RunCompleteBaselineOnly -or $RunCheckInPathsOnly -or $RunCheckInBaselineOnly -or $RunCheckInRoutedOnly -or $RunCheckInClosedOnly -or $RunCheckInActivityOnly){
                . (Join-Path $PSScriptRoot 'Slice4beProductionRunStock.ps1')
                . (Join-Path $PSScriptRoot 'Slice4beProductionRunWorksheet.ps1')
                . (Join-Path $PSScriptRoot 'Slice4beProductionCheckInBaseline.ps1')
                . (Join-Path $PSScriptRoot 'Slice4beProductionCheckInYield.ps1')
                Install-ProductionRunStockProbe
                Install-ProductionRunWorksheetProbe
                Install-ProductionCheckInBaselineProbe
                Install-ProductionCheckInYieldProbe
                if($RunCompleteBaselineOnly){
                    . (Join-Path $PSScriptRoot 'Slice4beProductionCompleteBaseline.ps1')
                    Install-ProductionCompleteBaselineProbe
                    if($RunCompleteClosedBoundary -ne 'All'){
                        . (Join-Path $PSScriptRoot 'Slice4beProductionCompleteClosureTrace.ps1')
                        Install-ProductionCompleteClosureTrace
                    }
                    if($RunCompleteActivityOnly -or $RunCompletePathsOnly){
                        . (Join-Path $PSScriptRoot 'Slice4beProductionCompleteWorksheet.ps1')
                        Install-ProductionCompleteWorksheetProbe
                        if($RunCompletePathsOnly){
                            . (Join-Path $PSScriptRoot 'Slice4beProductionCompletePaths.ps1')
                            Install-CompletePathsProbe
                        }
                        if($RunCompleteSubmissionFaultOnly){
                            . (Join-Path $PSScriptRoot 'Slice4beProductionCompleteSubmissionFault.ps1')
                            Install-ProductionCompleteSubmissionFaultProbe
                        }
                    }
                    if($RunCompletePreparationDiagnostic){
                        . (Join-Path $PSScriptRoot 'Slice4beProductionCompletePreparation.ps1')
                        Install-ProductionCompletePreparationTrace
                    }
                    if($RunNextBaselineOnly){
                        . (Join-Path $PSScriptRoot 'Slice4beProductionNextBaseline.ps1')
                        Install-ProductionNextBaselineProbe
                        if($RunNextActivityOnly -or $RunNextPathsOnly -or $RunNextYieldOnly -or $RunNextTerminalOnly){
                            . (Join-Path $PSScriptRoot 'Slice4beProductionNextActivity.ps1')
                            Install-ProductionNextActivityProbe
                            if($RunNextActivityOnly -or $RunNextYieldOnly -or $RunNextTerminalOnly){
                                . (Join-Path $PSScriptRoot 'Slice4beProductionNextPolicy.ps1')
                                Install-ProductionNextPolicyProbe
                            }
                            if($RunNextTerminalOnly){
                                . (Join-Path $PSScriptRoot 'Slice4beProductionCheckInTerminal.ps1')
                                Install-ProductionCheckInTerminalProbe -ControlId 'PRODUCTION_RUN_NEXT_BATCH'
                            }
                            if($RunNextYieldOnly){
                                . (Join-Path $PSScriptRoot 'Slice4beProductionNextYield.ps1')
                                Install-ProductionNextYieldProbe
                                if($RunNextClosedOnly){
                                    . (Join-Path $PSScriptRoot 'Slice4beProductionNextClosed.ps1')
                                    Install-ProductionNextClosedProbe
                                }
                            }
                        }
                    }
                }
                if($RunCheckInActivityOnly){
                    . (Join-Path $PSScriptRoot 'Slice4beProductionCheckInActivity.ps1')
                    Install-ProductionCheckInActivityProbe
                }
                if($RunCheckInClosedOnly -or $RunCheckInRoutedOnly){
                    . (Join-Path $PSScriptRoot 'Slice4beProductionCheckInClosed.ps1')
                    . (Join-Path $PSScriptRoot 'Slice4beProductionCheckInClosedYield.ps1')
                    Install-ProductionCheckInClosedProbe
                    Install-ProductionCheckInClosedYieldProbe
                }
                if($RunCheckInRoutedOnly){
                    . (Join-Path $PSScriptRoot 'Slice4beProductionCheckInRouted.ps1')
                    Install-ProductionCheckInRoutedProbe
                }
            }
            if($RunLocalContractOnly){
                . (Join-Path $PSScriptRoot 'Slice4beProductionRunContract.ps1')
                Install-ProductionRunContractProbe
            }
            if($RunLocalPolicyOnly){
                . (Join-Path $PSScriptRoot 'Slice4beProductionRunPresentation.ps1')
                . (Join-Path $PSScriptRoot 'Slice4beProductionRunPolicy.ps1')
                Install-ProductionRunPresentationProbe
                Install-ProductionRunPolicyProbe
            }
            if($RunAllocatePathsOnly -or $RunRefreshPathsOnly -or $RunLoadPathsOnly -or $RunClearPathsOnly -or $RunLocalFaultOnly -or $RunLocalYieldOnly -or $RunLocalStockOnly -or $RunLocalWorksheetOnly -or $RunLocalWorksheetOwnerOnly -or $RunLocalWorksheetScaleOnly){
                . (Join-Path $PSScriptRoot 'Slice4beProductionRunPresentation.ps1')
                . (Join-Path $PSScriptRoot 'Slice4beProductionRunFault.ps1')
                Install-ProductionRunPresentationProbe
                Install-ProductionRunFaultProbe
                if($RunLocalYieldOnly){
                    . (Join-Path $PSScriptRoot 'Slice4beProductionRunYield.ps1')
                    Install-ProductionRunYieldProbe
                }
                if($RunLocalStockOnly){
                    . (Join-Path $PSScriptRoot 'Slice4beProductionRunStock.ps1')
                    Install-ProductionRunStockProbe
                }
                if($RunRefreshPathsOnly -or $RunClearPathsOnly -or $RunLocalWorksheetOnly -or $RunLocalWorksheetOwnerOnly -or $RunLocalWorksheetScaleOnly){
                    . (Join-Path $PSScriptRoot 'Slice4beProductionRunWorksheet.ps1')
                    Install-ProductionRunWorksheetProbe
                }
                if($RunRefreshPathsOnly -or $RunClearPathsOnly -or $RunLocalWorksheetOwnerOnly -or $RunLocalWorksheetScaleOnly){
                    . (Join-Path $PSScriptRoot 'Slice4beProductionRunWorksheetOwner.ps1')
                    Install-ProductionRunWorksheetOwnerProbe
                }
                if($RunLocalWorksheetOwnerOnly){
                    . (Join-Path $PSScriptRoot 'Slice4beProductionRunClearYield.ps1')
                    Install-ProductionRunClearYieldProbe
                }
                if($RunLocalWorksheetScaleOnly){
                    . (Join-Path $PSScriptRoot 'Slice4beProductionRunWorksheetScale.ps1')
                    Install-ProductionRunWorksheetScaleProbe
                }
            }
        }
        if($CheckProductionRunPresentation){
            . (Join-Path $PSScriptRoot 'Slice4beProductionRunPresentation.ps1')
            Install-ProductionRunPresentationProbe
        }
        if($CheckProcessWorksheetHeaders){
            . (Join-Path $PSScriptRoot 'Slice4beProcessWorksheetHeaders.ps1')
            Install-ProcessWorksheetHeadersProbe
            if($CheckProcessWorksheetActivity){
                . (Join-Path $PSScriptRoot 'Slice4beProcessWorksheetActivity.ps1')
                Install-ProcessWorksheetActivityProbe
            }
            if($CheckProcessWorksheetPicker){
                . (Join-Path $PSScriptRoot 'Slice4beProcessWorksheetPicker.ps1')
                Install-ProcessWorksheetPickerProbe
            }
        }
        if($CheckProductionClose){
            . (Join-Path $PSScriptRoot 'Slice4beProductionCloseProbe.ps1')
            . (Join-Path $PSScriptRoot 'Slice4beProductionCloseActivity.ps1')
            Install-ProductionCloseProbe
            if($CheckInventoryQueryReadOnly){
                . (Join-Path $PSScriptRoot 'Slice4beInventoryQueryReadOnly.ps1')
                Install-InventoryQueryReadOnlyProbe
            }
        }
        if($CheckProductionUomStaging){
            . (Join-Path $PSScriptRoot 'Slice4beProductionUomStaging.ps1')
            Install-ProductionUomStagingProbe
            if($CheckProductionUomActivity){
                . (Join-Path $PSScriptRoot 'Slice4beProductionUomActivity.ps1')
                Install-ProductionUomActivityProbe
            }
            if($CheckProductionUomPublicClose){
                . (Join-Path $PSScriptRoot 'Slice4beProductionUomPublicClose.ps1')
                Install-ProductionUomPublicCloseProbe
            }
        }
        if($CheckProductionInstructions){
            . (Join-Path $PSScriptRoot 'Slice4beProductionInstructionProbe.ps1')
            . (Join-Path $PSScriptRoot 'Slice4beProductionInstructionActivity.ps1')
            Install-ProductionInstructionProbe
        }
        if($CheckProductionComponents){
            . (Join-Path $PSScriptRoot 'Slice4beProductionComponentProbe.ps1')
            . (Join-Path $PSScriptRoot 'Slice4beProductionComponentActivity.ps1')
            Install-ProductionComponentProbe
        }
        if($CheckProductionRegulation){
            . (Join-Path $PSScriptRoot 'Slice4beProductionRegulationProbe.ps1')
            . (Join-Path $PSScriptRoot 'Slice4beProductionRegulationActivity.ps1')
            Install-ProductionRegulationProbe
            if($CheckProductionRegulationPaths){
                . (Join-Path $PSScriptRoot 'Slice4beProductionRegulationPaths.ps1')
                Install-ProductionRegulationPathProbe
            }
        }
        if($CheckProductionDesignReads){
            . (Join-Path $PSScriptRoot 'Slice4beProductionDesignReadProbe.ps1')
            . (Join-Path $PSScriptRoot 'Slice4beProductionDesignReadActivity.ps1')
            Install-ProductionDesignReadProbe
            if($CheckProductionDesignReadPaths){
                . (Join-Path $PSScriptRoot 'Slice4beProductionDesignReadPaths.ps1')
                Install-ProductionDesignReadPathProbe
            }
        }
        if($CheckProductionAssignment){
            . (Join-Path $PSScriptRoot 'Slice4beProductionDesignReadProbe.ps1')
            . (Join-Path $PSScriptRoot 'Slice4beProductionAssignmentProbe.ps1')
            . (Join-Path $PSScriptRoot 'Slice4beProductionAssignmentActivity.ps1')
            Install-ProductionDesignReadProbe
            Install-ProductionAssignmentProbe
            if($CheckProductionAssignmentPaths -or $CheckProductionAssignmentSourcePaths){
                . (Join-Path $PSScriptRoot 'Slice4beProductionLifecycle.ps1')
                Install-ProductionLifecycleProbe
                if($CheckProductionAssignmentPaths){
                    . (Join-Path $PSScriptRoot 'Slice4beProductionAssignmentPaths.ps1')
                    Install-ProductionAssignmentPathProbe
                }
            }
            if($CheckProductionAssignmentSafety){
                . (Join-Path $PSScriptRoot 'Slice4beProductionLifecycle.ps1')
                . (Join-Path $PSScriptRoot 'Slice4beProductionAssignmentSafety.ps1')
                . (Join-Path $PSScriptRoot 'Slice4beProductionAssignmentPolicy.ps1')
                . (Join-Path $PSScriptRoot 'Slice4beProductionAssignmentYield.ps1')
                Install-ProductionLifecycleProbe
                Install-ProductionAssignmentPolicyProbe
                Install-ProductionAssignmentYieldProbe
            }
        }
        if($CheckProductionRecipeStructure){
            . (Join-Path $PSScriptRoot 'Slice4beProductionRecipeStructureProbe.ps1')
            . (Join-Path $PSScriptRoot 'Slice4beProductionRecipeStructureActivity.ps1')
            Install-ProductionRecipeStructureProbe
            if($CheckProductionRecipeStructureReleasedData -or $CheckProductionRecipeStructurePaths){
                . (Join-Path $PSScriptRoot 'Slice4beProductionRecipeStructureReleased.ps1')
                Install-ProductionRecipeStructureReleasedProbe
            }
            if($CheckProductionRecipeStructurePaths){
                . (Join-Path $PSScriptRoot 'Slice4beProductionRecipeStructurePaths.ps1')
                Install-ProductionRecipeStructurePathProbe
            }
        }
        if($CheckProductionRecipeOrder){
            . (Join-Path $PSScriptRoot 'Slice4beProductionRecipeOrderProbe.ps1')
            . (Join-Path $PSScriptRoot 'Slice4beProductionRecipeOrderActivity.ps1')
            Install-ProductionRecipeOrderProbe
        }
        if($CheckProductionLifecycle){
            . (Join-Path $PSScriptRoot 'Slice4beProductionLifecycle.ps1')
            Install-ProductionLifecycleProbe
        }
        if($RunNextPathsOnly -or $RunCompletePathsOnly -or $RunCheckInPathsOnly -or $RunAllocatePathsOnly -or $RunRefreshPathsOnly -or $RunLoadPathsOnly -or $RunClearPathsOnly -or $RunPresentationPathsOnly -or $CheckProductionDesignerPaths -or $CheckProductionLifecyclePaths -or $CheckProductionLifecyclePresentation -or $CheckProductionInstructionPaths -or $CheckProductionUomPaths -or $CheckProductionComponentPaths -or $CheckProductionRecipeOrderPaths -or $CheckProductionRecipeStructurePaths -or $CheckProductionDesignReadPaths -or $CheckProductionRegulationPaths -or $CheckProductionClosePaths -or $CheckProcessWorksheetPaths -or $CheckProductionAssignmentPaths -or $CheckProductionAssignmentSourcePaths){
            . (Join-Path $PSScriptRoot 'Slice4beProductionPathsProbe.ps1')
            Install-ProductionPathsProbe
        }
        if($CheckProductionLifecyclePaths){
            . (Join-Path $PSScriptRoot 'Slice4beDesignsSourceEvidence.ps1')
            Install-DesignsSourceEvidenceProbe
        }
        if($CheckProductionLifecycleNative){
            . (Join-Path $PSScriptRoot 'Slice4beProductionLifecycleNative.ps1')
            Install-ProductionLifecycleNativeProbe
        }
        if($RunPrintRecordedOnly -or $RunNextPathsOnly -or $RunCompletePathsOnly -or $RunCheckInPathsOnly -or $RunAllocatePathsOnly -or $RunRefreshPathsOnly -or $RunLoadPathsOnly -or $RunClearPathsOnly -or $RunPresentationPathsOnly -or $CheckProductionLifecyclePresentation -or $CheckProductionInstructionPaths -or $CheckProductionUomPaths -or $CheckProductionComponentPaths -or $CheckProductionRecipeOrderPaths -or $CheckProductionRecipeStructurePaths -or $CheckProductionDesignReadPaths -or $CheckProductionRegulationPaths -or $CheckProductionClosePaths -or $CheckProcessWorksheetPaths -or $CheckProductionAssignmentPaths){
            . (Join-Path $PSScriptRoot 'Slice4beGuideDraft.ps1')
            Install-GuideDraftProbe
        }
        if($RunPrintPathsDiagnostic){
            . (Join-Path $PSScriptRoot 'Slice4beGuideViewTrace.ps1')
            Install-Slice4beGuideViewTrace
        }
        . (Join-Path $PSScriptRoot 'Slice4beEvaluationNativeTrace.ps1')
        Compile-Slice4beEvaluationProbes
    }
    if(($RunCompleteBaselineOnly -or $RunPrintBaselineOnly -or $RunCheckInActivityOnly) -and $ViewerStartupPackageStateForTest -eq 'SavedCopies'){
        $noForms=[long](Run 'invSys.Admin.xlam' 'TestD5Commands.LoadedFormsForTest') -eq 0
        Check 'Harness.ProductionProbesInstalledBeforeForms' $noForms
        if(-not $noForms){throw 'Production probes must precede forms; not product RED.'}
        . (Join-Path $PSScriptRoot 'Slice4beProbePackages.ps1')
        Save-DisposableStartupProbes $packages $probeDeploy $ViewerStartupPackageStateForTest
    }
    if($CheckSettingsEditorActivity){
        . (Join-Path $PSScriptRoot 'Slice4beViewerPublishedRead.ps1')
        Test-Slice4beViewerPublishedRead $null $null $true
        . (Join-Path $PSScriptRoot 'Slice4beActionRecording.ps1')
        Test-Slice4beActionRecording $null $true
        . (Join-Path $PSScriptRoot 'Slice4beShippingCatalog.ps1')
        Install-Slice4beShippingCatalogProbe
        Install-SettingsActivityProbe
        if($GeneralSettingsOnly){Install-GeneralSettingsProbe}
        if($CheckSettingsDiagnostics){
            . (Join-Path $PSScriptRoot 'Slice4beRecordingReader.ps1')
            Test-Slice4beRecordingReader $null $null $true
            . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
            Install-RecordingEvaluationProbe
            . (Join-Path $PSScriptRoot 'Slice4beEvaluationContracts.ps1')
            . (Join-Path $PSScriptRoot 'Slice4beSettingsDiagnostics.ps1')
        }
        . (Join-Path $PSScriptRoot 'Slice4beEvaluationNativeTrace.ps1')
        Compile-Slice4beEvaluationProbes
        $noForms=[long](Run 'invSys.Admin.xlam' 'TestD5Commands.LoadedFormsForTest') -eq 0
        Check 'Harness.SettingsEditorProbesInstalledBeforeForms' $noForms
        if(-not $noForms){throw 'Settings probes must precede forms; not product RED.'}
    }
    if($CheckTrackingSettings -and $CompileEvaluationProbesForTest){
        . (Join-Path $PSScriptRoot 'Slice4beEvaluationNativeTrace.ps1')
        Compile-Slice4beEvaluationProbes
        $noForms=[long](Run 'invSys.Admin.xlam' 'TestD5Commands.LoadedFormsForTest') -eq 0
        Check 'Harness.SettingsRegressionProbesInstalledBeforeForms' $noForms
        if(-not $noForms){throw 'Settings regression probes must precede forms; not product RED.'}
    }
    if($CheckAuthReadOnly){
        Install-Slice4beAuthReadProbe
        . (Join-Path $PSScriptRoot 'Slice4beEvaluationNativeTrace.ps1')
        Compile-Slice4beEvaluationProbes
    }
    $templateRoot=Join-Path $repo 'deploy/current/templates'
    if($CheckReceivingReplay){$templateRoot=Join-Path $inputDeploy 'templates'}
    [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride' @($templateRoot))
    [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetLocalOperatorRootOverrideForAutomation' @((Join-Path $runRoot 'operators')))
    if($GuideCaptureSavedWorkbookForTest){
        . (Join-Path $PSScriptRoot 'Slice4beViewerStartup.ps1')
        Initialize-Slice4beViewerStartupWorkbook
        if($CheckGuideTransfer){Write-GuideTransferHostState 'Initialized'}
    }
    $step='Admin-generated fixtures'; Write-Output $step
    if($PrintSeedBoundaryForTest -ne 'None'){'BeforeFixtureA'|Add-Content (Join-Path $reportRoot 'print-seed-fixtures.txt')}
    $a=NewFixture 'a'
    if($PrintSeedBoundaryForTest -ne 'None'){'AfterFixtureA'|Add-Content (Join-Path $reportRoot 'print-seed-fixtures.txt')}
    $b=NewFixture 'b'
    if($CheckGuideTransfer){Write-GuideTransferHostState 'WarehousesCreated'}
    if($PrintSeedBoundaryForTest -ne 'None'){'AfterFixtureB'|Add-Content (Join-Path $reportRoot 'print-seed-fixtures.txt')}
    if($PrintSeedBoundaryForTest -in @('FreshSeed','BareSeed')){
        $step='Admin Seed before shared Config and UOM form exercises'
        if($PrintSeedBoundaryForTest -ceq 'BareSeed'){. (Join-Path $PSScriptRoot 'Slice4beProductionPrintBaseline.ps1')}
        Test-ProductionPrintBaseline $a $b $PrintSeedBoundaryForTest
    }
    if($PrepareShippingFixturesBeforeProbesForTest){
        $step='Admin-generated Shipping fixtures before Shipping probes'
        $templateRoot=Join-Path $repo 'deploy/current/templates'
        $template=Join-Path $templateRoot 'invSys.Data.Inventory.template.xlsb'
        foreach($entry in @(
            @('shipping-activity','shipping-operators'),
            @('shipping-context-other','shipping-operators'),
            @('shipping-submission','shipping-submission-operators')
        )){
            $suffix=$entry[0];$operatorRoot=Join-Path $runRoot $entry[1]
            [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetLocalOperatorRootOverrideForAutomation' @($operatorRoot))
            $roots=[string](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.TestFixtureBootstrapRoots' @($templateRoot,$operatorRoot))
            if($roots -cne 'True|True'){throw 'Prepared Shipping fixture roots are unavailable.'}
            $templateHash=Get-ShippingActivityHash $template
            $preparedShippingFixtures[$suffix]=NewFixture $suffix
            $unchanged=$templateHash -ceq (Get-ShippingActivityHash $template)
            Check ('Shipping.Prepared.'+$suffix+'.RootsAndTemplatePreserved') $unchanged
            if(-not $unchanged){throw 'Accepted template changed during Shipping fixture preparation.'}
            $preparedShippingBoundaries[$suffix]=[pscustomobject]@{Roots=$roots;TemplateHash=$templateHash}
        }
        [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetLocalOperatorRootOverrideForAutomation' @((Join-Path $runRoot 'operators')))
    }
    if($ShippingBeforeSharedFormsForTest){
        $step='Shipping handlers before shared form exercises'
        Test-Slice4beShippingActivity
        if($PrepareShippingFixturesBeforeProbesForTest){Check 'Shipping.Prepared.AllFixturesConsumedOnce' ($preparedShippingFixtures.Count -eq 0)}
        [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetLocalOperatorRootOverrideForAutomation' @((Join-Path $runRoot 'operators')))
    }
    SelectTarget $a
    if($CheckAuthReadOnly){
        $step='packaged ordinary Auth reads and explicit provisioning'
        Test-Slice4beAuthReadOnly $a $b
    }
    if($CheckAdminSettingsClose) {
        $step='real Admin Settings close and reopen'
        Test-AdminSettingsDefaultClose $a
    }
    if($CheckViewerRefreshFailure) {
        $step='packaged Viewer refresh failure'
        Test-Slice4beViewerRefreshFailure $a $b
    }
    if($CheckViewerEventDetail) {
        $step='packaged Viewer Event Detail'
        Test-Slice4beViewerEventDetail $a $b
    }
    if($CheckViewerEventGroups) {
        $step='packaged Viewer complete event groups'
        Test-Slice4beViewerEventGroups $a $b
        if($CheckViewerPublication) {
            $step='packaged Admin Events publication'
            Test-Slice4beViewerPublication $a $b
        }
    }
    if($CheckViewerPublishedRead) {
        if($CheckGuideTransfer){
            $step='packaged guide transfer entry and cancelled actions'
            Write-GuideTransferHostState 'BeforeGuideFixture'
            . (Join-Path $PSScriptRoot 'Slice4beGuideRestartFixture.ps1')
            Initialize-GuideRestartFixture $a -ForPublishedEdit
            . (Join-Path $PSScriptRoot 'Slice4beGuideTransfer.ps1')
            Test-GuideTransferEntry $a $b $guidePresentationRestartFixture.Guide
            if($CheckGuideTransferActivity){
                $step='packaged guide transfer observations'
                Test-GuideTransferActivity $a $b $guidePresentationRestartFixture.Guide
                if($CheckGuideTransferRecorded){
                    $step='packaged recorded transfers and explicit evaluation'
                    . (Join-Path $PSScriptRoot 'Slice4beGuideTransferRecorded.ps1')
                    Test-GuideTransferRecorded $a $b $guidePresentationRestartFixture.Guide
                }
                if($CheckGuideTransferPolicy){
                    $step='packaged guide transfer policy and storage'
                    . (Join-Path $PSScriptRoot 'Slice4beGuideTransferPolicy.ps1')
                    Test-GuideTransferPolicy $a $b $guidePresentationRestartFixture.Guide
                }
            }
            if($CheckGuideTransferRoundTrip){
                $step='packaged guide transfer round trip and guards'
                . (Join-Path $PSScriptRoot 'Slice4beGuideTransferRoundTrip.ps1')
                Test-GuideTransferRoundTrip $a $b $guidePresentationRestartFixture.Guide
            }
        } elseif($CheckReceivingReplay){
            $step='B0 actual Receiving recording and guide execution entry'
            Test-Slice4beReceivingReplay $a $b
        } elseif($GuideActionCurationOnly){
            $step='direct tracked-action curation through packaged handlers'
            . (Join-Path $PSScriptRoot 'Slice4beGuideActionCuration.ps1')
            Test-GuideActionCuration $a $b
        } elseif($PublishedGuideEditOnly){
            $step='focused published-guide edit through actual handlers'
            . (Join-Path $PSScriptRoot 'Slice4beGuideRestartFixture.ps1')
            Initialize-GuideRestartFixture $a -ForPublishedEdit
            . (Join-Path $PSScriptRoot 'Slice4bePublishedGuideEdit.ps1')
            Test-PublishedGuideEditFocused $a $b $guidePresentationRestartFixture
        } elseif($GuidePresentationRestartOnly){
            $step='actual guide and observed-run fixtures for paired preference restart'
            . (Join-Path $PSScriptRoot 'Slice4beGuideRestartFixture.ps1')
            Initialize-GuideRestartFixture $a
        } else {
        $step='packaged Viewer persisted publication read'
        Test-Slice4beViewerPublishedRead $a $b
        if($CheckActionRecording -and -not $TraceViewerStartupForTest) {
            $step='packaged Viewer recording lifecycle'
            . (Join-Path $PSScriptRoot 'Slice4beActionRecording.ps1')
            Test-Slice4beActionRecording $a
            if($CheckRecordingStorageBounds) {
                $step='recording journal serialized storage bounds'
                . (Join-Path $PSScriptRoot 'Slice4beRecordingStorageBounds.ps1')
                Test-Slice4beRecordingStorageBounds $a
            }
        }
        if($CheckViewerFilters) {
            $step='packaged Viewer loaded projection filters'
            . (Join-Path $PSScriptRoot 'Slice4beViewerFilters.ps1')
            Test-Slice4beViewerFilters $a
        }
        if($CheckViewerShippingState) {
            $step='packaged Viewer Shipping current-state presentation'
            . (Join-Path $PSScriptRoot 'Slice4beViewerShippingState.ps1')
            Test-Slice4beViewerShippingState $a
        }
        }
    }
    if(-not $AdminSettingsCloseOnly -and -not $CheckViewerRefreshFailure -and -not $CheckViewerEventDetail -and -not $CheckViewerEventGroups -and -not $CheckViewerPublishedRead -and $PrintSeedBoundaryForTest -notin @('FreshSeed','BareSeed')) {
    $step='unauthenticated command'
    [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
    $before=(Get-FileHash -LiteralPath $a.Config).Hash
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveDirect' @('BatchSize','777',$a.Warehouse,'S1'))
    Check 'Command.SignedOutDenied' (-not $ok -and $before -eq (Get-FileHash -LiteralPath $a.Config).Hash)
    SelectTarget $a 'config-reader'
    $before=(Get-FileHash -LiteralPath $a.Config).Hash
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveDirect' @('BatchSize','778',$a.Warehouse,'S1'))
    Check 'Command.MissingCapabilityDenied' (-not $ok -and $before -eq (Get-FileHash -LiteralPath $a.Config).Hash)
    SelectTarget $a
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
    if ($CheckTrackingSettings) {
        $step='packaged Settings tracking surface'
        Test-Slice4beTrackingSettingsSurface $a
        if ($CheckTrackingPolicy) { Test-Slice4beTrackingPolicy $a $b }
        if ($CheckUserTrackingPolicy) { Test-UserTrackingPolicy $a $b }
        if ($CheckDetailProfile) { Test-Slice4beDetailProfile $a $b }
        if ($CheckDetailColumns) { Test-DetailColumns $a }
        if ($CheckActionPathPreference) { Test-Slice4beActionPathPreference $a $b }
    }
    if ($CheckActivityEvidence) { $activityBefore = @(Get-Slice4beActivityFiles $a) }
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('BatchSize','601'))
    if($CheckDetailProfile -and -not $ok){
        $diagnostic=[string](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailGeneralDiagnostic')
        if($diagnostic -notmatch '^-?[0-9]+\|-?[0-9]+$'){throw 'Invalid fixed General diagnostic'}
        Write-Output ('General save error|stage: '+$diagnostic)
    }
    Check 'Settings.RealSaveHandler' ($ok -and [long](Run 'invSys.Core.xlam' 'modConfig.GetLong' @('BatchSize',0)) -eq 601)
    if ($CheckActivityEvidence) {
        if (-not $ok) { throw 'Activity fixture Settings action failed.' }
        Test-Slice4beObservedAction $a $activityBefore 'ADMIN_SETTINGS_SAVE_VALUE' `
            'CONFIG_SAVE_REQUESTED' 'CONFIG_SAVE_COMPLETED' 'Changed' 'Info' `
            'config-admin' 'Activity.AdminSettings'
        Test-Slice4beUnavailableStore $a
    }
    if($CaptureEvidence){
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.ShowSettings')
        Start-Sleep -Milliseconds 300
        CaptureOwnedFormByCaptionEvidence 'invSys Settings' 'settings-save.png'
    }
    SelectTarget $b
    $beforeA=(Get-FileHash -LiteralPath $a.Config).Hash; $beforeB=(Get-FileHash -LiteralPath $b.Config).Hash
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('BatchSize','602'))
    Check 'Settings.StaleCapturedTargetDenied' (-not $ok -and $beforeA -eq (Get-FileHash -LiteralPath $a.Config).Hash -and $beforeB -eq (Get-FileHash -LiteralPath $b.Config).Hash)
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
    SelectTarget $a
    $before=(Get-FileHash -LiteralPath $a.Config).Hash
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveDirect' @('BatchSize','not-a-number',$a.Warehouse,'S1'))
    Check 'Command.InvalidTypeNoWrite' (-not $ok -and $before -eq (Get-FileHash -LiteralPath $a.Config).Hash)
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveDirect' @('WarehouseId','different',$a.Warehouse,'S1'))
    Check 'Command.IdentityImmutable' (-not $ok -and $before -eq (Get-FileHash -LiteralPath $a.Config).Hash)
    $step='read non-mutation'
    $cfg=$excel.Workbooks.Open($a.Config,0,$false)
    $table=Table $cfg 'tblWarehouseConfig'
    $column=$table.ListColumns.Add(); $column.Name='Operator Extra'; $column.DataBodyRange.Value2='preserve'
    $table.ListColumns.Item('Timezone').Delete()
    $cfg.Save(); $count=$table.ListColumns.Count
    $ok=[bool](Run 'invSys.Core.xlam' 'modConfig.LoadConfig' @($a.Warehouse,'S1'))
    Check 'Read.OptionalHeaderNotRepaired' ($ok -and $table.ListColumns.Count -eq $count -and $cfg.Saved)
    Check 'Read.UnknownColumnPreserved' ($table.ListColumns.Item('Operator Extra').DataBodyRange.Cells.Item(1,1).Value2 -eq 'preserve')
    $cfg.Close($false)
    SelectTarget $a 'config-producer'
    $version=[long](Run 'invSys.Core.xlam' 'modConfig.GetLong' @('UomConversionCatalogVersion',1))
    if ($CheckActivityEvidence) { $activityBefore = @(Get-Slice4beActivityFiles $a) }
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.PublishUom')
    Check 'Production.ValidatedUomRouteRetained' ($ok -and [long](Run 'invSys.Core.xlam' 'modConfig.GetLong' @('UomConversionCatalogVersion',1)) -eq $version+1)
    if ($CheckActivityEvidence) {
        $activityAfter = @(Get-Slice4beActivityFiles $a)
        Check 'Activity.DirectServiceIsNotUserControl' (@($activityAfter | Where-Object { $_ -notin $activityBefore }).Count -eq 0)
    }
    $before=(Get-FileHash -LiteralPath $a.Config).Hash
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveDirect' @('BatchSize','603',$a.Warehouse,'S1'))
    Check 'Production.ArbitraryConfigDenied' (-not $ok -and $before -eq (Get-FileHash -LiteralPath $a.Config).Hash)
    $stage=$excel.Workbooks.Add()
    if ($CheckActivityEvidence) { $activityBefore = @(Get-Slice4beActivityFiles $a) }
    $ok=[bool](Run 'invSys.Operations.xlam' 'TestD5Uom.RoundTrip' @($stage.Name))
    Check 'Production.RealUomRetrieveHandler' $ok
    if ($CheckActivityEvidence) {
        if (-not $ok) { throw 'Activity fixture Production action failed.' }
        Test-Slice4beObservedAction $a $activityBefore 'PRODUCTION_UOM_RETRIEVE' `
            'UOM_RETRIEVE_REQUESTED' 'UOM_RETRIEVE_COMPLETED' 'Changed' 'Info' `
            'config-producer' 'Activity.ProductionRetrieve'
    }
    $stage.Close($false)
    SelectTarget $a 'config-reader'
    $before=(Get-FileHash -LiteralPath $a.Config).Hash
    $stage=$excel.Workbooks.Add()
    if ($CheckActivityEvidence) { $activityBefore = @(Get-Slice4beActivityFiles $a) }
    $ok=[bool](Run 'invSys.Operations.xlam' 'TestD5Uom.RoundTrip' @($stage.Name))
    Check 'Production.DeniedRetrievePreservesStaging' (-not $ok -and $stage.Worksheets.Item('invSys UOM Catalog').ListObjects.Count -eq 1 -and $before -eq (Get-FileHash -LiteralPath $a.Config).Hash)
    if ($CheckActivityEvidence) {
        if ($ok) { throw 'Activity fixture denied action unexpectedly succeeded.' }
        Test-Slice4beObservedAction $a $activityBefore 'PRODUCTION_UOM_RETRIEVE' `
            'UOM_RETRIEVE_REQUESTED' 'UOM_RETRIEVE_DENIED' 'Unchanged' 'Blocked' `
            'config-reader' 'Activity.ProductionDenied'
    }
    $stage.Close($false)
    SelectTarget $a
    $cfg=$excel.Workbooks.Open($a.Config,0,$true)
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveDirect' @('BatchSize','604',$a.Warehouse,'S1'))
    Check 'Command.ReadOnlyDenied' (-not $ok -and $cfg.Saved)
    $cfg.Close($false)
    $cfg=$excel.Workbooks.Open($a.Config,0,$false)
    $table=Table $cfg 'tblWarehouseConfig'
    $extra=$table.ListColumns.Item('Operator Extra').DataBodyRange.Cells.Item(1,1)
    $extra.Value2='unsaved user edit'
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveDirect' @('BatchSize','605',$a.Warehouse,'S1'))
    Check 'Command.DirtyWorkbookPreserved' (-not $ok -and -not $cfg.Saved -and $extra.Value2 -eq 'unsaved user edit')
    $cfg.Close($false)
    $before=(Get-FileHash -LiteralPath $a.Config).Hash
    $ok=[bool](Run 'invSys.Core.xlam' 'modConfig.LoadConfig' @($a.Warehouse,'S1'))
    Check 'Read.ClosedWorkbookBytesPreserved' ($ok -and $before -eq (Get-FileHash -LiteralPath $a.Config).Hash)
    if ($CheckActivityFoundation) { Test-Slice4beActivityFoundation $a $b }
    if($CheckProductionDesignerActivity){
        $step='Production designer observations through actual packaged handlers'
        if($CheckProcessWorksheetPicker){Test-ProcessWorksheetPicker $a}
        elseif($CheckProductionAssignmentSourcePaths){
            . (Join-Path $PSScriptRoot 'Slice4beProductionLifecyclePaths.ps1')
            Test-ProductionLifecyclePaths $a -Assignment
        }
        elseif($CheckProductionAssignmentPaths){
            . (Join-Path $PSScriptRoot 'Slice4beProcessWorksheetPaths.ps1')
            . (Join-Path $PSScriptRoot 'Slice4beProductionInstructionPaths.ps1')
            Test-ProductionInstructionPaths $a -Assignment
        }
        elseif($CheckProcessWorksheetPaths){
            . (Join-Path $PSScriptRoot 'Slice4beProcessWorksheetPaths.ps1')
            . (Join-Path $PSScriptRoot 'Slice4beProductionInstructionPaths.ps1')
            Test-ProductionInstructionPaths $a -Worksheet
        }
        elseif($CheckProcessWorksheetActivity){Test-ProcessWorksheetActivity $a $b}
        elseif($CheckProcessWorksheetHeaders){Test-ProcessWorksheetHeaders $a}
        elseif($CheckProductionInstructionPaths){
            . (Join-Path $PSScriptRoot 'Slice4beProductionInstructionPaths.ps1')
            Test-ProductionInstructionPaths $a
        }
        elseif($CheckProductionUomPaths){
            . (Join-Path $PSScriptRoot 'Slice4beProductionInstructionPaths.ps1')
            Test-ProductionInstructionPaths $a -Uom
        }
        elseif($CheckProductionComponentPaths){
            . (Join-Path $PSScriptRoot 'Slice4beProductionInstructionPaths.ps1')
            Test-ProductionInstructionPaths $a -Components
        }
        elseif($CheckProductionRecipeOrderPaths){
            . (Join-Path $PSScriptRoot 'Slice4beProductionInstructionPaths.ps1')
            Test-ProductionInstructionPaths $a -RecipeOrder
        }
        elseif($CheckProductionDesignReadPaths){
            . (Join-Path $PSScriptRoot 'Slice4beProductionInstructionPaths.ps1')
            Test-ProductionInstructionPaths $a -DesignReads
        }
        elseif($CheckProductionRegulationPaths){
            . (Join-Path $PSScriptRoot 'Slice4beProductionInstructionPaths.ps1')
            Test-ProductionInstructionPaths $a -Regulation
        }
        elseif($CheckProductionClose){
            if($CheckProductionClosePaths){
                . (Join-Path $PSScriptRoot 'Slice4beProductionInstructionPaths.ps1')
                Test-ProductionInstructionPaths $a -CloseActions
            }
            elseif(-not $InventoryQueriesOnly){Test-ProductionCloseActivity $a $b}
            if($CheckInventoryQueryReadOnly){try{Test-InventoryQueryReadOnly $b}finally{SelectTarget $a}}
        }
        elseif($CheckProductionRunLocal){
            if($RunPrintBaselineOnly){
                if($RunPrintRecordedOnly){Test-ProductionPrintRecorded $a $b -GuideDiagnostic:$RunPrintPathsDiagnostic}
                else{Test-ProductionPrintBaseline $a $b $PrintSeedBoundaryForTest}
            }
            elseif($RunCheckInPathsOnly){
                . (Join-Path $PSScriptRoot 'Slice4beProductionInstructionPaths.ps1')
                Test-ProductionInstructionPaths $a -RunCheckIn -CheckInMode $RunCheckInPathMode
            }
            elseif($RunNextPathsOnly){
                . (Join-Path $PSScriptRoot 'Slice4beProductionInstructionPaths.ps1')
                Test-ProductionInstructionPaths $a -RunNext -NextMode $RunNextPathMode
            }
            elseif($RunNextBaselineOnly){Test-ProductionNextBaseline $a $b}
            elseif($RunCompleteBaselineOnly){Test-ProductionCompleteBaseline $a $b -PreparationDiagnostic:$RunCompletePreparationDiagnostic -PreparationPrelude $RunCompletePreparationPrelude -SavedDecoy:$RunCompleteSavedDecoyForTest -VisibleHost:$RunCompleteVisibleHostForTest -Activity:$RunCompleteActivityOnly -SubmissionFaultOnly:$RunCompleteSubmissionFaultOnly -StatusOnly:$RunCompleteStatusOnly -Paths:$RunCompletePathsOnly -PathMode $RunCompletePathMode -ClosedBoundary $RunCompleteClosedBoundary}
            elseif($RunCheckInActivityOnly){Test-ProductionCheckInActivity $a $b}
            elseif($RunCheckInClosedOnly){Test-ProductionCheckInClosed $a $b}
            elseif($RunCheckInRoutedOnly){Test-ProductionCheckInRouted $a $b}
            elseif($RunCheckInBaselineOnly){Test-ProductionCheckInBaseline $a $b}
            elseif($RunAllocatePathsOnly){
                . (Join-Path $PSScriptRoot 'Slice4beProductionInstructionPaths.ps1')
                Test-ProductionInstructionPaths $a -RunAllocate
            }
            elseif($RunRefreshPathsOnly){
                . (Join-Path $PSScriptRoot 'Slice4beProductionInstructionPaths.ps1')
                Test-ProductionInstructionPaths $a -RunRefresh
            }
            elseif($RunLoadPathsOnly){
                . (Join-Path $PSScriptRoot 'Slice4beProductionInstructionPaths.ps1')
                Test-ProductionInstructionPaths $a -RunLoad
            }
            elseif($RunClearPathsOnly){
                . (Join-Path $PSScriptRoot 'Slice4beProductionInstructionPaths.ps1')
                Test-ProductionInstructionPaths $a -RunClear
            }
            elseif($RunLocalContractOnly){Test-ProductionRunContract}
            elseif($RunLocalPolicyOnly){Test-ProductionRunPolicy $a $b}
            elseif($RunLocalFaultOnly){Test-ProductionRunFault $a $b}
            elseif($RunLocalYieldOnly){Test-ProductionRunFault $a $b -YieldOnly}
            elseif($RunLocalStockOnly){Test-ProductionRunFault $a $b -StockOnly}
            elseif($RunLocalWorksheetOnly){Test-ProductionRunFault $a $b -WorksheetOnly}
            elseif($RunLocalWorksheetOwnerOnly){Test-ProductionRunFault $a $b -WorksheetOwnerOnly}
            elseif($RunLocalWorksheetScaleOnly){Test-ProductionRunFault $a $b -WorksheetScaleOnly}
            else{Test-ProductionRunLocalActivity $a $b}
        }
        elseif($RunPresentationPathsOnly){
            . (Join-Path $PSScriptRoot 'Slice4beProductionInstructionPaths.ps1')
            Test-ProductionInstructionPaths $a -RunPresentation
        }
        elseif($CheckProductionRunPresentation){Test-ProductionRunPresentation $a $b}
        elseif($CheckProductionRegulation){Test-ProductionRegulationActivity $a $b}
        elseif($CheckProductionDesignReads){Test-ProductionDesignReadActivity $a $b}
        elseif($CheckProductionAssignment){
            if($CheckProductionAssignmentSafety){Test-ProductionAssignmentSafety $a $b}
            else{Test-ProductionAssignmentActivity $a $b}
        }
        elseif($CheckProductionComponents){Test-ProductionComponentActivity $a $b}
        elseif($CheckProductionRecipeOrder){Test-ProductionRecipeOrderActivity $a $b}
        elseif($CheckProductionRecipeStructurePaths){
            . (Join-Path $PSScriptRoot 'Slice4beProductionInstructionPaths.ps1')
            Test-ProductionInstructionPaths $a -RecipeStructure
        }
        elseif($CheckProductionRecipeStructureReleasedData){Test-ProductionRecipeStructureReleased $a}
        elseif($CheckProductionRecipeStructure){Test-ProductionRecipeStructureActivity $a $b}
        elseif($CheckProductionUomPublicClose){Test-ProductionUomPublicClose $a}
        elseif($CheckProductionUomStaging){
            Test-ProductionUomStaging $a
            if($CheckProductionUomActivity){Test-ProductionUomActivity $a $b}
        }
        elseif($CheckProductionInstructions){Test-ProductionInstructionActivity $a $b}
        elseif($CheckProductionLifecyclePaths){
            Test-DesignsSourceEvidence
            . (Join-Path $PSScriptRoot 'Slice4beProductionLifecyclePaths.ps1')
            Test-ProductionLifecyclePaths $a
        }
        elseif($CheckProductionLifecycleNative){Test-ProductionLifecycleNative $a}
        elseif($CheckProductionLifecyclePresentation){
            . (Join-Path $PSScriptRoot 'Slice4beProductionLifecyclePresentation.ps1')
            Test-ProductionLifecyclePresentation $a
        }
        elseif($CheckProductionLifecycle){Test-ProductionLifecycle $a $b}
        else {Test-ProductionDesignerActivity $a $b}
    }
    if ($CheckAdminUomActivity) {
        $step='Admin UOM activity through actual packaged form handlers'
        Test-AdminUomCatalog
        Test-AdminUomActivity $a $b
        Test-AdminUomPublication $a
        Test-AdminUomTracking $a
    }
    if($CheckSettingsEditorActivity){
        $step='Settings observations through actual packaged callbacks'
        if($GeneralSettingsOnly){
            if(-not $GeneralSettingsPolicyOnly){Test-GeneralSettings $a $b}
            . (Join-Path $PSScriptRoot 'Slice4beGeneralSettingsPolicy.ps1')
            Test-GeneralSettingsPolicy $a $b
        }else{Test-SettingsEditorActivity $a $b}
    }
    if ($CheckShippingActivity) {
        $step='Shipping catalog and source-reference contract'
        Test-Slice4beShippingCatalog
        Test-Slice4beShippingCatalogPolicy $a
        $step='Shipping activity through packaged form handlers'
        if(-not $ShippingBeforeSharedFormsForTest){Test-Slice4beShippingActivity}
        SelectTarget $a
    }
    if ($CheckReceivingActivity) {
        $step='Receiving activity through packaged form handlers'
        Test-Slice4beReceivingActivity
        SelectTarget $a
    }
    $cfg=$excel.Workbooks.Open($a.Config,0,$false)
    $table=Table $cfg 'tblWarehouseConfig'
    $table.ListColumns.Item('WarehouseName').Delete()
    $cfg.Save(); $cfg.Close($false)
    $before=(Get-FileHash -LiteralPath $a.Config).Hash
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveDirect' @('BatchSize','607',$a.Warehouse,'S1'))
    Check 'Command.OtherRequiredHeaderDenied' (-not $ok -and $before -eq (Get-FileHash -LiteralPath $a.Config).Hash)
    $cfg=$excel.Workbooks.Open($a.Config,0,$false)
    $table=Table $cfg 'tblWarehouseConfig'
    $column=$table.ListColumns.Add(); $column.Name='WarehouseName'; $column.DataBodyRange.Value2='Config command fixture'
    $cfg.Save(); $cfg.Close($false)
    $cfg=$excel.Workbooks.Open($a.Config,0,$false)
    $table=Table $cfg 'tblWarehouseConfig'
    $table.ListColumns.Item('WarehouseId').Delete()
    $cfg.Save(); $count=$table.ListColumns.Count
    $ok=[bool](Run 'invSys.Core.xlam' 'modConfig.LoadConfig' @($a.Warehouse,'S1'))
    Check 'Read.RequiredHeaderNotRepaired' (-not $ok -and $table.ListColumns.Count -eq $count -and $cfg.Saved)
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveDirect' @('BatchSize','606',$a.Warehouse,'S1'))
    Check 'Command.MissingIdentityDenied' (-not $ok -and $table.ListColumns.Count -eq $count -and $cfg.Saved)
    $cfg.Close($false)
    if($CheckActionPathPreference -and $preferenceSaveVerified) {
        $step = 'personal preference Excel restart'
        Test-ActionPathPreferenceRestart $b
    }
    }
    if($CheckGuidePresentationRestart){
        $step='paired preference through fresh Operations-only Excel'
        if($null -eq $guidePresentationRestartFixture){throw 'Actual guide/run restart fixture was not established.'}
        . (Join-Path $PSScriptRoot 'Slice4beGuidePresentationRestart.ps1')
        Test-GuidePresentationRestart $guidePresentationRestartFixture
    }
}
catch {
    Check ('Harness.Exception.'+$step) $false
    $message=$_.Exception.Message -replace '(?i)[A-Z]:\\[^\r\n"'']+', '<path>'
    Write-Output $message
}
finally {
    if($null -ne $excel) {
        if($GuideCaptureSavedWorkbookForTest){
            try {
                if($CheckGuideTransfer){Write-GuideTransferHostState 'BeforeCleanup'}
                Check 'GuideCapture.SavedWorkbookIdentityPreserved' ($null -ne $script:ViewerStartupWorkbook -and $script:ViewerStartupWorkbook.FullName -ceq $script:ViewerStartupWorkbookPath -and $script:ViewerStartupWorkbook.Saved -is [bool] -and $script:ViewerStartupWorkbook.Saved)
                Close-Slice4beViewerStartupWorkbook 'GuideCapture.SavedWorkbookBytesPreservedAfterClose'
            } catch {Check 'Harness.Exception.guide capture workbook cleanup' $false}
        }
        try {
            if(Test-LoadedPackage 'invSys.Admin.xlam') {
                [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
            }
        } catch {}
        $hasSettingsHostEvidence=$CheckSettingsDiagnostics -and ($null -ne (Get-Command Write-SettingsHostEvidence -ErrorAction SilentlyContinue))
        try {
            foreach($book in @($excel.Workbooks)){ $book.Close($false) }
            if($hasSettingsHostEvidence){Write-SettingsHostEvidence 'BeforeQuit'}
            $excel.Quit()
            if($hasSettingsHostEvidence){Write-SettingsHostEvidence 'QuitReturned' $false}
        } catch {}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
        if($hasSettingsHostEvidence){Write-SettingsHostEvidence 'FinalReleaseReturned' $false}
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if(Test-Path -LiteralPath $settingsRoot) {
        foreach($key in @(Get-Item -LiteralPath $settingsRoot) + @(Get-ChildItem -LiteralPath $settingsRoot -Recurse)) {
            $path='Registry::'+$key.Name
            foreach($name in $key.GetValueNames()) {
                if(-not $registryBefore.ContainsKey($key.Name) -or -not $registryBefore[$key.Name].ContainsKey($name)) {
                    Remove-ItemProperty -LiteralPath $path -Name $name
                }
            }
        }
    }
    foreach($path in $registryBefore.Keys) {
        foreach($name in $registryBefore[$path].Keys) {
            $saved=$registryBefore[$path][$name]
            New-ItemProperty -LiteralPath ('Registry::'+$path) -Name $name -Value $saved[0] -PropertyType $saved[1] -Force | Out-Null
        }
    }
    $reportName=$Phase.ToLowerInvariant()+'.json'
    if($RunCompletePreparationDiagnostic){$reportName='diagnostic-preparation-'+$reportName}
    if($TraceReceivingRunCloseForTest){$reportName='diagnostic-receiving-close-'+$reportName}
    if($SettingsSafetyOnly){$reportName='diagnostic-settings-safety-'+$reportName}
    if($CheckShippingRecording){$reportName='shipping-recording-'+$reportName}
    if($CheckBoxingActivity){$reportName='boxing-activity-'+$reportName}
    if($OwnerCommandCompletionOnly){$reportName='owner-completion-'+$reportName}
    if($ViewerPublicationOnly){$reportName='diagnostic-publication-'+$reportName}
    if ($ShippingSubmissionOnly) { $reportName='diagnostic-submission-'+$reportName }
    if ($TraceBootstrapForTest) { $reportName='diagnostic-bootstrap-'+$reportName }
    if ($TraceSettingsOpenForTest) { $reportName='diagnostic-settings-open-'+$reportName }
    if ($ShippingBeforeSharedFormsForTest) { $reportName='diagnostic-shipping-first-'+$reportName }
    if ($PrepareShippingFixturesBeforeProbesForTest) { $reportName='diagnostic-prepared-fixtures-'+$reportName }
    if ($CheckReceivingNavigationActivity) { $reportName='navigation-'+$reportName }
    if ($ReceivingNavigationOnly) { $reportName='diagnostic-'+$reportName }
    if ($CheckReceivingSurfaceCoverage) { $reportName='surface-'+$Phase.ToLowerInvariant()+'.json' }
    if ($ReceivingSurfaceOnly) { $reportName='diagnostic-'+$reportName }
    if ($CheckReceivingLauncherDenial) { $reportName='launcher-denial-'+$Phase.ToLowerInvariant()+'.json' }
    if ($ReceivingLauncherDenialOnly) { $reportName='diagnostic-'+$reportName }
    if ($ReceivingLifecycleOnly) { $reportName='lifecycle-only-'+$reportName }
    if ($LifecycleDiagnostic -ne 'None') { $reportName='diagnostic-'+$LifecycleDiagnostic.ToLowerInvariant()+'.json' }
    if($RecordingEvaluationDiagnostic){$reportName='diagnostic-evaluation-'+$reportName}
    Complete-ResultEvidence (Join-Path $reportRoot $reportName) {
        # Runtime credentials stay only in disposable generated authority fixtures.
        $resolved=[IO.Path]::GetFullPath($runRoot)
        $temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
        if(-not $recordingFixtureTransferred -and $resolved.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase) -and (Split-Path $resolved -Leaf) -like 'invsys-config-command-*') {
            Remove-Item -LiteralPath $resolved -Recurse -Force
        }
    }
    if($null -ne $recordingHandoff){$recordingHandoff.Dispose()}
}
$failed=@($results | Where-Object { -not $_.Passed }).Count
Write-Output ("$Phase : {0} passed, {1} failed" -f ($results.Count-$failed),$failed)
if($failed -gt 0){exit 1}
