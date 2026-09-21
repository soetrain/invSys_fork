[CmdletBinding()]
param(
    [string]$RepoRoot = '',
    [string]$DeployRoot = 'deploy/current',
    [ValidateSet('RED','GREEN')][string]$Phase = 'GREEN',
    [switch]$CaptureEvidence,
    [switch]$CheckActivityEvidence,
    [switch]$CheckActivityFoundation,
    [switch]$CheckTrackingSettings,
    [switch]$CheckTrackingPolicy,
    [switch]$CheckDetailProfile,
    [switch]$CheckActionPathPreference,
    [switch]$CheckOperationsTrackingSettings,
    [switch]$CheckAdminSettingsClose,
    [switch]$AdminSettingsCloseOnly,
    [switch]$CheckViewerRefreshFailure,
    [switch]$CheckViewerEventDetail,
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
    [switch]$CheckRecordingEvaluation,
    [switch]$CheckEvaluationContracts,
    [switch]$CheckExpectationCompatibility,
    [switch]$CheckEvaluationVisualEvidence,
    [switch]$RecordingEvaluationDiagnostic,
    [Alias('CompileViewerProbesForTest')][switch]$CompileEvaluationProbesForTest,
    [switch]$CheckExpectationEditor,
    [string]$RecordingContinuationPipeName = '',
    [switch]$CheckViewerPublication,
    [switch]$ViewerPublicationOnly,
    [switch]$CheckShippingActivity,
    [switch]$CheckShippingRecording,
    [switch]$CheckBoxingActivity,
    [switch]$TraceBootstrapForTest,
    [switch]$TraceSettingsOpenForTest,
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
if($CheckGuidePresentation){$CheckGuideEvaluation=$true}
if($CheckGuideEvaluation){$CheckGuideExpectation=$true}
if($GuideCaptureVisibleExcelForTest -and -not $TraceViewerStartupForTest -and (-not $GuideDraftOnly -or -not $CaptureGuideEvidence)){
    throw 'Visible-Excel diagnosis requires Viewer startup tracing or GuideDraftOnly and CaptureGuideEvidence.'
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
    if($ViewerStartupPackageStateForTest -ne 'SavedCopies' -or -not $GuideDraftOnly -or -not $CheckGuideExpectation -or -not $CheckViewerPublishedRead -or -not $CompileEvaluationProbesForTest){
        throw 'Package-state comparison requires startup tracing; saved copies also support the compiled guide-expectation gate.'
    }
}
if($CheckBoxingActivity){$CheckShippingRecording=$true}
if($GuideDraftOnly){$CheckGuideDraft=$true}
if($CaptureGuideEvidence){$CheckGuideLibrary=$true}
if($CheckGuideExpectation){$CheckGuideLibrary=$true}
if($CheckGuideLibrary){$CheckGuideSave=$true}
if($CheckGuideSave){$CheckGuideDraft=$true}
if($CheckEvaluationVisualEvidence){$CheckExpectationCompatibility=$true}
if($RecordingEvaluationDiagnostic){
    if($Phase -ne 'RED' -or -not $CheckExpectationCompatibility){throw 'Evaluation isolation requires RED and the full evaluation contract probes; it is not regression acceptance.'}
    if($CheckRecordingLimits -or $CheckRecordingStorageBounds -or $CheckRecordingRestart -or $CheckExpectationEditor){throw 'Run limits, storage, restart and editor-only gates separately from evaluation diagnosis.'}
    Write-Output 'DIAGNOSTIC: evaluation fixture first; full regression gate remains required.'
}
if($CompileEvaluationProbesForTest -and -not ($RecordingEvaluationDiagnostic -or $CheckEvaluationVisualEvidence -or $CheckViewerPublishedRead)){
    throw 'Instrumented project compilation requires the published Viewer or evaluation gate.'
}
if($TraceSettingsOpenForTest -and ($Phase -ne 'RED' -or -not $CheckTrackingSettings)) {
    throw 'Settings constructor tracing requires RED and the Settings checks; it is not acceptance GREEN.'
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
if ($CheckShippingRecording -and (-not $CheckShippingActivity -or $ShippingSubmissionOnly)) { throw 'Shipping recording requires the complete Shipping activity route.' }
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
if ($CheckDetailProfile -and -not $CheckTrackingSettings) { throw 'Detail profile checks require the Settings route.' }
if ($CheckDetailProfile -and -not $CheckTrackingPolicy) { throw 'Detail profile checks retain the tracking policy baseline and cancellation observer.' }
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
    public delegate bool WindowCallback(IntPtr hwnd, IntPtr state);
    [StructLayout(LayoutKind.Sequential)] public struct Rect { public int Left, Top, Right, Bottom; }
    [StructLayout(LayoutKind.Sequential)] public struct Point { public int X,Y; }
    [StructLayout(LayoutKind.Sequential)] public struct Mouse { public int X,Y; public uint Data,Flags,Time; public UIntPtr Extra; }
    [StructLayout(LayoutKind.Sequential)] public struct Input { public uint Type; public Mouse Mouse; }
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr FindWindow(string cls, string title);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hwnd, out Rect rect);
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
    [DllImport("user32.dll")] static extern bool SetCursorPos(int x,int y);
    [DllImport("user32.dll")] static extern uint SendInput(uint count,Input[] inputs,int size);
    [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr window,IntPtr after,int x,int y,int width,int height,uint flags);
    [DllImport("user32.dll",EntryPoint="GetWindowLongPtrW")] static extern IntPtr GetWindowLong64(IntPtr window,int index);
    [DllImport("user32.dll",EntryPoint="GetWindowLongW")] static extern int GetWindowLong32(IntPtr window,int index);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd, System.Text.StringBuilder text, int length);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr hwnd, System.Text.StringBuilder text, int length);
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
            foreach(int offset in new int[] {36,(rect.Right-rect.Left)/3,(rect.Right-rect.Left)/2}) {
                var point=new Point {X=rect.Left+offset,Y=rect.Top+15};
                var target=WindowFromPoint(point);uint targetProcess;GetWindowThreadProcessId(target,out targetProcess);
                if(targetProcess!=formProcess || GetAncestor(target,2)!=form)continue;
                if(!SetCursorPos(point.X,point.Y))throw new Exception("Capture focus cursor positioning failed.");
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
            throw new Exception("No uncovered owned form caption point is available.");
        } finally {
            if(raised && (!SetWindowPos(form,new IntPtr(-2),0,0,0,0,0x13) || IsTopmost(form)!=wasTopmost))
                throw new Exception("Capture form topmost state could not be restored.");
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
    public static void SaveVisibleWindow(IntPtr hwnd, string path) {
        Rect r;
        if(hwnd==IntPtr.Zero || !GetWindowRect(hwnd,out r)) throw new Exception("Requested form window unavailable.");
        SetForegroundWindow(hwnd);
        System.Threading.Thread.Sleep(200);
        if(GetAncestor(GetForegroundWindow(),2)!=GetAncestor(hwnd,2)) throw new Exception("Requested form is not in the foreground.");
        using(var bitmap=new Bitmap(r.Right-r.Left,r.Bottom-r.Top)) {
            using(var graphics=Graphics.FromImage(bitmap)) { graphics.CopyFromScreen(r.Left,r.Top,0,0,bitmap.Size); }
            bitmap.Save(path,System.Drawing.Imaging.ImageFormat.Png);
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
function CaptureOwnedFormEvidence([string]$Title,[string]$FileName,[long]$WindowHandle) {
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
        $clicked=[InvSysSettingsCapture]::ActivateByCaptionClick([IntPtr]$WindowHandle,[IntPtr]$excel.Hwnd)
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
}
function Run([string]$Package,[string]$Macro,[object[]]$Values=@()) {
    $name="'$Package'!$Macro"
    try {
    switch($Values.Count) {
        0 { $excel.Run($name) }
        1 { $excel.Run($name,$Values[0]) }
        2 { $excel.Run($name,$Values[0],$Values[1]) }
        3 { $excel.Run($name,$Values[0],$Values[1],$Values[2]) }
        4 { $excel.Run($name,$Values[0],$Values[1],$Values[2],$Values[3]) }
        5 { $excel.Run($name,$Values[0],$Values[1],$Values[2],$Values[3],$Values[4]) }
        6 { $excel.Run($name,$Values[0],$Values[1],$Values[2],$Values[3],$Values[4],$Values[5]) }
        default { throw 'Unsupported macro argument count' }
    }
    } catch {
        Write-Host ('Packaged call failed: '+$Macro+'; argument count='+$Values.Count)
        $failurePath = Join-Path $reportRoot 'first-call-failure.json'
        if(-not (Test-Path -LiteralPath $failurePath)) {
            [pscustomobject]@{
                Macro=$Macro; HResult=$_.Exception.HResult
                InitialExcelProcessIds=$initialExcelProcessIds
                LiveExcelProcesses=@(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object Id,StartTime)
            } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $failurePath
        }
        throw
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
    if(-not $selected.StartsWith('OK|')){throw 'Fixture target selection failed.'}
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
    $created=[bool](Run 'invSys.Admin.xlam' 'modAdminConsole.BootstrapWarehouseLocalAdmin' @($wh,'Config command fixture','S1','config-admin',$root,$share))
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
    if($CheckGuideDraft){
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
    mBtnUomCatalogSend_Click
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
    if ($CheckTrackingSettings) { Install-Slice4beTrackingSettingsProbe $testModule }
    if ($CheckTrackingPolicy) {
        . (Join-Path $PSScriptRoot 'Slice4beTrackingPolicy.ps1')
        Install-Slice4beTrackingPolicyProbe $testModule $formCode
    }
    if ($CheckDetailProfile) {
        . (Join-Path $PSScriptRoot 'Slice4beDetailProfile.ps1')
        Install-Slice4beDetailProfileProbe $testModule $formCode
    }
    if ($CheckActionPathPreference) {
        . (Join-Path $PSScriptRoot 'Slice4beActionPathPreference.ps1')
        Install-Slice4beActionPathPreferenceProbe $testModule $formCode
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
            if($CheckGuideDraft){
                . (Join-Path $PSScriptRoot 'Slice4beGuideDraft.ps1')
                Install-GuideDraftProbe
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
        if($TraceViewerStartupForTest){
            . (Join-Path $PSScriptRoot 'Slice4beViewerStartup.ps1')
            Install-Slice4beViewerStartupProbe
        }
        if($CompileEvaluationProbesForTest -or $CheckShippingRecording){
            . (Join-Path $PSScriptRoot 'Slice4beEvaluationNativeTrace.ps1')
            Compile-Slice4beEvaluationProbes
        }
        if($ViewerStartupPackageStateForTest -ne 'OriginalReadOnly'){
            foreach($name in @('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam','invSys.Admin.xlam')){
                $book=$packages[$name]
                if($null -eq $book -or $book.ReadOnly -isnot [bool] -or $book.ReadOnly -or
                    -not [string]::Equals($book.FullName,(Join-Path $probeDeploy $name),[StringComparison]::OrdinalIgnoreCase)){
                    throw 'Only the writable disposable probe package may be saved.'
                }
                if($ViewerStartupPackageStateForTest -eq 'SavedCopies'){
                    $book.Save()
                    if($book.Saved -isnot [bool] -or -not $book.Saved){throw 'Disposable probe package save is not verified.'}
                }
            }
            if($ViewerStartupPackageStateForTest -eq 'SavedCopies'){
                Check 'Harness.DisposableStartupProbesSaved' $true
            } else {
                $unsaved=$packages['invSys.Operations.xlam'].Saved
                if($unsaved -isnot [bool] -or $unsaved){throw 'Writable Operations probes must remain unsaved for this comparison.'}
                Check 'Harness.DisposableStartupProbesUnsaved' $true
            }
        }
        $noForms=[long](Run 'invSys.Admin.xlam' 'TestD5Commands.LoadedFormsForTest') -eq 0
        Check 'Harness.RecordingProbesInstalledBeforeForms' $noForms
        if(-not $noForms){throw 'A form was already loaded during initial probe setup; not product RED.'}
    }
    [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride' @((Join-Path $repo 'deploy/current/templates')))
    [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetLocalOperatorRootOverrideForAutomation' @((Join-Path $runRoot 'operators')))
    if($GuideCaptureSavedWorkbookForTest){
        . (Join-Path $PSScriptRoot 'Slice4beViewerStartup.ps1')
        Initialize-Slice4beViewerStartupWorkbook
    }
    $step='Admin-generated fixtures'; Write-Output $step
    $a=NewFixture 'a'; $b=NewFixture 'b'
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
    if($CheckAdminSettingsClose) {
        $step='real Admin Settings close and reopen'
        . (Join-Path $PSScriptRoot 'Slice4beAdminSettingsClose.ps1')
        Test-AdminSettingsDefaultClose $a $testModule $formCode
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
    if(-not $AdminSettingsCloseOnly -and -not $CheckViewerRefreshFailure -and -not $CheckViewerEventDetail -and -not $CheckViewerEventGroups -and -not $CheckViewerPublishedRead) {
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
        if ($CheckDetailProfile) { Test-Slice4beDetailProfile $a $b }
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
        CaptureFormEvidence 'invSys Settings' 'settings-save.png'
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
                Check 'GuideCapture.SavedWorkbookIdentityPreserved' ($null -ne $script:ViewerStartupWorkbook -and $script:ViewerStartupWorkbook.FullName -ceq $script:ViewerStartupWorkbookPath -and $script:ViewerStartupWorkbook.Saved -is [bool] -and $script:ViewerStartupWorkbook.Saved)
                Close-Slice4beViewerStartupWorkbook 'GuideCapture.SavedWorkbookBytesPreservedAfterClose'
            } catch {Check 'Harness.Exception.guide capture workbook cleanup' $false}
        }
        try {
            if(Test-LoadedPackage 'invSys.Admin.xlam') {
                [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
            }
        } catch {}
        try { foreach($book in @($excel.Workbooks)){ $book.Close($false) }; $excel.Quit() } catch {}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
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
    # Runtime credentials stay only in disposable generated authority fixtures.
    $resolved=[IO.Path]::GetFullPath($runRoot)
    $temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if(-not $recordingFixtureTransferred -and $resolved.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase) -and (Split-Path $resolved -Leaf) -like 'invsys-config-command-*') {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
    $reportName=$Phase.ToLowerInvariant()+'.json'
    if($CheckShippingRecording){$reportName='shipping-recording-'+$reportName}
    if($CheckBoxingActivity){$reportName='boxing-activity-'+$reportName}
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
    $results | ConvertTo-Json | Set-Content -Encoding UTF8 -LiteralPath (Join-Path $reportRoot $reportName)
    if($null -ne $recordingHandoff){$recordingHandoff.Dispose()}
}
$failed=@($results | Where-Object { -not $_.Passed }).Count
Write-Output ("$Phase : {0} passed, {1} failed" -f ($results.Count-$failed),$failed)
if($failed -gt 0){exit 1}
