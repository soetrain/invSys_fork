[CmdletBinding()]
param([switch]$ValidateInputOnly)
# This developer-only worker receives its generated fixture exclusively on stdin.
# It never emits input, authority values, credentials, raw exceptions or paths.
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$excel=$null; $packages=@{}; $failed=$false; $newOwner=0; $owner=0
function Check([string]$Name,[bool]$Passed) {
    if(-not $Passed){$script:failed=$true}
    [pscustomobject]@{Type='Check';Name=$Name;Passed=$Passed}|ConvertTo-Json -Compress
}
function Stage([string]$Name) {
    [pscustomobject]@{Type='Stage';Name=$Name}|ConvertTo-Json -Compress
}
function PinMap($Value) {
    $map=@{}
    foreach($property in $Value.PSObject.Properties){$map[$property.Name]=[string]$property.Value}
    return $map
}
function Run([string]$Package,[string]$Macro,[object[]]$Values=@()) {
    $name="'$Package'!$Macro"
    switch($Values.Count){
        0 {$excel.Run($name)}
        1 {$excel.Run($name,$Values[0])}
        2 {$excel.Run($name,$Values[0],$Values[1])}
        3 {$excel.Run($name,$Values[0],$Values[1],$Values[2])}
        4 {$excel.Run($name,$Values[0],$Values[1],$Values[2],$Values[3])}
        default {throw 'Unsupported worker argument count.'}
    }
}
function SelectTarget($Fixture,[string]$User='config-admin') {
    [void](Run 'invSys.Core.xlam' 'modRuntimeWorkbooks.SetCoreDataRootOverride' @($Fixture.Root))
    $selected=[string](Run 'invSys.Core.xlam' 'modNasConnection.SelectWarehouseTargetForAutomation' @($Fixture.Root,$Fixture.Root,'S1',$false))
    if(-not $selected.StartsWith('OK|')){throw 'Worker fixture target unavailable.'}
    [void](Run 'invSys.Core.xlam' 'modNasConnection.SetCurrentTargetPathsForTest' @('\\fixture-host\config-command',$Fixture.Root))
    $signed=[string](Run 'invSys.Core.xlam' 'modAuth.SignInCurrentTargetForAutomation' @($User,$Fixture.Secret,''))
    if(-not $signed.StartsWith('OK|')){throw 'Worker fixture sign-in failed.'}
}
try {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingTransfer.ps1')
    $state=(Read-RecordingStandardInput)|ConvertFrom-Json
    $expected=@('Fixture','OtherFixture','Deploy','RunRoot','ReportRoot','PackageNames','PackagePins','Probes',
        'OldExcelId','ParentControllerId','CreatorControllerId','RequireCreatorExit','Action','Sequence','PathId','AuthorityBefore','OtherBefore','JournalBefore')
    if(@($state.PSObject.Properties).Count -ne $expected.Count){throw 'Invalid worker input.'}
    foreach($name in $expected){if($null -eq $state.$name){throw 'Missing worker input.'}}
    $Fixture=$state.Fixture; $b=$state.OtherFixture; $deploy=[string]$state.Deploy; $runRoot=[string]$state.RunRoot
    $rootPath=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    $tempPath=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if(-not $rootPath.StartsWith($tempPath,[StringComparison]::OrdinalIgnoreCase) -or
        (Split-Path $runRoot -Leaf) -notlike 'invsys-config-command-*'){throw 'Worker requires a generated fixture root.'}
    foreach($fixtureRoot in @($Fixture.Root,$b.Root)){
        if(-not [IO.Path]::GetFullPath([string]$fixtureRoot).StartsWith($rootPath,[StringComparison]::OrdinalIgnoreCase)){throw 'Fixture escaped the generated root.'}
    }
    if([string]::IsNullOrEmpty([string]$Fixture.Secret)){throw 'Fixture credential unavailable.'}
    if($ValidateInputOnly){Check 'RecordingRestart.WorkerInputValidated' $true; exit 0}
    if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Excel must be closed before the fresh reader.'}
    $packageNames=@('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam','invSys.Admin.xlam')
    if(($state.PackageNames -join '|') -cne ($packageNames -join '|')){throw 'Worker package set changed.'}
    $packagePins=PinMap $state.PackagePins
    foreach($name in $packageNames){
        $path=Join-Path $deploy $name
        if(-not $packagePins.ContainsKey($path) -or (Get-FileHash -LiteralPath $path).Hash -cne $packagePins[$path]){throw 'Worker package pin changed.'}
    }
    $probes=$state.Probes; $owner=[int]$state.OldExcelId; $action=$state.Action
    $sequence=[string]$state.Sequence; $pathId=[string]$state.PathId; $reportRoot=[string]$state.ReportRoot
    $authorityBefore=PinMap $state.AuthorityBefore; $otherBefore=PinMap $state.OtherBefore; $journalBefore=PinMap $state.JournalBefore
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    Add-Type @'
using System;using System.Runtime.InteropServices;
public static class RecordingRestartProcess {
 [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
 public static uint Owner(long h){uint p;if(h==0||GetWindowThreadProcessId(new IntPtr(h),out p)==0||p==0)throw new InvalidOperationException("Excel identity unavailable.");return p;}
}
'@
    $controller=Get-CimInstance Win32_Process -Filter ('ProcessId='+$PID)
    $separate=$PID -ne [int]$state.ParentControllerId -and $controller.ParentProcessId -eq [int]$state.ParentControllerId
    Check 'RecordingRestart.FreshControllerVerified' $separate
    if(-not $separate){throw 'Reader controller identity is not independent.'}
    if($state.RequireCreatorExit){
        $creatorExited=$PID -ne [int]$state.CreatorControllerId -and -not (Get-Process -Id ([int]$state.CreatorControllerId) -ErrorAction SilentlyContinue)
        Check 'RecordingRestart.CreatorControllerExitedBeforeReader' $creatorExited
        if(-not $creatorExited){throw 'Creator controller remains alive.'}
    }
$script:excel=New-Object -ComObject Excel.Application
$excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=1
foreach($name in $packageNames){$packages[$name]=$excel.Workbooks.Open((Join-Path $deploy $name),0,$true)}
$current=@(Get-Process EXCEL -ErrorAction Stop)
$newOwner=[RecordingRestartProcess]::Owner([long]$excel.Hwnd)
$distinct=$current.Count -eq 1 -and $current[0].Id -eq $newOwner -and $newOwner -ne $owner
Check 'RecordingRestart.FreshExcelProcessVerified' $distinct
if(-not $distinct){throw 'Fresh isolated Excel identity is unavailable.'}
[pscustomobject]@{OldProcessId=$owner;NewProcessId=$newOwner;ParentControllerId=[int]$state.ParentControllerId;CreatorControllerId=[int]$state.CreatorControllerId;
    FreshControllerId=$PID;InterruptionWasDeliberate=$true;FreshControllerVerified=$separate}|
    ConvertTo-Json|Set-Content -LiteralPath (Join-Path $reportRoot 'recording-restart-processes.json')
Stage 'Pristine Viewer'
SelectTarget $Fixture 'config-admin'
$pristine=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RunInventoryViewerActionForTest')
$openedPristine=$pristine.StartsWith('OK|')
Check 'RecordingRestart.PristinePackagedViewerLaunches' $openedPristine
if(-not $openedPristine){throw 'Pristine fresh-process Viewer baseline is unavailable.'}
CloseRecordingViewer
Stage 'Install drivers'
foreach($probe in $probes){
    if($probe.Module -eq 'TestD5Commands'){
        $module=$packages[$probe.Package].VBProject.VBComponents.Add(1)
        $module.Name=$probe.Module
        $module.CodeModule.AddFromString("Option Explicit`r`nPrivate mForm As frmAdminSettings`r`nPrivate mLastStatus As String")
    } else {$module=$packages[$probe.Package].VBProject.VBComponents.Item($probe.Module)}
    $module.CodeModule.AddFromString($probe.Text)
}
# Minimal fresh-process driver enters the same Events and Refresh handlers.
$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmInventoryViewer').CodeModule.AddFromString(@'
Public Function PublishedReadActionForTest(ByVal action As String, Optional ByVal expected As String = "") As Boolean
If action = "Events" Then mCboEventRange.Value = "All": mTabs.Value = 1
mBtnRefresh_Click
PublishedReadActionForTest = True
End Function
'@)
$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modInventoryViewer').CodeModule.AddFromString(@'
Public Function PublishedReadActionForTest(ByVal action As String, Optional ByVal expected As String = "") As Boolean
If mInventoryViewer Is Nothing Then Err.Raise 5, , "Restart Viewer fixture missing."
PublishedReadActionForTest = mInventoryViewer.PublishedReadActionForTest(action, expected)
End Function
'@)
SelectTarget $Fixture 'config-admin'
Stage 'Read journal'
OpenRecordingViewer
Check 'RecordingRestart.NoRecorderResumed' ((RecordingControl 'Start Recording') -ceq 'True|True' -and
    (RecordingControl 'Stop Recording') -ceq 'True|False' -and (RecordingStatus) -notmatch '(?i)recording.*1\s*/\s*256')
$opened=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Open',''))
$selected=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Select',$pathId))
Check 'RecordingRestart.RealUnclosedRunSelectable' ($opened -ceq 'DELIVERED' -and $selected -ceq 'SELECTED')
$evidence=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Evidence',''))
$status=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Status',''))
Check 'RecordingRestart.InterruptedNeverConcluded' (($status+"`n"+$evidence) -match '(?i)interrupted' -and
    ($status+"`n"+$evidence) -notmatch '(?i)conclusion observed')
Check 'RecordingRestart.OriginalOwningAttemptAndOutcomeVisible' ($evidence.Contains($action.Attempt.ActivityId) -and
    $evidence -match 'REQUESTED' -and $evidence -match 'COMPLETED')
Check 'RecordingRestart.ReadPreservesAllFixtureFiles' ((RestartPinsEqual $authorityBefore $Fixture.Root) -and
    (RestartPinsEqual $otherBefore $b.Root))
$ordinary=SaveRecordedSetting '632'
Check 'RecordingRestart.OrdinaryActionDoesNotJoinInterruptedRun' ($ordinary.Attempt.SequenceId -ceq '' -and
    $ordinary.Outcome.SequenceId -ceq '' -and (RestartPinsEqual $journalBefore $journalRoot))
[void](RecordingControl 'Start Recording' 'Click')
$next=SaveRecordedSetting '633'
[void](RecordingControl 'Stop Recording' 'Click')
Check 'RecordingRestart.NewExplicitRunHasIndependentIdentity' ((HasSequence $next 1) -and
    $next.Attempt.SequenceId -cne $sequence -and (JournalFact $next.Attempt.SequenceId 'Close' 1 'Stopped') -and
    (JournalChain $next.Attempt.SequenceId 4) -and (JournalChain $sequence 3))
$unchanged=$true
foreach($path in $journalBefore.Keys){if((Get-FileHash -LiteralPath $path).Hash -cne $journalBefore[$path]){$unchanged=$false}}
Check 'RecordingRestart.LaterWorkPreservesInterruptedEvidence' $unchanged
$unchanged=$true
foreach($path in $packagePins.Keys){if((Get-FileHash -LiteralPath $path).Hash -cne $packagePins[$path]){$unchanged=$false}}
Check 'RecordingRestart.FivePackageFilesUnchanged' $unchanged

} catch {
    Check 'RecordingRestart.WorkerHarnessFailure' $false
} finally {
    if($null -ne $excel){
        try{CloseRecordingViewer}catch{}
        try{[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')}catch{}
        try{foreach($book in @($excel.Workbooks)){$book.Close($false)};$excel.Quit()}catch{$failed=$true}
        try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}catch{}
    }
    $state=$null; $Fixture=$null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
if($failed){exit 1}
