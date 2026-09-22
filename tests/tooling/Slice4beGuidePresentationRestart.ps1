# Existing D18 persistence contract, through actual Operations form handlers.
# Only generated fixtures and saved disposable probe copies are used. No kill,
# authority repair, credential serialization, evaluation or preference shortcut.
function Test-GuidePresentationRestart($State) {
    $fixture=$State.Fixture;$guide=$State.Guide;$observed=$State.Observed
    function RestartControl([string]$Form,[string]$Name,[string]$Action,[string]$Value='') {
        [string](Run 'invSys.Operations.xlam' 'modInventoryViewer.GuideDraftControlForTest' @($Form,$Name,$Action,$Value))
    }
    function RestartLibrary([string]$Action,[string]$Value='') {
        [string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @($Action,$Value))
    }
    function RestartSettings {
        if((RestartControl 'frmInventoryViewer' 'btnSettings' 'Click') -cne 'DELIVERED'){throw 'Actual Operations Settings entry unavailable.'}
    }
    function RestartViewer {
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OpenInventoryViewer')
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('Events',''))
    }
    function RestartPair {
        if((RestartLibrary 'Select' $observed.ActionPathId) -cne 'SELECTED'){throw 'Restart observed recording fixture unavailable.'}
        if((RestartControl 'frmActionPaths' 'btnPublishedGuides' 'Click') -cne 'DELIVERED'){throw 'Actual published guide entry unavailable.'}
        $key=[string]$guide.ActionPathId+'|'+[string]$guide.Version+'|'+[string]$guide.ContentSha256
        $keys=@((RestartControl 'frmActionPathLibrary' 'lstPublishedGuides' 'Values') -split "`n"|Where-Object {$_ -cne ''})
        $index=[Array]::IndexOf($keys,$key)
        if($index -lt 0 -or (RestartControl 'frmActionPathLibrary' 'lstPublishedGuides' 'Select' ([string]$index)) -cne 'SELECTED'){
            throw 'Restart exact published guide fixture unavailable.'
        }
        if((RestartControl 'frmActionPathLibrary' 'btnUseGuideForRun' 'Click') -cne 'DELIVERED'){throw 'Actual Use for selected run unavailable.'}
        [void](RestartControl 'frmActionPathLibrary' 'btnCloseGuides' 'Click')
        if((RestartControl 'frmActionPaths' 'btnViewActionPath' 'Click') -cne 'DELIVERED'){throw 'Actual paired view entry unavailable.'}
    }
    function RestartPins {
        $pins=@{}
        foreach($file in Get-ChildItem -LiteralPath (Join-Path $fixture.Root 'Training') -File -Recurse){$pins[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash}
        $pins[$fixture.Config]=(Get-FileHash -LiteralPath $fixture.Config).Hash
        return $pins
    }
    function RestartSame($Before,$After) {
        if($Before.Count -ne $After.Count){return $false}
        foreach($key in $Before.Keys){if(-not $After.ContainsKey($key) -or $Before[$key] -cne $After[$key]){return $false}}
        return $true
    }
    function RestartCapture([string]$Caption,[string]$Name) {
        if($CaptureGuideEvidence){CaptureOwnedFormByCaptionEvidence $Caption ('paired-restart-'+$Name+'.png')}
    }
    if(-not ('InvSysPairedRestartOwner' -as [type])){
        Add-Type @'
using System; using System.Runtime.InteropServices;
public static class InvSysPairedRestartOwner {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
}
'@
    }
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.CloseInventoryViewerForTest')
    SelectTarget $fixture 'config-reader'
    RestartViewer;RestartSettings
    $before=RestartPins
    $saved=(RestartControl 'frmEventTrackingSettings' 'cmbPreferredActionPathView' 'Write' 'Compare both') -ceq 'DELIVERED' -and
        (RestartControl 'frmEventTrackingSettings' 'btnSaveMyPreference' 'Click') -ceq 'DELIVERED' -and
        (RestartControl 'frmEventTrackingSettings' 'lblPreferenceStatus' 'Label') -ceq 'Your Action Path preference was saved.'
    Check 'GuidePresentation.Restart.PreferenceSavedThroughOperationsHandler' $saved
    if(-not $saved){return}
    [void](RestartControl 'frmEventTrackingSettings' 'btnClose' 'Click')
    if((RestartLibrary 'Open') -cne 'DELIVERED'){throw 'Actual recording library entry unavailable.'}
    RestartPair
    Check 'GuidePresentation.Restart.BeforeCloseUsesSavedChoice' ((RestartControl 'frmActionPathView' 'cboActionPathView' 'Selected') -ceq 'Compare both')
    $switched=(RestartControl 'frmActionPathView' 'cboActionPathView' 'Write' 'How-To') -ceq 'DELIVERED' -and
        (RestartControl 'frmActionPathView' 'cboActionPathView' 'Selected') -ceq 'How-To'
    Check 'GuidePresentation.Restart.UnsavedViewSwitchEstablished' $switched
    RestartCapture 'Action Path view' 'unsaved-how-to'
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.CloseInventoryViewerForTest')
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
    [uint32]$ownerId=0
    [void][InvSysPairedRestartOwner]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$ownerId)
    $owners=@(Get-Process EXCEL -ErrorAction Stop)
    if($owners.Count -ne 1 -or $owners[0].Id -ne $ownerId -or $ownerId -notin $initialExcelProcessIds){throw 'Restart Excel ownership is ambiguous.'}
    $original=$owners[0]
    $packageNames=@($packages.Keys)
    $count=$excel.Workbooks.Count
    if($count -isnot [int]){throw 'Restart workbook collection unavailable.'}
    for($index=$count;$index -ge 1;$index--){
        $closing=$excel.Workbooks.Item($index)
        $closing.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($closing)
        $closing=$null
    }
    $count=$excel.Workbooks.Count
    if($count -isnot [int] -or $count -ne 0){throw 'Restart requires verified empty owned Excel.'}
    Check 'GuidePresentation.Restart.OwnedWorkbooksClosedNormally' $true
    $excel.Quit()
    # Release retained instrumentation references only after normal close/Quit.
    foreach($name in @('formCode','testModule','productionCode','productionTest','bootstrapCode','book')){
        $variable=Get-Variable -Name $name -Scope Script -ErrorAction SilentlyContinue
        if($null -ne $variable -and $null -ne $variable.Value -and [Runtime.InteropServices.Marshal]::IsComObject($variable.Value)){
            try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($variable.Value)}catch [Runtime.InteropServices.InvalidComObjectException]{}
            Set-Variable -Name $name -Value $null -Scope Script
        }
    }
    $script:packages=@{}
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    $script:excel=$null
    [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()
    $deadline=[DateTime]::UtcNow.AddMinutes(10)
    while(-not $original.HasExited -and [DateTime]::UtcNow -lt $deadline){
        Write-Output 'Paired restart: awaiting normal exit of verified empty owned Excel; no termination.'
        [void]$original.WaitForExit(30000)
    }
    if(-not $original.HasExited -or (Get-Process EXCEL -ErrorAction SilentlyContinue)){throw 'Normal Excel closure not established; restart was not attempted.'}
    Check 'GuidePresentation.Restart.OriginalProcessExitedNormally' $true
    # Excel holds writable probe XLAMs open. Hash only after verified closure,
    # before the fresh session opens the same saved copies read-only.
    $copies=@{}
    foreach($name in $packageNames){$copies[$name]=(Get-FileHash -LiteralPath (Join-Path $deploy $name)).Hash}
    $script:excel=New-Object -ComObject Excel.Application
    $script:initialExcelWindow=[long]$excel.Hwnd
    [uint32]$newOwner=0
    [void][InvSysPairedRestartOwner]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$newOwner)
    $current=@(Get-Process EXCEL -ErrorAction Stop)
    $fresh=$current.Count -eq 1 -and $current[0].Id -eq $newOwner -and $newOwner -ne $ownerId
    Check 'GuidePresentation.Restart.DifferentVerifiedExcelProcess' $fresh
    if(-not $fresh){throw 'Fresh Excel process identity unavailable.'}
    $excel.Visible=[bool]$GuideCaptureVisibleExcelForTest;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=1
    foreach($name in @('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam')){
        $script:packages[$name]=$excel.Workbooks.Open((Join-Path $deploy $name),0,$true)
    }
    SelectTarget $fixture 'config-reader'
    RestartViewer;RestartSettings
    Check 'GuidePresentation.Restart.OperationsSettingsRestoresSavedChoice' ((RestartControl 'frmEventTrackingSettings' 'cmbPreferredActionPathView' 'Selected') -ceq 'Compare both')
    Check 'GuidePresentation.Restart.SettingsEffectiveViewMatches' ((RestartControl 'frmEventTrackingSettings' 'lblEffectiveView' 'Label') -match '^Effective view: Compare both')
    RestartCapture 'Event Tracking Settings' 'saved-settings'
    [void](RestartControl 'frmEventTrackingSettings' 'btnClose' 'Click')
    if((RestartLibrary 'Open') -cne 'DELIVERED'){throw 'Fresh recording library unavailable.'}
    Check 'GuidePresentation.Restart.NoAutomaticPairOrConclusion' ((RestartControl 'frmActionPaths' 'btnViewActionPath' 'State') -ceq 'True|False' -and (RestartControl 'frmActionPaths' 'txtPathEvaluation' 'Text') -ceq '')
    RestartPair
    Check 'GuidePresentation.Restart.PairedViewRestoresSavedNotUnsavedChoice' ((RestartControl 'frmActionPathView' 'cboActionPathView' 'Selected') -ceq 'Compare both')
    $pair=RestartControl 'frmActionPathView' 'lblActionPathPair' 'Label'
    Check 'GuidePresentation.Restart.ExplicitPairRetainsExactGuideAndObservedRun' ($pair.Contains([string]$guide.ActionPathId) -and $pair.Contains([string]$guide.ContentSha256) -and $pair.Contains([string]$observed.ActionPathId))
    $diagnostic=RestartControl 'frmActionPathView' 'txtActionPathDiagnostic' 'Text'
    Check 'GuidePresentation.Restart.ReopenNeverInfersSavedEvaluation' ($diagnostic -match '(?i)not evaluated|no .*evaluation|choose Evaluate' -and -not $diagnostic.Contains('Conclusion observed'))
    Check 'GuidePresentation.Restart.OperationsOnlyWithoutAdminDependency' (-not (Test-LoadedPackage 'invSys.Admin.xlam'))
    RestartCapture 'Action Path view' 'restored-compare'
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.CloseInventoryViewerForTest')
    Check 'GuidePresentation.Restart.PreferenceAndViewsPreserveAllTrainingAndConfigBytes' (RestartSame $before (RestartPins))
    $unchanged=$true
    foreach($name in $copies.Keys){if((Get-FileHash -LiteralPath (Join-Path $deploy $name)).Hash -cne $copies[$name]){$unchanged=$false}}
    Check 'GuidePresentation.Restart.SavedProbePackagesUnchanged' $unchanged
}
