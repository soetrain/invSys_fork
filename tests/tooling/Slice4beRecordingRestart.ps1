# D18 interruption proof: retain a real unclosed journal across distinct Excel
# processes. Termination is restricted to the current disposable fixture runtime.
# No reset seam, deleted Close record, saved instrumentation or persisted secret.
function Test-Slice4beRecordingRestart($Fixture) {
    if(-not ('RecordingRestartProcess' -as [type])) {
        Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class RecordingRestartProcess {
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
    public static uint Owner(long window) {
        if(window == 0) throw new InvalidOperationException("Excel window unavailable.");
        uint process;
        if(GetWindowThreadProcessId(new IntPtr(window), out process) == 0 || process == 0)
            throw new InvalidOperationException("Excel process unavailable.");
        return process;
    }
}
'@
    }
    function RestartPins([string]$Root) {
        $pins=@{}
        foreach($file in Get-ChildItem -LiteralPath $Root -Recurse -File) {
            $pins[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash
        }
        return $pins
    }
    function RestartPinsEqual($Before,[string]$Root) {
        $after=RestartPins $Root
        if($after.Count -ne $Before.Count){return $false}
        foreach($path in $Before.Keys){if(-not $after.ContainsKey($path) -or $after[$path] -cne $Before[$path]){return $false}}
        return $true
    }
    function CopyRestartProbe([string]$Package,[string]$Module,[string[]]$Procedures) {
        $code=$packages[$Package].VBProject.VBComponents.Item($Module).CodeModule
        $parts=@(foreach($procedure in $Procedures){$code.Lines($code.ProcStartLine($procedure,0),$code.ProcCountLines($procedure,0))})
        [pscustomobject]@{Package=$Package;Module=$Module;Text=($parts -join "`r`n")}
    }
    CloseRecordingViewer
    SelectTarget $Fixture 'config-admin'
    SetRecordingPolicy $true
    OpenRecordingViewer
    if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'Restart fixture Start control is unavailable.'}
    $action=SaveRecordedSetting '631'
    $sequence=[string]$action.Attempt.SequenceId
    $entries=@(RecordingJournal $sequence|Sort-Object Version)
    $prepared=(HasSequence $action 1) -and (JournalChain $sequence 3) -and
        @($entries|Where-Object RecordType -CEQ 'Close').Count -eq 0 -and
        (RecordingControl 'Stop Recording') -ceq 'True|True'
    Check 'RecordingRestart.RealActionDurableBeforeInterruption' $prepared
    if(-not $prepared){throw 'Refusing process interruption without the real active journal fixture.'}
    $pathId=[string]$entries[0].ActionPathId
    $probes=@(
        (CopyRestartProbe 'invSys.Operations.xlam' 'frmInventoryViewer' @('RecordingControlForTest','RecordingStatusForTest')),
        (CopyRestartProbe 'invSys.Operations.xlam' 'modInventoryViewer' @('RecordingControlForTest','RecordingStatusForTest','RecordingLibraryForTest')),
        (CopyRestartProbe 'invSys.Admin.xlam' 'frmAdminSettings' @('D5TestSave')),
        (CopyRestartProbe 'invSys.Admin.xlam' 'TestD5Commands' @('OpenSettings','CloseSettings','SaveSettings'))
    )
    $rootPath=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    $tempPath=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if(-not $rootPath.StartsWith($tempPath,[StringComparison]::OrdinalIgnoreCase) -or
        (Split-Path $runRoot -Leaf) -notlike 'invsys-config-command-*') {throw 'Restart requires the generated disposable fixture root.'}
    $packageNames=@('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam','invSys.Admin.xlam')
    $packagePins=@{}
    foreach($name in $packageNames){$path=Join-Path $deploy $name;$packagePins[$path]=(Get-FileHash -LiteralPath $path).Hash}
    # Verify every loaded project and workbook, including hidden add-ins, before
    # intentionally interrupting this runtime. Never terminate an unknown project.
    $projects=$excel.VBE.VBProjects
    if($projects.Count -ne 5){throw 'Restart requires exactly the five candidate projects.'}
    foreach($project in $projects){
        $path=[IO.Path]::GetFullPath([string]$project.FileName)
        if(-not $packagePins.ContainsKey($path)){throw 'Unknown loaded project prevents interruption.'}
    }
    foreach($book in @($excel.Workbooks)){
        $path=[IO.Path]::GetFullPath([string]$book.FullName)
        if($packagePins.ContainsKey($path)){continue}
        if(-not $path.StartsWith($rootPath,[StringComparison]::OrdinalIgnoreCase) -or -not $book.Saved){
            throw 'Unknown or unsaved non-package workbook prevents interruption.'
        }
    }
    $owned=@(Get-Process EXCEL -ErrorAction Stop)
    $owner=[RecordingRestartProcess]::Owner([long]$excel.Hwnd)
    if($owned.Count -ne 1 -or $initialExcelProcessIds.Count -ne 1 -or
        $owned[0].Id -ne $owner -or $owner -ne $initialExcelProcessIds[0]){throw 'Excel ownership changed before interruption.'}
    [void]$owned[0].Handle
    $authorityBefore=RestartPins $Fixture.Root
    $otherBefore=RestartPins $b.Root
    $journalBefore=RestartPins $journalRoot
    # Keep the exact Process handle; PID reuse or a new unrelated instance is
    # not a termination target. No Quit/form close/SignOut is called beforehand.
    if($owned[0].HasExited -or [RecordingRestartProcess]::Owner([long]$excel.Hwnd) -ne $owner -or
        @(Get-Process EXCEL -ErrorAction Stop).Count -ne 1){throw 'Excel ownership changed at the interruption boundary.'}
    $owned[0].Kill()
    if(-not $owned[0].WaitForExit(5000)){throw 'Interrupted fixture process has not exited.'}
    try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}catch{}
    $script:excel=$null; $script:packages=@{}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel process prevents cold restart.'}
    Check 'RecordingRestart.InterruptionPreservesUnclosedJournal' (
        (RestartPinsEqual $journalBefore $journalRoot) -and (JournalChain $sequence 3))
    $script:excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=1
    foreach($name in $packageNames){$packages[$name]=$excel.Workbooks.Open((Join-Path $deploy $name),0,$true)}
    $current=@(Get-Process EXCEL -ErrorAction Stop)
    $newOwner=[RecordingRestartProcess]::Owner([long]$excel.Hwnd)
    $distinct=$current.Count -eq 1 -and $current[0].Id -eq $newOwner -and $newOwner -ne $owner
    Check 'RecordingRestart.FreshExcelProcessVerified' $distinct
    if(-not $distinct){throw 'Fresh isolated Excel identity is unavailable.'}
    Write-Output 'Recording restart: calibrate pristine packaged Viewer before fresh-process probes.'
    SelectTarget $Fixture 'config-admin'
    $pristine=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RunInventoryViewerActionForTest')
    $openedPristine=$pristine.StartsWith('OK|')
    Check 'RecordingRestart.PristinePackagedViewerLaunches' $openedPristine
    if(-not $openedPristine){throw 'Pristine fresh-process Viewer baseline is unavailable.'}
    CloseRecordingViewer
    Write-Output 'Recording restart: install only the unsaved control drivers.'
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
    Write-Output 'Recording restart: open instrumented Viewer and inspect the original journal.'
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
    [pscustomobject]@{OldProcessId=$owner;NewProcessId=$newOwner;OldExited=$owned[0].HasExited;
        InterruptionWasDeliberate=$true;PackageFilesUnchanged=$unchanged}|
        ConvertTo-Json|Set-Content -LiteralPath (Join-Path $reportRoot 'recording-restart-processes.json')
}
