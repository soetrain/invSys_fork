[CmdletBinding()]
param([string]$DeployRoot='deploy/validation-tester-recording-02',[ValidateSet('RED','GREEN')][string]$Phase='RED',[string]$DesktopMonitorReport='',[switch]$PathComparison,[switch]$BoundaryMatrix,[switch]$WriteBoundariesOnly,[switch]$RestartOnly,[switch]$CompletionOnly,[ValidateSet('All','Inspection','Write','Ownership')][string]$BoundaryGroup='All',[ValidateSet('All','UnknownOutputColumn')][string]$BoundaryCase='All',
    [ValidateSet('','RedFocus','PhaseRedFocus','FeedReopen','ConnectionFailed','ConnectionSucceeded','Confirmation','ConfirmationCancelled','ResultCompleted','ResultCompletedSelectionChanged','ResultFailed','BeforeSave','AfterSave','AfterPromotion')][string]$YieldGroup='',
    [switch]$SuppressPhaseRechecksForTest,[switch]$RecordingOnly,[switch]$GuideOnly,[switch]$RecordingOutcomesOnly)
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
if($GuideOnly -and $RecordingOutcomesOnly){throw 'Choose the guide or outcome recording gate.'}
if(($RecordingOnly -or $GuideOnly -or $RecordingOutcomesOnly) -and ($YieldGroup -or $BoundaryMatrix -or $WriteBoundariesOnly -or $RestartOnly -or $CompletionOnly -or $PathComparison -or $BoundaryGroup -ne 'All' -or $BoundaryCase -ne 'All')){throw 'Aggregation recording and guide gates exclude other focused modes.'}
if($BoundaryCase -ne 'All' -and $BoundaryGroup -ne 'Write'){throw 'The focused unknown-column case requires the Write boundary group.'}
if($YieldGroup -and ($BoundaryGroup -ne 'All' -or $BoundaryCase -ne 'All' -or $PathComparison -or $WriteBoundariesOnly -or $RestartOnly -or $CompletionOnly)){throw 'Named Aggregation yields exclude other focused modes.'}
if($SuppressPhaseRechecksForTest -and ($Phase -ne 'RED' -or $YieldGroup -notin @('PhaseRedFocus','BeforeSave','AfterSave','AfterPromotion'))){throw 'The isolated missing-phase-check seam is restricted to its RED phase tests.'}
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before the isolated Aggregation owner gate.'}
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$snapshot=Get-InvSysTestSettingsSnapshot
$root=Join-Path 'reports/runtime/admin-aggregation-controller' ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$pins=@(Get-ChildItem -LiteralPath $DeployRoot -File -Filter '*.xlam'|ForEach-Object {[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash -LiteralPath $_.FullName).Hash}})
if($pins.Count -ne 5){throw 'Five frozen packages required.'}
$testPins=@(Get-ChildItem -LiteralPath $PSScriptRoot -File -Filter '*.ps1'|ForEach-Object {[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash -LiteralPath $_.FullName).Hash}})
$pins|ConvertTo-Json|Set-Content (Join-Path $root 'package-pins.json')
$testPins|ConvertTo-Json|Set-Content (Join-Path $root 'tooling-pins.json')
$start=[DateTimeOffset]::UtcNow;$code=1;$worker=$null;$ownedFixture=''
Write-Output ('Aggregation owner controller: '+$root)
try {
    if($RestartOnly){
        $ownedFixture=Join-Path ([IO.Path]::GetTempPath()) ('invsys-config-command-'+[guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $ownedFixture -ErrorAction Stop|Out-Null
    }
    $workerArgs=@('-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',
        (Join-Path $PSScriptRoot 'Test-Slice4beConfigCommands.ps1'),'-DeployRoot',$DeployRoot,
        '-Phase',$Phase,'-AdminAggregationOnly','-CompileEvaluationProbesForTest',
        '-WaitForExcelReadyForTest','-ExcelReadyReadLimitForTest','40')
    if($DesktopMonitorReport){$workerArgs+=@('-DesktopMonitorReport',$DesktopMonitorReport)}
    if($RecordingOnly -or $GuideOnly -or $RecordingOutcomesOnly){$workerArgs+='-AdminAggregationRecording'}
    if($GuideOnly){$workerArgs+='-AdminAggregationGuideOnly'}
    if($RecordingOutcomesOnly){$workerArgs+='-AdminAggregationOutcomeOnly'}
    if($PathComparison){$workerArgs+='-AdminAggregationPathComparison'}
    if($BoundaryMatrix -or $WriteBoundariesOnly -or $RestartOnly -or $CompletionOnly -or ($BoundaryGroup -ne 'All') -or $YieldGroup){$workerArgs+='-AdminAggregationBoundaryMatrix'}
    if($BoundaryGroup -ne 'All'){$workerArgs+=@('-AdminAggregationBoundaryGroup',$BoundaryGroup)}
    if($BoundaryCase -ne 'All'){$workerArgs+=@('-AdminAggregationBoundaryCase',$BoundaryCase)}
    if($YieldGroup){$workerArgs+=@('-AdminAggregationYieldGroup',$YieldGroup)}
    if($SuppressPhaseRechecksForTest){$workerArgs+='-AdminAggregationSuppressPhaseRechecksForTest'}
    if($WriteBoundariesOnly){$workerArgs+='-AdminAggregationWriteBoundariesOnly'}
    if($RestartOnly){$workerArgs+=@('-AdminAggregationRestartOnly','-ViewerStartupPackageStateForTest','SavedCopies','-AggregationOwnedFixtureRootForTest',$ownedFixture)}
    if($CompletionOnly){$workerArgs+='-AdminAggregationCompletionOnly'}
    $quotedArgs=($workerArgs|ForEach-Object {'"'+$_+'"'}) -join ' '
    $launch=[Diagnostics.ProcessStartInfo]::new((Join-Path $PSHOME 'powershell.exe'))
    $launch.Arguments=$quotedArgs;$launch.WorkingDirectory=$PWD.Path
    $launch.UseShellExecute=$false;$launch.CreateNoWindow=$true
    $launch.RedirectStandardOutput=$true;$launch.RedirectStandardError=$true
    $outputFile=[IO.FileStream]::new((Join-Path $PWD.Path (Join-Path $root 'worker.log')),[IO.FileMode]::Create,[IO.FileAccess]::Write,[IO.FileShare]::Read,1,$true)
    $errorFile=[IO.FileStream]::new((Join-Path $PWD.Path (Join-Path $root 'worker-error.log')),[IO.FileMode]::Create,[IO.FileAccess]::Write,[IO.FileShare]::Read,1,$true)
    try{
        $worker=[Diagnostics.Process]::Start($launch)
        $outputCopy=$worker.StandardOutput.BaseStream.CopyToAsync($outputFile,256)
        $errorCopy=$worker.StandardError.BaseStream.CopyToAsync($errorFile,256)
        $worker.WaitForExit();$outputCopy.Wait();$errorCopy.Wait()
        if(-not $worker.HasExited){throw 'Aggregation worker termination was not confirmed.'}
        $code=$worker.ExitCode
    }finally{$outputFile.Dispose();$errorFile.Dispose()}
} finally {
    Wait-RecordingCleanup -Creator $null -Worker $worker
    $fixtureRemoved=$true
    if($ownedFixture){
        try{
            if(Test-Path -LiteralPath $ownedFixture){
                if((Get-Item -LiteralPath $ownedFixture).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Owned fixture became a reparse point.'}
                Remove-Item -LiteralPath $ownedFixture -Recurse -Force -ErrorAction Stop
            }
            $fixtureRemoved=-not(Test-Path -LiteralPath $ownedFixture)
        }catch{$fixtureRemoved=$false}
    }
    $restored=Restore-InvSysTestSettingsSnapshot $snapshot
    $same=@($pins|Where-Object {(Get-FileHash -LiteralPath $_.File).Hash -cne $_.Hash}).Count -eq 0
    $testsSame=@($testPins|Where-Object {(Get-FileHash -LiteralPath $_.File).Hash -cne $_.Hash}).Count -eq 0
    [pscustomobject]@{StartUTC=$start.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');ExitCode=$code;SettingsRestored=$restored;PackagesPreserved=$same;ToolingPreserved=$testsSame;OwnedFixtureRemoved=$fixtureRemoved;ExcelClosed=$true;FullAcceptance=$false;NativeInputAcceptance=$false}|ConvertTo-Json|Set-Content (Join-Path $root 'closure.json')
    if(-not($restored -and $same -and $testsSame -and $fixtureRemoved)){throw 'Aggregation diagnostic preservation or owned fixture cleanup failed.'}
}
# Delayed native audit: process exit alone does not prove clean native shutdown.
Start-Sleep -Seconds 20
$audit=[DateTimeOffset]::UtcNow
$events=@(Get-WinEvent -FilterHashtable @{LogName='Application';Id=1000,1001,1002;StartTime=$start.LocalDateTime;EndTime=$audit.LocalDateTime} -ErrorAction SilentlyContinue -ErrorVariable queryErrors|Where-Object {$_.Message -match '(?i)excel\.exe'})
if(@($queryErrors|Where-Object FullyQualifiedErrorId -notlike 'NoMatchingEventsFound*').Count){throw 'Aggregation native audit unavailable.'}
$closed=@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -eq 0
[pscustomobject]@{StartUTC=$start.ToString('o');AuditUTC=$audit.ToString('o');ExcelEventCount=$events.Count;ExcelClosed=$closed}|ConvertTo-Json|Set-Content (Join-Path $root 'native-audit.json')
if($events.Count -or -not $closed){throw 'Aggregation native closure did not pass.'}
Get-Content (Join-Path $root 'worker.log') -Tail 14
exit $code
