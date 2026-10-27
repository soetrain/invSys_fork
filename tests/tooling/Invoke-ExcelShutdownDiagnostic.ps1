[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$RepoRoot,
    [Parameter(Mandatory=$true)][string]$WorkerScript,
    [Parameter(Mandatory=$true)][string]$DeployRoot,
    [Parameter(Mandatory=$true)][string]$DesktopMonitorReport
)
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot 'ExcelShutdownTiming.ps1')
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
$workerPath=(Resolve-Path -LiteralPath $WorkerScript).Path
$deploy=(Resolve-Path -LiteralPath $DeployRoot).Path
if([IO.Path]::GetFileName($workerPath) -cne 'Test-Slice4beConfigCommands.ps1'){throw 'Known Config command worker required.'}
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Existing Excel prevents isolated diagnostic.'}
$desktop=Get-Content -Raw -LiteralPath $DesktopMonitorReport|ConvertFrom-Json
if($desktop.LockSamples -or $desktop.MouseLockSamples -or ([DateTimeOffset]::UtcNow-[DateTimeOffset]::Parse($desktop.UTC)).TotalSeconds -gt 15){throw 'Fresh unlocked desktop and cursor proof required.'}
$root=Join-Path $repo ('reports/runtime/excel-shutdown-timing/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$module=(Join-Path $PSScriptRoot 'ExcelShutdownTiming.ps1')
$source=[IO.File]::ReadAllText($workerPath)
$generated=ConvertTo-ExcelShutdownTimedWorker $source (Split-Path -Parent $workerPath) $module $root
$generatedPath=Join-Path $root 'timed-worker.ps1'
[IO.File]::WriteAllText($generatedPath,$generated,[Text.UTF8Encoding]::new($false))
$pins=@(Get-ChildItem -LiteralPath $deploy -Filter '*.xlam' -File|ForEach-Object {[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash -LiteralPath $_.FullName).Hash}})
if($pins.Count -ne 5){throw 'Five frozen packages required.'}
$toolingPins=@(Get-ChildItem -LiteralPath (Split-Path -Parent $workerPath) -Filter '*.ps1' -File|ForEach-Object {[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash -LiteralPath $_.FullName).Hash}})
[pscustomobject]@{WorkerHash=(Get-FileHash $workerPath).Hash;GeneratedWorkerHash=(Get-FileHash $generatedPath).Hash;
    TimingModuleHash=(Get-FileHash $module).Hash;Mode='ObservationOnly';Workload='InventoryWorksheetEvidence';
    FullAcceptance=$false;AssistedTeardown=$false}|ConvertTo-Json|Set-Content (Join-Path $root 'manifest.json')
$pins|ConvertTo-Json|Set-Content (Join-Path $root 'package-pins.json')
$toolingPins|ConvertTo-Json|Set-Content (Join-Path $root 'tooling-pins.json')
$settings=Get-InvSysTestSettingsSnapshot
$worker=$null;$watching=$false;$owner=$null;$started=[DateTimeOffset]::UtcNow;$code=1
Write-Output ('Shutdown diagnostic: '+$root)
Initialize-ExcelExitObserver
try {
    $arguments=@('-NoProfile','-NonInteractive','-File',('"'+$generatedPath+'"'),'-RepoRoot',('"'+$repo+'"'),
        '-DeployRoot',('"'+$deploy+'"'),'-Phase','RED','-AdminInventoryEntryOnly','-AdminInventoryRecordingOnly',
        '-AdminInventoryWorksheetRecordingOnly','-AdminInventoryWorksheetRecordingGroup','Evidence',
        '-CompileEvaluationProbesForTest','-WaitForExcelReadyForTest','-ExcelReadyReadLimitForTest','40',
        '-DesktopMonitorReport',('"'+$DesktopMonitorReport+'"'))
    $worker=Start-Process powershell.exe -WindowStyle Hidden -PassThru -ArgumentList $arguments -WorkingDirectory $repo `
        -RedirectStandardOutput (Join-Path $root 'worker.log') -RedirectStandardError (Join-Path $root 'worker.stderr.log')
    while(-not $worker.WaitForExit(500)){
        if(-not $watching -and (Test-Path -LiteralPath (Join-Path $root 'owner.json'))){
            $owner=Get-Content -Raw -LiteralPath (Join-Path $root 'owner.json')|ConvertFrom-Json
            if($owner.WorkerId -ne $worker.Id -or [DateTime]::FromFileTimeUtc($owner.CreatedFileTime) -lt $worker.StartTime.ToUniversalTime()){
                throw 'Owned process provenance failed.'
            }
            $owned=Get-Process -Id $owner.ProcessId -ErrorAction Stop
            if($owned.ProcessName -cne 'EXCEL'){throw 'Window owner is not Excel.'}
            [InvSysExcelExitObserver]::Watch($owner.ProcessId,$owner.CreatedFileTime,(Join-Path $root 'native-exit.json'))
            $watching=$true
        }
    }
    $code=$worker.ExitCode
    $script:ExcelShutdownTimingRoot=$root
    Write-ExcelShutdownTiming WorkerExit
} finally {
    # Keep restoration state even when worker/observer setup fails. Baseline has
    # no timeout/kill policy: do not relabel an assisted exit as normal evidence.
    Wait-RecordingCleanup -Creator $null -Worker $worker
    $restored=Restore-InvSysTestSettingsSnapshot $settings
    $same=@($pins|Where-Object {(Get-FileHash -LiteralPath $_.File).Hash -cne $_.Hash}).Count -eq 0
    $toolsSame=@($toolingPins|Where-Object {(Get-FileHash -LiteralPath $_.File).Hash -cne $_.Hash}).Count -eq 0
    [pscustomobject]@{StartUTC=$started.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');WorkerExitCode=$code;
        ExcelClosed=$true;SettingsRestored=$restored;PackagesPreserved=$same;OriginalToolingPreserved=$toolsSame;
        NativeObserverStarted=$watching;AssistedTeardown=$false;FullAcceptance=$false}|ConvertTo-Json|Set-Content (Join-Path $root 'closure.json')
    if(-not($restored -and $same -and $toolsSame)){throw 'Diagnostic preservation failed.'}
}
# Give the observer time to flush its native exit record, without querying COM.
for($attempt=0;$attempt -lt 20 -and -not(Test-Path (Join-Path $root 'native-exit.json'));$attempt++){Start-Sleep -Milliseconds 100}
$rows=@(Get-Content (Join-Path $root 'phases.jsonl')|ForEach-Object {$_|ConvertFrom-Json})
$timings=@()
foreach($pair in @(@('Quit','BeforeQuit','AfterQuit'),@('ComRelease','BeforeComRelease','AfterComRelease'),
    @('GC','BeforeGC','AfterGC'),@('FinalizerWait','BeforeFinalizerWait','AfterFinalizerWait'))){
    $before=@($rows|Where-Object Stage -ceq $pair[1]);$after=@($rows|Where-Object Stage -ceq $pair[2])
    $complete=$before.Count -eq 1 -and $after.Count -eq 1
    $timings+=[pscustomobject]@{Phase=$pair[0];Complete=$complete;Milliseconds=$(if($complete){1000.0*($after[0].Ticks-$before[0].Ticks)/$before[0].Frequency}else{$null})}
}
$timings|ConvertTo-Json|Set-Content (Join-Path $root 'timings.json')
$timings|Format-Table -AutoSize
if(-not $watching -or -not(Test-Path (Join-Path $root 'native-exit.json'))){throw 'Independent process-exit observation missing; diagnostic incomplete.'}
$exit=Get-Content -Raw (Join-Path $root 'native-exit.json')|ConvertFrom-Json
if($exit.PSObject.Properties['ObserverFailed'] -or $exit.WaitResult -ne 0 -or -not $exit.TimesAvailable){throw 'Independent native exit proof failed.'}
Get-Content (Join-Path $root 'closure.json')
