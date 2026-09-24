# Isolated restart diagnostic. Retain original settings in memory through closure.
[CmdletBinding()]
param([string]$RepoRoot='.',[string]$DeployRoot='deploy/validation-settings-diagnostic',[switch]$SavedWorkbook,[switch]$TraceViewCalls,[switch]$CheckLayoutStability,[switch]$SkipCapture,[ValidateSet('RED','GREEN')][string]$Phase='GREEN')
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before isolated resource diagnostic.'}
if($TraceViewCalls -and -not $SavedWorkbook){throw 'View-call tracing requires the saved-workbook control.'}
if($CheckLayoutStability -and -not $TraceViewCalls){throw 'Layout stability requires the instrumented handler trace.'}
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
. (Join-Path $PSScriptRoot 'Slice4beGuideResourceTrace.ps1')
$root=Join-Path $repo ('reports/runtime/guide-resource-diagnostic/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
Write-Output ('Diagnostic report: '+$root)
$settings=Get-InvSysTestSettingsSnapshot
$start=[DateTimeOffset]::UtcNow
$packages=@(Get-ChildItem -LiteralPath (Join-Path $repo $DeployRoot) -Filter '*.xlam' -File|ForEach-Object {
    [pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash -LiteralPath $_.FullName).Hash}
})
if($packages.Count -ne 5){throw 'Expected five candidate packages.'}
$worker=$null
try {
    $arguments=@('-NoProfile','-ExecutionPolicy','Bypass','-File',('"'+(Join-Path $PSScriptRoot 'Test-Slice4beConfigCommands.ps1')+'"'),
        '-RepoRoot',('"'+$repo+'"'),'-DeployRoot',('"'+$DeployRoot+'"'),'-Phase',$Phase,
        '-GuideDraftOnly','-GuidePresentationRestartOnly','-TraceGuideResourcesForTest',
        '-GuideCaptureVisibleExcelForTest','-CheckViewerPublishedRead',
        '-CompileEvaluationProbesForTest','-ViewerStartupPackageStateForTest','SavedCopies',
        '-WaitForExcelReadyForTest','-ExcelReadyReadLimitForTest','40')
    if($SavedWorkbook){$arguments+='-GuideResourceSavedWorkbookForTest'}
    if(-not $SkipCapture){$arguments+='-CaptureGuideEvidence'}
    if($TraceViewCalls){$arguments+='-TraceGuideViewCallsForTest'}
    if($CheckLayoutStability){$arguments+='-CheckGuideLayoutStabilityForTest'}
    $worker=Start-Process powershell.exe -ArgumentList $arguments -WindowStyle Hidden -PassThru -WorkingDirectory $repo -RedirectStandardOutput (Join-Path $root 'worker.stdout.log') -RedirectStandardError (Join-Path $root 'worker.stderr.log')
    $null=$worker.Handle
    while(-not $worker.HasExited){
        if(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -eq 1){
            $sample=Get-GuideResourceSample 'Passive' -AllowExited
            if($null -ne $sample){$sample|ConvertTo-Json -Compress|Add-Content (Join-Path $root 'passive.jsonl')}
        }
        Start-Sleep -Milliseconds 1000
    }
    $worker.Refresh()
    [pscustomobject]@{ExitCode=$worker.ExitCode;UTC=[DateTimeOffset]::UtcNow.ToString('o')}|ConvertTo-Json|Set-Content (Join-Path $root 'worker-exit.json')
} finally {
    Wait-RecordingCleanup -Creator $null -Worker $worker
    $restored=Restore-InvSysTestSettingsSnapshot $settings
    $preserved=$true
    foreach($pin in $packages){if((Get-FileHash -LiteralPath $pin.File).Hash -cne $pin.Hash){$preserved=$false}}
    [pscustomobject]@{StartUTC=$start.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');SettingsRestored=$restored;PackagesPreserved=$preserved;ExcelClosed=@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -eq 0;ReleaseAccepted=$false}|ConvertTo-Json|Set-Content (Join-Path $root 'closure.json')
    if(-not $restored -or -not $preserved){throw 'Diagnostic restoration verification failed.'}
}
if($null -ne $worker -and $worker.ExitCode -ne 0){exit $worker.ExitCode}
