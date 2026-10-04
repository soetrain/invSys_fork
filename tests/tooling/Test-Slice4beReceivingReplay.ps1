[CmdletBinding()]
param([string]$DeployRoot='deploy/validation-warehouse-purpose-02',[ValidateSet('RED','GREEN')][string]$Phase='RED')
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before isolated Receiving replay validation.'}
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$settings=Get-InvSysTestSettingsSnapshot
$root=Join-Path 'reports/runtime/receiving-replay-controller' ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$pins=@(Get-ChildItem -LiteralPath $DeployRoot -Filter '*.xlam' -File|ForEach-Object{[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash -LiteralPath $_.FullName).Hash}})
if($pins.Count -ne 5){throw 'Five frozen packages required.'}
$pins|ConvertTo-Json|Set-Content (Join-Path $root 'package-pins.json')
$start=[DateTimeOffset]::UtcNow;$code=1
Write-Output ('Receiving replay controller: '+$root)
try {
    # Reuse the accepted guide fixture's compiled, saved disposable package copies.
    $ErrorActionPreference='Continue'
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-Slice4beConfigCommands.ps1') -DeployRoot $DeployRoot -Phase $Phase -CheckReceivingReplay -GuideDraftOnly -CheckGuideExpectation -CheckViewerPublishedRead -ViewerStartupPackageStateForTest SavedCopies -CompileEvaluationProbesForTest -WaitForExcelReadyForTest -ExcelReadyReadLimitForTest 40 *> (Join-Path $root 'worker.log')
    $code=$LASTEXITCODE
    $ErrorActionPreference='Stop'
} finally {
    Wait-RecordingCleanup -Creator $null -Worker $null
    $restored=Restore-InvSysTestSettingsSnapshot $settings
    $same=@($pins|Where-Object {(Get-FileHash -LiteralPath $_.File).Hash -cne $_.Hash}).Count -eq 0
    [pscustomobject]@{Phase=$Phase;StartUTC=$start.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');ExitCode=$code;SettingsRestored=$restored;PackagesPreserved=$same;ExcelClosed=(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -eq 0);B0Accepted=$false;ReleaseAccepted=$false}|ConvertTo-Json|Set-Content (Join-Path $root 'closure.json')
    if(-not $restored -or -not $same){throw 'Replay test preservation failed.'}
}
Get-Content (Join-Path $root 'worker.log') -Tail 10
exit $code
