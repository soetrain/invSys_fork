[CmdletBinding()]
param([string]$DeployRoot='deploy/validation-production-header',[ValidateSet('RED','GREEN')][string]$Phase='RED')
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before the isolated lifecycle gate.'}
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$snapshot=Get-InvSysTestSettingsSnapshot
$controller=Join-Path 'reports/runtime/production-lifecycle-controller' ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $controller|Out-Null
$pins=@(Get-ChildItem -LiteralPath $DeployRoot -Filter '*.xlam' -File|ForEach-Object {[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash -LiteralPath $_.FullName).Hash}})
if($pins.Count -ne 5){throw 'Five packages required.'}
$pins|ConvertTo-Json|Set-Content (Join-Path $controller 'package-pins.json')
$start=[DateTimeOffset]::UtcNow;$code=1
Write-Output ('Lifecycle controller: '+$controller)
try {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-Slice4beConfigCommands.ps1') -DeployRoot $DeployRoot -Phase $Phase -CheckProductionDesignerActivity -CheckProductionLifecycle -CompileEvaluationProbesForTest -WaitForExcelReadyForTest -ExcelReadyReadLimitForTest 40 *> (Join-Path $controller 'worker.log')
    $code=$LASTEXITCODE
} finally {
    Wait-RecordingCleanup -Creator $null -Worker $null
    $restored=Restore-InvSysTestSettingsSnapshot $snapshot
    $preserved=@($pins|Where-Object {(Get-FileHash -LiteralPath $_.File).Hash -cne $_.Hash}).Count -eq 0
    [pscustomobject]@{StartUTC=$start.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');ExitCode=$code;SettingsRestored=$restored;PackagesPreserved=$preserved;ExcelClosed=$true;FullAcceptance=$false}|ConvertTo-Json|Set-Content (Join-Path $controller 'closure.json')
    if(-not $restored -or -not $preserved){throw 'Lifecycle preservation failed.'}
}
Get-Content (Join-Path $controller 'worker.log') -Tail 8
exit $code
