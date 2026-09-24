[CmdletBinding()]
param([string]$DeployRoot='deploy/validation-approved-detail-scroll',[ValidateSet('RED','GREEN')][string]$Phase='RED')
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Excel must be closed before the isolated Auth gate.'}
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$snapshot=Get-InvSysTestSettingsSnapshot
$controller=Join-Path 'reports/runtime/auth-read-controller' ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $controller|Out-Null
$pins=@(Get-ChildItem -LiteralPath $DeployRoot -Filter '*.xlam' -File|ForEach-Object {[pscustomobject]@{Path=$_.FullName;Hash=(Get-FileHash -LiteralPath $_.FullName).Hash}})
if($pins.Count -ne 5){throw 'Five packages required.'}
$testPins=@(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.ps1' -File|ForEach-Object {[pscustomobject]@{Path=$_.FullName;Hash=(Get-FileHash -LiteralPath $_.FullName).Hash}})
$testPins|ConvertTo-Json|Set-Content (Join-Path $controller tooling-pins.json)
$pins|ConvertTo-Json|Set-Content (Join-Path $controller package-pins.json)
$start=[DateTimeOffset]::UtcNow;$code=1
Write-Output ('Controller: '+$controller)
try {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-Slice4beConfigCommands.ps1') -DeployRoot $DeployRoot -Phase $Phase -CheckAuthReadOnly -CompileEvaluationProbesForTest -WaitForExcelReadyForTest -ExcelReadyReadLimitForTest 40 *> (Join-Path $controller worker.log)
    $code=$LASTEXITCODE
} finally {
    Wait-RecordingCleanup -Creator $null -Worker $null
    $restored=Restore-InvSysTestSettingsSnapshot $snapshot
    $same=@($pins|Where-Object {(Get-FileHash -LiteralPath $_.Path).Hash -cne $_.Hash}).Count -eq 0
    $testsSame=@($testPins|Where-Object {(Get-FileHash -LiteralPath $_.Path).Hash -cne $_.Hash}).Count -eq 0
    [pscustomobject]@{StartUTC=$start.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');ExitCode=$code;SettingsRestored=$restored;PackagesPreserved=$same;ToolingPreserved=$testsSame;ExcelClosed=$true;ReleaseAccepted=$false}|ConvertTo-Json|Set-Content (Join-Path $controller closure.json)
    if(-not $restored -or -not $same -or -not $testsSame){throw 'Auth gate preservation failed.'}
}
Get-Content (Join-Path $controller worker.log) -Tail 8
exit $code
