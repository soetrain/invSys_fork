[CmdletBinding()]
param([string]$DeployRoot='deploy/validation-detail-columns',[ValidateSet('RED','GREEN')][string]$Phase='RED')
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before isolated Production dismissal validation.'}
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$settings=Get-InvSysTestSettingsSnapshot
$root=Join-Path 'reports/runtime/production-close-controller' ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$pins=@(Get-ChildItem -LiteralPath $DeployRoot -File -Filter '*.xlam'|ForEach-Object{[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash $_.FullName).Hash}})
if($pins.Count -ne 5){throw 'Five packages required.'}
$pins|ConvertTo-Json|Set-Content (Join-Path $root 'package-pins.json')
$start=[DateTimeOffset]::UtcNow;$code=1
Write-Output ('Production Close controller: '+$root)
try {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-Slice4beConfigCommands.ps1') -DeployRoot $DeployRoot -Phase $Phase -CheckProductionDesignerActivity -CheckProductionClose -CaptureEvidence -CompileEvaluationProbesForTest -WaitForExcelReadyForTest -ExcelReadyReadLimitForTest 40 *> (Join-Path $root 'worker.log')
    $code=$LASTEXITCODE
} finally {
    Wait-RecordingCleanup -Creator $null -Worker $null
    $restored=Restore-InvSysTestSettingsSnapshot $settings
    $same=@($pins|Where-Object{(Get-FileHash -LiteralPath $_.File).Hash -cne $_.Hash}).Count -eq 0
    [pscustomobject]@{StartUTC=$start.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');ExitCode=$code;SettingsRestored=$restored;PackagesPreserved=$same;ExcelClosed=$true;HumanAcceptance=$false}|ConvertTo-Json|Set-Content (Join-Path $root 'closure.json')
    if(-not $restored -or -not $same){throw 'Production Close preservation failed.'}
}
Get-Content (Join-Path $root 'worker.log') -Tail 8
exit $code
