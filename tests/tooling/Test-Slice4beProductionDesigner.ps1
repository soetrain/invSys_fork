[CmdletBinding()]
param([string]$DeployRoot='deploy/validation-production-palette',[ValidateSet('RED','GREEN')][string]$Phase='RED',[switch]$CheckPaths,[switch]$CapturePaths)
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Excel must be closed before this isolated gate.'}
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$snapshot=Get-InvSysTestSettingsSnapshot
$controller=Join-Path 'reports/runtime/production-designer-controller' ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $controller|Out-Null
$pins=@(Get-ChildItem -LiteralPath $DeployRoot -Filter '*.xlam' -File|ForEach-Object {[pscustomobject]@{Path=$_.FullName;Hash=(Get-FileHash -LiteralPath $_.FullName).Hash}})
if($pins.Count -ne 5){throw 'Five packages required.'}
$start=[DateTimeOffset]::UtcNow;$code=1
Write-Output ('Controller: '+$controller)
$flags=@('-WaitForExcelReadyForTest','-ExcelReadyReadLimitForTest','40');if($CheckPaths){$flags+='-CheckProductionDesignerPaths'}
if($CapturePaths){if(-not $CheckPaths){throw 'Path capture requires the path gate.'};$flags+='-CaptureProductionDesignerPaths'}
try {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-Slice4beConfigCommands.ps1') -DeployRoot $DeployRoot -Phase $Phase -CheckProductionDesignerActivity -CompileEvaluationProbesForTest @flags *> (Join-Path $controller 'worker.log')
    $code=$LASTEXITCODE
} finally {
    Wait-RecordingCleanup -Creator $null -Worker $null
    $restored=Restore-InvSysTestSettingsSnapshot $snapshot
    $same=@($pins|Where-Object {(Get-FileHash -LiteralPath $_.Path).Hash -cne $_.Hash}).Count -eq 0
    [pscustomobject]@{StartUTC=$start.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');ExitCode=$code;SettingsRestored=$restored;PackagesPreserved=$same;ExcelClosed=$true;ReleaseAccepted=$false}|ConvertTo-Json|Set-Content (Join-Path $controller 'closure.json')
    if(-not $restored -or -not $same){throw 'Preservation failed.'}
}
Get-Content (Join-Path $controller 'worker.log') -Tail 8
exit $code
