[CmdletBinding()]
param([string]$DeployRoot='deploy/validation-production-uom-extent',[ValidateSet('RED','GREEN')][string]$Phase='RED',[switch]$CaptureEvidence)
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Excel must be closed before this isolated gate.'}
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$snapshot=Get-InvSysTestSettingsSnapshot
$root=Join-Path 'reports/runtime/production-component-controller' ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$pins=@(Get-ChildItem -LiteralPath $DeployRoot -Filter '*.xlam' -File|ForEach-Object{[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash -LiteralPath $_.FullName).Hash}})
if($pins.Count -ne 5){throw 'Five packages required.'}
$pins|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $root 'package-pins.json')
$start=[DateTimeOffset]::UtcNow;$code=1
Write-Output ('Component controller: '+$root)
try {
    $extra=@();if($CaptureEvidence){$extra+=@('-CaptureEvidence')}
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-Slice4beConfigCommands.ps1') -DeployRoot $DeployRoot -Phase $Phase -CheckProductionDesignerActivity -CheckProductionComponents -CompileEvaluationProbesForTest -WaitForExcelReadyForTest -ExcelReadyReadLimitForTest 40 @extra *> (Join-Path $root 'worker.log')
    $code=$LASTEXITCODE
}finally{
    Wait-RecordingCleanup -Creator $null -Worker $null
    $restored=Restore-InvSysTestSettingsSnapshot $snapshot
    $same=@($pins|Where-Object{(Get-FileHash -LiteralPath $_.File).Hash -cne $_.Hash}).Count -eq 0
    [pscustomobject]@{Phase=$Phase;CaptureEvidence=[bool]$CaptureEvidence;StartUTC=$start.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');ExitCode=$code;SettingsRestored=$restored;PackagesPreserved=$same;ExcelClosed=$true;ReleaseAccepted=$false}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $root 'closure.json')
    if(-not $restored -or -not $same){throw 'Component gate preservation failed.'}
}
Get-Content -LiteralPath (Join-Path $root 'worker.log') -Tail 6
exit $code
