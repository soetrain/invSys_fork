[CmdletBinding()]
param([string]$DeployRoot='deploy/validation-production-assignment-01',[ValidateSet('RED','GREEN')][string]$Phase='RED',[switch]$ClosedDiagnostic,[switch]$ContractOnly,[switch]$PolicyOnly,[switch]$FaultOnly)
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
if(@($ClosedDiagnostic,$ContractOnly,$PolicyOnly,$FaultOnly|Where-Object{$_}).Count -gt 1){throw 'Choose one isolated Run companion gate.'}
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before isolated Run local validation.'}
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$snapshot=Get-InvSysTestSettingsSnapshot
$root=Join-Path 'reports/runtime/production-run-local-controller' ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$pins=@(Get-ChildItem -LiteralPath $DeployRoot -Filter '*.xlam' -File|ForEach-Object{[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash -LiteralPath $_.FullName).Hash}})
if($pins.Count -ne 5){throw 'Five packages required.'}
$pins|ConvertTo-Json|Set-Content (Join-Path $root 'package-pins.json')
$start=[DateTimeOffset]::UtcNow;$code=1
Write-Output ('Run local controller: '+$root)
try{
    $extra=@();if($ClosedDiagnostic){$extra=@('-RunLocalClosedDiagnostic')}
    if($ContractOnly){$extra=@('-RunLocalContractOnly')}
    if($PolicyOnly){$extra=@('-RunLocalPolicyOnly')}
    if($FaultOnly){$extra=@('-RunLocalFaultOnly')}
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-Slice4beConfigCommands.ps1') -DeployRoot $DeployRoot -Phase $Phase -CheckProductionDesignerActivity -CheckProductionRunLocal -CompileEvaluationProbesForTest -WaitForExcelReadyForTest -ExcelReadyReadLimitForTest 40 @extra *> (Join-Path $root 'worker.log')
    $code=$LASTEXITCODE
}finally{
    Wait-RecordingCleanup -Creator $null -Worker $null
    $restored=Restore-InvSysTestSettingsSnapshot $snapshot
    $same=@($pins|Where-Object{(Get-FileHash -LiteralPath $_.File).Hash -cne $_.Hash}).Count -eq 0
    [pscustomobject]@{Phase=$Phase;ClosedDiagnostic=[bool]$ClosedDiagnostic;ContractOnly=[bool]$ContractOnly;PolicyOnly=[bool]$PolicyOnly;FaultOnly=[bool]$FaultOnly;StartUTC=$start.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');ExitCode=$code;SettingsRestored=$restored;PackagesPreserved=$same;ExcelClosed=$true;ReleaseAccepted=$false}|ConvertTo-Json|Set-Content (Join-Path $root 'closure.json')
    if(-not $restored -or -not $same){throw 'Run local preservation failed.'}
}
Get-Content (Join-Path $root 'worker.log') -Tail 8
exit $code
