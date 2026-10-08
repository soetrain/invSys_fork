[CmdletBinding()]
param([string]$DeployRoot='deploy/validation-guide-transfer-notice-01',[ValidateSet('RED','GREEN')][string]$Phase='RED',[switch]$PolicyOnlyDiagnostic)
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
if($PolicyOnlyDiagnostic -and $Phase -cne 'RED'){throw 'Policy-only diagnosis cannot claim full GREEN.'}
$diagnosticArgs=@()
if($PolicyOnlyDiagnostic){$diagnosticArgs+='-GeneralSettingsPolicyOnly'}
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before General Settings validation.'}
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$settings=Get-InvSysTestSettingsSnapshot
$root=Join-Path 'reports/runtime/general-settings-controller' ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$pins=@(Get-ChildItem -LiteralPath $DeployRoot -Filter '*.xlam' -File|ForEach-Object{[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash -LiteralPath $_.FullName).Hash}})
if($pins.Count -ne 5){throw 'Five frozen packages required.'}
$pins|ConvertTo-Json|Set-Content (Join-Path $root 'package-pins.json')
$start=[DateTimeOffset]::UtcNow;$code=1
Write-Output ('General Settings controller: '+$root)
try {
    $ErrorActionPreference='Continue'
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-Slice4beConfigCommands.ps1') -DeployRoot $DeployRoot -Phase $Phase -CheckSettingsEditorActivity -GeneralSettingsOnly -CompileEvaluationProbesForTest -CaptureEvidence -WaitForExcelReadyForTest -ExcelReadyReadLimitForTest 40 @diagnosticArgs *> (Join-Path $root 'worker.log')
    $code=$LASTEXITCODE
    $ErrorActionPreference='Stop'
} finally {
    Wait-RecordingCleanup -Creator $null -Worker $null
    $restored=Restore-InvSysTestSettingsSnapshot $settings
    $same=@($pins|Where-Object {(Get-FileHash -LiteralPath $_.File).Hash -cne $_.Hash}).Count -eq 0
    [pscustomobject]@{Phase=$Phase;PolicyOnlyDiagnostic=[bool]$PolicyOnlyDiagnostic;StartUTC=$start.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');ExitCode=$code;SettingsRestored=$restored;PackagesPreserved=$same;ExcelClosed=(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -eq 0);Accepted=$false}|ConvertTo-Json|Set-Content (Join-Path $root 'closure.json')
    if(-not $restored -or -not $same){throw 'General Settings preservation failed.'}
}
Get-Content (Join-Path $root 'worker.log') -Tail 12
exit $code
