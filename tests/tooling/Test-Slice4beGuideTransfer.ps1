[CmdletBinding()]
param([string]$DeployRoot='deploy/validation-print-recorded-01',[ValidateSet('RED','GREEN')][string]$Phase='RED',[switch]$CaptureEvidence,[switch]$SavedHost,[switch]$RoundTrip)
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before isolated guide transfer validation.'}
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$settings=Get-InvSysTestSettingsSnapshot
$root=Join-Path 'reports/runtime/guide-transfer-controller' ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$pins=@(Get-ChildItem -LiteralPath $DeployRoot -Filter '*.xlam' -File|ForEach-Object{[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash -LiteralPath $_.FullName).Hash}})
if($pins.Count -ne 5){throw 'Five frozen packages required.'}
$pins|ConvertTo-Json|Set-Content (Join-Path $root 'package-pins.json')
$start=[DateTimeOffset]::UtcNow;$code=1
Write-Output ('Guide transfer controller: '+$root)
try {
    $extra=@();if($CaptureEvidence){$extra+='-CaptureEvidence'}
    if($RoundTrip){$extra+='-CheckGuideTransferRoundTrip'}
    if($SavedHost){$extra+=@('-CaptureGuideEvidence','-GuideCaptureVisibleExcelForTest','-GuideCaptureSavedWorkbookForTest')}
    $ErrorActionPreference='Continue'
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-Slice4beConfigCommands.ps1') -DeployRoot $DeployRoot -Phase $Phase -CheckGuideTransfer -GuideDraftOnly -CheckGuideExpectation -CheckViewerPublishedRead -ViewerStartupPackageStateForTest SavedCopies -CompileEvaluationProbesForTest -WaitForExcelReadyForTest -ExcelReadyReadLimitForTest 40 @extra *> (Join-Path $root 'worker.log')
    $code=$LASTEXITCODE
    $ErrorActionPreference='Stop'
} finally {
    Wait-RecordingCleanup -Creator $null -Worker $null
    $restored=Restore-InvSysTestSettingsSnapshot $settings
    $same=@($pins|Where-Object {(Get-FileHash -LiteralPath $_.File).Hash -cne $_.Hash}).Count -eq 0
    [pscustomobject]@{Phase=$Phase;CaptureEvidence=[bool]$CaptureEvidence;SavedHost=[bool]$SavedHost;RoundTrip=[bool]$RoundTrip;StartUTC=$start.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');ExitCode=$code;SettingsRestored=$restored;PackagesPreserved=$same;ExcelClosed=(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -eq 0);TransferAccepted=$false}|ConvertTo-Json|Set-Content (Join-Path $root 'closure.json')
    if(-not $restored -or -not $same){throw 'Guide transfer preservation failed.'}
}
Get-Content (Join-Path $root 'worker.log') -Tail 12
exit $code
