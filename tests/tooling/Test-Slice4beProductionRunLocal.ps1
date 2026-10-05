[CmdletBinding()]
param([string]$DeployRoot='deploy/validation-production-assignment-01',[ValidateSet('RED','GREEN')][string]$Phase='RED',[ValidateSet('None','Submission','Entry','Interruptions','Initial','PriorSequence')][string]$CompletePreparationPrelude='None',[switch]$CompletePreparationDiagnostic,[switch]$SavedProbeCopiesForTest,[switch]$SavedDecoyForTest,[switch]$ClosedDiagnostic,[switch]$ContractOnly,[switch]$PolicyOnly,[switch]$FaultOnly,[switch]$YieldOnly,[switch]$StockOnly,[switch]$WorksheetOnly,[switch]$WorksheetOwnerOnly,[switch]$WorksheetScaleOnly,[switch]$ClearPathsOnly,[switch]$LoadPathsOnly,[switch]$RefreshPathsOnly,[switch]$RefillOnly,[switch]$AllocatePathsOnly,[switch]$CompleteBaselineOnly,[switch]$NextBaselineOnly,[switch]$NextActivityOnly,[switch]$CheckInBaselineOnly,[switch]$CheckInRoutedOnly,[switch]$CheckInClosedOnly,[switch]$CheckInActivityOnly,[switch]$CheckInPathsOnly,[ValidateSet('Reusable','Worksheet')][string]$CheckInPathMode='Reusable')
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
if($CompletePreparationPrelude -ne 'None' -and -not $CompletePreparationDiagnostic){throw 'Preparation prelude requires explicit diagnostic mode.'}
if($CompletePreparationDiagnostic -and (-not $CompleteBaselineOnly -or $NextBaselineOnly)){throw 'Preparation diagnosis requires the isolated Complete Run fixture.'}
if($SavedProbeCopiesForTest -and (-not $CompleteBaselineOnly -or $NextBaselineOnly)){throw 'Saved probe copies require the isolated completion gate.'}
if($SavedDecoyForTest -and (-not $CompleteBaselineOnly -or $NextBaselineOnly)){throw 'Saved decoy requires the isolated completion gate.'}
if($NextActivityOnly -and -not $NextBaselineOnly){throw 'Next Batch activity requires NextBaselineOnly.'}
if($NextBaselineOnly -and -not $CompleteBaselineOnly){throw 'Next Batch baseline requires CompleteBaselineOnly fixture setup.'}
if(@($ClosedDiagnostic,$ContractOnly,$PolicyOnly,$FaultOnly,$YieldOnly,$StockOnly,$WorksheetOnly,$WorksheetOwnerOnly,$WorksheetScaleOnly,$ClearPathsOnly,$LoadPathsOnly,$RefreshPathsOnly,$RefillOnly,$AllocatePathsOnly,$CompleteBaselineOnly,$CheckInBaselineOnly,$CheckInRoutedOnly,$CheckInClosedOnly,$CheckInActivityOnly,$CheckInPathsOnly|Where-Object{$_}).Count -gt 1){throw 'Choose one isolated Run companion gate.'}
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
    if($YieldOnly){$extra=@('-RunLocalYieldOnly')}
    if($StockOnly){$extra=@('-RunLocalStockOnly')}
    if($WorksheetOnly){$extra=@('-RunLocalWorksheetOnly')}
    if($WorksheetOwnerOnly){$extra=@('-RunLocalWorksheetOwnerOnly')}
    if($WorksheetScaleOnly){$extra=@('-RunLocalWorksheetScaleOnly')}
    if($ClearPathsOnly){$extra=@('-RunClearPathsOnly','-CaptureEvidence')}
    if($LoadPathsOnly){$extra=@('-RunLoadPathsOnly','-CaptureEvidence')}
    if($RefreshPathsOnly){$extra=@('-RunRefreshPathsOnly','-CaptureEvidence')}
    if($AllocatePathsOnly){$extra=@('-RunAllocatePathsOnly','-CaptureEvidence')}
    if($CompleteBaselineOnly){$extra=@('-RunCompleteBaselineOnly','-CaptureEvidence');if($NextBaselineOnly){$extra+='-RunNextBaselineOnly'};if($NextActivityOnly){$extra+='-RunNextActivityOnly'}}
    if($CheckInBaselineOnly){$extra=@('-RunCheckInBaselineOnly','-CaptureEvidence')}
    if($CheckInRoutedOnly){$extra=@('-RunCheckInRoutedOnly','-CaptureEvidence')}
    if($CheckInClosedOnly){$extra=@('-RunCheckInClosedOnly','-CaptureEvidence')}
    if($CheckInActivityOnly){$extra=@('-RunCheckInActivityOnly','-CaptureEvidence')}
    if($CheckInPathsOnly){$extra=@('-RunCheckInPathsOnly','-RunCheckInPathMode',$CheckInPathMode,'-CaptureEvidence')}
    if($CompletePreparationDiagnostic){$extra+=@('-RunCompletePreparationDiagnostic','-RunCompletePreparationPrelude',$CompletePreparationPrelude)}
    if($SavedProbeCopiesForTest){$extra+=@('-ViewerStartupPackageStateForTest','SavedCopies')}
    if($SavedDecoyForTest){$extra+='-RunCompleteSavedDecoyForTest'}
    if($RefillOnly){$extra=@('-RunLocalRefillDiagnostic')}
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-Slice4beConfigCommands.ps1') -DeployRoot $DeployRoot -Phase $Phase -CheckProductionDesignerActivity -CheckProductionRunLocal -CompileEvaluationProbesForTest -WaitForExcelReadyForTest -ExcelReadyReadLimitForTest 40 @extra *> (Join-Path $root 'worker.log')
    $code=$LASTEXITCODE
}finally{
    Wait-RecordingCleanup -Creator $null -Worker $null
    $restored=Restore-InvSysTestSettingsSnapshot $snapshot
    $same=@($pins|Where-Object{(Get-FileHash -LiteralPath $_.File).Hash -cne $_.Hash}).Count -eq 0
    [pscustomobject]@{Phase=$Phase;CompletePreparationDiagnostic=[bool]$CompletePreparationDiagnostic;CompletePreparationPrelude=$CompletePreparationPrelude;SavedProbeCopiesForTest=[bool]$SavedProbeCopiesForTest;SavedDecoyForTest=[bool]$SavedDecoyForTest;NextBaselineOnly=[bool]$NextBaselineOnly;NextActivityOnly=[bool]$NextActivityOnly;CompleteBaselineOnly=[bool]$CompleteBaselineOnly;CheckInPathsOnly=[bool]$CheckInPathsOnly;CheckInPathMode=$CheckInPathMode;CheckInActivityOnly=[bool]$CheckInActivityOnly;CheckInClosedOnly=[bool]$CheckInClosedOnly;CheckInRoutedOnly=[bool]$CheckInRoutedOnly;CheckInBaselineOnly=[bool]$CheckInBaselineOnly;AllocatePathsOnly=[bool]$AllocatePathsOnly;RefillOnly=[bool]$RefillOnly;ClosedDiagnostic=[bool]$ClosedDiagnostic;ContractOnly=[bool]$ContractOnly;PolicyOnly=[bool]$PolicyOnly;FaultOnly=[bool]$FaultOnly;YieldOnly=[bool]$YieldOnly;StockOnly=[bool]$StockOnly;WorksheetOnly=[bool]$WorksheetOnly;WorksheetOwnerOnly=[bool]$WorksheetOwnerOnly;WorksheetScaleOnly=[bool]$WorksheetScaleOnly;ClearPathsOnly=[bool]$ClearPathsOnly;LoadPathsOnly=[bool]$LoadPathsOnly;RefreshPathsOnly=[bool]$RefreshPathsOnly;StartUTC=$start.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');ExitCode=$code;SettingsRestored=$restored;PackagesPreserved=$same;ExcelClosed=$true;ReleaseAccepted=$false}|ConvertTo-Json|Set-Content (Join-Path $root 'closure.json')
    if(-not $restored -or -not $same){throw 'Run local preservation failed.'}
}
Get-Content (Join-Path $root 'worker.log') -Tail 8
exit $code
