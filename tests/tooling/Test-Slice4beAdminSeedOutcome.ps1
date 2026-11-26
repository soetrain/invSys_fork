[CmdletBinding()]
param([string]$DeployRoot='deploy/validation-detail-glyph-01',[ValidateSet('RED','GREEN')][string]$Phase='RED',[switch]$NativeFeedback,[switch]$WorkflowEdges,[switch]$Recording,[switch]$RecordingBoundaries,[switch]$GuidePaths,[switch]$GuideTiming,[switch]$SeedReplayProfile,[switch]$SeedReplayRun,[switch]$SeedReplayLostAck,[ValidateSet('Start','Reopen')][string]$SeedReplayNested)
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
if($SeedReplayLostAck -and $SeedReplayNested){throw 'Select a single Seed return case.'}
if($SeedReplayLostAck -or $SeedReplayNested){$SeedReplayRun=$true}
if($SeedReplayRun){$SeedReplayProfile=$true}
if($SeedReplayProfile){$Recording=$true;$RecordingBoundaries=$true;$GuidePaths=$true}
if($RecordingBoundaries -and -not $Recording){throw 'Recording boundaries require the retained recording gate.'}
if($GuidePaths -and -not $RecordingBoundaries){throw 'Seed guides require the qualified boundary probes.'}
if($GuideTiming -and -not $GuidePaths){throw 'Guide timing requires the Seed guide fixture.'}
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before the isolated Admin Seed outcome gate.'}
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$snapshot=Get-InvSysTestSettingsSnapshot
$root=Join-Path 'reports/runtime/admin-seed-outcome-controller' ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$pins=@(Get-ChildItem -LiteralPath $DeployRoot -File -Filter '*.xlam'|ForEach-Object {[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash $_.FullName).Hash}})
if($pins.Count -ne 5){throw 'Five packages required.'}
$testPins=@(Get-ChildItem -LiteralPath $PSScriptRoot -File -Filter '*.ps1'|ForEach-Object {[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash $_.FullName).Hash}})
$pins|ConvertTo-Json|Set-Content (Join-Path $root 'package-pins.json')
$testPins|ConvertTo-Json|Set-Content (Join-Path $root 'tooling-pins.json')
$start=[DateTimeOffset]::UtcNow;$code=1
Write-Output ('Admin Seed outcome controller: '+$root)
try{
    $extra=@();if($NativeFeedback){$extra+='-AdminSeedNativeFeedback'}
    if($SeedReplayRun){$extra+='-AdminSeedReplayRun'}
    if($SeedReplayLostAck){$extra+='-AdminSeedReplayLostAck'}
    if($SeedReplayNested){$extra+=@('-AdminSeedReplayNested',$SeedReplayNested)}
    if($SeedReplayProfile){$extra+='-AdminSeedReplayProfile';$extra+='-ViewerStartupPackageStateForTest';$extra+='SavedCopies'}
    if($WorkflowEdges){$extra+='-AdminSeedWorkflowEdges'}
    if($Recording){$extra+='-AdminSeedRecording'}
    if($RecordingBoundaries){$extra+='-AdminSeedRecordingBoundaries'}
    if($GuidePaths){$extra+='-AdminSeedGuidePaths';$extra+='-CaptureEvidence'}
    if($GuideTiming){$extra+='-AdminSeedGuideTimingOnly';$extra+='-ViewerStartupPackageStateForTest';$extra+='SavedCopies'}
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-Slice4beConfigCommands.ps1') -DeployRoot $DeployRoot -Phase $Phase -AdminSeedOutcomeOnly -CompileEvaluationProbesForTest -WaitForExcelReadyForTest -ExcelReadyReadLimitForTest 40 @extra *> (Join-Path $root 'worker.log')
    $code=$LASTEXITCODE
}finally{
    Wait-RecordingCleanup -Creator $null -Worker $null
    $restored=Restore-InvSysTestSettingsSnapshot $snapshot
    $same=@($pins|Where-Object {(Get-FileHash $_.File).Hash -cne $_.Hash}).Count -eq 0
    $testsSame=@($testPins|Where-Object {(Get-FileHash $_.File).Hash -cne $_.Hash}).Count -eq 0
    [pscustomobject]@{StartUTC=$start.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');ExitCode=$code;SettingsRestored=$restored;PackagesPreserved=$same;ToolingPreserved=$testsSame;ExcelClosed=$true;FullAcceptance=$false;NativeInputAcceptance=$false}|ConvertTo-Json|Set-Content (Join-Path $root 'closure.json')
    if(-not($restored -and $same -and $testsSame)){throw 'Admin Seed outcome preservation failed.'}
}
Get-Content (Join-Path $root 'worker.log') -Tail 14
exit $code
