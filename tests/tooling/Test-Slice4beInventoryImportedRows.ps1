[CmdletBinding()]
param(
    [string]$DeployRoot='deploy/validation-tester-recording-02',
    [ValidateSet('RED','GREEN')][string]$Phase='RED',
    [Parameter(Mandatory=$true)][string]$DesktopMonitorReport,
    [switch]$TesterBaselineOverlayForTest,
    [switch]$TraceResourcesForTest,
    [switch]$WorksheetResultsRegression,
    [string]$ReceiptScenarioForTest='',
    [ValidateSet('ImportedRows','ImportedEdits','EditReceiptPhases','EditReceiptResults','EditReceiptSave','EditReceiptAll','EditBindings','EditOwnership','WorksheetResults','PrimitivePayload','CatalogCompletion','Headless','CreatePhases','CreateName','ConfigMetadata')]
    [string]$InventoryCase='ImportedRows'
)
# Source-overlay diagnostic only. Preserve frozen packages, tooling, product
# source and local settings; audit native shutdown independently of worker exit.
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
if($WorksheetResultsRegression){
    if($InventoryCase -notin @('ImportedRows','WorksheetResults')){throw 'Conflicting Inventory case selectors.'}
    $InventoryCase='WorksheetResults'
}
if($InventoryCase -in @('WorksheetResults','CatalogCompletion','Headless') -and $Phase -ne 'GREEN'){throw 'Retained Inventory regressions require the GREEN phase.'}
if($ReceiptScenarioForTest -ne ''){
    $cohorts=@{EditReceiptPhases='Phase';EditReceiptResults='Result';EditReceiptSave='Canonical'}
    if(-not $cohorts.ContainsKey($InventoryCase)){throw 'Receipt scenario requires a focused receipt cohort.'}
    . (Join-Path $PSScriptRoot 'Slice4beInventoryReceiptScenarios.ps1')
    [void](Get-InventoryReceiptScenarios -Cohort $cohorts[$InventoryCase] -Scenario $ReceiptScenarioForTest)
}
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before the isolated Inventory gate.'}
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$repo=Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Set-Location -LiteralPath $repo
$snapshot=Get-InvSysTestSettingsSnapshot
$root=Join-Path $repo ('reports/runtime/inventory-imported-controller/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$pins=@(Get-ChildItem -LiteralPath $DeployRoot -File -Filter '*.xlam'|ForEach-Object {[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash -LiteralPath $_.FullName).Hash}})
if($pins.Count -ne 5){throw 'Five frozen packages required.'}
$testPins=@(Get-ChildItem -LiteralPath $PSScriptRoot -File -Filter '*.ps1'|ForEach-Object {[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash -LiteralPath $_.FullName).Hash}})
$sourcePins=@(Get-ChildItem -LiteralPath (Join-Path $repo 'src') -File -Recurse|ForEach-Object {[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash -LiteralPath $_.FullName).Hash}})
$pins|ConvertTo-Json|Set-Content (Join-Path $root 'package-pins.json')
$testPins|ConvertTo-Json|Set-Content (Join-Path $root 'tooling-pins.json')
$sourcePins|ConvertTo-Json|Set-Content (Join-Path $root 'source-pins.json')
$start=[DateTimeOffset]::UtcNow;$code=1;$worker=$null
Write-Output ('Inventory imported controller: '+$root)
try{
    $workerArgs=@('-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',
        (Join-Path $PSScriptRoot 'Test-Slice4beConfigCommands.ps1'),'-DeployRoot',$DeployRoot,
        '-Phase',$Phase,'-AdminInventoryEntryOnly',
        '-AdminInventorySourceOverlayForTest',
        '-CompileEvaluationProbesForTest','-WaitForExcelReadyForTest','-ExcelReadyReadLimitForTest','40',
        '-DesktopMonitorReport',$DesktopMonitorReport)
    if($TesterBaselineOverlayForTest){$workerArgs+='-AdminInventoryTesterBaselineOverlayForTest'}
    if($TraceResourcesForTest){$workerArgs+='-TraceGuideResourcesForTest'}
    if($ReceiptScenarioForTest -ne ''){$workerArgs+=@('-AdminInventoryReceiptScenarioForTest',$ReceiptScenarioForTest)}
    switch($InventoryCase){
        'ImportedRows' {$workerArgs+=@('-AdminInventoryWorksheetResultsOnly','-AdminInventoryImportedRowsOnly')}
        'ImportedEdits' {$workerArgs+=@('-AdminInventoryWorksheetResultsOnly','-AdminInventoryImportedRowsOnly','-AdminInventoryImportedEditsOnly')}
        'EditReceiptPhases' {$workerArgs+=@('-AdminInventoryWorksheetResultsOnly','-AdminInventoryEditReceiptPhasesOnly')}
        'EditReceiptResults' {$workerArgs+=@('-AdminInventoryWorksheetResultsOnly','-AdminInventoryEditReceiptPhasesOnly','-AdminInventoryEditReceiptResultsOnly')}
        'EditReceiptSave' {$workerArgs+=@('-AdminInventoryWorksheetResultsOnly','-AdminInventoryEditReceiptPhasesOnly','-AdminInventoryEditCanonicalSaveOnly')}
        'EditReceiptAll' {$workerArgs+=@('-AdminInventoryWorksheetResultsOnly','-AdminInventoryEditReceiptPhasesOnly','-AdminInventoryEditReceiptAllOnly')}
        'EditBindings' {$workerArgs+=@('-AdminInventoryWorksheetResultsOnly','-AdminInventoryEditReceiptPhasesOnly','-AdminInventoryEditBindingsOnly')}
        'EditOwnership' {$workerArgs+=@('-AdminInventoryWorksheetResultsOnly','-AdminInventoryEditReceiptPhasesOnly','-AdminInventoryEditOwnershipOnly')}
        'ConfigMetadata' {$workerArgs+=@('-AdminInventoryWorksheetResultsOnly','-AdminInventoryImportedRowsOnly','-AdminInventoryConfigMetadataOnly')}
        'WorksheetResults' {$workerArgs+='-AdminInventoryWorksheetResultsOnly'}
        'PrimitivePayload' {$workerArgs+='-AdminInventoryPrimitivePayloadOnly'}
        'CatalogCompletion' {$workerArgs+='-AdminInventoryOwnResultsOnly'}
        'Headless' {$workerArgs+='-AdminInventoryHeadlessWorksheetOnly'}
        'CreatePhases' {$workerArgs+='-AdminInventoryCreatePhasesOnly'}
        'CreateName' {$workerArgs+=@('-AdminInventoryCreatePhasesOnly','-AdminInventoryCreateNameOnly')}
    }
    if(@($workerArgs|Where-Object {$_ -match '["\r\n]'}).Count){throw 'Unsupported character in worker argument.'}
    $launch=[Diagnostics.ProcessStartInfo]::new((Join-Path $PSHOME 'powershell.exe'))
    $launch.Arguments=($workerArgs|ForEach-Object {'"'+$_+'"'}) -join ' '
    $launch.WorkingDirectory=$repo;$launch.UseShellExecute=$false;$launch.CreateNoWindow=$true
    $launch.RedirectStandardOutput=$true;$launch.RedirectStandardError=$true
    $outputFile=[IO.FileStream]::new((Join-Path $root 'worker.log'),[IO.FileMode]::Create,[IO.FileAccess]::Write,[IO.FileShare]::Read,1,$true)
    $errorFile=[IO.FileStream]::new((Join-Path $root 'worker-error.log'),[IO.FileMode]::Create,[IO.FileAccess]::Write,[IO.FileShare]::Read,1,$true)
    try{
        $worker=[Diagnostics.Process]::Start($launch)
        [pscustomobject]@{ProcessId=$worker.Id;StartUTC=$worker.StartTime.ToUniversalTime().ToString('o');InventoryCase=$InventoryCase;ReceiptScenario=$ReceiptScenarioForTest}|ConvertTo-Json|Set-Content (Join-Path $root 'worker-identity.json')
        $outputCopy=$worker.StandardOutput.BaseStream.CopyToAsync($outputFile,256)
        $errorCopy=$worker.StandardError.BaseStream.CopyToAsync($errorFile,256)
        $worker.WaitForExit();$outputCopy.Wait();$errorCopy.Wait()
        if(-not $worker.HasExited){throw 'Inventory worker termination was not confirmed.'}
        $code=$worker.ExitCode
    }finally{$outputFile.Dispose();$errorFile.Dispose()}
}finally{
    Wait-RecordingCleanup -Creator $null -Worker $worker
    $restored=Restore-InvSysTestSettingsSnapshot $snapshot
    $same=@($pins|Where-Object {(Get-FileHash -LiteralPath $_.File).Hash -cne $_.Hash}).Count -eq 0
    $testsSame=@($testPins|Where-Object {(Get-FileHash -LiteralPath $_.File).Hash -cne $_.Hash}).Count -eq 0
    $sourceSame=@($sourcePins|Where-Object {(Get-FileHash -LiteralPath $_.File).Hash -cne $_.Hash}).Count -eq 0
    [pscustomobject]@{StartUTC=$start.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');ExitCode=$code;SettingsRestored=$restored;PackagesPreserved=$same;ToolingPreserved=$testsSame;SourcePreserved=$sourceSame;ExcelClosed=$true;FullAcceptance=$false;PackagedAction=$false;NativeInputAcceptance=$false}|ConvertTo-Json|Set-Content (Join-Path $root 'closure.json')
    if(-not($restored -and $same -and $testsSame -and $sourceSame)){throw 'Inventory diagnostic preservation failed.'}
}
Start-Sleep -Seconds 20
$audit=[DateTimeOffset]::UtcNow
$events=@(Get-WinEvent -FilterHashtable @{LogName='Application';Id=1000,1001,1002;StartTime=$start.LocalDateTime;EndTime=$audit.LocalDateTime} -ErrorAction SilentlyContinue -ErrorVariable queryErrors|Where-Object {$_.Message -match '(?i)excel\.exe'})
if(@($queryErrors|Where-Object FullyQualifiedErrorId -notlike 'NoMatchingEventsFound*').Count){throw 'Inventory native audit unavailable.'}
$closed=@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -eq 0
[pscustomobject]@{StartUTC=$start.ToString('o');AuditUTC=$audit.ToString('o');ExcelEventCount=$events.Count;ExcelClosed=$closed}|ConvertTo-Json|Set-Content (Join-Path $root 'native-audit.json')
if($events.Count -or -not $closed){throw 'Inventory native closure did not pass.'}
if((Test-Path (Join-Path $root 'forced-test-excel-stop.json')) -or (Test-Path (Join-Path $root 'native-owned-stop.json'))){throw 'Process intervention was recorded; normal closure is not qualified.'}
Get-Content (Join-Path $root 'worker.log') -Tail 14
exit $code
