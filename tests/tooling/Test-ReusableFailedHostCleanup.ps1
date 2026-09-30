# Execute the real validator's final cleanup block against disposable host doubles.
# No Excel process, workbook, Windows policy or runtime VBA is changed.
[CmdletBinding()]
param([string]$RepoRoot='.')
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
. (Join-Path $PSScriptRoot 'ProductionReusableCleanup.ps1')
$validator=[IO.File]::ReadAllText((Join-Path $repo 'tools/validate_plan022_packaged_launchers.ps1'))
$start=$validator.IndexOf('    $paletteBooksClosed=$false',[StringComparison]::Ordinal)
$end=$validator.IndexOf('    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())',$start,[StringComparison]::Ordinal)
if($start -lt 0 -or $end -le $start){throw 'Final cleanup anchor missing; no calibration RED.'}
$block=[scriptblock]::Create($validator.Substring($start,$end-$start))
$root=Join-Path $repo ('reports/runtime/reusable-failed-host-calibration/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$checks=[Collections.Generic.List[object]]::new()
function Check([string]$Name,[bool]$Passed){$checks.Add([pscustomobject]@{Check=$Name;Passed=$Passed})}
function Release-ComObject($Value){$script:releaseCalls++}
function Wait-ReusableAutomationExit {
    param([int]$ProcessId,$Variables)
    $script:exitCalls++
    [pscustomobject]@{Failure=$null;UnassistedExitObserved=$true;TestDouble=$true}
}
function Add-Evidence {
    param($Rows,[string]$Callback,[string]$Expected,[bool]$Passed,[string]$Observed)
    $Rows.Add([pscustomobject]@{Callback=$Callback;Expected=$Expected;Passed=$Passed;Observed=$Observed})
}
foreach($case in @('DeadWorkbookMetadata','MissingPackageMetadata','SuccessfulEmptyHost')){
    $outputPath=Join-Path $root $case;New-Item -ItemType Directory -Path $outputPath|Out-Null
    $runtimeRoot=Join-Path ([IO.Path]::GetTempPath()) ('invsys-plan022-launcher-red-'+[guid]::NewGuid().ToString('N'))
    $deployPath=Join-Path $root 'packages'
    $ProductionPaletteProbe=$false;$WorkbookState='ProductionReusable';$excelProcessId=0
    $script:releaseCalls=0;$script:exitCalls=0;$script:quitCalls=0
    $books=@();if($case -eq 'DeadWorkbookMetadata'){$books=@([pscustomobject]@{})}
    $excel=[pscustomobject]@{Workbooks=$books}
    $excel|Add-Member ScriptMethod Quit {$script:quitCalls++}
    $packages=@{};$packageNames=@()
    if($case -eq 'MissingPackageMetadata'){$packageNames=@('invSys.Operations.xlam')}
    $opened=[Collections.Generic.List[object]]::new()
    $evidence=[Collections.Generic.List[object]]::new()
    Add-Evidence $evidence 'PRIMARY_FAILURE' 'Original workflow evidence retained.' $false 'Synthetic primary failure'
    $reportPath=Join-Path $outputPath 'report.md';[IO.File]::WriteAllText($reportPath,'Primary failure retained.')
    $threw=$false;try{. $block}catch{$threw=$true}
    $closurePath=Join-Path $outputPath 'final-workbook-closure.json'
    $closure=$null;if(Test-Path $closurePath){$closure=Get-Content $closurePath -Raw|ConvertFrom-Json}
    $isFailure=$case -ne 'SuccessfulEmptyHost'
    Check "$case.NoSecondaryEscape" (-not $threw)
    Check "$case.QuitAndReleaseAttempted" ($script:quitCalls -eq 1 -and $script:releaseCalls -ge 1 -and $script:exitCalls -eq 1)
    Check "$case.FinalReceiptsWritten" ((Test-Path (Join-Path $outputPath 'final-release.json')) -and (Test-Path (Join-Path $outputPath 'final-cleanup-observation.json')))
    $state=$null -ne $closure -and $null -ne $closure.PSObject.Properties['Completed']
    if($state){$state=[bool]$closure.Completed -eq (-not $isFailure) -and ($null -ne $closure.Failure) -eq $isFailure}
    Check "$case.ClosureStateTruthful" $state
    $extra=@($evidence|Where-Object Callback -CEQ 'HARNESS_CLEANUP')
    Check "$case.CleanupFailureRemainsFailure" ($(if($isFailure){$extra.Count -eq 1 -and -not $extra[0].Passed -and [IO.File]::ReadAllText($reportPath).Contains('| HARNESS_CLEANUP | RED |')}else{$extra.Count -eq 0}))
    Check "$case.PrimaryFailureRetained" ($evidence[0].Callback -ceq 'PRIMARY_FAILURE' -and -not $evidence[0].Passed -and $evidence[0].Observed -ceq 'Synthetic primary failure')
    $safe=$state
    if($safe -and $isFailure){$safe=@($closure.Failure.PSObject.Properties.Name|Where-Object {$_ -cnotin @('Stage','ExceptionType','HResult')}).Count -eq 0 -and $closure.Failure.Stage -cin @('Workbooks','Packages') -and [IO.File]::ReadAllText($closurePath) -notmatch '[A-Za-z]:\\|StackTrace|Message|FullName|IsAddin'}
    Check "$case.FailureReceiptRedacted" $safe
}
$checks|ConvertTo-Json|Set-Content (Join-Path $root 'checks.json')
$failed=@($checks|Where-Object {-not $_.Passed}).Count
[pscustomobject]@{Root=$root;Passed=$checks.Count-$failed;Failed=$failed;ExcelOpened=$false}|ConvertTo-Json
if($failed){exit 1}
