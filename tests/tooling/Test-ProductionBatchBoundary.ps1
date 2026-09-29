# Scoped diagnostic of the standard launcher sequence through batch scale only.
[CmdletBinding()]
param([string]$RepoRoot='.',[string]$DeployRoot='deploy/validation-production-paths',
      [string]$PackagePinsPath='reports/runtime/production-lifecycle-native-controller/65e8e27585ca47289cf27e87e147e83e/package-pins.json',
      [switch]$TraceBoundaries)
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before the isolated diagnostic.'}
$deploy=(Resolve-Path -LiteralPath (Join-Path $repo $DeployRoot)).Path
$rawPins=Get-Content -LiteralPath (Join-Path $repo $PackagePinsPath) -Raw|ConvertFrom-Json
$pins=@($rawPins)
if($pins.Count -ne 5){throw 'Five package pins required.'}
foreach($pin in $pins){
    $name=if($pin.PSObject.Properties['Package']){$pin.Package}else{Split-Path -Leaf $pin.File}
    if((Get-FileHash -LiteralPath (Join-Path $deploy $name)).Hash -cne $pin.Hash){throw 'Candidate package differs.'}
}
$root=Join-Path $repo ('reports/runtime/production-batch-boundary/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
Write-Output ('Diagnostic: '+$root)
$validator=Join-Path $repo 'tools/validate_plan022_packaged_launchers.ps1'
$validatorHash=(Get-FileHash -LiteralPath $validator).Hash
$source=[IO.File]::ReadAllText($validator)
function Replace-Once([string]$Text,[string]$Anchor,[string]$Replacement){
    $pattern='(?m)^'+[regex]::Escape($Anchor)+'\r?$'
    if([regex]::Matches($Text,$pattern).Count -ne 1){throw 'Diagnostic harness anchor missing/ambiguous.'}
    [regex]::Replace($Text,$pattern,[Text.RegularExpressions.MatchEvaluator]{param($match) $Replacement})
}
if($TraceBoundaries){
    $anchor='    $coreName = [string]$packages["invSys.Core.xlam"].Name'
    $install=@'
    . (Join-Path $repo 'tests/tooling/ProductionBatchBoundaryTrace.ps1')
    Install-ProductionBatchTrace -Excel $excel -Packages $packages -PackageRoot $deployPath
'@
    $source=Replace-Once $source $anchor ($install+"`r`n"+$anchor)
    $anchor='    $currentStep = "invoke packaged callbacks for state $WorkbookState"'
    $arm=@'
    $tracePath=Join-Path $outputPath 'stages.txt'
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName -MacroName 'TestProductionBatchTrace.Arm' -Arguments @($tracePath))
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName -MacroName 'TestProductionBatchTrace.Mark' -Arguments @('UNREGISTERED_CALIBRATION_VALUE'))
    if(-not (Test-Path -LiteralPath $tracePath) -or [IO.File]::ReadAllText($tracePath).Trim() -cne 'Arm'){throw 'Trace transport/redaction calibration failed; not product RED.'}
'@
    $source=Replace-Once $source $anchor ($arm+"`r`n"+$anchor)
}
$anchor='                $observedText += " || PRODUCTION_BATCH_SCALE=" + $workflowControlReport'
$cut=@'
                $localPaths=@($newWorkbookPaths|Where-Object {[IO.Path]::GetFullPath($_).StartsWith([IO.Path]::GetFullPath($operatorRoot),[StringComparison]::OrdinalIgnoreCase)})
                [pscustomobject]@{NewWorkbookCount=$newWorkbooks.Count;NewPathCount=$newWorkbookPaths.Count;StationLocalCount=$localPaths.Count;SecondNewWorkbookCount=$secondNewWorkbooks.Count}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $outputPath 'workbook-counts.json')
                $focused=[ordered]@{
                    Setup=(@($setupEvidence|Where-Object {$_ -notmatch '=True$'}).Count -eq 0)
                    FirstLauncher=([string]::IsNullOrWhiteSpace($capture.Error))
                    FirstSavedWorkbook=($localPaths.Count -eq 1 -and (Test-Path -LiteralPath $localPaths[0]))
                    SecondLauncher=([string]::IsNullOrWhiteSpace($secondCapture.Error))
                    SecondReusesWorkbook=($secondNewWorkbooks.Count -eq 0)
                    OriginalScaleContract=$workflowControlPassed
                    ExactScaleResult=($workflowControlReport -ceq 'OK|Min=.001%|Default=100%|Max=1000%|BoundsRejected=True|ListOnly=True')
                }
                foreach($name in $focused.Keys){Add-Evidence -Rows $evidence -Callback ('BatchBoundary.'+$name) -Expected 'Existing packaged launcher/scale contract holds.' -Passed ([bool]$focused[$name]) -Observed 'Details omitted'}
                if(@($focused.Values|Where-Object {-not $_}).Count){throw 'Focused launcher/scale contract failed.'}
                [IO.File]::WriteAllText((Join-Path $outputPath 'cut-reached.txt'),'AfterBatchScale')
                return
'@
$source=Replace-Once $source $anchor $cut
$anchor='        $observed = ([string]$row.Observed).Replace("|", "/")'
$source=Replace-Once $source $anchor '        $observed = "Details omitted"'
$source=$source.Replace('# Plan 022 Slice 4x Packaged Reusable Production Evidence','# Scoped Production batch-boundary diagnostic; not full reusable acceptance')
$generated=Join-Path $root 'scoped-validator.ps1'
[IO.File]::WriteAllText($generated,$source,[Text.UTF8Encoding]::new($false))
$tokens=$null;$errors=$null
[void][Management.Automation.Language.Parser]::ParseFile($generated,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Scoped diagnostic does not parse.'}
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$settings=Get-InvSysTestSettingsSnapshot
$restored=$false;$code=1;$start=[DateTimeOffset]::UtcNow
try {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $generated -RepoRoot $repo -DeployRoot $DeployRoot -OutputDirectory ($root.Substring($repo.Length+1)) -CallbackFilter Production -WorkbookState ProductionReusable *> (Join-Path $root 'worker.log')
    $code=$LASTEXITCODE
} finally {
    Wait-RecordingCleanup -Creator $null -Worker $null
    $restored=Restore-InvSysTestSettingsSnapshot $settings
    $preserved=$true
    foreach($pin in $pins){
        $name=if($pin.PSObject.Properties['Package']){$pin.Package}else{Split-Path -Leaf $pin.File}
        $preserved=$preserved -and (Get-FileHash -LiteralPath (Join-Path $deploy $name)).Hash -ceq $pin.Hash
    }
    [pscustomobject]@{StartUTC=$start.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');ExitCode=$code;SettingsRestored=$restored;PackagesPreserved=$preserved;ExcelClosed=$true;ValidatorPreserved=((Get-FileHash -LiteralPath $validator).Hash -ceq $validatorHash);TraceBoundaries=[bool]$TraceBoundaries;FullProductionAccepted=$false}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $root 'closure.json')
    if(-not $restored -or -not $preserved){throw 'Diagnostic preservation failed.'}
}
Start-Sleep -Seconds 12
$queryErrors=@();$audit=[DateTimeOffset]::UtcNow
$events=@(Get-WinEvent -FilterHashtable @{LogName='Application';Id=1000,1001,1002;StartTime=$start.LocalDateTime;EndTime=$audit.LocalDateTime} -ErrorAction SilentlyContinue -ErrorVariable queryErrors|Where-Object {$_.Message -match '(?i)excel\.exe'})
if(@($queryErrors|Where-Object FullyQualifiedErrorId -NotLike 'NoMatchingEventsFound*').Count){throw 'Application audit unavailable.'}
$checks=@();$report=Join-Path $root 'production-reusable-production.md'
if(Test-Path -LiteralPath $report){$checks=@([regex]::Matches([IO.File]::ReadAllText($report),'(?m)^\| ([^|]+) \| (PASS|RED) \|')|ForEach-Object {[pscustomobject]@{Check=$_.Groups[1].Value;Passed=$_.Groups[2].Value -ceq 'PASS'}})}
$cutReached=Test-Path -LiteralPath (Join-Path $root 'cut-reached.txt')
$cleanup=$null;$cleanupPath=Join-Path $root 'final-cleanup-observation.json'
if(Test-Path -LiteralPath $cleanupPath){$cleanup=Get-Content -LiteralPath $cleanupPath -Raw|ConvertFrom-Json}
$traceValid=$true;$stages=@()
if($TraceBoundaries){
    . (Join-Path $PSScriptRoot 'ProductionBatchBoundaryTrace.ps1')
    $allow=@('Arm')+@(Get-ProductionBatchTracePlan|ForEach-Object Stage)
    $stagePath=Join-Path $root 'stages.txt'
    if(Test-Path -LiteralPath $stagePath){$stages=@(Get-Content -LiteralPath $stagePath)}
    $traceValid=$stages.Count -gt 1 -and @($stages|Where-Object {$_ -cnotin $allow}).Count -eq 0
}
$failed=@($checks|Where-Object {-not $_.Passed}).Count
$result=[pscustomobject]@{TraceBoundaries=[bool]$TraceBoundaries;PassedChecks=$checks.Count-$failed;FailedChecks=$failed;Checks=$checks;CutReached=$cutReached;TraceAllowlistValid=$traceValid;TraceEntries=$stages.Count;ExcelApplicationEvents=$events.Count;FinalCleanup=$cleanup;DiagnosticPassed=($code -eq 0 -and $cutReached -and $checks.Count -eq 7 -and $failed -eq 0 -and $traceValid -and $events.Count -eq 0);FullProductionAccepted=$false}
$result|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $root 'result.json')
$result|Select-Object TraceBoundaries,PassedChecks,FailedChecks,CutReached,TraceAllowlistValid,TraceEntries,ExcelApplicationEvents,FinalCleanup,DiagnosticPassed|ConvertTo-Json -Depth 4
if(-not $result.DiagnosticPassed){exit 1}
