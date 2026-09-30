# Offline generation calibration. No Excel, credentials, or business values are emitted.
[CmdletBinding()]
param([string]$RepoRoot='.',[string]$DeployRoot='deploy/validation-process-worksheet-picker',[Parameter(Mandatory=$true)][string]$PackagePinsPath)
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
$output=@(& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-ProductionRestartDiagnostic.ps1') -RepoRoot $repo -DeployRoot $DeployRoot -PackagePinsPath $PackagePinsPath -GenerateOnly)
if($LASTEXITCODE -ne 0){throw 'Diagnostic generation failed; no product RED.'}
$roots=@($output|Where-Object{$_ -like 'Restart diagnostic: *'})
if($roots.Count -ne 1){throw 'Diagnostic root unavailable.'}
$root=$roots[0].Substring('Restart diagnostic: '.Length)
$source=[IO.File]::ReadAllText((Join-Path $root 'restart-validator.ps1'))
$generation=Get-Content (Join-Path $root 'generation.json') -Raw|ConvertFrom-Json
$checks=[Collections.Generic.List[object]]::new()
$restart=$source.IndexOf('$currentStep = "restart reusable Production in a clean Excel process"',[StringComparison]::Ordinal)
$attach=$source.IndexOf('$restartObserver=Start-Process',[StringComparison]::Ordinal)
$trace=$source.IndexOf('Install-ProductionRestartTrace -Excel',[StringComparison]::Ordinal)
$callback=$source.IndexOf('$restartCapture = Invoke-PackagedCallback',[StringComparison]::Ordinal)
$checks.Add([pscustomobject]@{Check='Observer.AttachesOnlyToFreshRestart';Passed=($restart -ge 0 -and $attach -gt $restart -and $attach -lt $trace -and [regex]::Matches($source,[regex]::Escape('$restartObserver=Start-Process')).Count -eq 1)})
$checks.Add([pscustomobject]@{Check='Trace.AfterPackagesBeforeCallback';Passed=($trace -gt $attach -and $callback -gt $trace)})
$checks.Add([pscustomobject]@{Check='Generation.OriginalWorkflowPreserved';Passed=($generation.OriginalStatementsPreserved -and $generation.ParseErrors -eq 0 -and -not $generation.GenerationOpenedExcel)})
$checks.Add([pscustomobject]@{Check='Credentials.InheritedOnlyInMemory';Passed=($generation.CredentialInheritedOnlyInMemory -and $source.Contains('$testPin = $RestartDiagnosticPin'))})
. (Join-Path $PSScriptRoot 'ProductionRestartTrace.ps1')
$sources=@{mProduction=[IO.File]::ReadAllText((Join-Path $repo 'src/Production/Modules/mProduction.bas'));frmProduction=[IO.File]::ReadAllText((Join-Path $repo 'src/Production/Forms/frmProduction.frm'))}
$edits=@(Get-ProductionRestartTraceEdits $sources)
foreach($group in $edits|Group-Object Module){
    $lines=[Collections.Generic.List[string]]::new()
    foreach($line in ($sources[$group.Name] -split '\r?\n')){$lines.Add($line)}
    foreach($edit in $group.Group|Sort-Object Line -Descending){$lines.Insert($edit.Line-1,$edit.Code)}
    $stripped=($lines|Where-Object{$_ -notmatch '^    TestProductionRestartTrace[.]Mark "[A-Za-z.]+"$'}) -join "`n"
    $checks.Add([pscustomobject]@{Check=('Trace.PreservesEveryOriginalLine.'+$group.Name);Passed=($stripped -ceq $sources[$group.Name].Replace("`r`n","`n"))})
    # The original error handler must read Err without an intervening logger call.
    $instrumented=$lines -join "`n"
    $errorPattern='(?ms)^Failed:\n    TestProductionRestartTrace[.]Mark'
    $checks.Add([pscustomobject]@{Check=('Trace.DoesNotInterceptErr.'+$group.Name);Passed=(-not [regex]::IsMatch($instrumented,$errorPattern))})
}
$logger=Get-ProductionRestartTraceLogger
$stages=@('Arm')+@(Get-ProductionRestartTracePlan|ForEach-Object Stage)
$actual=@([regex]::Matches($logger,'(?m)^        Case "([A-Za-z.]+)"')|ForEach-Object{$_.Groups[1].Value})
$checks.Add([pscustomobject]@{Check='Trace.FixedAllowlistAndUnknownRejection';Passed=(($actual -join '|') -ceq ($stages -join '|') -and $logger.Contains('Case Else: Exit Sub'))})
$checks.Add([pscustomobject]@{Check='Trace.NoBusinessOrCredentialReads';Passed=($logger -notmatch '(?i)Workbook|Worksheet|Range\(|Value2|Password|Credential|PinHash')})
$checks|ConvertTo-Json|Set-Content (Join-Path $root 'calibration.json')
$failed=@($checks|Where-Object{-not $_.Passed}).Count
[pscustomobject]@{Root=$root;Passed=$checks.Count-$failed;Failed=$failed;TraceMarkers=$edits.Count;ExcelOpened=$false}|ConvertTo-Json
if($failed){exit 1}
