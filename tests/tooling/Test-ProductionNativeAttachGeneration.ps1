# Offline calibration of native-observer placement; never opens Excel.
[CmdletBinding()]
param([string]$RepoRoot='.',[Parameter(Mandatory=$true)][string]$DeployRoot,
      [Parameter(Mandatory=$true)][string]$PackagePinsPath)
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
$helper=Join-Path $PSScriptRoot 'Test-ProductionBatchBoundary.ps1'
$root=Join-Path $repo ('reports/runtime/native-attach-calibration/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$checks=[Collections.Generic.List[object]]::new()
$supportsLate=(Get-Command $helper).Parameters.ContainsKey('NativeBeforeRun')
$supportsFull=(Get-Command $helper).Parameters.ContainsKey('FullProductionFlow')
$helperSource=[IO.File]::ReadAllText($helper)
foreach($mode in @('Early','BeforeRun','Full')){
    $extra=@();if($mode -eq 'BeforeRun' -and $supportsLate){$extra+='-NativeBeforeRun'}
    if($mode -eq 'Full' -and $supportsFull){$extra+='-FullProductionFlow'}
    $output=@(& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $helper -RepoRoot $repo -DeployRoot $DeployRoot -PackagePinsPath $PackagePinsPath -StandardRunFlow -NativeExceptions -NativeFaultsOnly -GenerateOnly @extra)
    if($LASTEXITCODE -ne 0){throw 'Generation harness failed; no behavioral RED.'}
    $generatedRoot=(@($output|Where-Object {$_ -like 'Diagnostic: *'})[0]).Substring(12)
    $generation=Get-Content (Join-Path $generatedRoot 'generation.json') -Raw|ConvertFrom-Json
    $source=[IO.File]::ReadAllText((Join-Path $generatedRoot 'standard-run-validator.ps1'))
    $attach=$source.IndexOf('$nativeObserver=Start-Process',[StringComparison]::Ordinal)
    $setup=$source.IndexOf('$currentStep = "create isolated config and auth"',[StringComparison]::Ordinal)
    $run=$source.IndexOf('[IO.File]::WriteAllText($progressPath, "invoke Production reusable run actions")',[StringComparison]::Ordinal)
    $correctPlacement=if($mode -eq 'BeforeRun'){$attach -gt $setup -and $attach -lt $run}else{$attach -ge 0 -and $attach -lt $setup}
    $checks.Add([pscustomobject]@{Check="$mode.ObserverPlacement";Passed=$correctPlacement})
    $checks.Add([pscustomobject]@{Check="$mode.OriginalStatementsPreserved";Passed=[bool]$generation.OriginalStatementsPreserved})
    $checks.Add([pscustomobject]@{Check="$mode.ParsesWithoutExcel";Passed=($generation.ParseErrors -eq 0 -and -not $generation.GenerationOpenedExcel)})
    $checks.Add([pscustomobject]@{Check="$mode.ExactlyOneObserver";Passed=([regex]::Matches($source,[regex]::Escape('$nativeObserver=Start-Process')).Count -eq 1)})
    $checks.Add([pscustomobject]@{Check="$mode.NoVBEPreparation";Passed=(-not $generation.TraceBoundaries -and -not $generation.CompileOnly)})
    $checks.Add([pscustomobject]@{Check="$mode.PlacementDeclared";Passed=($null -ne $generation.PSObject.Properties['NativeBeforeRun'] -and [bool]$generation.NativeBeforeRun -eq ($mode -eq 'BeforeRun'))})
    if($mode -eq 'Full'){
        # Execute the helper's actual argument selection, not a duplicate test implementation.
        $selection=[regex]::Matches($helperSource,'(?m)^\s*\$flowArguments=@\(\);if\([^\r\n]+')
        if($selection.Count -ne 1){throw 'Flow selection anchor ambiguous; no behavioral RED.'}
        $selectFlow=[scriptblock]::Create('param($StandardRunFlow,$FullProductionFlow)'+"`n"+$selection[0].Value+"`n"+'return ,$flowArguments')
        $fullArguments=& $selectFlow $true $true
        $shortArguments=& $selectFlow $true $false
        $checks.Add([pscustomobject]@{Check='Full.OriginalFullFlowSelected';Passed=($supportsFull -and $fullArguments.Count -eq 0)})
        $checks.Add([pscustomobject]@{Check='Full.RunOnlyControlPreserved';Passed=($shortArguments.Count -eq 1 -and $shortArguments[0] -ceq '-ProductionRunOnly')})
        $checks.Add([pscustomobject]@{Check='Full.BothAggregatesRequired';Passed=($null -ne $generation.PSObject.Properties['ExpectedChecks'] -and $generation.ExpectedChecks -eq 2)})
    }
}
$checks|ConvertTo-Json|Set-Content (Join-Path $root 'checks.json')
$failed=@($checks|Where-Object {-not $_.Passed}).Count
[pscustomobject]@{Root=$root;Passed=$checks.Count-$failed;Failed=$failed;ExcelOpened=$false}|ConvertTo-Json
if($failed){exit 1}
