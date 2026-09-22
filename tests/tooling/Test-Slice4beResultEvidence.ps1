# Offline result-retention check. Never creates or attaches to Excel.
[CmdletBinding()]
param([string]$RepoRoot='.')
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
$root=Join-Path $repo ('reports/runtime/result-evidence/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$tokens=$null;$errors=$null
$tree=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'tests/tooling/Test-Slice4beConfigCommands.ps1'),[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Shared harness does not parse.'}
foreach($name in @('Check','Complete-ResultEvidence')){
    $definition=$tree.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq $name},$true)
    if($null -eq $definition){throw 'Required shared result helper missing.'}
    Invoke-Expression $definition.Extent.Text
}
$checks=[Collections.Generic.List[object]]::new()
function Assert([string]$Name,[bool]$Passed){$checks.Add([pscustomobject]@{Check=$Name;Passed=$Passed});Write-Output ($Name+': '+$Passed)}
foreach($kind in @('Success','Failure')){
    $results=[Collections.Generic.List[object]]::new()
    $results.Add([pscustomobject]@{Check='Existing.Pass';Passed=$true})
    $reportPath=Join-Path $root ($kind+'.json')
    $script:beforeCleanup=$false
    $canary='unlogged cleanup detail '+[guid]::NewGuid().ToString('N')
    $output=Complete-ResultEvidence $reportPath {
        $before=Get-Content -LiteralPath $reportPath -Raw|ConvertFrom-Json
        $script:beforeCleanup=$before.Check -ceq 'Existing.Pass' -and $before.Passed -is [bool] -and $before.Passed
        if($kind -ceq 'Failure'){throw $canary}
    }|Out-String
    $text=Get-Content -LiteralPath $reportPath -Raw
    $saved=$text|ConvertFrom-Json
    $saved=@($saved)
    Assert ($kind+'.TypedSnapshotPrecedesCleanup') $script:beforeCleanup
    Assert ($kind+'.PriorResultPreserved') ($saved[0].Check -ceq 'Existing.Pass' -and $saved[0].Passed -is [bool] -and $saved[0].Passed)
    $expected=if($kind -ceq 'Failure'){2}else{1}
    $outcome=$saved.Count -eq $expected
    if($kind -ceq 'Failure'){$outcome=$outcome -and $saved[1].Check -ceq 'Harness.Exception.disposable fixture cleanup' -and $saved[1].Passed -is [bool] -and -not $saved[1].Passed}
    Assert ($kind+'.CleanupOutcomeRetained') $outcome
    Assert ($kind+'.ExceptionDetailExcluded') (-not $text.Contains($canary) -and -not $output.Contains($canary))
}
$checks|ConvertTo-Json -Depth 3|Set-Content (Join-Path $root 'results.json')
if(@($checks|Where-Object {-not $_.Passed}).Count){exit 1}
Write-Output ('PASS: '+$checks.Count+' offline result-retention checks.')
