[CmdletBinding()]
param([string]$RepoRoot='.')
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
. (Join-Path $PSScriptRoot 'ProductionBatchBoundaryTrace.ps1')
$sources=@{
    mProduction=[IO.File]::ReadAllText((Join-Path $repo 'src/Production/Modules/mProduction.bas'))
    frmProduction=[IO.File]::ReadAllText((Join-Path $repo 'src/Production/Forms/frmProduction.frm'))
}
$root=Join-Path $repo ('reports/runtime/production-batch-trace-placement/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$checks=[Collections.Generic.List[object]]::new()
function Check([string]$Name,[bool]$Passed){$checks.Add([pscustomobject]@{Check=$Name;Passed=$Passed})}
$plan=@(Get-ProductionBatchTracePlan);$edits=@(Get-ProductionBatchTraceEdits $sources)
Check 'FixedUniqueStages' ($plan.Count -eq 31 -and @($plan.Stage|Sort-Object -Unique).Count -eq 31)
foreach($edit in $edits){
    $step=@($plan|Where-Object Stage -CEQ $edit.Stage)[0]
    $lines=$sources[$edit.Module] -split '\r?\n'
    Check ($edit.Stage+'.ExactStatement') ($lines[$edit.Line-1].Trim().StartsWith($step.Anchor,[StringComparison]::Ordinal))
    Check ($edit.Stage+'.ConstantOnly') ($edit.Code -cmatch '^    TestProductionBatchTrace.Mark "[A-Za-z0-9.]+"$')
}
foreach($name in $sources.Keys){
    $lines=[Collections.Generic.List[string]]::new()
    foreach($line in ($sources[$name] -split '\r?\n')){$lines.Add($line)}
    foreach($edit in $edits|Where-Object Module -CEQ $name|Sort-Object Line -Descending){$lines.Insert($edit.Line-1,$edit.Code)}
    $restored=@($lines|Where-Object {$_ -cnotmatch '^    TestProductionBatchTrace\.Mark "'}) -join "`n"
    Check ($name+'.OriginalStatementsPreserved') ($restored -ceq (($sources[$name] -split '\r?\n') -join "`n"))
}
$normalized=@{};foreach($name in $sources.Keys){$normalized[$name]=$sources[$name].ToUpperInvariant()}
$upper=@(Get-ProductionBatchTraceEdits $normalized)
Check 'VbeCaseNormalizationAccepted' (($upper|ConvertTo-Json -Compress) -ceq ($edits|ConvertTo-Json -Compress))
$first=$edits[0]
foreach($case in @('Missing','Duplicate')){
    $changed=$sources.Clone();$lines=[Collections.Generic.List[string]]::new()
    foreach($line in ($changed[$first.Module] -split '\r?\n')){$lines.Add($line)}
    if($case -ceq 'Missing'){$lines[$first.Line-1]='    Rem Removed calibration anchor'}else{$lines.Insert($first.Line-1,$lines[$first.Line-1])}
    $changed[$first.Module]=$lines -join "`r`n";$rejected=$false
    try{$null=@(Get-ProductionBatchTraceEdits $changed)}catch{$rejected=$_.Exception.Message -like 'Trace anchor missing/ambiguous:*'}
    Check ($case+'.FailsClosed') $rejected
}
$logger=Get-ProductionBatchTraceLogger
$cases=@([regex]::Matches($logger,'(?m)^        Case "([^"]+)"')|ForEach-Object {$_.Groups[1].Value})
Check 'LoggerFixedAllowlist' ($cases.Count -eq 32 -and @($cases|Where-Object {$_ -cnotin (@('Arm')+$plan.Stage)}).Count -eq 0)
Check 'LoggerRejectsUnknownBeforeWriting' ($logger.IndexOf('Case Else: Exit Sub') -lt $logger.IndexOf('Open mPath For Append'))
$checks|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $root 'checks.json')
$summary=[pscustomobject]@{Passed=@($checks|Where-Object Passed).Count;Failed=@($checks|Where-Object {-not $_.Passed}).Count;Stages=$plan.Count;NoLiveExcel=$true;Root=$root}
$summary|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $root 'summary.json');$summary|ConvertTo-Json
if($summary.Failed){exit 1}
