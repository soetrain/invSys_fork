# Test the actual chain prefix, without starting the next Excel stage.
[CmdletBinding()]
param([string]$RepoRoot='.',[string]$DeployRoot='deploy/validation-production-complete-entry-02',
      [string]$PackagePinsPath='reports/runtime/production-run-local-controller/44800f286a65445c9019d51478a09419/package-pins.json',
      [ValidateSet('RED','GREEN')][string]$Phase='RED',
      [ValidateSet('Admin','Source')][string]$Boundary='Admin')
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
$deployPath=(Resolve-Path -LiteralPath (Join-Path $repo $DeployRoot)).Path
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before the stage-exit test.'}
$id=[guid]::NewGuid().ToString('N')
$root=Join-Path $repo ('reports/runtime/chain-stage-exit/'+$id)
$tempRoot=Join-Path ([IO.Path]::GetTempPath()) ('invsys-chain-stage-exit-'+$id)
if(Test-Path -LiteralPath $tempRoot){throw 'Preserve existing fixture.'}
New-Item -ItemType Directory -Path $root|Out-Null
Write-Output ('Report: '+$root)
$chain=Join-Path $repo 'tools/validate_release1_full_chain.ps1'
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($chain,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Chain parse failure is not meaningful RED.'}
foreach($name in 'Release-ComObject','Run-WorkbookMacro','Invoke-AdminEntryGate','Invoke-RepositoryScript'){
    $definition=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -ceq $name},$false)
    if($null -eq $definition){throw 'Actual setup helper unavailable; not meaningful RED.'}
    . ([scriptblock]::Create($definition.Extent.Text))
}
$outer=@($ast.EndBlock.Statements|Where-Object{$_ -is [Management.Automation.Language.TryStatementAst] -and $_.Body.Extent.Text.Contains('Invoke-AdminEntryGate')})
if($outer.Count -ne 1){throw 'Actual chain body unavailable; not meaningful RED.'}
$prefix=[Collections.Generic.List[string]]::new();$boundaryFound=$false
$nextStage=if($Boundary -ceq 'Admin'){'$createWarehouse'}else{'$orderedValidator'}
foreach($statement in $outer[0].Body.Statements){
    if($statement -is [Management.Automation.Language.AssignmentStatementAst] -and $statement.Left.Extent.Text -ceq $nextStage){$boundaryFound=$true;break}
    $prefix.Add($statement.Extent.Text)
}
if(-not $boundaryFound -or @($prefix|Where-Object{$_ -ceq 'Invoke-AdminEntryGate'}).Count -ne 1){throw 'Source boundary unavailable; not meaningful RED.'}
$checks=[Collections.Generic.List[object]]::new()
function Add-Result([string]$Check,[bool]$Passed,[string]$Detail){
    $checks.Add([pscustomobject]@{Check=$Check;Passed=$Passed})
    Write-Output ($Check+': '+$(if($Passed){'PASS'}else{'FAIL'}))
}
$pins=Get-Content (Join-Path $repo $PackagePinsPath) -Raw|ConvertFrom-Json
if($pins.Count -ne 5){throw 'Five pinned packages required.'}
foreach($pin in $pins){if((Get-FileHash $pin.File).Hash -cne $pin.Hash -or [IO.Path]::GetDirectoryName($pin.File) -ine $deployPath){throw 'Candidate differs.'}}
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$settings=Get-InvSysTestSettingsSnapshot
$sourceReport=Join-Path $repo 'tests/integration/create-warehouse-results.md'
$reportBytes=if($Boundary -ceq 'Source'){[IO.File]::ReadAllBytes($sourceReport)}else{$null}
$reportRestored=$true
$start=[DateTimeOffset]::UtcNow;$restored=$false;$remaining=@()
try{
    . ([scriptblock]::Create(($prefix -join "`r`n")))
    $expectedSetup=if($Boundary -ceq 'Admin'){3}else{4}
    if($checks.Count -ne $expectedSetup -or @($checks|Where-Object{-not $_.Passed}).Count){throw 'Setup fixture did not pass; not meaningful RED.'}
    $remaining=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object{[pscustomobject]@{Id=$_.Id;StartUTC=$_.StartTime.ToUniversalTime().ToString('o');Responding=$_.Responding}})
    [pscustomobject]@{UTC=[DateTimeOffset]::UtcNow.ToString('o');Boundary=$Boundary;Excel=$remaining;NextSourceStageNotStarted=($Boundary -ceq 'Admin');NextStageNotStarted=$true}|ConvertTo-Json -Depth 4|Set-Content (Join-Path $root 'boundary.json')
    $check=if($Boundary -ceq 'Admin'){'AdminEntry.ExcelExitedBeforeSource'}else{'Source.ExcelExitedBeforeOrderedLive'}
    Add-Result $check ($remaining.Count -eq 0) ''
}catch{
    Add-Result 'Harness.Exception' $false ''
    [pscustomobject]@{Type=$_.Exception.GetType().FullName;HResult=$_.Exception.HResult}|ConvertTo-Json|Set-Content (Join-Path $root 'failure-facts.json')
}finally{
    Wait-RecordingCleanup -Creator $null -Worker $null
    $restored=Restore-InvSysTestSettingsSnapshot $settings
    if($null -ne $reportBytes){
        [IO.File]::WriteAllBytes($sourceReport,$reportBytes)
        $reportRestored=[Convert]::ToBase64String([IO.File]::ReadAllBytes($sourceReport)) -ceq [Convert]::ToBase64String($reportBytes)
    }
    [pscustomobject]@{UTC=[DateTimeOffset]::UtcNow.ToString('o');SettingsRestored=$restored;ExcelClosed=@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -eq 0}|ConvertTo-Json|Set-Content (Join-Path $root 'closure.json')
}
Start-Sleep -Seconds 12
$audit=[DateTimeOffset]::UtcNow;$queryErrors=@()
$events=@(Get-WinEvent -FilterHashtable @{LogName='Application';Id=1000,1001,1002;StartTime=$start.LocalDateTime;EndTime=$audit.LocalDateTime} -ErrorAction SilentlyContinue -ErrorVariable queryErrors|Where-Object{$_.Message -match '(?i)EXCEL.EXE'})
if(@($queryErrors|Where-Object FullyQualifiedErrorId -notlike 'NoMatchingEventsFound*').Count){throw 'Native audit unavailable.'}
$same=$true;foreach($pin in $pins){$same=$same -and (Get-FileHash $pin.File).Hash -ceq $pin.Hash}
$failed=@($checks|Where-Object{-not $_.Passed}).Count
[pscustomobject]@{Phase=$Phase;Boundary=$Boundary;StartUTC=$start.ToString('o');AuditUTC=$audit.ToString('o');Checks=$checks.ToArray();Pass=$checks.Count-$failed;Fail=$failed;OriginalChainPrefix=$true;ChainHash=(Get-FileHash $chain).Hash;SourceStageNotStarted=($Boundary -ceq 'Admin');NextStageNotStarted=$true;RemainingExcelAtBoundary=$remaining.Count;SettingsRestored=$restored;SourceReportRestored=$reportRestored;PackagesPreserved=$same;ExcelClosed=@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -eq 0;DelayedNativeEvents=$events.Count;ProductRED=$false;FullChainAccepted=$false}|ConvertTo-Json -Depth 5|Set-Content (Join-Path $root ($Phase.ToLowerInvariant()+'.json'))
if(-not $restored -or -not $reportRestored -or -not $same -or $events.Count -or $failed){exit 1}
exit 0
