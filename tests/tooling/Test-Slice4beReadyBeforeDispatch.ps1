# Offline test of the actual shared Run boundary; never attaches to Excel.
[CmdletBinding()]
param([string]$RepoRoot='.')
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
$root=Join-Path (Join-Path $repo 'reports/runtime') ('ready-before-dispatch/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$tokens=$null;$errors=$null
$tree=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'tests/tooling/Test-Slice4beConfigCommands.ps1'),[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Controller parse failed.'}
foreach($name in @('Wait-ExcelReadyForTest','Run')){
    $definition=$tree.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq $name},$true)
    if($null -ne $definition){Invoke-Expression $definition.Extent.Text}elseif($name -ceq 'Run'){throw 'Actual Run boundary missing.'}
}
$CheckGuidePresentation=$false;$RetryActionPathViewCountForTest=$false;$RetryGuideObservationForTest=$false
$initialExcelProcessIds=@();$initialExcelWindow=0
$checks=[Collections.Generic.List[object]]::new()
function Check([string]$Name,[bool]$Passed){$checks.Add([pscustomobject]@{Check=$Name;Passed=$Passed});Write-Output ($Name+': '+$Passed)}
$canary='unlogged field '+[guid]::NewGuid().ToString('N')
foreach($case in @(
    @{Name='Ready';Plan=@('Ready');Reads=1;Dispatches=1;Throws=$false;Flag=$true},
    @{Name='BusyThenReady';Plan=@('False','Ready');Reads=2;Dispatches=1;Throws=$false;Flag=$true},
    @{Name='UnavailableThenReady';Plan=@('Unavailable','Ready');Reads=2;Dispatches=1;Throws=$false;Flag=$true},
    @{Name='BusyBounded';Plan=@('False');Reads=8;Dispatches=0;Throws=$true;Flag=$true},
    @{Name='UnavailableBounded';Plan=@('Unavailable');Reads=8;Dispatches=0;Throws=$true;Flag=$true},
    @{Name='InvalidTypeStops';Plan=@('Invalid');Reads=1;Dispatches=0;Throws=$true;Flag=$true},
    @{Name='DispatchFailureNotReplayed';Plan=@('Ready');Reads=1;Dispatches=1;Throws=$true;Flag=$true;RunThrows=$true},
    @{Name='FlagOff';Plan=@('False');Reads=0;Dispatches=1;Throws=$false;Flag=$false}
)){
    $reportRoot=Join-Path $root $case.Name;New-Item -ItemType Directory -Path $reportRoot|Out-Null
    $WaitForExcelReadyForTest=$case.Flag
    $excel=[pscustomobject]@{Reads=0;Calls=0;ReadsAtDispatch=-1;Plan=$case.Plan;RunThrows=($case.ContainsKey('RunThrows') -and $case.RunThrows)}
    $excel|Add-Member -MemberType ScriptProperty -Name Ready -Value {
        $index=[Math]::Min($this.Reads,$this.Plan.Count-1);$this.Reads++
        switch($this.Plan[$index]){
            'Ready' {return $true}
            'False' {return $false}
            'Invalid' {return 'not a Boolean'}
            default {return $null}
        }
    }
    $excel|Add-Member -MemberType ScriptMethod -Name Run -Value {
        param($macro,$field,$value)
        $this.Calls++;$this.ReadsAtDispatch=$this.Reads
        if($this.RunThrows){throw [Runtime.InteropServices.COMException]::new('Isolated dispatch fixture',[Convert]::ToInt32('800AC472',16))}
        return 'OK'
    }
    $thrown=$false;$result=$null
    try{$result=Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('BatchSize',$canary)}catch{$thrown=$true}
    Check ($case.Name+'.ExactReadsAndSingleDispatch') ($excel.Reads -eq $case.Reads -and $excel.Calls -eq $case.Dispatches -and ($excel.Calls -eq 0 -or $excel.ReadsAtDispatch -eq $case.Reads))
    Check ($case.Name+'.Outcome') ($thrown -eq $case.Throws -and ($thrown -or $result -ceq 'OK'))
    $tracePath=Join-Path $reportRoot 'readiness-before-dispatch.jsonl';$trace=@();$traceText=''
    if(Test-Path $tracePath){$traceText=Get-Content $tracePath -Raw;$trace=@(Get-Content $tracePath|ForEach-Object {$_|ConvertFrom-Json})}
    Check ($case.Name+'.ReadinessTrace') ($trace.Count -eq $case.Reads -and ($case.Reads -eq 0 -or ($trace[-1].ReadAttempt -eq $case.Reads -and ($case.Dispatches -eq 0 -or $trace[-1].Status -ceq 'Ready'))))
    $allText=@(Get-ChildItem $reportRoot -File|ForEach-Object {Get-Content $_.FullName -Raw}) -join "`n"
    Check ($case.Name+'.ArgumentsExcluded') (-not $allText.Contains($canary) -and -not $allText.Contains('BatchSize'))
}
$checks|ConvertTo-Json -Depth 4|Set-Content (Join-Path $root results.json)
Write-Output ('Report: '+$root)
if(@($checks|Where-Object {-not $_.Passed}).Count){exit 1}
Write-Output ('PASS: '+$checks.Count+' offline readiness checks.')
