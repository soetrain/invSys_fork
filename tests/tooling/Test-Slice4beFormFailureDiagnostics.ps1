# Offline privacy/replay check. Never creates or attaches to Excel.
[CmdletBinding()]
param([string]$RepoRoot='.')
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
$root=Join-Path $repo ('reports/runtime/form-failure-diagnostics/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$tokens=$null;$errors=$null
$tree=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'tests/tooling/Test-Slice4beConfigCommands.ps1'),[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Shared harness does not parse.'}
$definition=$tree.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq 'Run'},$true)
if($null -eq $definition){throw 'Shared Run boundary missing.'}
Invoke-Expression $definition.Extent.Text
$CheckGuidePresentation=$false
$RetryActionPathViewCountForTest=$true
$initialExcelProcessIds=@();$initialExcelWindow=0
$canary='unlogged field content '+[guid]::NewGuid().ToString('N')
$checks=[Collections.Generic.List[object]]::new()
function Check([string]$Name,[bool]$Passed){$checks.Add([pscustomobject]@{Check=$Name;Passed=$Passed});Write-Output ($Name+': '+$Passed)}
foreach($kind in @('Control','Other')){
    $reportRoot=Join-Path $root $kind
    New-Item -ItemType Directory -Path $reportRoot|Out-Null
    $excel=[pscustomobject]@{Calls=0}
    $excel|Add-Member -MemberType ScriptMethod -Name Run -Value {
        param($macro,$first,$second,$third,$fourth)
        $this.Calls++
        throw [Runtime.InteropServices.COMException]::new('Isolated rejected-call fixture',[Convert]::ToInt32('800AC472',16))
    }
    $thrown=$false
    try {
        if($kind -ceq 'Control'){
            [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.GuideDraftControlForTest' @('frmActionPathView','btnCloseActionPathView','Click',$canary))
        } else {[void](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('BatchSize',$canary))}
    } catch {$thrown=$true}
    $text=Get-Content (Join-Path $reportRoot 'first-call-failure.json') -Raw
    $record=$text|ConvertFrom-Json
    Check ($kind+'.FailurePropagatesWithoutReplay') ($thrown -and $excel.Calls -eq 1)
    Check ($kind+'.FieldContentExcluded') (-not $text.Contains($canary) -and -not $text.Contains('unlogged field content'))
    Check ($kind+'.InnerComCodeRetained') ('0x800AC472' -cin $record.ExceptionHResults)
    Check ($kind+'.NoCaptureRequested') (-not $record.OwnedForegroundCapture -and @((Get-ChildItem -LiteralPath $reportRoot -File -Filter '*.png')).Count -eq 0)
    if($kind -ceq 'Control'){
        Check 'Control.FixedIdentityRetained' ($record.FormControl.Form -ceq 'frmActionPathView' -and $record.FormControl.Control -ceq 'btnCloseActionPathView' -and $record.FormControl.Action -ceq 'Click')
    } else {Check 'Other.ArgumentsExcluded' ($null -eq $record.FormControl -and -not $text.Contains('BatchSize'))}
}
foreach($case in @(
    @{Name='CountRecovered';Failures=1;Code='800AC472';Form='frmActionPathView';Control='';Action='Count';Value='';Flag=$true;Calls=2;Throws=$false;RetryLines=2},
    @{Name='CountRecoveredCallRejected';Failures=1;Code='80010001';Form='frmActionPathView';Control='';Action='Count';Value='';Flag=$true;Calls=2;Throws=$false;RetryLines=2},
    @{Name='CountBounded';Failures=9;Code='800AC472';Form='frmActionPathView';Control='';Action='Count';Value='';Flag=$true;Calls=4;Throws=$true;RetryLines=3},
    @{Name='CountUnlistedCode';Failures=9;Code='80004005';Form='frmActionPathView';Control='';Action='Count';Value='';Flag=$true;Calls=1;Throws=$true;RetryLines=0},
    @{Name='CountOtherForm';Failures=9;Code='800AC472';Form='frmActionPathLibrary';Control='';Action='Count';Value='';Flag=$true;Calls=1;Throws=$true;RetryLines=0},
    @{Name='CountWithFieldContent';Failures=9;Code='800AC472';Form='frmActionPathView';Control='';Action='Count';Value=$canary;Flag=$true;Calls=1;Throws=$true;RetryLines=0},
    @{Name='CountFlagOff';Failures=9;Code='800AC472';Form='frmActionPathView';Control='';Action='Count';Value='';Flag=$false;Calls=1;Throws=$true;RetryLines=0}
)){
    $reportRoot=Join-Path $root $case.Name
    New-Item -ItemType Directory -Path $reportRoot|Out-Null
    $RetryActionPathViewCountForTest=$case.Flag
    $excel=[pscustomobject]@{Calls=0;Failures=$case.Failures;Code=$case.Code}
    $excel|Add-Member -MemberType ScriptMethod -Name Run -Value {
        param($macro,$first,$second,$third,$fourth)
        $this.Calls++
        if($this.Calls -le $this.Failures){throw [Runtime.InteropServices.COMException]::new('Isolated rejected-call fixture',[Convert]::ToInt32($this.Code,16))}
        return '1'
    }
    $thrown=$false;$value=$null
    try {$value=Run 'invSys.Operations.xlam' 'modInventoryViewer.GuideDraftControlForTest' @($case.Form,$case.Control,$case.Action,$case.Value)}catch{$thrown=$true}
    Check ($case.Name+'.ExactAttemptBound') ($excel.Calls -eq $case.Calls -and $thrown -eq $case.Throws -and ($thrown -or $value -ceq '1'))
    $firstText=Get-Content (Join-Path $reportRoot 'first-call-failure.json') -Raw
    $first=$firstText|ConvertFrom-Json
    Check ($case.Name+'.FirstFailureRetained') ($first.FormControl.Action -ceq 'Count' -and ('0x'+$case.Code) -cin $first.ExceptionHResults)
    $retryPath=Join-Path $reportRoot 'readonly-count-retries.jsonl'
    $retryText='';$retries=@()
    if(Test-Path -LiteralPath $retryPath){$retryText=Get-Content -LiteralPath $retryPath -Raw;$retries=@(Get-Content -LiteralPath $retryPath|ForEach-Object {$_|ConvertFrom-Json})}
    $trace=$retries.Count -eq $case.RetryLines
    if(-not $thrown){$trace=$trace -and $retries[-1].Recovered -is [bool] -and $retries[-1].Recovered}
    Check ($case.Name+'.RetryTraceMatchesOutcome') $trace
    Check ($case.Name+'.FieldContentExcluded') (-not $firstText.Contains($canary) -and -not $retryText.Contains($canary))
}
$checks|ConvertTo-Json -Depth 4|Set-Content (Join-Path $root 'results.json')
if(@($checks|Where-Object {-not $_.Passed}).Count){exit 1}
Write-Output ('PASS: '+$checks.Count+' offline diagnostics checks.')
