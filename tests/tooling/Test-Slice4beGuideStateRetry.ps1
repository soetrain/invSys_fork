# Exercise the real PowerShell dispatcher with controlled COM transport faults.
# This calibrates the harness only; packaged handlers remain the product proof.
[CmdletBinding()]
param([string]$RepoRoot='.',[ValidateSet('RED','GREEN')][string]$Phase='GREEN')
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
$root=Join-Path $repo ('reports/runtime/guide-state-retry/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$tokens=$null;$parseErrors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Test-Slice4beConfigCommands.ps1'),[ref]$tokens,[ref]$parseErrors)
if($parseErrors.Count){throw 'Actual dispatcher does not parse.'}
$definition=$ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq 'Run'},$false)
if($null -eq $definition){throw 'Actual dispatcher missing.'}
. ([scriptblock]::Create($definition.Extent.Text))
$WaitForExcelReadyForTest=$false;$RetryActionPathViewCountForTest=$true
$CheckGuidePresentation=$false;$initialExcelProcessIds=@()
$rows=[Collections.Generic.List[object]]::new()
function Check([string]$Name,[bool]$Passed){$rows.Add([pscustomobject]@{Check=$Name;Passed=$Passed});Write-Output ($Name+': '+$(if($Passed){'PASS'}else{'FAIL'}))}
$cases=@(
    @{Name='HowToBusy';Control='txtActionPathHowTo';Action='State';Failures=1;Calls=2;Code='800AC472';Success=$true},
    @{Name='DiagnosticRejectedCall';Control='txtActionPathDiagnostic';Action='State';Failures=1;Calls=2;Code='80010001';Success=$true},
    @{Name='PermanentBusyBounded';Control='txtActionPathHowTo';Action='State';Failures=9;Calls=4;Code='800AC472';Success=$false},
    @{Name='ClickNeverReplayed';Control='btnRefreshActionPathView';Action='Click';Failures=1;Calls=1;Code='800AC472';Success=$false},
    @{Name='WriteNeverReplayed';Control='cboActionPathView';Action='Write';Value='Diagnostic';Failures=1;Calls=1;Code='800AC472';Success=$false},
    @{Name='UnknownControlNotRetried';Control='unknown';Action='State';Failures=1;Calls=1;Code='800AC472';Success=$false},
    @{Name='WrongPackageNotRetried';Package='invSys.Admin.xlam';Control='txtActionPathHowTo';Action='State';Failures=1;Calls=1;Code='800AC472';Success=$false},
    @{Name='OtherErrorNotRetried';Control='txtActionPathHowTo';Action='State';Failures=1;Calls=1;Code='80020009';Success=$false},
    @{Name='DisabledNotRetried';Enabled=$false;Control='txtActionPathHowTo';Action='State';Failures=1;Calls=1;Code='800AC472';Success=$false},
    @{Name='ExistingViewCount';Control='';Action='Count';Failures=1;Calls=2;Code='800AC472';Success=$true},
    @{Name='ExistingGuideCount';Form='frmActionPathGuide';Control='';Action='Count';Failures=1;Calls=2;Code='800AC472';Success=$true},
    @{Name='ExistingPickerValues';Form='frmGuideActionPicker';Control='lstGuideActions';Action='Values';Failures=1;Calls=2;Code='800AC472';Success=$true}
)
foreach($case in $cases){
    $reportRoot=Join-Path $root $case.Name
    New-Item -ItemType Directory -Path $reportRoot|Out-Null
    $RetryGuideObservationForTest=-not ($case.ContainsKey('Enabled') -and -not $case.Enabled)
    $excel=[pscustomobject]@{Calls=0;Failures=$case.Failures;FailureCode=[Convert]::ToInt32($case.Code,16)}
    $excel|Add-Member -MemberType ScriptMethod -Name Run -Value {
        param($Macro,$Form,$Control,$Action,$Value)
        $this.Calls++
        if($this.Calls -le $this.Failures){throw [Runtime.InteropServices.COMException]::new('Controlled transport fault.',$this.FailureCode)}
        return 'OBSERVED'
    }
    $package=if($case.ContainsKey('Package')){$case.Package}else{'invSys.Operations.xlam'}
    $form=if($case.ContainsKey('Form')){$case.Form}else{'frmActionPathView'}
    $value=if($case.ContainsKey('Value')){$case.Value}else{''}
    $succeeded=$false;$result=''
    try {$result=Run $package 'modInventoryViewer.GuideDraftControlForTest' @($form,$case.Control,$case.Action,$value);$succeeded=$true}catch{}
    Check ($case.Name+'.ExactAttempts') ($excel.Calls -eq $case.Calls)
    Check ($case.Name+'.ExpectedTransportOutcome') ($succeeded -eq $case.Success -and (-not $succeeded -or $result -ceq 'OBSERVED'))
    Check ($case.Name+'.FirstFailureRetained') (Test-Path -LiteralPath (Join-Path $reportRoot 'first-call-failure.json'))
}
ConvertTo-Json -InputObject @($rows.ToArray())|Set-Content (Join-Path $root ($Phase.ToLowerInvariant()+'.json'))
Write-Output ('Report: '+$root)
if(@($rows|Where-Object {-not $_.Passed}).Count){exit 1}
