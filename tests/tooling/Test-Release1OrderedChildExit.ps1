[CmdletBinding()]
param(
    [string]$RepoRoot='',
    [ValidateSet('RED','GREEN')][string]$Phase='GREEN'
)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
if(-not $RepoRoot){$RepoRoot=Split-Path -Parent (Split-Path -Parent $PSScriptRoot)}
$sourceRoot=(Resolve-Path -LiteralPath $RepoRoot).Path
$path=Join-Path $sourceRoot 'tools/validate_release1_full_chain.ps1'
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Full-chain source does not parse.'}
foreach($name in @('Add-Result','Test-LiveResultCheck')){
    $definition=@($ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name},$true))
    if($definition.Count -ne 1){throw 'Full-chain result helper is ambiguous.'}
    . ([scriptblock]::Create($definition[0].Extent.Text))
}
$main=@($ast.EndBlock.Statements|Where-Object {$_ -is [Management.Automation.Language.TryStatementAst]})
if($main.Count -ne 1){throw 'Full-chain orchestration block is ambiguous.'}
$statements=$main[0].Body.Statements
$begin=-1;$end=-1
for($i=0;$i -lt $statements.Count;$i++){
    if($statements[$i].Extent.Text -match '^\$orderedValidator\s*='){$begin=$i}
    if($statements[$i].Extent.Text -match '^Invoke-RestartReconciliation\b'){$end=$i}
}
if($begin -lt 0 -or $end -le $begin){throw 'Ordered child decision boundary is unavailable.'}
# Execute the actual orchestration statements and report parser without editing
# their decision logic. Only external process creation is replaced by its result.
$decision=[scriptblock]::Create((@($statements[$begin..($end-1)]|ForEach-Object {$_.Extent.Text}) -join "`n"))
function New-OrderedLiveValidator { 'unused-isolated-child.ps1' }
function Invoke-RepositoryScript([string]$Path,[string[]]$Arguments){
    [pscustomobject]@{ExitCode=$script:fixtureChildExit;Output=@();Text=''}
}
$fixture=Join-Path ([IO.Path]::GetTempPath()) ('invsys-child-exit-'+[guid]::NewGuid().ToString('N'))
$repo=$fixture;$DeployRoot='unused-isolated-package-root'
$report=Join-Path $fixture 'tests/unit/phase6_live_role_workflow_results.md'
$checks=[Collections.Generic.List[object]]::new()
function Invoke-DecisionCase([string]$Name,[int]$ExitCode,[bool]$ExpectedAccepted){
    $script:fixtureChildExit=$ExitCode
    $results=[Collections.Generic.List[object]]::new()
    . $decision
    $accepted=$results.Count -gt 0 -and @($results|Where-Object {-not $_.Passed}).Count -eq 0
    $checks.Add([pscustomobject]@{Check=$Name;Passed=($accepted -eq $ExpectedAccepted);ObservedAssertions=$results.Count})
}
try{
    New-Item -ItemType Directory -Path (Split-Path $report -Parent) -Force|Out-Null
    $names=@(
        'Receiving.Form.Stage','Receiving.FormAction.ConfirmWrites.CapturedWorkbook',
        'Receiving.ConfirmWrites.Queue','Receiving.ConfirmWrites.Process','Receiving.ConfirmWrites.InventoryLog',
        'InventoryDomain.ProjectionRecovery.RunBatch','InventoryDomain.ProjectionRecovery.NonAuthoritative',
        'Production.FormActions.TwoConsecutiveBatches.CapturedWorkbook','Production.Form.CheckIn',
        'Production.Form.CompleteRun.Process','Production.Form.CompleteRun.InventoryLog',
        'Boxing.FormAction.Release1','Boxing.BomVersionAndIdentity','Shipping.Form.Stage',
        'Shipping.FormAction.ShipmentsSent.CapturedWorkbook','Shipping.BtnShipmentsSent.Queue',
        'Shipping.BtnShipmentsSent.Process','Shipping.BtnShipmentsSent.InventoryLog',
        'InventoryDomain.ProjectionRecovery.Balances'
    )
    $lines=@($names|ForEach-Object {'| '+$_+' | PASS | fixture |'})
    $lines+='Payload={"SKU":"SKU-BOX","Location":"BIN-B","BomVersionLabel":"v1"}'
    Set-Content -LiteralPath $report -Value $lines
    Invoke-DecisionCase 'OrderedChild.SuccessWithPassingReportAccepted' 0 $true
    if(-not $checks[0].Passed){throw 'Passing-report calibration failed; not behavioral RED.'}
    Invoke-DecisionCase 'OrderedChild.NonzeroExitRejectsPassingReport' 1 $false
    Invoke-DecisionCase 'OrderedChild.AbnormalExitRejectsPassingReport' 99 $false
    Remove-Item -LiteralPath $report
    Invoke-DecisionCase 'OrderedChild.MissingReportRejected' 0 $false
}finally{
    $resolved=[IO.Path]::GetFullPath($fixture)
    $temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if(-not $resolved.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolved -Leaf) -notlike 'invsys-child-exit-*'){throw 'Fixture cleanup escaped its temporary root.'}
    if(Test-Path -LiteralPath $resolved){Remove-Item -LiteralPath $resolved -Recurse -Force}
}
$output=Join-Path $sourceRoot 'reports/runtime/slice4be-shipping-activity'
New-Item -ItemType Directory -Path $output -Force|Out-Null
$checks|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $output ('ordered-child-exit-'+$Phase.ToLowerInvariant()+'.json'))
$checks|ForEach-Object {Write-Output ($_.Check+': '+$(if($_.Passed){'PASS'}else{'FAIL'}))}
if(@($checks|Where-Object {-not $_.Passed}).Count){exit 1}
