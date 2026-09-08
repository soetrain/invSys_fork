[CmdletBinding()]
param(
    [string]$RepoRoot='',
    [Parameter(Mandatory=$true)][string]$OutputRoot,
    [Parameter(Mandatory=$true)][string]$EvidenceDirectory
)
$ErrorActionPreference='Stop'
if($RepoRoot -eq ''){$RepoRoot=Split-Path -Parent (Split-Path -Parent $PSScriptRoot)}
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
$output=[IO.Path]::GetFullPath((Join-Path $repo $OutputRoot))
$evidence=[IO.Path]::GetFullPath((Join-Path $repo $EvidenceDirectory))
$deployPrefix=[IO.Path]::GetFullPath((Join-Path $repo 'deploy'))+[IO.Path]::DirectorySeparatorChar
$evidencePrefix=[IO.Path]::GetFullPath((Join-Path $repo 'reports/runtime'))+[IO.Path]::DirectorySeparatorChar
if(-not $output.StartsWith($deployPrefix,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path -Leaf $output) -notlike 'validation-*'){throw 'Use a new isolated validation package directory.'}
if(-not $evidence.StartsWith($evidencePrefix,[StringComparison]::OrdinalIgnoreCase)){throw 'Keep machine evidence under reports/runtime.'}
if((Test-Path -LiteralPath $output) -or (Test-Path -LiteralPath $evidence)){throw 'Preserve existing candidate and evidence directories.'}
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before packaged build-boundary validation.'}
New-Item -ItemType Directory -Path $evidence | Out-Null
$sourcePath=Join-Path $repo 'tools/build-xlam.ps1'
$beforeHash=(Get-FileHash -LiteralPath $sourcePath).Hash
$source=[IO.File]::ReadAllText($sourcePath)
$recordPath=Join-Path $evidence 'external-edit-boundaries.json'
$quotedRecordPath="'"+$recordPath.Replace("'","''")+"'"
$state='$script:externalEditRows=[Collections.Generic.List[object]]::new()'+[Environment]::NewLine+
    '$script:externalEditEvidence='+$quotedRecordPath
$needle='$ErrorActionPreference = "Stop"'
if(([regex]::Matches($source,[regex]::Escape($needle))).Count -ne 1){throw 'Expected builder error-policy boundary.'}
$source=$source.Replace($needle,$needle+[Environment]::NewLine+$state)
$guard=@'
function Assert-ExternalEditWorkbookClosed {
    param([string]$WorkbookPath,[string]$EditStage)
    # Open add-ins are omitted from Workbooks enumeration; use exact named lookup.
    $loaded=$false;$candidateWorkbook=$null
    try {$candidateWorkbook=$excel.Workbooks.Item([IO.Path]::GetFileName($WorkbookPath))}catch{}
    if($null -ne $candidateWorkbook){$loaded=[string]::Equals([string]$candidateWorkbook.FullName,$WorkbookPath,[StringComparison]::OrdinalIgnoreCase)}
    $script:externalEditRows.Add([pscustomobject]@{Package=[IO.Path]::GetFileName($WorkbookPath);Stage=$EditStage;ClosedBeforeExternalEdit=(-not $loaded)})
    $script:externalEditRows | ConvertTo-Json | Set-Content -LiteralPath $script:externalEditEvidence
    if($loaded){throw 'XLAM_EXTERNAL_EDIT_WHILE_OPEN'}
}
function Install-RibbonCustomUi {
    param([string]$WorkbookPath,[hashtable]$RibbonConfig)
    Assert-ExternalEditWorkbookClosed $WorkbookPath 'RibbonX'
    Install-RibbonCustomUiObserved $WorkbookPath $RibbonConfig
}
function Install-RibbonCustomUiObserved {
'@
$needle='function Install-RibbonCustomUi {'
if(([regex]::Matches($source,[regex]::Escape($needle))).Count -ne 1){throw 'Expected actual RibbonX package writer.'}
$source=$source.Replace($needle,$guard)
$guard=@'
function Set-PackageBuildIdentity {
    param([string]$WorkbookPath)
    Assert-ExternalEditWorkbookClosed $WorkbookPath 'BuildIdentity'
    Set-PackageBuildIdentityObserved $WorkbookPath
}
function Set-PackageBuildIdentityObserved {
'@
$needle='function Set-PackageBuildIdentity {'
if(([regex]::Matches($source,[regex]::Escape($needle))).Count -ne 1){throw 'Expected actual build-identity package writer.'}
$source=$source.Replace($needle,$guard)
$probePath=Join-Path $evidence 'observed-build.ps1'
[IO.File]::WriteAllText($probePath,$source,[Text.UTF8Encoding]::new($false))
$tokens=$null;$errors=$null;$null=[Management.Automation.Language.Parser]::ParseFile($probePath,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Observed builder did not parse.'}
$buildError=''
try {& $probePath -RepoRoot $repo -OutputRoot $output -Apply *> (Join-Path $evidence 'build.log')}
catch {if($_.Exception.Message -match 'XLAM_EXTERNAL_EDIT_WHILE_OPEN'){$buildError='XLAM_EXTERNAL_EDIT_WHILE_OPEN'}else{$buildError='UNEXPECTED_BUILD_FAILURE'}}
if((Get-FileHash -LiteralPath $sourcePath).Hash -cne $beforeHash){throw 'Boundary instrumentation modified the checked-in builder.'}
if(-not (Test-Path -LiteralPath $recordPath)){throw 'No package-boundary observations; this is not behavioral RED.'}
$parsed=Get-Content -LiteralPath $recordPath -Raw | ConvertFrom-Json
$rows=@($parsed)
$failures=@($rows | Where-Object {-not $_.ClosedBeforeExternalEdit}).Count
$passes=$rows.Count-$failures
$complete=($buildError -eq '' -and $rows.Count -eq 7 -and $failures -eq 0)
[pscustomobject]@{Pass=$passes;Fail=$failures;Observed=$rows.Count;FullBuildCompleted=$complete;BuildError=$buildError;BuilderSourcePreserved=$true} | ConvertTo-Json | Set-Content (Join-Path $evidence 'summary.json')
Write-Output ('PACKAGED_EXTERNAL_EDIT_BOUNDARIES PASS='+$passes+' FAIL='+$failures+' OBSERVED='+$rows.Count+' COMPLETE='+$complete)
if($buildError -eq 'UNEXPECTED_BUILD_FAILURE'){throw 'Unexpected build failure; inspect ignored build log. Not behavioral RED.'}
if(-not $complete){exit 1}
