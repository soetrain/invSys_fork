[CmdletBinding()]
param([string]$RepoRoot='', [ValidateSet('RED','GREEN')][string]$Phase='GREEN')
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
if(-not $RepoRoot){$RepoRoot=Split-Path -Parent (Split-Path -Parent $PSScriptRoot)}
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'tools/validate_release1_full_chain.ps1'),[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Full-chain source does not parse; not behavioral RED.'}
$checks=[Collections.Generic.List[object]]::new()
function Check([string]$Name,[bool]$Passed){$checks.Add([pscustomobject]@{Check=$Name;Passed=$Passed})}
function Release-ComObject($Object) {}
function Run-WorkbookMacro($Excel,$WorkbookName,$MacroName){
    if($MacroName -cne 'modRuntimeWorkbooks.ClearCoreDataRootOverride'){throw 'Unexpected cleanup command.'}
    $script:overrideBeforeClose=($script:attempts.Count -eq 0)
}
foreach($phaseName in @('Invoke-RestartReconciliation','Invoke-RuntimeFivePackageEvidence')){
    $definition=@($ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $phaseName},$true))
    if($definition.Count -ne 1){throw 'Cleanup phase is ambiguous.'}
    $outer=@($definition[0].Body.EndBlock.Statements|Where-Object {$_ -is [Management.Automation.Language.TryStatementAst] -and $null -ne $_.Finally})
    if($outer.Count -ne 1){throw 'Cleanup boundary is ambiguous.'}
    $cleanup=[scriptblock]::Create((@($outer[0].Finally.Statements|ForEach-Object {$_.Extent.Text}) -join "`n"))
    $script:open=@{};$script:attempts=[Collections.Generic.List[string]]::new()
    $script:rejected=0;$script:saves=0;$script:openAtQuit=-1;$script:overrideBeforeClose=$false
    $localBooks=[Collections.Generic.List[object]]::new();$packageMap=@{}
    # Actual phase acquisition order: five packages, then disposable workbooks.
    $names=@('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam','invSys.Admin.xlam','fixture-projection.xlsb','fixture-operator.xlsm')
    foreach($name in $names){
        $book=[pscustomobject]@{Name=$name;IsAddin=$name.EndsWith('.xlam')}
        $book|Add-Member ScriptMethod Close {
            param([bool]$SaveChanges)
            $script:attempts.Add($this.Name)
            if($SaveChanges){$script:saves++}
            $dependentOpen=($this.Name -eq 'invSys.Core.xlam' -and @($script:open.Keys|Where-Object {$_ -ne $this.Name}).Count -gt 0) -or
                ($this.Name -like 'invSys.*.Domain.xlam' -and $script:open.ContainsKey('invSys.Operations.xlam'))
            if($dependentOpen){$script:rejected++;throw 'Fixture rejects closing a dependency while its consumers remain open.'}
            $script:open.Remove($this.Name)
        }
        $localBooks.Add($book);$script:open.Add($name,$true)
        if($book.IsAddin){$packageMap[$name]=$book}
    }
    $localExcel=[pscustomobject]@{}
    $localExcel|Add-Member ScriptMethod Quit {$script:openAtQuit=$script:open.Count}
    $isolationExcel=$null;$restoreExcel=$null;$suspendedAddinPaths=@()
    . $cleanup
    Check ($phaseName+'.EveryOwnedBookClosesBeforeQuit') ($script:openAtQuit -eq 0)
    Check ($phaseName+'.NoDependencyCloseRejected') ($script:rejected -eq 0)
    $firstPackage=@($script:attempts|Where-Object {$_ -like '*.xlam'})[0]
    $firstPackageIndex=$script:attempts.IndexOf($firstPackage)
    Check ($phaseName+'.OrdinaryBooksCloseBeforePackages') ($script:attempts.IndexOf('fixture-projection.xlsb') -lt $firstPackageIndex -and $script:attempts.IndexOf('fixture-operator.xlsm') -lt $firstPackageIndex)
    Check ($phaseName+'.CoreClosesAfterConsumers') ($script:attempts[$script:attempts.Count-1] -eq 'invSys.Core.xlam')
    Check ($phaseName+'.NoImplicitSave') ($script:saves -eq 0)
    Check ($phaseName+'.EveryOwnedBookAttemptedOnce') ($script:attempts.Count -eq $names.Count -and @($script:attempts|Group-Object|Where-Object Count -NE 1).Count -eq 0)
    if($phaseName -eq 'Invoke-RestartReconciliation'){Check ($phaseName+'.OverrideClearedBeforePackageClosure') $script:overrideBeforeClose}
}
$root=Join-Path $repo 'reports/runtime';[void](New-Item -ItemType Directory -Path $root -Force)
$checks|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $root ('events-chain-close-order-'+$Phase.ToLowerInvariant()+'.json'))
$checks|ForEach-Object {Write-Output ($_.Check+': '+$(if($_.Passed){'PASS'}else{'FAIL'}))}
if(@($checks|Where-Object {-not $_.Passed}).Count){exit 1}
