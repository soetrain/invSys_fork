[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$WorkerScript)
$ErrorActionPreference='Stop'; Set-StrictMode -Version 2.0
. (Join-Path $PSScriptRoot 'ExcelShutdownTiming.ps1')
$checks=[Collections.Generic.List[object]]::new()
function Check([string]$Name,[bool]$Passed) {
    $checks.Add([pscustomobject]@{Check=$Name;Passed=$Passed})
    Write-Output (('{0} {1}' -f $(if($Passed){'PASS'}else{'FAIL'}),$Name))
}
$source=[IO.File]::ReadAllText((Resolve-Path -LiteralPath $WorkerScript))
$combined='    [GC]::Collect(); [GC]::WaitForPendingFinalizers()'
$split="    [GC]::Collect()`r`n    Write-InventoryShutdownTrace 'AfterCollectBeforeFinalizers'`r`n    [GC]::WaitForPendingFinalizers()"
$splitPattern="(?m)^    \[GC\]::Collect\(\)\r?\n    Write-InventoryShutdownTrace 'AfterCollectBeforeFinalizers'\r?\n    \[GC\]::WaitForPendingFinalizers\(\)"
$matches=[regex]::Matches($source,$splitPattern)
if($matches.Count -eq 1) {
    $combinedSource=$source.Replace($matches[0].Value,$combined)
    $splitSource=$source
} elseif([regex]::Matches($source,[regex]::Escape($combined)).Count -eq 1) {
    $combinedSource=$source
    $splitSource=$source.Replace($combined,$split)
} else { throw 'Known worker fixture required; no runtime dispatched.' }
foreach($fixture in @(@('Combined',$combinedSource),@('Split',$splitSource))) {
    $ok=$false; $generated=''
    try {
        $generated=ConvertTo-ExcelShutdownTimedWorker $fixture[1] 'C:\synthetic tooling' 'C:\timing.ps1' 'C:\synthetic report'
        $ok=$true
    } catch {}
    Check ($fixture[0]+'.Accepted') $ok
    $preserved=$ok
    foreach($call in @('[GC]::Collect()','[GC]::WaitForPendingFinalizers()','$excel.Quit()',
        '[Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)')) {
        $preserved=$preserved -and ([regex]::Matches($fixture[1],[regex]::Escape($call)).Count -eq
            [regex]::Matches($generated,[regex]::Escape($call)).Count)
    }
    Check ($fixture[0]+'.CleanupCallsPreserved') $preserved
    if($fixture[0] -eq 'Split') {
        Check 'Split.ExistingMarkerPreserved' ($ok -and $generated.Contains("Write-InventoryShutdownTrace 'AfterCollectBeforeFinalizers'"))
    }
}
$rejected=$false
try {
    [void](ConvertTo-ExcelShutdownTimedWorker ($splitSource.Replace("'AfterCollectBeforeFinalizers'","'UnknownIntermediateWork'")) 'x' 'x' 'x')
} catch {$rejected=$true}
Check 'UnknownGapRefuses' $rejected
$rejected=$false
try {
    [void](ConvertTo-ExcelShutdownTimedWorker ($combinedSource.Replace($combined,$combined+"`r`n"+$combined)) 'x' 'x' 'x')
} catch {$rejected=$true}
Check 'DuplicateGcPairRefuses' $rejected

$variantModule=Join-Path $PSScriptRoot 'ExcelShutdownTableVariant.ps1'
$available=Test-Path -LiteralPath $variantModule
Check 'TableVariantAvailable' $available
if($available) {
    . $variantModule
    $baseline=ConvertTo-ExcelShutdownTableVariant -Source $source -Mode Baseline
    Check 'BaselineByteForByte' ([string]::Equals($baseline,$source,[StringComparison]::Ordinal))
    $variant=ConvertTo-ExcelShutdownTableVariant -Source $source -Mode OwnedTableTraversal
    $tokens=$null; $errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseInput($variant,[ref]$tokens,[ref]$errors)
    Check 'VariantParses' ($errors.Count -eq 0)
    $fns=@($ast.FindAll({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and
        $n.Name -in @('Table','Release-ExcelShutdownTableReference')},$true))
    Check 'ExactlyTwoOwnedFunctions' ($fns.Count -eq 2)
    $table=@($fns|Where-Object Name -eq 'Table')[0].Extent.Text
    Check 'NoBorrowedWorkbookRelease' (-not $table.Contains('Release-ExcelShutdownTableReference $Workbook'))
    Check 'ReturnTransfersOwnership' ($table.Contains('$candidate=$null') -and $table.Contains('return $result'))
    Check 'NoComEnumeration' (-not $table.Contains('foreach'))
    $originalAst=[Management.Automation.Language.Parser]::ParseInput($source,[ref]$tokens,[ref]$errors)
    $originalTable=@($originalAst.FindAll({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Table'},$true))[0]
    $variantOnly=$variant.Remove($originalTable.Extent.StartOffset,
        $variant.Length-($source.Length-$originalTable.Extent.Text.Length))
    $originalOnly=$source.Remove($originalTable.Extent.StartOffset,$originalTable.Extent.Text.Length)
    Check 'OnlyTableRegionChanged' ([string]::Equals($variantOnly,$originalOnly,[StringComparison]::Ordinal))
    $rejected=$false
    try { [void](ConvertTo-ExcelShutdownTableVariant ($source.Replace('foreach ($sheet in $Workbook.Worksheets)','foreach ($sheet in $Workbook.OtherSheets)')) OwnedTableTraversal) }
    catch {$rejected=$true}
    Check 'TableAnchorDriftRefuses' $rejected
    $timed=ConvertTo-ExcelShutdownTimedWorker $variant 'C:\synthetic tooling' 'C:\timing.ps1' 'C:\synthetic report'
    $preserved=$true
    foreach($call in @('[GC]::Collect()','[GC]::WaitForPendingFinalizers()','$excel.Quit()',
        '[Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)')) {
        $preserved=$preserved -and ([regex]::Matches($source,[regex]::Escape($call)).Count -eq
            [regex]::Matches($timed,[regex]::Escape($call)).Count)
    }
    Check 'CombinedTransformKeepsFinalCleanup' $preserved

    # Managed stand-ins test lookup semantics only, never native COM lifetime.
    function Collection([object[]]$Items) {
        $c=[pscustomobject]@{Entries=$Items;Count=$Items.Count}
        $c|Add-Member ScriptMethod Item {param([int]$Index) return $this.Entries[$Index-1]}
        return $c
    }
    $first=[pscustomobject]@{Name='TARGET';Note='USER NOTE'}
    $second=[pscustomobject]@{Name='target';Note='SECOND'}
    $other=[pscustomobject]@{Name='other';Note='UNCHANGED'}
    $wb=[pscustomobject]@{Worksheets=(Collection @(
        [pscustomobject]@{ListObjects=(Collection @($other,$first))},
        [pscustomobject]@{ListObjects=(Collection @($second))}
    ))}
    $definitions=($fns|ForEach-Object {$_.Extent.Text}) -join "`n"
    $lookup=[scriptblock]::Create('param($Workbook,$Name)'+"`n"+$definitions+"`n"+'Table $Workbook $Name')
    $value=& $lookup $wb 'TaRgEt'
    Check 'CaseInsensitiveFirstMatch' ([object]::ReferenceEquals($first,$value))
    Check 'CallerDataPreserved' ($other.Note -ceq 'UNCHANGED' -and $first.Note -ceq 'USER NOTE' -and $wb.Worksheets.Count -eq 2)
    $rejected=$false
    try { [void](& $lookup $wb 'missing') } catch {$rejected=$_.Exception.Message -ceq 'Fixture table missing: missing'}
    Check 'MissingTableSameFailure' $rejected
    $rejected=$false
    try { [void](& $lookup ([pscustomobject]@{Worksheets=(Collection @())}) 'missing') }
    catch {$rejected=$_.Exception.Message -ceq 'Fixture table missing: missing'}
    Check 'EmptyWorkbookSameFailure' $rejected
    # Replace only the release primitive in the test scope to observe ownership.
    # This is not COM evidence: packaged native A/B remains mandatory.
    $tracked=[scriptblock]::Create(@'
param($Workbook,$Name,$Releases)
function Release-ExcelShutdownTableReference($Value) {
    if($null -ne $Value){$Releases.Add($Value)}
}
'@+"`n"+$table+"`n"+'Table $Workbook $Name')
    $releases=[Collections.Generic.List[object]]::new()
    $value=& $tracked $wb 'target' $releases
    Check 'FirstMatchReleasesOnlyOwnedIntermediates' ($releases.Count -eq 4 -and
        [object]::ReferenceEquals($releases[0],$other) -and
        [object]::ReferenceEquals($releases[1],$wb.Worksheets.Item(1).ListObjects) -and
        [object]::ReferenceEquals($releases[2],$wb.Worksheets.Item(1)) -and
        [object]::ReferenceEquals($releases[3],$wb.Worksheets))
    Check 'ReturnedTableAndBorrowedWorkbookNotReleased' (@($releases|Where-Object {
        [object]::ReferenceEquals($_,$first) -or [object]::ReferenceEquals($_,$wb)}).Count -eq 0)
    $releases.Clear()
    $rejected=$false
    try{[void](& $tracked $wb 'missing' $releases)}catch{$rejected=$true}
    Check 'MissingLookupReleasesAllAcquisitions' ($rejected -and $releases.Count -eq 8)
    $releases.Clear()
    $wb.Worksheets.Item(1).ListObjects|Add-Member -Force ScriptMethod Item {
        param([int]$Index)
        if($Index -eq 2){throw 'synthetic acquisition failure'}
        return $this.Entries[$Index-1]
    }
    $rejected=$false
    try{[void](& $tracked $wb 'target' $releases)}catch{$rejected=$true}
    Check 'AcquisitionFailureReleasesOuterOwners' ($rejected -and $releases.Count -eq 4 -and
        [object]::ReferenceEquals($releases[3],$wb.Worksheets))
}
$controller=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'Invoke-ExcelShutdownDiagnostic.ps1'))
$tokens=$null;$errors=$null
$controllerAst=[Management.Automation.Language.Parser]::ParseInput($controller,[ref]$tokens,[ref]$errors)
Check 'ControllerParses' ($errors.Count -eq 0)
$selector=@($controllerAst.ParamBlock.Parameters|Where-Object {$_.Name.VariablePath.UserPath -ceq 'TableTraversal'})
Check 'ControllerDefaultsToBaseline' ($selector.Count -eq 1 -and $selector[0].DefaultValue.SafeGetValue() -ceq 'Baseline')
Check 'ControllerTransformsBeforeTiming' ($controller.IndexOf('$source=ConvertTo-ExcelShutdownTableVariant') -ge 0 -and
    $controller.IndexOf('$source=ConvertTo-ExcelShutdownTableVariant') -lt $controller.IndexOf('$generated=ConvertTo-ExcelShutdownTimedWorker'))
Check 'ControllerLabelsDiagnosticMode' ($controller.Contains("'DiagnosticHelperLifetime'") -and $controller.Contains('TableTraversal=$TableTraversal'))
Check 'ControllerVerifiesDiagnosticPins' ($controller.Contains('diagnostic-tooling-pins.json') -and
    $controller.Contains('DiagnosticToolingPreserved=$diagnosticSame') -and
    $controller.Contains('$restored -and $same -and $toolsSame -and $diagnosticSame'))
$failed=@($checks|Where-Object {-not $_.Passed}).Count
Write-Output ('SHUTDOWN_VARIANTS passed={0} failed={1} total={2}' -f ($checks.Count-$failed),$failed,$checks.Count)
if($failed){exit 1}
