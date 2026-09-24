# Exercise the builder's actual import helper; no Excel or operational files.
[CmdletBinding()]
param([string]$RepoRoot='.',[ValidateSet('RED','GREEN')][string]$Phase='GREEN')
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
$root=Join-Path $repo ('reports/runtime/vba-import-test-regions/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'tools/build-xlam.ps1'),[ref]$tokens,[ref]$errors)
if($errors){throw 'Builder does not parse; not behavioral RED.'}
foreach($name in @('Remove-VbaTestOnlyRegions','New-NormalizedImportFile')){
    $definition=$ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq $name},$true)
    if($null -eq $definition){throw 'Actual import helper missing; not behavioral RED.'}
    . ([scriptblock]::Create($definition.Extent.Text))
}
$rows=New-Object 'System.Collections.Generic.List[object]'
$runtime="Attribute VB_Name = `"FixtureRuntime`"`nOption Explicit`nPublic Function RuntimeValue() As Long`n    RuntimeValue = 7`nEnd Function`n"
$region="'@TestOnlyBegin`nPublic Sub TestOnlyProbe()`nEnd Sub`n'@TestOnlyEnd`n"
function Import-Text([string]$Name,[string]$Text){
    $path=Join-Path $root ($Name+'.bas')
    [IO.File]::WriteAllText($path,$Text,[Text.Encoding]::ASCII)
    $import=New-NormalizedImportFile (Get-Item -LiteralPath $path)
    # Retain exact import bytes in the ignored report directory for inspection.
    Copy-Item -LiteralPath $import -Destination (Join-Path $root ($Name+'.import.bas'))
    return [IO.File]::ReadAllText($import)
}
function Check-Import([string]$Name,[string]$Text,[string]$Expected,[bool]$Reject=$false){
    $rejected=$false;$actual=''
    try {$actual=Import-Text $Name $Text}
    catch {
        if($_.Exception.Message -notmatch 'Unbalanced VBA test-only markers|VBA test-only marker remained'){throw}
        $rejected=$true
    }
    $passed=if($Reject){$rejected}else{-not $rejected -and $actual -ceq ($Expected -replace "`r?`n","`r`n")}
    $rows.Add([pscustomobject]@{Check=$Name;Passed=[bool]$passed;Rejected=$rejected})
    Write-Output ($Name+': '+$(if($passed){'PASS'}else{'FAIL'}))
}
$lf=$runtime+$region
$crlf=$lf.Replace("`n","`r`n")
Check-Import 'Import.LF' $lf $runtime
Check-Import 'Import.CRLF' $crlf $runtime
Check-Import 'Import.Mixed' ($runtime.Replace("`n","`r`n")+$region) $runtime
Check-Import 'Import.CRLFMarkerAtEOF' $crlf.TrimEnd([char[]]"`r`n") $runtime
$indented=$crlf.Replace("'@TestOnlyBegin"," `t'@TestOnlyBegin `t").Replace("'@TestOnlyEnd"," `t'@TestOnlyEnd `t")
Check-Import 'Import.IndentedCRLF' $indented $runtime
Check-Import 'Import.MultipleCRLFRegions' ($crlf+$region.Replace("`n","`r`n")) $runtime
Check-Import 'Reject.MissingEndLF' ($runtime+"'@TestOnlyBegin`nPublic Sub TestOnlyProbe()`nEnd Sub`n") '' $true
Check-Import 'Reject.MissingEndCRLF' ($runtime.Replace("`n","`r`n")+"'@TestOnlyBegin`r`nPublic Sub TestOnlyProbe()`r`nEnd Sub`r`n") '' $true
Check-Import 'Reject.StrayEndCRLF' ($runtime.Replace("`n","`r`n")+"'@TestOnlyEnd`r`n") '' $true
Check-Import 'Reject.NestedCRLF' ($runtime.Replace("`n","`r`n")+"'@TestOnlyBegin`r`n"+$region.Replace("`n","`r`n")+"'@TestOnlyEnd`r`n") '' $true
$literal=$runtime+"Private Const LabelText As String = `"'@TestOnlyBegin`"`n"
Check-Import 'Import.MarkerInLiteralPreserved' $literal.Replace("`n","`r`n") $literal
$production=Get-Content -LiteralPath (Join-Path $repo 'src/Production/Modules/mProduction.bas') -Raw
$productionLF=$production -replace "`r`n","`n"
$expected=Import-Text 'ProductionLF' $productionLF
$actual=Import-Text 'ProductionCRLF' $productionLF.Replace("`n","`r`n")
$clean=$actual -notmatch '(?im)^\s*Public\s+(Function|Sub)\s+Test' -and $actual -notmatch "'@TestOnly(Begin|End)"
$rows.Add([pscustomobject]@{Check='Import.ProductionLineEndingsEquivalent';Passed=($actual -ceq $expected -and $clean);Rejected=$false})
Write-Output ('Import.ProductionLineEndingsEquivalent: '+$(if($rows[$rows.Count-1].Passed){'PASS'}else{'FAIL'}))
$rows.ToArray()|ConvertTo-Json|Set-Content (Join-Path $root ($Phase.ToLowerInvariant()+'.json'))
$failed=@($rows|Where-Object {-not $_.Passed}).Count
Write-Output ("$Phase : "+($rows.Count-$failed)+' passed, '+$failed+' failed; '+$root)
if($failed){exit 1}
