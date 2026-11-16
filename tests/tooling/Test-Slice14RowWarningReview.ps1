[CmdletBinding()]
param([string]$RepoRoot='.')

$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
$helper=Join-Path $repo 'tools/Review-RetiredRowWarnings.ps1'
if(-not(Test-Path -LiteralPath $helper)){throw 'The focused static warning review helper is missing.'}
. $helper

$paths=@(
    'src/Admin/Modules/modInventoryUiText.bas',
    'src/Core/ClassModules/cInventoryCreatePlan.cls',
    'src/Core/ClassModules/cInventoryEditRow.cls',
    'src/Core/ClassModules/cInventoryTableInput.cls',
    'src/Core/Modules/modAggregationSources.bas',
    'src/Core/Modules/modInventoryEditCommands.bas'
)
$warnings=@($paths|ForEach-Object {[pscustomobject]@{code='RETIRED_ROW_HEADER';sourcePath=$_}})
$checks=[System.Collections.Generic.List[object]]::new()
function Check([string]$Name,[bool]$Passed){$checks.Add([pscustomobject]@{Name=$Name;Passed=$Passed});Write-Output ("${Name}: "+$(if($Passed){'PASS'}else{'FAIL'}))}
$reviewed=Classify-StaticWarningDelta -ActualWarnings $warnings -BaselineWarnings @() -RepoRoot $repo
Check 'StaticRowReview.ExactReviewedSix' (@($reviewed.Reviewed).Count -eq 6 -and @($reviewed.Unreviewed).Count -eq 0)
$extra=@($warnings)+@([pscustomobject]@{code='RETIRED_ROW_HEADER';sourcePath='src/Core/Modules/newRowIdentity.bas'})
$positive=Classify-StaticWarningDelta -ActualWarnings $extra -BaselineWarnings @() -RepoRoot $repo
Check 'StaticRowReview.NewRowUseStillFails' (@($positive.Reviewed).Count -eq 6 -and @($positive.Unreviewed).Count -eq 1)
$changed=Classify-StaticWarningDelta -ActualWarnings $warnings -BaselineWarnings @() -RepoRoot (Join-Path $repo 'tests')
Check 'StaticRowReview.SourceChangeRequiresReview' (@($changed.Reviewed).Count -eq 0 -and @($changed.Unreviewed).Count -eq 6)
$old=Classify-StaticWarningDelta -ActualWarnings $warnings -BaselineWarnings @($warnings[0]) -RepoRoot $repo
Check 'StaticRowReview.HistoricalBaselinePreserved' (@($old.Reviewed).Count -eq 5 -and @($old.Unreviewed).Count -eq 0)
$wrongCode=Classify-StaticWarningDelta -ActualWarnings @([pscustomobject]@{code='OTHER_WARNING';sourcePath=$paths[0]}) -BaselineWarnings @() -RepoRoot $repo
Check 'StaticRowReview.OtherCodeNotSuppressed' (@($wrongCode.Reviewed).Count -eq 0 -and @($wrongCode.Unreviewed).Count -eq 1)
$failed=@($checks|Where-Object {-not $_.Passed}).Count
Write-Output ("Static warning review: $($checks.Count-$failed) passed, $failed failed")
if($failed){exit 1}
