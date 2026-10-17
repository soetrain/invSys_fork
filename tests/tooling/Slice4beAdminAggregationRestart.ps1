# Real Core/Admin package unload/reload in the same owned Excel host. Fixture
# credentials remain solely in the existing worker's memory. This verifies an
# add-in restart; an Excel-process or machine crash is a separate qualification.
function Test-AggregationRestart($Fixture,$Other,$HostBook) {
    function Call([string]$Name,[object[]]$Values=@()){Run 'invSys.Admin.xlam' ('TestAggregation.'+$Name) $Values}
    function Core([string]$Name,[object[]]$Values=@()){Run 'invSys.Core.xlam' ('TestAggregationBoundary.'+$Name) $Values}
    function HashIfPresent([string]$Path){if(-not(Test-Path -LiteralPath $Path)){return ''};(Get-FileHash -LiteralPath $Path).Hash}
    SelectTarget $Fixture
    [void](Run 'invSys.Core.xlam' 'TestAggregationRoots.SetRoots' @($Fixture.Root,$Other.Root))
    [void](Run 'invSys.Core.xlam' 'TestAggregationRoots.PathComparison' @($true,$Fixture.Output))
    if(-not [bool](Run 'invSys.Core.xlam' 'TestAggregationRoots.NormalizedControl' @($Fixture.Snapshot,$Fixture.Output))){throw 'Restart output fixture did not persist.'}
    if(-not [bool](Call 'OpenForm' @($HostBook.Name)) -or [int](Call 'SelectSource' @($Fixture.Warehouse)) -ne 1){throw 'Restart fixture did not enter the actual form.'}
    $beforeBooks=@($excel.Workbooks|ForEach-Object {[string]$_.Name})
    [void](Core 'Arm' @('WRITE_FAILURE','BEFORE_SAVE',$Fixture.Snapshot,'',''))
    [void](Call 'Aggregate')
    if(-not [bool](Core 'Injected')){throw 'Restart interruption boundary was not reached; not behavioral RED.'}
    Check 'Aggregation.Restart.InterruptedAttemptHasNoFalseCompletion' (-not [bool](Call 'OwnerSuccess'))
    [void](Call 'CloseForm')
    $outputAtInterruption=HashIfPresent $Fixture.Output
    $retry=@(Get-ChildItem -LiteralPath (Split-Path -Parent $Fixture.Output) -Recurse -File -Filter '*aggregation*retry*.json' -ErrorAction SilentlyContinue)
    $receiptPins=@{};foreach($file in $retry){$receiptPins[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash}
    Check 'Aggregation.Restart.DurablePendingEvidenceExists' ($retry.Count -gt 0)
    # Close only workbooks created by the interrupted disposable command.
    foreach($book in @($excel.Workbooks)){if([string]$book.Name -cnotin $beforeBooks){$book.Close($false)}}
    foreach($name in @('invSys.Admin.xlam','invSys.Operations.xlam','invSys.Designs.Domain.xlam','invSys.Inventory.Domain.xlam','invSys.Core.xlam')){
        $packages[$name].Close($false)
        $packages.Remove($name)
    }
    $remaining=@($excel.Workbooks|Where-Object { [string]$_.Name -ceq 'invSys.Core.xlam' })
    if($remaining.Count){throw 'Core did not unload; not a meaningful restart test.'}
    foreach($name in @('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam','invSys.Admin.xlam')){
        $packages[$name]=$excel.Workbooks.Open((Join-Path $deploy $name),0,$false)
    }
    $fresh=[int](Run 'invSys.Core.xlam' 'TestAggregationRoots.Calls') -eq 0
    if(-not $fresh){throw 'Core runtime state was retained; not a meaningful restart test.'}
    Check 'Aggregation.Restart.CoreRuntimeWasReloaded' $fresh
    $same=$retry.Count -gt 0
    foreach($path in $receiptPins.Keys){$same=$same -and (Test-Path -LiteralPath $path) -and (HashIfPresent $path) -ceq $receiptPins[$path]}
    Check 'Aggregation.Restart.PendingEvidenceSurvivesReload' $same
    SelectTarget $Fixture
    [void](Run 'invSys.Core.xlam' 'TestAggregationRoots.SetRoots' @($Fixture.Root,$Other.Root))
    [void](Run 'invSys.Core.xlam' 'TestAggregationRoots.PathComparison' @($true,$Fixture.Output))
    if(-not [bool](Call 'OpenForm' @($HostBook.Name)) -or [int](Call 'SelectSource' @($Fixture.Warehouse)) -ne 1){throw 'Fresh restart form did not open.'}
    [void](Call 'Aggregate')
    Check 'Aggregation.Restart.ReopenedFormDoesNotAutoPromoteUnresolvedAttempt' (-not [bool](Call 'OwnerSuccess') -and [int](Run 'invSys.Core.xlam' 'TestAggregationRoots.Calls') -eq 0)
    Check 'Aggregation.Restart.UnresolvedFeedBytesPreserved' ((HashIfPresent $Fixture.Output) -ceq $outputAtInterruption)
    [void](Call 'CloseForm')
}
