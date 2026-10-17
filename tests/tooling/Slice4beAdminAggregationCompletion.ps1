# Actual handlers on owner-generated fixtures. Partial bytes are injected only
# after the baseline's writer closes its owned output, before its success return.
function Test-AggregationCompletion($Fixture,$Other,$HostBook) {
    function Call([string]$Name,[object[]]$Values=@()){Run 'invSys.Admin.xlam' ('TestAggregation.'+$Name) $Values}
    function Core([string]$Name,[object[]]$Values=@()){Run 'invSys.Core.xlam' ('TestAggregationBoundary.'+$Name) $Values}
    function Hash([string]$Path){if(-not(Test-Path -LiteralPath $Path)){return ''};(Get-FileHash -LiteralPath $Path).Hash}
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    if(-not [IO.Path]::GetFullPath($Fixture.Output).StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Completion output escapes the owned fixture.'}
    SelectTarget $Fixture
    [void](Run 'invSys.Core.xlam' 'TestAggregationRoots.SetRoots' @($Fixture.Root,$Other.Root))
    [void](Run 'invSys.Core.xlam' 'TestAggregationRoots.PathComparison' @($true,$Fixture.Output))
    Check 'Aggregation.Legacy.AuthorizedPositiveControl' ([bool](Call 'LegacyAggregate' @($HostBook.Name,$Fixture.Warehouse)))
    $before=Hash $Fixture.Output
    [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
    [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.SetCurrentUserId' @('config-admin'))
    Check 'Aggregation.Legacy.SignedOutPrecondition' (-not [bool](Run 'invSys.Core.xlam' 'modAuth.IsSignedIn'))
    $wrote=[bool](Call 'LegacyAggregate' @($HostBook.Name,$Fixture.Warehouse))
    Check 'Aggregation.Legacy.SavedHintDeniedBeforeCoreWriter' (-not $wrote -and [int](Run 'invSys.Core.xlam' 'TestAggregationRoots.Calls') -eq 0)
    Check 'Aggregation.Legacy.SignedOutFeedBytesPreserved' ((Hash $Fixture.Output) -ceq $before)
    SelectTarget $Fixture
    if(-not [bool](Call 'OpenForm' @($HostBook.Name)) -or [int](Call 'SelectSource' @($Fixture.Warehouse)) -ne 1){throw 'Actual completion handler unavailable.'}
    [void](Core 'Arm' @('PARTIAL_OUTPUT','AFTER_PROMOTION','','',$Fixture.Output))
    [void](Call 'Aggregate')
    if(-not [bool](Core 'Injected')){throw 'Partial persistence boundary was not reached; not meaningful RED.'}
    $bytes=[IO.File]::ReadAllBytes($Fixture.Output)
    $partial=[Text.Encoding]::ASCII.GetString($bytes) -ceq 'Incomplete disposable aggregation output'
    if(-not $partial){throw 'Partial artifact injection did not persist; not meaningful RED.'}
    Check 'Aggregation.Completion.PartialPersistencePrecondition' $partial
    Check 'Aggregation.Completion.PartialOutputNeverReportsComplete' (-not [bool](Call 'OwnerSuccess'))
    $partialHash=Hash $Fixture.Output
    $retry=@(Get-ChildItem -LiteralPath (Split-Path -Parent $Fixture.Output) -Recurse -File -Filter '*aggregation*retry*.json' -ErrorAction SilentlyContinue)
    Check 'Aggregation.Completion.UnknownPersistenceRetainsRetryEvidence' ($retry.Count -gt 0)
    [void](Core 'Arm' @('','','','',''))
    [void](Call 'Aggregate')
    Check 'Aggregation.Completion.SameFormCannotOverwriteUnknownFeed' (-not [bool](Call 'OwnerSuccess') -and (Hash $Fixture.Output) -ceq $partialHash)
    [void](Call 'CloseForm')
    if(-not [bool](Call 'OpenForm' @($HostBook.Name)) -or [int](Call 'SelectSource' @($Fixture.Warehouse)) -ne 1){throw 'Fresh completion form unavailable.'}
    [void](Call 'Aggregate')
    Check 'Aggregation.Completion.ReopenedFormCannotAutoPromoteUnknownFeed' (-not [bool](Call 'OwnerSuccess') -and (Hash $Fixture.Output) -ceq $partialHash)
    [void](Call 'ConfirmRetryReview' @($true))
    [void](Call 'Aggregate')
    Check 'Aggregation.Completion.ConfirmedReviewCannotEraseUnknownFeed' (-not [bool](Call 'OwnerSuccess') -and (Hash $Fixture.Output) -ceq $partialHash)
    if($retry.Count){
        Check 'Aggregation.Completion.ConfirmedUnknownReviewRetainsReceipt' (@($retry|Where-Object {Test-Path -LiteralPath $_.FullName}).Count -eq $retry.Count)
    }
    [void](Call 'CloseForm')
}
