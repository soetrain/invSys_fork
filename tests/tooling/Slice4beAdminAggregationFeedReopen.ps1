# A feed mismatch already observed by the command must require a fresh capture,
# even if another actor later restores the saved setting. All files are owned fixtures.
function Install-AggregationFeedReopenProbe {
    $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestAggregationBoundary').CodeModule.AddFromString(@'
Public Function LiveContextForFeedReopen() As String
    LiveContextForFeedReopen = modAuthUserCommands.CaptureContext()
End Function
'@)
}

function Test-AggregationFeedReopen($Fixture,$Other,$HostBook) {
    function Call([string]$Name,[object[]]$Values=@()){Run 'invSys.Admin.xlam' ('TestAggregation.'+$Name) $Values}
    function Core([string]$Name,[object[]]$Values=@()){Run 'invSys.Core.xlam' ('TestAggregationBoundary.'+$Name) $Values}
    function Hash([string]$Path){(Get-FileHash -LiteralPath $Path).Hash}
    function OpenActual {
        if(-not [bool](Call 'OpenForm' @($HostBook.Name)) -or [int](Call 'SelectSource' @($Fixture.Warehouse)) -ne 1){throw 'Actual feed-reopen handler unavailable; not behavioral RED.'}
    }
    $configBytes=[IO.File]::ReadAllBytes($Fixture.Config)
    $changedFeed=Join-Path $runRoot 'feed-reopen-changed'
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    if(-not [IO.Path]::GetFullPath($changedFeed).StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Feed-reopen fault escapes disposable scope.'}
    $expected=ReadEntities $Fixture.Snapshot 'tblInventorySnapshot'
    try {
        foreach($connection in @($false,$true)){
            $label='Aggregation.FeedReopen.'+$(if($connection){'SuccessfulConnection'}else{'FailedConnection'})
            SelectTarget $Fixture
            [void](Call 'ArmYield' @('',$false));[void](Core 'Arm' @('','','','',''))
            [void](Run 'invSys.Core.xlam' 'TestAggregationRoots.PathComparison' @($true,$Fixture.Output))
            if(-not [bool](Run 'invSys.Core.xlam' 'TestAggregationRoots.NormalizedControl' @($Fixture.Snapshot,$Fixture.Output))){throw 'Feed-reopen control output unavailable.'}
            OpenActual
            $session=[string](Core 'LiveContextForFeedReopen')
            [void](Call 'ArmYield' @('CONNECTION',$connection,'FEED_CHANGED',$Fixture.Snapshot,$Other.Snapshot,$changedFeed))
            [void](Call 'Connect')
            $observed=[bool](Call 'YieldReturned') -and [bool](Core 'Injected') -and [bool](Call 'NeedsReopen')
            Check ($label+'.ChangedFeedObservedThroughActualReturn') $observed
            if(-not $observed){throw 'Changed feed was not observed by the actual return guard; not meaningful RED.'}
            $before=Hash $Fixture.Output
            # Restore only the deliberate saved-config fault. Do not sign in,
            # reselect the target or replace the form before testing the old capture.
            [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
            $sameSession=$session -ceq [string](Core 'LiveContextForFeedReopen')
            Check ($label+'.SameLiveSessionAfterSettingRestoration') $sameSession
            if(-not $sameSession){throw 'Feed restoration changed the live session; not meaningful feed-reopen RED.'}
            [void](Call 'ArmYield' @('',$false));[void](Core 'Arm' @('','','','',''))
            [void](Call 'Aggregate')
            Check ($label+'.OldCaptureStillRefusedBeforeWriter') ([int](Run 'invSys.Core.xlam' 'TestAggregationRoots.Calls') -eq 0)
            Check ($label+'.OldCaptureCannotReportCompletion') (-not [bool](Call 'OwnerSuccess'))
            Check ($label+'.OldCapturePreservesFeedBytes') ((Hash $Fixture.Output) -ceq $before)
            Check ($label+'.ReopeningRemainsRequired') ([bool](Call 'NeedsReopen'))
            $pending=@(Get-ChildItem -LiteralPath (Split-Path -Parent $Fixture.Output) -File -Filter 'invSys.Aggregation.*.retry.json' -ErrorAction SilentlyContinue)
            Check ($label+'.NoUnexpectedRetryArtifact') ($pending.Count -eq 0)
            if($pending.Count){throw 'Unexpected retained attempt requires review; independent case preparation stopped.'}
            [void](Call 'CloseForm')
            SelectTarget $Fixture;OpenActual
            [void](Call 'Aggregate')
            $saved=[bool](Call 'OwnerSuccess')
            Check ($label+'.FreshCaptureCanComplete') $saved
            $actual=ReadEntities $Fixture.Output 'tblGlobalInventorySnapshot'
            $exact=$expected.Count -eq $actual.Count
            foreach($key in $expected.Keys){$exact=$exact -and $actual.ContainsKey($key) -and $actual[$key] -eq $expected[$key]}
            Check ($label+'.FreshCapturePersistsExactSelectedSet') $exact
            Check ($label+'.FeedbackExcludesEnteredPaths') ([bool](Call 'FeedbackRedacted' @($Fixture.Root,$changedFeed)))
            Check ($label+'.NoRedirectedFeed') (-not(Test-Path -LiteralPath (Join-Path $changedFeed 'Global/invSys.Global.InventorySnapshot.xlsb')))
            [void](Call 'CloseForm')
        }
    }finally{
        [void](Call 'CloseForm');[void](Call 'ArmYield' @('',$false));[void](Core 'Arm' @('','','','',''))
        [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
        SelectTarget $Fixture
    }
}
