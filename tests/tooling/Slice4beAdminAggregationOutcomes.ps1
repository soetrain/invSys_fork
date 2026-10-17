# Actual failed connection and post-save audit failure; no activity is fabricated.
function Test-AggregationOutcomeRecording($Fixture,$Other,$HostBook){
    function PendingReceipts {
        @(Get-ChildItem -LiteralPath (Split-Path -Parent $Fixture.Output) -File -Filter 'invSys.Aggregation.*.retry.json' -ErrorAction SilentlyContinue)
    }
    try{
        SelectTarget $Fixture
        if(-not [bool](Activity 'SetPolicy')){throw 'Outcome recording policy prerequisite unavailable.'}
        if(-not [bool](Invoke 'OpenForm' @($HostBook.Name))){throw 'Actual connection form unavailable.'}
        [void](Invoke 'SetConnectionResult' @($false))
        [void](Run 'invSys.Core.xlam' 'TestAggregationRoots.ResetError')
        Observe 'CONNECT' 'FAILED' {[void](Invoke 'RecordingGesture' @('Connect',$Fixture.Root,$Fixture.Secret))} 'ConnectionNoOutput' -NoRetryClaim
        Check 'Aggregation.Recording.ConnectionNoOutput.NoAggregationWriter' ([int](Run 'invSys.Core.xlam' 'TestAggregationRoots.Calls') -eq 0)
        Check 'Aggregation.Recording.ConnectionNoOutput.NoOutputOrReceipt' (-not(Test-Path -LiteralPath $Fixture.Output) -and @(PendingReceipts).Count -eq 0)
        [void](Invoke 'CloseForm')
        if(-not [bool](Invoke 'OpenForm' @($HostBook.Name)) -or [int](Invoke 'SelectSource' @($Fixture.Warehouse)) -ne 1){throw 'Actual saved-output form/selection unavailable.'}
        [void](Invoke 'SetAuditFailure' @($true))
        Observe 'AGGREGATE' 'COMPLETED_STOPPED' {[void](Invoke 'Aggregate')} 'AuditStopped' -NoRetryClaim
        Check 'Aggregation.Recording.AuditStopped.ActualAuditFailureReached' ([bool](Invoke 'AuditFailureReached'))
        Check 'Aggregation.Recording.AuditStopped.NoContinuedOwnerSuccess' (-not [bool](Invoke 'OwnerSuccess'))
        Check 'Aggregation.Recording.AuditStopped.OneWriterAndExactPersistedOutput' ([int](Run 'invSys.Core.xlam' 'TestAggregationRoots.Calls') -eq 1 -and (ExactOutput))
        Check 'Aggregation.Recording.AuditStopped.CompletedCoreReceiptAlreadyResolved' (@(PendingReceipts).Count -eq 0)
        Check 'Aggregation.Recording.AuditStopped.KnownPersistenceAcknowledged' ([bool](Invoke 'RecordingAcknowledgesPersistence'))
    }finally{[void](Invoke 'SetAuditFailure' @($false));[void](Invoke 'CloseForm');SelectTarget $Fixture}
}
