# The parent fixture has disabled this actor through the actual Admin handlers.
# Exercise ordinary Receiving while optional recording stays disabled.
function Test-ReceivingUserAudit($Fixture) {
    $activity=ActivityPins
    $config=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $auth=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $authHash=(Get-FileHash -LiteralPath $auth).Hash
    $operatorName=[string](Run 'invSys.Operations.xlam' 'modTS_Received.ReplaySourceForTest' @('Launch'))
    $operator=$excel.Workbooks.Item($operatorName)
    try {
        $staging=Table $operator 'ReceivedTally'
        $initialRows=$staging.ListRows.Count;$initialEvents=0;$initialKeys=0
        foreach($row in $staging.ListRows){
            if([string]$row.Range.Cells.Item(1,$staging.ListColumns.Item('EventId').Index).Value2 -ne ''){$initialEvents++}
            if([string]$row.Range.Cells.Item(1,$staging.ListColumns.Item('System_Key').Index).Value2 -ne ''){$initialKeys++}
        }
        $prepared=[string](Run 'invSys.Operations.xlam' 'modTS_Received.ReplaySourceForTest' @('Refresh'))
        if($prepared -notmatch '^[1-9][0-9]*$'){throw 'Unrecorded Receiving item fixture unavailable.'}
        $business=BusinessPins
        [void](Run 'invSys.Operations.xlam' 'modTS_Received.ReplaySourceForTest' @('Clear'))
        $cleared=$staging.ListRows.Count -eq 0
        [pscustomobject]@{InitialRows=$initialRows;InitialEventCells=$initialEvents;InitialIdentityCells=$initialKeys;RowsAfterClear=$staging.ListRows.Count}|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'unrecorded-receiving-staging-facts.json')
        Check 'ReceivingRun.UserPolicy.OrdinaryClearPreservesBusiness' ($cleared -and (BoundSame $business (BusinessPins)))
        if(-not $cleared){return}
        $added=(Run 'invSys.Operations.xlam' 'modTS_Received.ReplaySourceForTest' @('Add')) -ceq 'True'
        Check 'ReceivingRun.UserPolicy.OrdinaryAddStillAuthorized' $added
        if(-not $added){return}
        $eventId=[string]$staging.DataBodyRange.Cells.Item(1,$staging.ListColumns.Item('EventId').Index).Value2
        $entityId=[string]$staging.DataBodyRange.Cells.Item(1,$staging.ListColumns.Item('System_Key').Index).Value2
        if($eventId -ceq '' -or $entityId -ceq ''){throw 'Ordinary Add did not supply exact source identities.'}
        $confirmed=(Run 'invSys.Operations.xlam' 'modTS_Received.ReplaySourceForTest' @('Confirm')) -ceq 'True'
        Check 'ReceivingRun.UserPolicy.OrdinaryConfirmStillAuthorized' ($confirmed -and $staging.ListRows.Count -eq 0)
        Check 'ReceivingRun.UserPolicy.CustomHeaderPreserved' (@($staging.ListColumns|Where-Object Name -CEQ 'B0 Custom Display').Count -eq 1)
        $authority=Open-ReceivingEvidenceBook (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb'))
        try {
            $applied=@(Get-ReceivingFixtureRows (Table $authority 'tblAppliedEvents')|Where-Object EventID -CEQ $eventId)
            $logged=@(Get-ReceivingFixtureRows (Table $authority 'tblInventoryLog')|Where-Object {$_.EventID -ceq $eventId -and $_.System_Key -ceq $entityId -and $_.QtyDelta -eq 1})
            Check 'ReceivingRun.UserPolicy.ExactRequiredAppliedEventAndAuditRemain' ($confirmed -and $applied.Count -eq 1 -and $logged.Count -eq 1)
        } finally {$authority.Close($false)}
    } finally {[void](Run 'invSys.Operations.xlam' 'modTS_Received.ReplaySourceForTest' @('Close'))}
    Check 'ReceivingRun.UserPolicy.NoOptionalActivityForOrdinaryReceipt' (BoundSame $activity (ActivityPins))
    Check 'ReceivingRun.UserPolicy.AuthAndConfigUnchanged' ($config -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash -and $authHash -ceq (Get-FileHash -LiteralPath $auth).Hash)
}
