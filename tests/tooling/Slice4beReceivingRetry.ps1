# A second explicit Confirm action exercises the unchanged owning retry path.
# The activity layer never starts a retry. All data is disposable fixture data.
function Test-ReceivingRetryAfterUncertainSubmission($Fixture,$Operator,$Other,$Expected,[string]$OtherHash) {
    $before = @(Get-Slice4beActivityFiles $Fixture)
    $snapshot = @{}; foreach ($path in $before) { $snapshot[$path] = Get-ReceivingFixtureHash $path }
    [void](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.SetLoseAcknowledgement' @($false))
    [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.Reopen' @($Operator.Name))
    $status = [string](Run 'invSys.Operations.xlam' 'TestReceivingActivity.Confirm' @($Operator.Name,$Other.Name,[bool]$CaptureEvidence))
    if ($CaptureEvidence) {
        Start-Sleep -Milliseconds 300
        CaptureFormEvidence 'Receiving' 'receiving-explicit-retry.png'
        [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.CloseForm')
    }
    $inboxPath = [string](Run 'invSys.Core.xlam' 'modRoleEventWriter.ResolveInboxWorkbookPath' @('RECEIVE',$Fixture.Warehouse,'S1',''))
    if (-not [IO.Path]::GetFullPath($inboxPath).StartsWith($Fixture.Root.TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Retry inbox escaped fixture root.' }
    $inbox = Open-ReceivingEvidenceBook $inboxPath
    $authority = Open-ReceivingEvidenceBook (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb'))
    $inboxRows = @(Get-ReceivingFixtureRows (Table $inbox 'tblInboxReceive'))
    $appliedRows = @(Get-ReceivingFixtureRows (Table $authority 'tblAppliedEvents'))
    $loggedRows = @(Get-ReceivingFixtureRows (Table $authority 'tblInventoryLog'))
    $exact = $status.StartsWith('Succeeded=True') -and (Table $Operator 'ReceivedTally').ListRows.Count -eq 0
    foreach ($item in $Expected) {
        $queued = @($inboxRows | Where-Object { $_.EventID -ceq $item.EventId })
        $applied = @($appliedRows | Where-Object { $_.EventID -ceq $item.EventId })
        $logged = @($loggedRows | Where-Object { $_.EventID -ceq $item.EventId -and $_.System_Key -ceq $item.System_Key -and $_.QtyDelta -eq $item.QUANTITY })
        $exact = $exact -and $queued.Count -eq 1 -and $applied.Count -eq 1 -and $logged.Count -eq 1
    }
    Check 'Receiving.ExplicitRetry.AppliedOnceWithOriginalIdentities' $exact
    Check 'Receiving.ExplicitRetry.CapturedWorkbookAndQuietUi' ($status.Contains('BoundWorkbook='+$Operator.Name) -and $status.Contains('QuietDuring=True') -and $status.Contains('QuietRestored=True'))
    Check 'Receiving.ExplicitRetry.OtherWorkbookPreserved' ($OtherHash -ceq (Get-ReceivingFixtureHash $Other.FullName) -and $Other.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'unrelated workbook sentinel')
    Test-ReceivingObservations $Fixture $before $Expected $false 'ExplicitRetry' 'CONFIRMED' 'Submitted'
    $unchanged = $true; foreach ($path in $snapshot.Keys) { $unchanged = $unchanged -and $snapshot[$path] -ceq (Get-ReceivingFixtureHash $path) }
    Check 'Receiving.ExplicitRetry.PriorObservationsImmutable' $unchanged
    $priorIds = @($before | ForEach-Object { ([IO.File]::ReadAllText($_) | ConvertFrom-Json).ActivityId })
    $newIds = @(Get-Slice4beActivityFiles $Fixture | Where-Object { $_ -notin $before } | ForEach-Object { ([IO.File]::ReadAllText($_) | ConvertFrom-Json).ActivityId } | Select-Object -Unique)
    Check 'Receiving.ExplicitRetry.DistinctActivityId' ($newIds.Count -eq 1 -and $newIds[0] -cnotin $priorIds)
    foreach ($book in $receivingEvidenceOpened) { $book.Close($false) }
    $receivingEvidenceOpened.Clear()
}
