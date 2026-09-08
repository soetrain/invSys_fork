# D18 next control coverage. New cases exercise actual Add/Confirm handlers;
# direct staging is a supplemental negative user-attribution check.
function Test-ReceivingControlRecords($Fixture,[string[]]$Before,[string]$ControlId,
    [string]$Owner,[string]$Prefix,[string]$Outcome,[int]$Actions,[string]$Effect,$ExpectedSources,[string]$Label,[string]$Severity='Info',[int]$MinimumCatalog=3) {
    $records = @(); $payloads = @()
    foreach ($path in @(Get-Slice4beActivityFiles $Fixture)) {
        if ($path -in $Before) { continue }
        $raw = [IO.File]::ReadAllText($path); $record = $raw | ConvertFrom-Json
        if ($record.ControlId -ceq $ControlId) { $records += $record; $payloads += $raw }
    }
    $attempts = @($records | Where-Object OutcomeCode -eq 'REQUESTED')
    $outcomes = @($records | Where-Object OutcomeCode -eq $Outcome)
    $complete = $attempts.Count -eq $Actions -and $outcomes.Count -eq $Actions -and $records.Count -eq ($Actions*2)
    Check "$Label.AttemptsAndOutcomes" $complete
    $correlated = $complete; $owned = $complete; $references = $complete; $redacted = $complete
    $ids = @($attempts | ForEach-Object ActivityId | Select-Object -Unique)
    $correlated = $correlated -and $ids.Count -eq $Actions
    foreach ($attempt in $attempts) {
        $matched = @($outcomes | Where-Object { $_.ActivityId -ceq $attempt.ActivityId })
        $correlated = $correlated -and $matched.Count -eq 1
        if ($matched.Count -ne 1) { continue }
        $last = $matched[0]
        $correlated = $correlated -and $attempt.RecordId -cne $last.RecordId
        $owned = $owned -and $attempt.EventCode -ceq ($Prefix+'REQUESTED') -and $attempt.DataEffect -ceq 'Unknown' -and
            $last.EventCode -ceq ($Prefix+$Outcome) -and $last.DataEffect -ceq $Effect -and $last.Severity -ceq $Severity
        $references = $references -and @($attempt.SourceEventRefs).Count -eq 0 -and @($last.SourceEventRefs).Count -eq $ExpectedSources.Count
        foreach ($source in $ExpectedSources) {
            $match = @($last.SourceEventRefs | Where-Object { $_.EventId -ceq $source.EventId -and $_.WarehouseId -ceq $Fixture.Warehouse -and $_.SourceKind -ceq 'Inventory' -and $_.SubmissionState -ceq 'Submitted' })
            $references = $references -and $match.Count -eq 1
        }
    }
    $readable = $complete
    foreach ($record in $records) {
        $owned = $owned -and $record.OwnerId -ceq $Owner -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.UserId -ceq 'config-reader' -and $record.SourceRole -ceq 'Receiving'
        $readable = $readable -and $record.CatalogVersion -in @(3,4) -and $record.CatalogVersion -ge $MinimumCatalog -and (Get-ActivityRead $record.RecordId).StartsWith('OK|')
    }
    foreach ($raw in $payloads) {
        foreach ($value in @($Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'ACTIVITY-PRIVATE','mBtnAdd_Click','mBtnConfirm_Click','Err.Description')) {
            if ($raw.IndexOf($value,[StringComparison]::OrdinalIgnoreCase) -ge 0) { $redacted = $false }
        }
    }
    Check "$Label.DistinctCorrelatedActions" $correlated
    Check "$Label.OwnerAndEffect" $owned
    Check "$Label.ExactSubmissionReferencesOnly" $references
    Check "$Label.InputValuesExcluded" $redacted
    Check "$Label.SupportedCatalogRead" $readable
}

function Test-ReceivingStagingCoverage($Fixture) {
    foreach ($disposition in @($false,$true)) {
        $label = if ($disposition) { 'Disposition' } else { 'ReceiptAdd' }
        $operator = $excel.Workbooks.Add()
        $operator.SaveAs((Join-Path $runRoot ('coverage-'+$label+'.xlsm')),52)
        $other = $excel.Workbooks.Add()
        $other.Worksheets.Item(1).Cells.Item(1,1).Value2 = 'unrelated staging sentinel'
        $other.SaveAs((Join-Path $runRoot ('coverage-other-'+$label+'.xlsm')),52)
        $otherHash = Get-ReceivingFixtureHash $other.FullName
        $before = @(Get-Slice4beActivityFiles $Fixture)
        $staged = [bool](Run 'invSys.Operations.xlam' 'TestReceivingActivity.Stage' @($operator.Name,$disposition,$other.Name))
        $staging = Table $operator 'ReceivedTally'
        $rows = @(Get-ReceivingFixtureRows $staging)
        if (-not $staged -or $rows.Count -ne 2) { throw 'Staging coverage actual Add handlers did not stage two rows.' }
        $uniqueKeys = @($rows | ForEach-Object System_Key | Select-Object -Unique)
        $uniqueEvents = @($rows | ForEach-Object EventId | Select-Object -Unique)
        $identities = $uniqueEvents.Count -eq 2 -and '' -notin $uniqueEvents -and '' -notin $uniqueKeys
        if ($disposition) {
            $identities = $identities -and $uniqueKeys.Count -eq 1 -and $rows[0].RECEIPT_TYPE -ceq 'RETURN' -and $rows[1].RECEIPT_TYPE -ceq 'DUMP'
        } else { $identities = $identities -and $uniqueKeys.Count -eq 2 }
        Check "Coverage.$label.OwningStageIdentities" $identities
        $control = if ($disposition) { 'DISPOSITION_ADD_SELECTED' } else { 'RECEIVING_ADD_SELECTED' }
        $owner = if ($disposition) { 'RECEIVING_DISPOSITION' } else { 'RECEIVING_STAGING' }
        $prefix = if ($disposition) { 'DISPOSITION_ADD_' } else { 'RECEIVE_ADD_' }
        Test-ReceivingControlRecords $Fixture $before $control $owner $prefix 'STAGED' 2 'Changed' @() "Coverage.$label"
        Show-ReceivingStagingEvidence ($label+'-staged')
        foreach ($rejected in @($true,$false)) {
            $case = if ($rejected) { 'Rejected' } else { 'Failed' }
            $beforeCheck = @(Get-Slice4beActivityFiles $Fixture)
            $stagingBefore = @(Get-ReceivingFixtureRows $staging) | ConvertTo-Json -Depth 5 -Compress
            if (-not $rejected) { $staging.Parent.Protect() }
            try { $status = [string](Run 'invSys.Operations.xlam' 'TestReceivingActivity.AddCheck' @($rejected)) }
            finally { if (-not $rejected) { $staging.Parent.Unprotect() } }
            $stagingAfter = @(Get-ReceivingFixtureRows $staging) | ConvertTo-Json -Depth 5 -Compress
            $cause = if ($rejected) { 'Quantity must be greater than zero.' } else { 'staging failed' }
            Check "Coverage.$label.$case.BusinessUnchangedAndCauseVisible" ($stagingBefore -ceq $stagingAfter -and $status.Contains($cause))
            Show-ReceivingStagingEvidence ($label+'-'+$case)
            $effect = if ($rejected) { 'Unchanged' } else { 'Unknown' }
            $severity = if ($rejected) { 'Warning' } else { 'Error' }
            Test-ReceivingControlRecords $Fixture $beforeCheck $control $owner $prefix $case.ToUpperInvariant() 1 $effect @() "Coverage.$label.$case" $severity
        }
        $beforeUnavailable = @(Get-Slice4beActivityFiles $Fixture)
        $status = Invoke-ReceivingStagingStoreFault $Fixture
        Check "Coverage.$label.StoreFailureDoesNotBlockStaging" ($staging.ListRows.Count -eq 3 -and $status.Contains('Staged ') -and $status.Contains('Tracking unavailable'))
        Check "Coverage.$label.StoreFailureDoesNotInventEvidence" (@(Get-Slice4beActivityFiles $Fixture).Count -eq $beforeUnavailable.Count)
        Show-ReceivingStagingEvidence ($label+'-tracking-unavailable')
        $beforeDirect = @(Get-Slice4beActivityFiles $Fixture)
        $direct = [bool](Run 'invSys.Operations.xlam' 'TestReceivingActivity.DirectStage' @($operator.Name,$disposition))
        Check "Coverage.$label.DirectServiceIsNotUserAction" ($direct -and $staging.ListRows.Count -eq 4 -and @(Get-Slice4beActivityFiles $Fixture).Count -eq $beforeDirect.Count)
        if (-not $direct -or $staging.ListRows.Count -ne 4) { throw 'Direct staging fixture did not establish its business effect.' }
        $expected = @(Get-ReceivingFixtureRows $staging)
        $extra = $staging.ListColumns.Add(); $extra.Name = 'Coverage Extra'
        $beforeConfirm = @(Get-Slice4beActivityFiles $Fixture)
        $status = [string](Run 'invSys.Operations.xlam' 'TestReceivingActivity.Confirm' @($operator.Name,$other.Name,[bool]$CaptureEvidence))
        if ($CaptureEvidence) {
            Start-Sleep -Milliseconds 300
            CaptureFormEvidence 'Receiving' ('coverage-'+$label.ToLowerInvariant()+'.png')
            [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.CloseForm')
        }
        $authority = Open-ReceivingEvidenceBook (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb'))
        $logged = @(Get-ReceivingFixtureRows (Table $authority 'tblInventoryLog'))
        $applied = @(Get-ReceivingFixtureRows (Table $authority 'tblAppliedEvents'))
        $business = $status.StartsWith('Succeeded=True') -and $staging.ListRows.Count -eq 0
        foreach ($item in $expected) {
            $delta = if ($disposition) { -[double]$item.QUANTITY } else { [double]$item.QUANTITY }
            $type = if ($disposition) { $item.RECEIPT_TYPE } else { 'RECEIVE' }
            $matching = @($logged | Where-Object { $_.EventID -ceq $item.EventId -and $_.System_Key -ceq $item.System_Key -and $_.QtyDelta -eq $delta -and $_.EventType -ceq $type })
            $business = $business -and $matching.Count -eq 1 -and @($applied | Where-Object { $_.EventID -ceq $item.EventId }).Count -eq 1
        }
        Check "Coverage.$label.IndependentAppliedBusinessEvidence" $business
        if (-not $business) { throw 'Staging coverage did not establish the intended owning command outcome.' }
        Check "Coverage.$label.CapturedWorkbookAndUnknownColumn" ($status.Contains('BoundWorkbook='+$operator.Name) -and $extra.Name -ceq 'Coverage Extra')
        Check "Coverage.$label.OtherWorkbookPreserved" ($otherHash -ceq (Get-ReceivingFixtureHash $other.FullName) -and $other.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'unrelated staging sentinel')
        if ($disposition) {
            Test-ReceivingControlRecords $Fixture $beforeConfirm 'DISPOSITION_CONFIRM' 'RECEIVING_DISPOSITION' 'DISPOSITION_CONFIRM_' 'CONFIRMED' 1 'Unknown' $expected 'Coverage.DispositionConfirm'
        } else { Test-ReceivingObservations $Fixture $beforeConfirm $expected $false 'ReceiptAddConfirm' 'CONFIRMED' 'Submitted' }
        foreach ($book in $receivingEvidenceOpened) { $book.Close($false) }
        $receivingEvidenceOpened.Clear()
        $operator.Close($false); $other.Close($false)
    }
    $operator = $excel.Workbooks.Add()
    $operator.SaveAs((Join-Path $runRoot 'coverage-stale-add.xlsm'),52)
    if (-not [bool](Run 'invSys.Operations.xlam' 'TestReceivingActivity.Stage' @($operator.Name))) { throw 'Stale Add fixture staging failed.' }
    $staging = Table $operator 'ReceivedTally'
    $stagedBefore = @(Get-ReceivingFixtureRows $staging) | ConvertTo-Json -Depth 5 -Compress
    $before = @(Get-Slice4beActivityFiles $Fixture)
    [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
    SelectTarget $Fixture 'config-reader'
    $status = [string](Run 'invSys.Operations.xlam' 'TestReceivingActivity.AddCheck' @($false))
    Check 'Coverage.StaleAdd.StagingPreserved' ($stagedBefore -ceq (@(Get-ReceivingFixtureRows $staging) | ConvertTo-Json -Depth 5 -Compress))
    Check 'Coverage.StaleAdd.VisibleRejectionWithoutAttribution' ($status.Contains('Session or warehouse changed') -and @(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count)
    Show-ReceivingStagingEvidence 'stale-add'
    [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.CloseForm')
    $operator.Close($false)
}

function Show-ReceivingStagingEvidence([string]$Name) {
    if (-not $CaptureEvidence) { return }
    [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.ShowForm' @($true))
    Start-Sleep -Milliseconds 300
    CaptureFormEvidence 'Receiving' ('coverage-'+$Name.ToLowerInvariant()+'.png')
    [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.ShowForm' @($false))
}

function Invoke-ReceivingStagingStoreFault($Fixture) {
    $leaf = Join-Path $Fixture.Root ('Training\Activity\'+$Fixture.Warehouse)
    $held = $leaf+'-staging-fixture-held'
    foreach ($path in @($leaf,$held)) {
        if (-not [IO.Path]::GetFullPath($path).StartsWith($Fixture.Root.TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Staging store fault escaped its fixture.' }
    }
    if (-not (Test-Path -LiteralPath $leaf -PathType Container) -or (Test-Path -LiteralPath $held)) { throw 'Staging store fault fixture is not ready.' }
    Move-Item -LiteralPath $leaf -Destination $held
    try {
        [IO.File]::WriteAllText($leaf,'Blocked disposable activity path')
        return [string](Run 'invSys.Operations.xlam' 'TestReceivingActivity.AddCheck' @($false))
    } finally {
        if (Test-Path -LiteralPath $leaf -PathType Leaf) { Remove-Item -LiteralPath $leaf }
        Move-Item -LiteralPath $held -Destination $leaf
    }
}
