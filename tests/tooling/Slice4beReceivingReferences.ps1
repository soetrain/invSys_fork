# Supplemental wire validation: synthetic activity records are never presented
# as captured operator evidence. Real handler scenarios protect their producers.
function Test-ReceivingReferenceRead($Fixture,$TemplateFixture,$Expected) {
    $templatePath = @(Get-Slice4beActivityFiles $TemplateFixture)[0]
    $record = [IO.File]::ReadAllText($templatePath) | ConvertFrom-Json
    $record.PSObject.Properties.Remove('ContentSha256')
    $record.RecordId = [guid]::NewGuid().ToString()
    $record.ActivityId = [guid]::NewGuid().ToString()
    $record.WarehouseId = $Fixture.Warehouse
    $record.CatalogVersion = 2
    $record.UserId = 'config-reader'
    $record.ControlId = 'RECEIVING_CONFIRM_WRITES'
    $record.OwnerId = 'RECEIVING_WORKFLOW'
    $record.SourceRole = 'Receiving'
    $record.Caption = 'Confirm Writes'
    $record.Surface = 'Operations > Receiving'
    $record.EventCode = 'RECEIVE_CONFIRM_CONFIRMED'
    $record.OutcomeCode = 'CONFIRMED'
    $record.Severity = 'Info'
    $record.DataEffect = 'Unknown'
    $record.UserMessage = 'Receiving command completed; Domain application not asserted.'
    $record.NextStep = 'Inspect published outcomes for every related event.'
    $record.SourceEventRefs = @($Expected | ForEach-Object {
        [pscustomobject][ordered]@{WarehouseId=$Fixture.Warehouse;SourceKind='Inventory';EventId=[string]$_.EventId;SubmissionState='Submitted'}
    })
    $root = Join-Path $Fixture.Root ('Training\Activity\'+$Fixture.Warehouse)
    New-Item -ItemType Directory -Path $root -Force | Out-Null
    $path = Join-Path $root ($record.RecordId+'.json')
    $body = $record | ConvertTo-Json -Depth 8 -Compress
    try {
        $legacy = [IO.File]::ReadAllText($templatePath) | ConvertFrom-Json
        $legacy.PSObject.Properties.Remove('ContentSha256')
        $legacy.RecordId = $record.RecordId
        $legacy.WarehouseId = $Fixture.Warehouse
        $legacy.CatalogVersion = 1
        Save-ActivityFixtureBody $path ($legacy | ConvertTo-Json -Depth 8 -Compress)
        Check 'Receiving.Reference.SupportedCatalogOneStillReadable' ((Get-ActivityRead $record.RecordId).StartsWith('OK|'))
        Save-ActivityFixtureBody $path $body
        $configBefore = Get-ReceivingFixtureHash $Fixture.Config
        Check 'Receiving.Reference.ValidSupportedRecordRead' ((Get-ActivityRead $record.RecordId).StartsWith('OK|'))
        Check 'Receiving.Reference.PolicyReadPreservesConfigBytes' ($configBefore -ceq (Get-ReceivingFixtureHash $Fixture.Config))
        # Supplemental delivery checks: synthetic gateway calls are not operator evidence.
        $context = [string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
        $action = [string](Run 'invSys.Core.xlam' 'modActivity.BeginAction' @('RECEIVING_CONFIRM_WRITES',$context))
        $references = ConvertTo-Json -InputObject @($record.SourceEventRefs) -Depth 5 -Compress
        $finished = [bool](Run 'invSys.Core.xlam' 'modActivity.FinishAction' @($action,'CONFIRMED','',$references))
        Check 'Receiving.Reference.SupportedCompletionAccepted' ($action -ne '' -and $finished)
        $snapshot = @{}; foreach ($p in @(Get-Slice4beActivityFiles $Fixture)) { $snapshot[$p] = Get-ReceivingFixtureHash $p }
        $again = [bool](Run 'invSys.Core.xlam' 'modActivity.FinishAction' @($action,'CONFIRMED','',$references))
        Check 'Receiving.Reference.SameCompletionIdempotent' $again
        $changedReferences = $references.Replace([string]$record.SourceEventRefs[0].EventId,'DIFFERENT-VALID-EVENT')
        Check 'Receiving.Reference.ConflictingSourceRejected' (-not [bool](Run 'invSys.Core.xlam' 'modActivity.FinishAction' @($action,'CONFIRMED','',$changedReferences)))
        $unchanged = $snapshot.Count -eq @(Get-Slice4beActivityFiles $Fixture).Count
        foreach ($p in $snapshot.Keys) { $unchanged = $unchanged -and $snapshot[$p] -ceq (Get-ReceivingFixtureHash $p) }
        Check 'Receiving.Reference.RepeatedDeliveryPreservesEveryByte' $unchanged
        $config = $excel.Workbooks.Open($Fixture.Config,0,$false)
        try {
            $cell = (Table $config 'tblWarehouseConfig').ListColumns.Item('WarehouseName').DataBodyRange.Cells.Item(1,1)
            $cell.Value2 = 'unsaved policy reader fixture'
            $read = Get-ActivityRead $record.RecordId
            Check 'Receiving.Reference.DirtyOpenConfigRejectedAndPreserved' ($read.StartsWith('UNAVAILABLE|') -and -not $config.Saved -and $cell.Value2 -ceq 'unsaved policy reader fixture')
        } finally { $config.Close($false) }
        $cases = @(
            @('CrossWarehouse', { param($r) $r.SourceEventRefs[0].WarehouseId='OTHER' }),
            @('UnknownSource', { param($r) $r.SourceEventRefs[0].SourceKind='Unknown' }),
            @('UnknownState', { param($r) $r.SourceEventRefs[0].SubmissionState='Applied' }),
            @('UnknownField', { param($r) $r.SourceEventRefs[0] | Add-Member -NotePropertyName RawPayload -NotePropertyValue 'forbidden' }),
            @('DuplicateReference', { param($r) $r.SourceEventRefs=@($r.SourceEventRefs[0],$r.SourceEventRefs[0]) }),
            @('PathIdentity', { param($r) $r.SourceEventRefs[0].EventId='C:\forbidden\event' }),
            @('NumericIdentity', { param($r) $r.SourceEventRefs[0].EventId=7 }),
            @('MissingReferences', { param($r) $r.SourceEventRefs=@() }),
            @('UncertainConfirmation', { param($r) $r.SourceEventRefs[0].SubmissionState='Unknown' }),
            @('UnsupportedCatalog', { param($r) $r.CatalogVersion=999 }),
            @('ControlAbsentFromCatalog', { param($r) $r.CatalogVersion=1 })
        )
        foreach ($case in $cases) {
            $changed = $body | ConvertFrom-Json
            & $case[1] $changed
            Save-ActivityFixtureBody $path ($changed | ConvertTo-Json -Depth 8 -Compress)
            Check ('Receiving.Reference.Reject'+$case[0]) ((Get-ActivityRead $record.RecordId).StartsWith('UNAVAILABLE|'))
        }
    } finally {
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    }
}
