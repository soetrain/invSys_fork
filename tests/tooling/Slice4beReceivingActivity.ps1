# D18: actual Receiving controls, generated disposable warehouse and independent
# source evidence. Unsaved VBA instrumentation only exposes existing handlers.
# Row values, actor credentials, raw activity and paths never enter reports.
function Get-ReceivingFixtureRows($List) {
    foreach ($row in $List.ListRows) {
        $values = @{}
        foreach ($column in $List.ListColumns) {
            $values[$column.Name] = $row.Range.Cells.Item(1,$column.Index).Value2
        }
        [pscustomobject]$values
    }
}

function Open-ReceivingEvidenceBook([string]$Path) {
    foreach ($book in $excel.Workbooks) {
        if ($book.FullName -eq $Path) { return $book }
    }
    $opened = $excel.Workbooks.Open($Path,0,$true)
    $receivingEvidenceOpened.Add($opened)
    return $opened
}

function Get-ReceivingFixtureHash([string]$Path) {
    $stream = [IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash($stream)) }
    finally { $sha.Dispose(); $stream.Dispose() }
}

function Test-ReceivingObservations($Fixture,[string[]]$Before,$Expected,[bool]$Pending) {
    $records = @()
    foreach ($path in @(Get-Slice4beActivityFiles $Fixture)) {
        if ($path -in $Before) { continue }
        $record = [IO.File]::ReadAllText($path) | ConvertFrom-Json
        if ((Get-Slice4beField $record 'ControlId') -eq 'RECEIVING_CONFIRM_WRITES') { $records += $record }
    }
    $attempts = @($records | Where-Object { (Get-Slice4beField $_ 'OutcomeCode') -eq 'REQUESTED' })
    $outcomes = @($records | Where-Object { (Get-Slice4beField $_ 'OutcomeCode') -ne 'REQUESTED' })
    $pair = $attempts.Count -eq 1 -and $outcomes.Count -eq 1
    $label = if ($Pending) { 'Pending' } else { 'Applied' }
    Check "Receiving.Activity.$label.AttemptAndOutcome" $pair
    $correlated = $false; $sources = $false; $truthful = $false
    if ($pair) {
        $first = $attempts[0]; $last = $outcomes[0]
        $correlated = (Get-Slice4beField $first 'ActivityId') -ne '' -and
            (Get-Slice4beField $first 'ActivityId') -ceq (Get-Slice4beField $last 'ActivityId') -and
            (Get-Slice4beField $first 'RecordId') -cne (Get-Slice4beField $last 'RecordId') -and
            (Get-Slice4beField $last 'WarehouseId') -ceq $Fixture.Warehouse -and
            (Get-Slice4beField $last 'UserId') -ceq 'config-reader'
        $property = $last.PSObject.Properties['SourceEventRefs']
        if ($null -ne $property) {
            $refs = @($property.Value)
            $ids = @($refs | ForEach-Object { Get-Slice4beField $_ 'EventId' })
            $sources = $ids.Count -eq $Expected.Count
            foreach ($item in $Expected) { $sources = $sources -and ($item.EventId -cin $ids) }
            $sources = $sources -and @($ids | Select-Object -Unique).Count -eq $Expected.Count
        }
        $truthful = (Get-Slice4beField $first 'DataEffect') -ceq 'Unknown'
        if ($Pending) {
            $truthful = $truthful -and (Get-Slice4beField $last 'DataEffect') -ceq 'Unknown' -and
                (Get-Slice4beField $last 'OutcomeCode') -cnotin @('APPLIED','COMPLETED')
        } else {
            # A known applied result may report Changed; otherwise keep explicit
            # uncertainty until owning per-event evidence is available.
            $truthful = $truthful -and (Get-Slice4beField $last 'DataEffect') -cin @('Changed','Unknown')
        }
    }
    Check "Receiving.Activity.$label.StableCorrelation" $correlated
    Check "Receiving.Activity.$label.EveryExactSourceEvent" $sources
    Check "Receiving.Activity.$label.NoInferredApplication" $truthful
}

function Test-Slice4beReceivingActivity {
    $receivingEvidenceOpened = [Collections.Generic.List[object]]::new()
    # Install fault instrumentation before creating any live form or capturing
    # authentication. Editing a referenced project resets existing VBA globals.
    $gate = $packages['invSys.Core.xlam'].VBProject.VBComponents.Add(1)
    $gate.Name = 'TestReceivingActivityGate'
    $gate.CodeModule.AddFromString(@'
Option Explicit
Public Pending As Boolean
Public Sub SetPending(ByVal value As Boolean)
    Pending = value
End Sub
'@)
    $bridge = $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modOperationsPrimitiveBridge').CodeModule
    $procedureStart = $bridge.ProcStartLine('RunBatchAndRefreshOperatorWorkbook',0)
    $count = $bridge.ProcCountLines('RunBatchAndRefreshOperatorWorkbook',0)
    $savedProcedure = $bridge.Lines($procedureStart,$count)
    $changed = $savedProcedure -replace '(?m)^(\s*)Dim wb As Workbook', ('$1Dim wb As Workbook' + "`r`n" + '    If TestReceivingActivityGate.Pending Then report = "Fixture application withheld.": Exit Function')
    if ($changed -eq $savedProcedure) { throw 'Pending fixture bridge seam not found.' }
    $bridge.DeleteLines($procedureStart,$count)
    $bridge.InsertLines($procedureStart,$changed)
    [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride' @((Join-Path $repo 'deploy/current/templates')))
    [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetLocalOperatorRootOverrideForAutomation' @((Join-Path $runRoot 'operators')))
    $fixture = NewFixture 'receiving-activity'
    SelectTarget $fixture
    $seeded = [string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($fixture.Warehouse,'S1','config-admin'))
    if (-not $seeded.StartsWith('OK|')) { throw 'Receiving fixture Seed failed.' }
    SelectTarget $fixture 'config-reader'
    $formCode = $packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmReceiving').CodeModule
    $formCode.AddFromString(@'
Public Function ActivityTestStage(ByVal operatorWb As Workbook) As Boolean
    Dim report As String, i As Long, lo As ListObject
    SetOperatorWorkbook operatorWb
    InitializeFromReceiving
    mBtnRefresh_Click
    mTabs.Value = 0
    ApplyReceivingTab
    If mLstReceiveItems.ListCount = 0 Then Exit Function
    For i = 1 To 2
        mLstReceiveItems.ListIndex = 0
        mLstReceiveItems_Click
        mTxtRef.Value = "ACTIVITY-PRIVATE-REFERENCE-" & CStr(i)
        mTxtQty.Value = CStr(i)
        mTxtReceiveLocation.Value = "ACTIVITY-PRIVATE-LOCATION"
        mCboCondition.Value = "GOOD"
        mBtnAdd_Click
    Next i
    Set lo = operatorWb.Worksheets("ReceivedTally").ListObjects("ReceivedTally")
    ActivityTestStage = (lo.ListRows.Count = 2)
End Function
'@)
    $helper = $packages['invSys.Operations.xlam'].VBProject.VBComponents.Add(1)
    $helper.Name = 'TestReceivingActivity'
    $helper.CodeModule.AddFromString(@'
Option Explicit
Private mForm As frmReceiving
Public Function Stage(ByVal workbookName As String) As Boolean
    Set mForm = New frmReceiving
    Stage = mForm.ActivityTestStage(Application.Workbooks(workbookName))
End Function
Public Function Confirm(ByVal workbookName As String, ByVal otherName As String) As String
    Application.Workbooks(otherName).Activate
    Confirm = mForm.TestRunConfirmWritesActionForWorkbook(Application.Workbooks(workbookName))
    Unload mForm
    Set mForm = Nothing
End Function
'@)
    foreach ($pending in @($false,$true)) {
        $label = if ($pending) { 'Pending' } else { 'Applied' }
        Write-Output "Receiving fixture: $label"
        $operator = $excel.Workbooks.Add()
        $operator.SaveAs((Join-Path $runRoot ("receiving-$label.xlsm")),52)
        $other = $excel.Workbooks.Add()
        $other.Worksheets.Item(1).Cells.Item(1,1).Value2 = 'unrelated workbook sentinel'
        $other.SaveAs((Join-Path $runRoot ("other-$label.xlsm")),52)
        $otherHash = Get-ReceivingFixtureHash $other.FullName
        $staged = [bool](Run 'invSys.Operations.xlam' 'TestReceivingActivity.Stage' @($operator.Name))
        if (-not $staged) { throw 'Receiving fixture actual Add handlers did not stage two rows.' }
        Write-Output "Receiving fixture: $label staged through Add"
        $staging = Table $operator 'ReceivedTally'
        $expected = @(Get-ReceivingFixtureRows $staging)
        $keys = @($expected | ForEach-Object { $_.System_Key } | Select-Object -Unique)
        $ids = @($expected | ForEach-Object { $_.EventId } | Select-Object -Unique)
        if ($keys.Count -ne 2 -or $ids.Count -ne 2 -or '' -in $keys -or '' -in $ids) { throw 'Receiving fixture identity generation failed.' }
        $extra = $staging.ListColumns.Add(); $extra.Name = 'Operator Extra'
        [void](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.SetPending' @($pending))
        $before = @(Get-Slice4beActivityFiles $fixture)
        $otherSavedBefore = $other.Saved
        $status = [string](Run 'invSys.Operations.xlam' 'TestReceivingActivity.Confirm' @($operator.Name,$other.Name))
        Write-Output "Receiving fixture: $label Confirm handler returned"
        $inboxPath = [string](Run 'invSys.Core.xlam' 'modRoleEventWriter.ResolveInboxWorkbookPath' @('RECEIVE',$fixture.Warehouse,'S1',''))
        if (-not [IO.Path]::GetFullPath($inboxPath).StartsWith($fixture.Root.TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)) {
            throw 'Receiving fixture inbox path escaped its generated runtime.'
        }
        $inbox = Open-ReceivingEvidenceBook $inboxPath
        $authority = Open-ReceivingEvidenceBook (Join-Path $fixture.Root ($fixture.Warehouse+'.invSys.Data.Inventory.xlsb'))
        $inboxRows = @(Get-ReceivingFixtureRows (Table $inbox 'tblInboxReceive'))
        $appliedRows = @(Get-ReceivingFixtureRows (Table $authority 'tblAppliedEvents'))
        $logRows = @(Get-ReceivingFixtureRows (Table $authority 'tblInventoryLog'))
        $queued = @($inboxRows | Where-Object { $_.EventID -cin $ids })
        $applied = @($appliedRows | Where-Object { $_.EventID -cin $ids })
        $logged = @($logRows | Where-Object { $_.EventID -cin $ids })
        $business = $queued.Count -eq 2
        if ($pending) {
            $business = $business -and $status.StartsWith('Succeeded=False') -and $applied.Count -eq 0 -and $logged.Count -eq 0 -and $staging.ListRows.Count -eq 2
        } else {
            $business = $business -and $status.StartsWith('Succeeded=True') -and $applied.Count -eq 2 -and $logged.Count -eq 2 -and $staging.ListRows.Count -eq 0
            foreach ($item in $expected) {
                $matching = @($logged | Where-Object { $_.EventID -ceq $item.EventId -and $_.System_Key -ceq $item.System_Key -and $_.QtyDelta -eq $item.QUANTITY })
                $business = $business -and $matching.Count -eq 1
            }
        }
        Check "Receiving.$label.IndependentBusinessEvidence" $business
        if (-not $business) { throw 'Receiving fixture business evidence did not establish the intended scenario.' }
        Check "Receiving.$label.CapturedWorkbook" ($status.Contains('BoundWorkbook='+$operator.Name))
        Check "Receiving.$label.QuietUi" ($status.Contains('QuietDuring=True') -and $status.Contains('QuietRestored=True'))
        Check "Receiving.$label.OtherWorkbookBytes" ($otherHash -eq (Get-ReceivingFixtureHash $other.FullName))
        Check "Receiving.$label.OtherWorkbookContent" ($other.Worksheets.Count -eq 1 -and $other.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'unrelated workbook sentinel' -and $other.Worksheets.Item(1).ListObjects.Count -eq 0)
        Write-Output "Receiving fixture: $label other workbook Saved before=$otherSavedBefore after=$($other.Saved)"
        Check "Receiving.$label.UnknownColumnPreserved" ($staging.ListColumns.Item('Operator Extra').Name -ceq 'Operator Extra')
        Test-ReceivingObservations $fixture $before $expected $pending
        foreach ($book in $receivingEvidenceOpened) { $book.Close($false) }
        $receivingEvidenceOpened.Clear()
        $operator.Close($false); $other.Close($false)
    }
    [void](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.SetPending' @($false))
}
