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

function Test-ReceivingObservations($Fixture,[string[]]$Before,$Expected,[bool]$Pending,
    [string]$CaseLabel='', [string]$Outcome='', [string]$ReferenceState='Submitted') {
    $records = @()
    $payloads = @()
    foreach ($path in @(Get-Slice4beActivityFiles $Fixture)) {
        if ($path -in $Before) { continue }
        $raw = [IO.File]::ReadAllText($path)
        $record = $raw | ConvertFrom-Json
        if ((Get-Slice4beField $record 'ControlId') -eq 'RECEIVING_CONFIRM_WRITES') { $records += $record; $payloads += $raw }
    }
    $attempts = @($records | Where-Object { (Get-Slice4beField $_ 'OutcomeCode') -eq 'REQUESTED' })
    $outcomes = @($records | Where-Object { (Get-Slice4beField $_ 'OutcomeCode') -ne 'REQUESTED' })
    $pair = $attempts.Count -eq 1 -and $outcomes.Count -eq 1
    $label = if ($CaseLabel) { $CaseLabel } elseif ($Pending) { 'Pending' } else { 'Applied' }
    Write-Output "Receiving observation counts: $label attempts=$($attempts.Count) outcomes=$($outcomes.Count)"
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
            if ($ReferenceState -eq 'Empty') { $sources = $refs.Count -eq 0 }
            else {
                $sources = $ids.Count -eq $Expected.Count
                foreach ($item in $Expected) { $sources = $sources -and ($item.EventId -cin $ids) }
                $sources = $sources -and @($ids | Select-Object -Unique).Count -eq $Expected.Count
            }
            foreach ($reference in $refs) {
                $sources = $sources -and (Get-Slice4beField $reference 'WarehouseId') -ceq $Fixture.Warehouse -and
                    (Get-Slice4beField $reference 'SourceKind') -ceq 'Inventory' -and
                    (Get-Slice4beField $reference 'SubmissionState') -ceq $ReferenceState -and
                    @($reference.PSObject.Properties).Count -eq 4
            }
        }
        $truthful = (Get-Slice4beField $first 'DataEffect') -ceq 'Unknown'
        $expectedOutcome = if ($Outcome) { $Outcome } elseif ($Pending) { 'PENDING' } else { 'CONFIRMED' }
        $expectedEffect = if ($expectedOutcome -eq 'DENIED') { 'Unchanged' } else { 'Unknown' }
        $truthful = $truthful -and (Get-Slice4beField $last 'DataEffect') -ceq $expectedEffect -and
            (Get-Slice4beField $last 'OutcomeCode') -ceq $expectedOutcome -and
            (Get-Slice4beField $last 'EventCode') -ceq ('RECEIVE_CONFIRM_'+$expectedOutcome)
    }
    Check "Receiving.Activity.$label.StableCorrelation" $correlated
    Check "Receiving.Activity.$label.EveryExactSourceEvent" $sources
    Check "Receiving.Activity.$label.NoInferredApplication" $truthful
    $redacted = $pair; $integrity = $pair
    foreach ($raw in $payloads) {
        foreach ($forbidden in @($Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,
            'ACTIVITY-PRIVATE-REFERENCE','ACTIVITY-PRIVATE-LOCATION','mBtnConfirm_Click','PinHash','Err.Description')) {
            if ($raw.IndexOf($forbidden,[StringComparison]::OrdinalIgnoreCase) -ge 0) { $redacted = $false }
        }
        foreach ($item in $Expected) {
            if ($raw.Contains([string]$item.System_Key)) { $redacted = $false }
        }
        $match = [regex]::Match($raw,'^(?<body>\{.*),"ContentSha256":"(?<hash>[a-f0-9]{64})"\}$')
        if (-not $match.Success) { $integrity = $false; continue }
        $sha = [Security.Cryptography.SHA256]::Create()
        try {
            $digest = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($match.Groups['body'].Value+'}'))).Replace('-','').ToLowerInvariant()
            $integrity = $integrity -and $digest -ceq $match.Groups['hash'].Value
        } finally { $sha.Dispose() }
    }
    Check "Receiving.Activity.$label.RedactedPayload" $redacted
    Check "Receiving.Activity.$label.ContentIntegrity" $integrity
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
Public LoseAcknowledgement As Boolean
Public RefreshFault As Boolean
Public RefreshCalls As Long
Public Sub SetRefreshFault(ByVal value As Boolean)
    RefreshFault = value
    RefreshCalls = 0
End Sub
Public Function RefreshCallCount() As Long
    RefreshCallCount = RefreshCalls
End Function
Public Sub SetPending(ByVal value As Boolean)
    Pending = value
End Sub
Public Sub SetLoseAcknowledgement(ByVal value As Boolean)
    LoseAcknowledgement = value
End Sub
Public Function PolicyStatus() As String
    Dim target As WarehouseTarget, version As Long, collect As Boolean, visible As Boolean, notice As String
    Set target = modNasConnection.GetCurrentTarget()
    PolicyStatus = CStr(modActivityPolicy.ReadPolicy(target, "RECEIVING_CONFIRM_WRITES", version, collect, visible, notice)) & ";Loaded=" & CStr(modConfig.IsLoaded()) & ";" & notice
End Function
'@)
    $bridge = $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modOperationsPrimitiveBridge').CodeModule
    $procedureStart = $bridge.ProcStartLine('RunBatchAndRefreshOperatorWorkbook',0)
    $count = $bridge.ProcCountLines('RunBatchAndRefreshOperatorWorkbook',0)
    $savedProcedure = $bridge.Lines($procedureStart,$count)
    $changed = $savedProcedure -replace '(?m)^(\s*)Dim wb As Workbook', ('$1Dim wb As Workbook' + "`r`n" + '    If TestReceivingActivityGate.Pending Then report = "Fixture application withheld.": Exit Function')
    if ($changed -eq $savedProcedure) { throw 'Pending fixture bridge seam not found.' }
    $bridge.DeleteLines($procedureStart,$count)
    $bridge.InsertLines($procedureStart,$changed)
    if ($CheckReceivingLocalActivity) { Install-ReceivingLocalCoreSeam $bridge }
    if ($CheckReceivingLifecycleActivity) { Install-ReceivingLifecycleCoreSeam $bridge $gate.CodeModule }
    # The real queue persists normally; withhold only its acknowledgement.
    # This models an uncertain response after durable submission, not a rollback.
    $writer = $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modRoleEventWriter').CodeModule
    $queueStart = $writer.ProcStartLine('QueueReceiveEventBatchServer',0)
    $queueCount = $writer.ProcCountLines('QueueReceiveEventBatchServer',0)
    $queueSource = $writer.Lines($queueStart,$queueCount)
    $queueChanged = $queueSource -replace '(?im)^[ \t]*acceptedCount[ \t]*=[ \t]*rows\.Count[ \t]*\r?$', ('    If TestReceivingActivityGate.LoseAcknowledgement Then errorMessage = "Fixture queue acknowledgement withheld.": Exit Function' + "`r`n" + '    acceptedCount = rows.Count')
    if ($queueChanged -eq $queueSource) { throw 'Queue acknowledgement fixture seam not found.' }
    $writer.DeleteLines($queueStart,$queueCount)
    $writer.InsertLines($queueStart,$queueChanged)
    [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride' @((Join-Path $repo 'deploy/current/templates')))
    [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetLocalOperatorRootOverrideForAutomation' @((Join-Path $runRoot 'operators')))
    $fixture = NewFixture 'receiving-activity'
    Write-Output 'Receiving fixture: generated; signing in setup actor'
    SelectTarget $fixture
    $configWasOpen = @($excel.Workbooks | Where-Object { $_.FullName -eq $fixture.Config }).Count -gt 0
    $seeded = [string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($fixture.Warehouse,'S1','config-admin'))
    if (-not $seeded.StartsWith('OK|')) { throw 'Receiving fixture Seed failed.' }
    Check 'Receiving.Setup.ImplicitConfigOwnershipReleased' (-not $configWasOpen -and @($excel.Workbooks | Where-Object { $_.FullName -eq $fixture.Config }).Count -eq 0)
    SelectTarget $fixture 'config-reader'
    $formCode = $packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmReceiving').CodeModule
    $formCode.AddFromString(@'
Public Function ActivityTestStage(ByVal operatorWb As Workbook, Optional ByVal disposition As Boolean = False, Optional ByVal otherWb As Workbook = Nothing) As Boolean
    Dim report As String, i As Long, lo As ListObject
    SetOperatorWorkbook operatorWb
    InitializeFromReceiving
    mBtnRefresh_Click
    If disposition Then mTabs.Value = 1 Else mTabs.Value = 0
    ApplyReceivingTab
    If mLstReceiveItems.ListCount = 0 Then Exit Function
    For i = 1 To 2
        mLstReceiveItems.ListIndex = 0
        mLstReceiveItems_Click
        mTxtRef.Value = "ACTIVITY-PRIVATE-REFERENCE-" & CStr(i)
        mTxtQty.Value = CStr(i)
        mTxtReceiveLocation.Value = "ACTIVITY-PRIVATE-LOCATION"
        mCboCondition.Value = "GOOD"
        If disposition Then
            mCboDisposition.ListIndex = i - 1
            mTxtReturnReason.Value = "ACTIVITY-PRIVATE-DISPOSITION-REASON"
        End If
        If Not otherWb Is Nothing Then otherWb.Activate
        mBtnAdd_Click
    Next i
    Set lo = operatorWb.Worksheets("ReceivedTally").ListObjects("ReceivedTally")
    ActivityTestStage = (lo.ListRows.Count = 2)
End Function
Public Function ActivityTestDirectStage(ByVal operatorWb As Workbook, ByVal disposition As Boolean) As Boolean
    Dim report As String, receiptType As String
    receiptType = "RECEIPT"
    If disposition Then receiptType = "RETURN"
    ActivityTestDirectStage = modTS_Received.StageReceivingFormItemForWorkbook( _
        operatorWb, "ACTIVITY-PRIVATE-DIRECT-REFERENCE", SelectedReceiveItemSystemKey(0), _
        NzText(mLstReceiveItems.List(0, 0)), 1, report, _
        CStr(mTxtReceiveLocation.Value), "", "GOOD", receiptType, "ACTIVITY-PRIVATE-DIRECT-REASON")
End Function
Public Function ActivityTestAddCheck(ByVal rejectInput As Boolean) As String
    mTxtRef.Value = "ACTIVITY-PRIVATE-EXTRA-REFERENCE"
    mTxtQty.Value = "1"
    If rejectInput Then mTxtQty.Value = "0"
    mBtnAdd_Click
    ActivityTestAddCheck = CStr(mTxtStatus.Value)
End Function
'@)
    $helper = $packages['invSys.Operations.xlam'].VBProject.VBComponents.Add(1)
    $helper.Name = 'TestReceivingActivity'
    $helper.CodeModule.AddFromString(@'
Option Explicit
Private mForm As frmReceiving
Public Function Stage(ByVal workbookName As String, Optional ByVal disposition As Boolean = False, Optional ByVal otherName As String = "") As Boolean
    Dim otherWb As Workbook
    If otherName <> "" Then Set otherWb = Application.Workbooks(otherName)
    Set mForm = New frmReceiving
    Stage = mForm.ActivityTestStage(Application.Workbooks(workbookName), disposition, otherWb)
End Function
Public Function DirectStage(ByVal workbookName As String, ByVal disposition As Boolean) As Boolean
    DirectStage = mForm.ActivityTestDirectStage(Application.Workbooks(workbookName), disposition)
End Function
Public Function AddCheck(ByVal rejectInput As Boolean) As String
    AddCheck = mForm.ActivityTestAddCheck(rejectInput)
End Function
Public Sub ShowForm(ByVal visible As Boolean)
    If visible Then mForm.Show 0 Else mForm.Hide
End Sub
Public Sub Reopen(ByVal workbookName As String)
    Set mForm = New frmReceiving
    mForm.SetOperatorWorkbook Application.Workbooks(workbookName)
End Sub
Public Function Confirm(ByVal workbookName As String, ByVal otherName As String, ByVal showEvidence As Boolean) As String
    Application.Workbooks(otherName).Activate
    Confirm = mForm.TestRunConfirmWritesActionForWorkbook(Application.Workbooks(workbookName))
    If showEvidence Then
        mForm.Show 0
    Else
        CloseForm
    End If
End Function
Public Sub CloseForm()
    Unload mForm
    Set mForm = Nothing
End Sub
'@)
    if ($CheckReceivingLocalActivity) { Install-ReceivingLocalFormSeams $formCode $helper.CodeModule }
    if ($CheckReceivingLifecycleActivity) { Install-ReceivingLifecycleSeams $packages['invSys.Operations.xlam'].VBProject $formCode }
    if ($ReceivingLifecycleOnly) { Test-ReceivingLifecycleActivity $fixture; return }
    foreach ($label in @('Applied','Pending','Stale','StoreFailure','Denied','Rejected','UnknownSubmission')) {
        $pending = $label -eq 'Pending'
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
        $extra.DataBodyRange.Cells.Item(1,1).Value2 = 'preserve first extra'
        $extra.DataBodyRange.Cells.Item(2,1).Value2 = 'preserve second extra'
        [void](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.SetPending' @($pending))
        [void](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.SetLoseAcknowledgement' @($label -eq 'UnknownSubmission'))
        $before = @(Get-Slice4beActivityFiles $fixture)
        $otherSavedBefore = $other.Saved
        if ($label -eq 'Stale') {
            [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
            SelectTarget $fixture 'config-reader'
        }
        $blockedLeaf = ''; $heldLeaf = ''; $movedLeaf = $false
        $authPath = Join-Path $fixture.Root ($fixture.Warehouse+'.invSys.Auth.xlsb')
        $authBefore = $null
        try {
            if ($label -eq 'Denied') {
                $authBefore = [IO.File]::ReadAllBytes($authPath)
                $auth = $excel.Workbooks.Open($authPath,0,$false)
                $caps = Table $auth 'tblCapabilities'
                $revoked = 0
                foreach ($row in $caps.ListRows) {
                    if ($row.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-reader' -and
                        $row.Range.Cells.Item(1,$caps.ListColumns.Item('Capability').Index).Value2 -ceq 'RECEIVE_POST') {
                        $row.Range.Cells.Item(1,$caps.ListColumns.Item('Status').Index).Value2 = 'Inactive'; $revoked++
                    }
                }
                $auth.Save(); $auth.Close($false)
                if ($revoked -ne 1) { throw 'Receiving denial fixture capability was not unique.' }
            }
            if ($label -eq 'Rejected') {
                $staging.ListColumns.Item('QUANTITY').DataBodyRange.Cells.Item(2,1).Value2 = -1.0
            }
            if ($label -eq 'StoreFailure') {
                $parent = Join-Path $fixture.Root 'Training\Activity'
                $blockedLeaf = Join-Path $parent $fixture.Warehouse
                $heldLeaf = $blockedLeaf+'-fixture-held'
                foreach ($path in @($blockedLeaf,$heldLeaf)) {
                    if (-not [IO.Path]::GetFullPath($path).StartsWith($fixture.Root.TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Tracking fault escaped fixture root.' }
                }
                New-Item -ItemType Directory -Path $parent -Force | Out-Null
                if (Test-Path -LiteralPath $blockedLeaf) { Move-Item -LiteralPath $blockedLeaf -Destination $heldLeaf; $movedLeaf=$true }
                [IO.File]::WriteAllText($blockedLeaf,'Blocked test activity path')
            }
            $status = [string](Run 'invSys.Operations.xlam' 'TestReceivingActivity.Confirm' @($operator.Name,$other.Name,[bool]$CaptureEvidence))
            if ($CaptureEvidence) {
                Start-Sleep -Milliseconds 300
                CaptureFormEvidence 'Receiving' ('receiving-'+$label.ToLowerInvariant()+'.png')
                [void](Run 'invSys.Operations.xlam' 'TestReceivingActivity.CloseForm')
            }
        } finally {
            if ($blockedLeaf -ne '' -and (Test-Path -LiteralPath $blockedLeaf -PathType Leaf)) { Remove-Item -LiteralPath $blockedLeaf -Force }
            if ($movedLeaf) { Move-Item -LiteralPath $heldLeaf -Destination $blockedLeaf }
            if ($null -ne $authBefore) { [IO.File]::WriteAllBytes($authPath,$authBefore) }
        }
        Write-Output "Receiving fixture: $label Confirm handler returned"
        foreach ($notice in @('source references are invalid','training record could not be saved','action context is no longer current','completion is not authorized','saved tracking policy is invalid','configuration could not be validated','action could not be recorded')) {
            if ($status.Contains($notice)) { Write-Output ("Receiving tracking diagnostic: " + $notice) }
        }
        Write-Output ('Receiving policy diagnostic: ' + [string](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.PolicyStatus'))
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
        if ($label -in @('Stale','Denied','Rejected')) {
            $business = $status.StartsWith('Succeeded=False') -and $queued.Count -eq 0 -and $applied.Count -eq 0 -and $logged.Count -eq 0 -and $staging.ListRows.Count -eq 2
        } elseif ($pending -or $label -eq 'UnknownSubmission') {
            $business = $business -and $status.StartsWith('Succeeded=False') -and $applied.Count -eq 0 -and $logged.Count -eq 0 -and $staging.ListRows.Count -eq 2
        } else {
            $business = $business -and $status.StartsWith('Succeeded=True') -and $applied.Count -eq 2 -and $logged.Count -eq 2 -and $staging.ListRows.Count -eq 0
            foreach ($item in $expected) {
                $matching = @($logged | Where-Object { $_.EventID -ceq $item.EventId -and $_.System_Key -ceq $item.System_Key -and $_.QtyDelta -eq $item.QUANTITY })
                $business = $business -and $matching.Count -eq 1
            }
        }
        Check "Receiving.$label.IndependentBusinessEvidence" $business
        if (-not $business -and $label -ne 'Stale') { throw 'Receiving fixture business evidence did not establish the intended scenario.' }
        Check "Receiving.$label.CapturedWorkbook" ($status.Contains('BoundWorkbook='+$operator.Name))
        Check "Receiving.$label.QuietUi" ($status.Contains('QuietDuring=True') -and $status.Contains('QuietRestored=True'))
        Check "Receiving.$label.OtherWorkbookBytes" ($otherHash -eq (Get-ReceivingFixtureHash $other.FullName))
        Check "Receiving.$label.OtherWorkbookContent" ($other.Worksheets.Count -eq 1 -and $other.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'unrelated workbook sentinel' -and $other.Worksheets.Item(1).ListObjects.Count -eq 0)
        Write-Output "Receiving fixture: $label other workbook Saved before=$otherSavedBefore after=$($other.Saved)"
        Check "Receiving.$label.UnknownColumnPreserved" ($staging.ListColumns.Item('Operator Extra').Name -ceq 'Operator Extra')
        if ($label -eq 'Stale') {
            Check 'Receiving.Stale.VisibleContextRejection' ($status.Contains('Session or warehouse changed'))
            Check 'Receiving.Stale.NoNewSessionAttribution' (@(Get-Slice4beActivityFiles $fixture).Count -eq $before.Count)
        } elseif ($label -eq 'StoreFailure') {
            Check 'Receiving.StoreFailure.TrackingUnavailableVisible' ($status.Contains('Tracking unavailable'))
            Check 'Receiving.StoreFailure.NoDurableLocalFallback' (@(Get-Slice4beActivityFiles $fixture).Count -eq $before.Count)
        } elseif ($label -in @('Denied','Rejected','UnknownSubmission')) {
            $outcome = @{Denied='DENIED';Rejected='REJECTED';UnknownSubmission='FAILED'}[$label]
            $referenceState = if ($label -eq 'UnknownSubmission') { 'Unknown' } else { 'Empty' }
            Test-ReceivingObservations $fixture $before $expected $false $label $outcome $referenceState
            Check "Receiving.$label.UnknownValuesPreserved" ($extra.DataBodyRange.Cells.Item(1,1).Value2 -ceq 'preserve first extra' -and $extra.DataBodyRange.Cells.Item(2,1).Value2 -ceq 'preserve second extra')
            $current = @(Get-ReceivingFixtureRows $staging)
            $identities = $current.Count -eq $expected.Count
            for ($i=0; $i -lt $current.Count; $i++) { $identities = $identities -and $current[$i].System_Key -ceq $expected[$i].System_Key -and $current[$i].EventId -ceq $expected[$i].EventId }
            Check "Receiving.$label.StagingIdentitiesPreserved" $identities
            if ($label -eq 'Rejected') {
                Check 'Receiving.Rejected.PartialValidationIsNotRollback' ($current[0].WorkflowState -cne $expected[0].WorkflowState -and $current[1].WorkflowState -ceq $expected[1].WorkflowState)
            }
            $visibleCause = @{Denied='lacks RECEIVE_POST';Rejected='Receiving validation failed';UnknownSubmission='queue acknowledgement withheld'}[$label]
            Check "Receiving.$label.VisibleOwnerFailure" ($status.Contains($visibleCause) -and -not $status.Contains('Tracking unavailable'))
        } else {
            Test-ReceivingObservations $fixture $before $expected $pending
        }
        foreach ($book in $receivingEvidenceOpened) { $book.Close($false) }
        $receivingEvidenceOpened.Clear()
        if ($label -eq 'UnknownSubmission') {
            Test-ReceivingRetryAfterUncertainSubmission $fixture $operator $other $expected $otherHash
        }
        if ($label -eq 'Denied') { SelectTarget $fixture 'config-reader' }
        $operator.Close($false); $other.Close($false)
    }
    [void](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.SetPending' @($false))
    [void](Run 'invSys.Core.xlam' 'TestReceivingActivityGate.SetLoseAcknowledgement' @($false))
    Test-ReceivingReferenceRead $fixture $a $expected
    $existingConfig = $excel.Workbooks.Open($fixture.Config,0,$false)
    $extra = (Table $existingConfig 'tblWarehouseConfig').ListColumns.Add()
    $extra.Name = 'Setup Existing Extra'; $extra.DataBodyRange.Value2 = 'preserve existing setup value'
    $existingConfig.Save()
    $ok = [bool](Run 'invSys.Core.xlam' 'modConfig.EnsureStationInbox' @($fixture.Warehouse,'S1','RECEIVE',''))
    $stillOpen = @($excel.Workbooks | Where-Object { $_.FullName -eq $fixture.Config }).Count -eq 1
    Check 'Receiving.Setup.PreExistingConfigRemainsOpen' ($ok -and $stillOpen)
    if ($stillOpen) {
        Check 'Receiving.Setup.PreExistingUnknownColumnPreserved' ((Table $existingConfig 'tblWarehouseConfig').ListColumns.Item('Setup Existing Extra').DataBodyRange.Cells.Item(1,1).Value2 -ceq 'preserve existing setup value')
        $existingConfig.Close($false)
    } else { Check 'Receiving.Setup.PreExistingUnknownColumnPreserved' $false }
    if ($CheckReceivingStagingActivity) { Test-ReceivingStagingCoverage $fixture }
    if ($CheckReceivingLocalActivity) { Test-ReceivingLocalActivity $fixture }
    if ($CheckReceivingLifecycleActivity) { Test-ReceivingLifecycleActivity $fixture }
}
