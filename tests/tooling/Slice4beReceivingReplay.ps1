# B0 entry RED uses a real successful Receiving recording and authored guide.
# It does not manufacture a journal/profile or mistake original success for replay.
function Install-ReceivingReplayProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    # Deliver the existing input handlers, never set the provenance token or
    # manufacture an activity. Native input behavior has its own packaged gate.
    $project.VBComponents.Item('cReceivingSelectionInput').CodeModule.AddFromString(@'
Public Sub ReplaySelectionDownForTest()
    mList_MouseDown 1, 0, 0, 0
End Sub
Public Sub ReplaySelectionUpForTest()
    mList_MouseUp 1, 0, 0, 0
End Sub
'@)
    $ribbon=$project.VBComponents.Add(2);$ribbon.Name='TestReceivingReplayRibbon'
    $ribbon.CodeModule.AddFromString(@'
Option Explicit
Implements Office.IRibbonControl
Private Property Get IRibbonControl_Id() As String
    IRibbonControl_Id = "btnOperationsReceivingForm"
End Property
Private Property Get IRibbonControl_Tag() As String
    IRibbonControl_Tag = ""
End Property
Private Property Get IRibbonControl_Context() As Object
    Set IRibbonControl_Context = Application.ActiveWindow
End Property
'@)
    $project.VBComponents.Item('frmReceiving').CodeModule.AddFromString(@'
Public Function ReplaySourceActionForTest(ByVal action As String) As String
    Dim lo As ListObject, before As Long
    Dim inputState As cReceivingSelectionInput
    Set lo = mOperatorWorkbook.Worksheets("ReceivedTally").ListObjects("ReceivedTally")
    Select Case action
        Case "Refresh": mBtnRefresh_Click
        Case "Clear": mBtnClear_Click
        Case "Add"
            If mLstReceiveItems.ListCount = 0 Then ReplaySourceActionForTest = "NO_ITEMS": Exit Function
            before = lo.ListRows.Count
            mLstReceiveItems.ListIndex = -1
            Set inputState = mNavigationInputs("lstReceiveItems")
            inputState.ReplaySelectionDownForTest
            mLstReceiveItems.ListIndex = 0
            inputState.ReplaySelectionUpForTest
            mTxtRef.Value = "B0-ORIGINAL-REFERENCE"
            mTxtQty.Value = "1"
            mTxtReceiveLocation.Value = "B0-ORIGINAL-LOCATION"
            mCboCondition.Value = "GOOD"
            mBtnAdd_Click
            ReplaySourceActionForTest = CStr(lo.ListRows.Count = before + 1): Exit Function
        Case "Confirm"
            mBtnConfirm_Click
            ReplaySourceActionForTest = CStr(modTS_Received.LastConfirmWritesSucceeded()): Exit Function
    End Select
    ReplaySourceActionForTest = CStr(mLstReceiveItems.ListCount)
End Function
'@)
    $project.VBComponents.Item('modTS_Received').CodeModule.AddFromString(@'
Public Function ReplaySourceForTest(ByVal action As String) As String
    Dim control As New TestReceivingReplayRibbon
    Select Case action
        Case "Launch"
            modRibbonGenerated.RibbonOnActionOperations control
            ReplaySourceForTest = mReceivingLauncherWorkbookName
        Case "Close"
            If Not mReceivingLauncherForm Is Nothing Then Unload mReceivingLauncherForm
            Set mReceivingLauncherForm = Nothing
        Case Else
            If mReceivingLauncherForm Is Nothing Then Err.Raise 5, , "Missing Receiving source fixture."
            ReplaySourceForTest = mReceivingLauncherForm.ReplaySourceActionForTest(action)
    End Select
End Function
'@)
}

function Test-Slice4beReceivingReplay($Fixture,$Other) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beReceivingActivity.ps1')
    $receivingEvidenceOpened=[Collections.Generic.List[object]]::new()
    function Deliver([string]$Form,[string]$Name,[string]$Action,[string]$Value='') {
        if((BoundControl $Name $Action $Value $Form) -cne 'DELIVERED'){throw 'Existing authoring handler unavailable; not replay RED.'}
    }
    CloseRecordingViewer
    SelectTarget $Fixture 'config-admin'
    $config=$excel.Workbooks.Open($Fixture.Config,0,$true)
    try {
        $warehouse=Table $config 'tblWarehouseConfig'
        $purpose=[string]$warehouse.DataBodyRange.Cells.Item(1,$warehouse.ListColumns.Item('WarehousePurpose').Index).Value2
    } finally {$config.Close($false)}
    if($purpose -cne 'Training'){throw 'B0 requires an actually generated Training warehouse.'}
    Check 'ReplayPrerequisite.GeneratedTrainingRuntime' $true
    $allowed=Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-admin',$Fixture.Warehouse,'S1')
    if($allowed -isnot [bool] -or -not $allowed){throw 'Guide-author capability fixture missing.'}
    SetRecordingPolicy $true
    OpenRecordingViewer
    $before=ActivityPins
    if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'Existing Start Recording handler unavailable.'}
    $operator=$null
    try {
        $operatorName=[string](Run 'invSys.Operations.xlam' 'modTS_Received.ReplaySourceForTest' @('Launch'))
        $operator=$excel.Workbooks.Item($operatorName)
        $staging=Table $operator 'ReceivedTally'
        if($staging.ListRows.Count -ne 0){throw 'Receiving source staging must begin empty.'}
        $custom=$staging.ListColumns.Add();$custom.Name='B0 Custom Display'
        $prepared=[string](Run 'invSys.Operations.xlam' 'modTS_Received.ReplaySourceForTest' @('Refresh'))
        if($prepared -notmatch '^[1-9][0-9]*$'){throw 'Actual Refresh could not prepare source items.'}
        [void](Run 'invSys.Operations.xlam' 'modTS_Received.ReplaySourceForTest' @('Clear'))
        if((Run 'invSys.Operations.xlam' 'modTS_Received.ReplaySourceForTest' @('Add')) -cne 'True'){throw 'Actual source Add failed.'}
        $eventId=[string]$staging.DataBodyRange.Cells.Item(1,$staging.ListColumns.Item('EventId').Index).Value2
        $entityId=[string]$staging.DataBodyRange.Cells.Item(1,$staging.ListColumns.Item('System_Key').Index).Value2
        if($eventId -ceq '' -or $entityId -ceq ''){throw 'Owning Add did not supply exact source identities.'}
        if((Run 'invSys.Operations.xlam' 'modTS_Received.ReplaySourceForTest' @('Confirm')) -cne 'True'){throw 'Actual source Confirm failed; not replay RED.'}
        Check 'ReplayPrerequisite.SuccessfulConfirmClearsStaging' ($staging.ListRows.Count -eq 0)
        Check 'ReplayPrerequisite.CustomHeaderPreserved' (@($staging.ListColumns|Where-Object Name -CEQ 'B0 Custom Display').Count -eq 1)
        $authority=Open-ReceivingEvidenceBook (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb'))
        try {
            $applied=@(Get-ReceivingFixtureRows (Table $authority 'tblAppliedEvents')|Where-Object EventID -CEQ $eventId)
            $logged=@(Get-ReceivingFixtureRows (Table $authority 'tblInventoryLog')|Where-Object {$_.EventID -ceq $eventId -and $_.System_Key -ceq $entityId -and $_.QtyDelta -eq 1})
            Check 'ReplayPrerequisite.OriginalExactEntityApplied' ($applied.Count -eq 1 -and $logged.Count -eq 1)
        } finally {$authority.Close($false)}
        if((RecordingControl 'Stop Recording' 'Click') -cne 'DELIVERED'){throw 'Existing Stop Recording handler unavailable.'}
        $observations=@(foreach($file in Get-ChildItem -LiteralPath $activityRoot -File -Filter '*.json'){
            if(-not $before.ContainsKey($file.Name)){Get-Content -Raw -LiteralPath $file.FullName|ConvertFrom-Json}
        })
        $confirm=@($observations|Where-Object {$_.ControlId -ceq 'RECEIVING_CONFIRM_WRITES' -and $_.OutcomeCode -ceq 'CONFIRMED'})
        if($confirm.Count -ne 1){throw 'Source confirmation lacks a unique positive owner observation.'}
        $sequence=[string]$confirm[0].SequenceId
        $journal=@(RecordingJournal $sequence|Sort-Object Version)
        if($journal.Count -lt 2 -or $journal[-1].RecordType -cne 'Close' -or -not (JournalChain $sequence $journal.Count)){throw 'Recorded Receiving journal is not intact.'}
        $source=$journal[-1]
        $refs=@($confirm[0].SourceEventRefs)
        Check 'ReplayPrerequisite.OriginalExactSourceEvent' ($refs.Count -eq 1 -and $refs[0].EventId -ceq $eventId -and $refs[0].WarehouseId -ceq $Fixture.Warehouse)
        Check 'ReplayPrerequisite.InputsRemainRedacted' (($source.Observations|ConvertTo-Json -Depth 12 -Compress) -notmatch 'B0-ORIGINAL-REFERENCE|B0-ORIGINAL-LOCATION')
        if((BoundLibrary 'Open') -cne 'DELIVERED' -or (BoundLibrary 'Select' ([string]$source.ActionPathId)) -cne 'SELECTED'){throw 'Recorded source cannot be selected.'}
        $guideRoot=Join-Path $journalRoot 'Guides';$guideBefore=BoundPins $guideRoot
        Deliver 'frmActionPaths' 'btnCreateGuide' 'Click'
        Deliver 'frmActionPathGuide' 'txtGuideName' 'Write' 'Receive one training item and prove application'
        Deliver 'frmActionPathGuide' 'txtGuideInstructions' 'Write' 'Open Receiving, refresh and clear staging, select an item, add the training receipt, then confirm writes. Verify the fresh source event.'
        Deliver 'frmActionPathGuide' 'btnGuideExpectedConclusion' 'Click'
        Deliver 'frmActionPathExpectation' 'cboExpectedControl' 'Write' 'RECEIVING_CONFIRM_WRITES'
        Deliver 'frmActionPathExpectation' 'cboExpectedOutcome' 'Write' 'CONFIRMED'
        Deliver 'frmActionPathExpectation' 'chkExpectedRetry' 'Check' 'False'
        Deliver 'frmActionPathExpectation' 'btnAddExpectedStep' 'Click'
        if((BoundControl 'cboTerminalStep' 'Select' '0' 'frmActionPathExpectation') -cne 'SELECTED'){throw 'Guide terminal step unavailable.'}
        Deliver 'frmActionPathExpectation' 'cboTerminalKind' 'Write' 'SourceEventsApplied'
        Deliver 'frmActionPathExpectation' 'btnUseExpectation' 'Click'
        Deliver 'frmActionPathGuide' 'btnSaveGuide' 'Click'
        $newGuides=@(Get-ChildItem -LiteralPath $guideRoot -File -Filter '*.json'|Where-Object {-not $guideBefore.ContainsKey($_.FullName)})
        if($newGuides.Count -ne 1){throw 'Actual Guide Save did not create exactly one guide.'}
        $guide=ReadGuideExpectationRecord $newGuides[0].FullName
        if($null -eq $guide){throw 'Guide integrity validation failed.'}
        Check 'ReplayPrerequisite.GuideBindsActualRecording' ($guide.SourceRun.ActionPathId -ceq $source.ActionPathId -and $guide.SourceRun.ContentSha256 -ceq $source.ContentSha256)
        Check 'ReplayPrerequisite.AuthoredAppliedEventConclusion' ($guide.ExpectedConclusion.TerminalKind -ceq 'SourceEventsApplied' -and @($guide.ExpectedConclusion.Steps).Count -eq 1)
        $guideControls=@($guide.Steps|ForEach-Object ControlId)
        foreach($control in @('RECEIVING_OPEN','RECEIVING_REFRESH','RECEIVING_CLEAR','RECEIVING_SELECT_ITEM','RECEIVING_ADD_SELECTED','RECEIVING_CONFIRM_WRITES')){
            if($control -cnotin $guideControls){throw ('Source guide lacks deliberate action '+$control+'.')}
        }
        if(@($results|Where-Object {$_.Check -like 'ReplayPrerequisite.*' -and -not $_.Passed}).Count){throw 'Receiving replay prerequisites failed; no product RED established.'}
        Deliver 'frmActionPathGuide' 'btnCancelGuide' 'Click'
        BoundOpen;BoundSelect $guide
        $original=BoundPins $journalRoot;$activity=ActivityPins
        Check 'ReceivingReplay.ConfigureExecutionAvailable' ((BoundControl 'btnConfigureExecution' 'State') -ceq 'True|True')
        $opened=(BoundControl 'btnConfigureExecution' 'Click') -ceq 'DELIVERED'
        Check 'ReceivingReplay.ConfigureOpensEditor' ($opened -and (BoundControl '' 'Count' '' 'frmActionPathExecution') -ceq '1')
        $binding=BoundControl 'lblExecutionGuide' 'Label' '' 'frmActionPathExecution'
        Check 'ReceivingReplay.EditorBindsExactGuide' ($opened -and $binding.Contains([string]$guide.ActionPathId) -and $binding.Contains([string]$guide.ContentSha256))
        Check 'ReceivingReplay.RunHowToControlPresent' ((BoundControl 'btnRunHowTo' 'State') -match '^True\|')
        Check 'ReceivingReplay.SetupEntryDoesNotExecute' ((BoundSame $original (BoundPins $journalRoot)) -and (PinsRetained $activity) -and (ActivityPins).Count -eq $activity.Count -and $staging.ListRows.Count -eq 0)
        . (Join-Path $PSScriptRoot 'Slice4beExecutionProfile.ps1')
        Test-ReceivingExecutionProfile $guide $staging $journalRoot (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')) $Fixture $Other
        [pscustomobject]@{SourceGuideCreated=$true;SourceSteps=$guide.Steps.Count;SourceExactEventObserved=$true;ReplayExecuted=$false;FreshReplayProof=$false;B0Accepted=$false}|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'receiving-replay-scope.json')
    } finally {
        CloseRecordingViewer
        [void](Run 'invSys.Operations.xlam' 'modTS_Received.ReplaySourceForTest' @('Close'))
    }
}
