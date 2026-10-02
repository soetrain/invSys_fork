# D18 action-entry suppression through the real packaged completion handler.
function Install-ProductionCompleteEntryProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $start=$form.ProcStartLine('mBtnManagerApplyOutput_Click',0)
    $body=$form.ProcBodyLine('mBtnManagerApplyOutput_Click',0)
    if($body -lt $start){throw 'Completion handler body unavailable; not product RED.'}
    $form.InsertLines($body+1,'    TestProductionDesigner.CompleteEntryHit')
    $start=$form.ProcStartLine('CompleteProductionRun',0);$end=$start+$form.ProcCountLines('CompleteProductionRun',0)
    $anchor='ShowPersistencePending "Saving the reusable Process run to the warehouse server..."'
    $hits=@(for($line=$start;$line -lt $end;$line++){if($form.Lines($line,1).Trim() -ieq $anchor){$line}})
    if($hits.Count -ne 1){throw 'Completion pending-yield anchor changed; not product RED.'}
    $form.InsertLines($hits[0]+1,'    TestProductionDesigner.CompleteEntryPending')
    $form.InsertLines(1,'Private mCompleteEntryGuardsRestored As Boolean')
    $form.AddFromString(@'
Public Function CompleteEntryActForTest(ByVal guard As String) As Boolean
    Dim priorLoading As Boolean, priorBusy As Boolean, entryLoading As Boolean, entryBusy As Boolean
    On Error GoTo Failed
    priorLoading = mLoading: priorBusy = mDesignerActionInProgress
    mCompleteEntryGuardsRestored = False
    If guard = "Loading" Then mLoading = True
    If guard = "Busy" Then mDesignerActionInProgress = True
    entryLoading = mLoading: entryBusy = mDesignerActionInProgress
    mBtnManagerApplyOutput_Click
    mCompleteEntryGuardsRestored = (mLoading = entryLoading And mDesignerActionInProgress = entryBusy)
    CompleteEntryActForTest = True
Failed:
    mLoading = priorLoading: mDesignerActionInProgress = priorBusy
End Function
Public Function CompleteEntryNestedForTest() As Boolean
    On Error GoTo Failed
    mBtnManagerApplyOutput_Click
    CompleteEntryNestedForTest = True
Failed:
End Function
Public Function CompleteEntryGuardsForTest() As Boolean
    CompleteEntryGuardsForTest = mCompleteEntryGuardsRestored
End Function
Public Function CompleteEntrySuccessMessageForTest() As Boolean
    CompleteEntrySuccessMessageForTest = _
        InStr(1, mTxtStatus.Text, "Production batch completed and persisted.", vbBinaryCompare) = 1
End Function
'@)
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,@'
Private mCompleteEntryCount As Long, mCompleteEntryArm As Boolean, mCompleteEntryReached As Boolean
Private mCompleteEntryNestedReturn As Boolean, mCompleteEntryOwnerSame As Boolean
Private mCompleteEntryProjectionSame As Boolean, mCompleteEntryMessageSame As Boolean
'@)
    $adapter.AddFromString(@'
Public Sub CompleteEntryReset(ByVal nested As Boolean)
    mCompleteEntryCount = 0: mCompleteEntryArm = nested: mCompleteEntryReached = False
    mCompleteEntryNestedReturn = False: mCompleteEntryOwnerSame = False
    mCompleteEntryProjectionSame = False: mCompleteEntryMessageSame = False
End Sub
Public Sub CompleteEntryHit()
    mCompleteEntryCount = mCompleteEntryCount + 1
End Sub
Public Sub CompleteEntryPending()
    Dim ownerState As String, projection As String, status As String
    If Not mCompleteEntryArm Then Exit Sub
    mCompleteEntryArm = False: mCompleteEntryReached = True
    ownerState = modProductionReusableRun.RunLocalStateForTest()
    projection = mForm.CheckYieldProjectionForTest(): status = mForm.TestStatusText()
    mCompleteEntryNestedReturn = mForm.CompleteEntryNestedForTest()
    mCompleteEntryOwnerSame = (modProductionReusableRun.RunLocalStateForTest() = ownerState)
    mCompleteEntryProjectionSame = (mForm.CheckYieldProjectionForTest() = projection)
    mCompleteEntryMessageSame = (mForm.TestStatusText() = status)
End Sub
Public Function CompleteEntryAct(ByVal guard As String) As Boolean
    CompleteEntryAct = mForm.CompleteEntryActForTest(guard)
End Function
Public Function CompleteEntryCount() As Long
    CompleteEntryCount = mCompleteEntryCount
End Function
Public Function CompleteEntryProjection() As String
    CompleteEntryProjection = mForm.CheckYieldProjectionForTest()
End Function
Public Function CompleteEntryStatus() As String
    CompleteEntryStatus = mForm.TestStatusText()
End Function
Public Function CompleteEntryFact(ByVal fact As String) As Boolean
    Select Case fact
        Case "Reached": CompleteEntryFact = mCompleteEntryReached
        Case "NestedReturned": CompleteEntryFact = mCompleteEntryNestedReturn
        Case "OwnerSame": CompleteEntryFact = mCompleteEntryOwnerSame
        Case "ProjectionSame": CompleteEntryFact = mCompleteEntryProjectionSame
        Case "MessageSame": CompleteEntryFact = mCompleteEntryMessageSame
        Case "GuardsRestored": CompleteEntryFact = mForm.CompleteEntryGuardsForTest()
        Case "SuccessMessage": CompleteEntryFact = mForm.CompleteEntrySuccessMessageForTest()
    End Select
End Function
'@)
}

function Test-ProductionCompleteEntry($Fixture,$Book,$Decoy,[string]$Canary){
    $sheet=$Book.Worksheets.Item(1);$decoySheet=$Decoy.Worksheets.Item(1)
    foreach($mode in @('Loading','Busy','Nested')){
        [void](Probe 'CheckYieldReset')
        SelectTarget $Fixture 'config-producer'
        [void](Probe 'RunLocalReopen' @($Book.Name))
        if(-not [bool](Probe 'CompleteBaselinePrepare' @($true))){throw 'Completion entry prerequisite unavailable; not product RED.'}
        [void](Probe 'RunLocalShowAndCapture' @($Book.Name,'CHECK_IN'));$Decoy.Activate()
        $before=[string](Probe 'RunLocalOwnerState');$projection=[string](Probe 'CompleteEntryProjection');$status=[string](Probe 'CompleteEntryStatus')
        $nested=$mode -ceq 'Nested';$guard=if($nested){''}else{$mode}
        [void](Probe 'CompleteEntryReset' @($nested))
        $label='CompleteEntry.'+$mode
        Check ($label+'.ActualHandlerReturned') ([bool](Probe 'CompleteEntryAct' @($guard)))
        Check ($label+'.ActualHandlerEntries') ([int](Probe 'CompleteEntryCount') -eq $(if($nested){2}else{1}))
        Check ($label+'.GuardsRestored') ([bool](Probe 'CompleteEntryFact' @('GuardsRestored')))
        if($nested){
            foreach($fact in @('Reached','NestedReturned','OwnerSame','ProjectionSame','MessageSame','SuccessMessage')){
                Check ($label+'.'+$fact) ([bool](Probe 'CompleteEntryFact' @($fact)))
            }
            Check ($label+'.OnlyOneCompletionOwnerEntered') ([string](Probe 'CompleteBaselineOwners') -ceq '1|0')
            Check ($label+'.ProcessAndBatchCompleted') ([bool](Probe 'CompleteBaselineCompleted'))
            Check ($label+'.ExactAllocatedEntitiesConsumedOnce') ([bool](Owner 'CompleteBaselineBalancesForTest' @($true)))
            Check ($label+'.FreshOutputEntityQuantity') ([bool](Owner 'CompleteBaselineOutputForTest'))
        }else{
            Check ($label+'.NoCompletionOwnerEntered') ([string](Probe 'CompleteBaselineOwners') -ceq '0|0')
            Check ($label+'.OwnerStagingPreserved') ([string](Probe 'RunLocalOwnerState') -ceq $before)
            Check ($label+'.ProjectionPreserved') ([string](Probe 'CompleteEntryProjection') -ceq $projection)
            Check ($label+'.ActiveMessagePreserved') ([string](Probe 'CompleteEntryStatus') -ceq $status)
            Check ($label+'.ExactInputBalancesPreserved') ([bool](Owner 'CompleteBaselineBalancesForTest' @($false)))
        }
        Check ($label+'.CapturedBookCustomValueAndFormula') ($sheet.Cells.Item(2,1).Value2 -ceq $Canary -and $sheet.Cells.Item(2,2).Formula -ceq '=1+2')
        Check ($label+'.DecoyPreserved') ($decoySheet.Cells.Item(1,1).Value2 -ceq $Canary -and $Decoy.Worksheets.Count -eq 1)
        CaptureOwnedFormByCaptionEvidence 'Production' ('complete-entry-'+$mode.ToLowerInvariant()+'.png')
        [void](Probe 'CompleteEntryReset' @($false))
    }
}
