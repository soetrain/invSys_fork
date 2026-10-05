# Interrupt after the real consume processor returns; retain its exact event.
# No source records, outcomes or business results are manufactured by this probe.
function Install-ProductionCompleteSubmissionProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $owner=$project.VBComponents.Item('modProductionReusableRun').CodeModule
    $start=$owner.ProcStartLine('CompleteReusableProcess',0)
    $end=$start+$owner.ProcCountLines('CompleteReusableProcess',0)
    $returns=@();$queues=@()
    for($line=$start;$line -lt $end;$line++){
        $text=$owner.Lines($line,1).Trim()
        if($text -ieq 'AppendProcessorReport mProcessorReports, processorReport'){$returns+=$line+1}
        if($text.StartsWith('If Not modRoleEventWriter.QueuePayloadEventCurrent(',[StringComparison]::OrdinalIgnoreCase)){$queues+=$line}
    }
    if($returns.Count -ne 2 -or $queues.Count -ne 2){throw 'Complete submission anchors changed; not product RED.'}
    $edits=@(@{Line=$returns[0];Code='        TestProductionDesigner.CompleteSubmissionReturned eventId, processedNow'})
    foreach($line in $queues){$edits+=@{Line=$line;Code='    TestProductionDesigner.CompleteSubmissionBeforeQueue'}}
    foreach($edit in @($edits|Sort-Object Line -Descending)){$owner.InsertLines($edit.Line,$edit.Code)}
    $owner.AddFromString(@'
Public Function CompleteSubmissionInputForTest() As Variant
    Dim key As Variant
    If mAllocations.Count <> 1 Then Exit Function
    For Each key In mAllocations.Keys
        CompleteSubmissionInputForTest = Array(AllocationSystemKey(CStr(key)), CDbl(mAllocations(key)))
    Next key
End Function
Public Function CompleteSubmissionOutputsAbsentForTest() As Boolean
    Dim key As Variant
    If mOutputKeys.Count <> 1 Or mCompletedNodes.Count <> 0 Or mCompleted Then Exit Function
    For Each key In mOutputKeys.Keys
        If Abs(ReusableRunExactEntityQty(CStr(mOutputKeys(key)))) > 0.0000001 Then Exit Function
    Next key
    CompleteSubmissionOutputsAbsentForTest = True
End Function
'@)
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,'Private mCompleteSubmissionArmed As Boolean, mCompleteSubmissionQueues As Long, mCompleteSubmissionEvent As String')
    $adapter.AddFromString(@'
Public Sub CompleteSubmissionArm(ByVal interruption As String, ByVal authPath As String)
    mCompleteSubmissionQueues = 0: mCompleteSubmissionEvent = ""
    CheckYieldArm "ConsumeProcessed", 1, interruption, authPath
    mCompleteSubmissionArmed = True
End Sub
Public Sub CompleteSubmissionBeforeQueue()
    If mCompleteSubmissionArmed Then mCompleteSubmissionQueues = mCompleteSubmissionQueues + 1
End Sub
Public Sub CompleteSubmissionReturned(ByVal eventId As String, ByVal processed As Long)
    If Not mCompleteSubmissionArmed Then Exit Sub
    mCompleteSubmissionEvent = eventId
    CheckYieldReturned "ConsumeProcessed", (processed > 0 And eventId <> "")
End Sub
Public Function CompleteSubmissionEvent() As String
    CompleteSubmissionEvent = mCompleteSubmissionEvent
End Function
Public Function CompleteSubmissionQueues() As Long
    CompleteSubmissionQueues = mCompleteSubmissionQueues
End Function
Public Sub CompleteSubmissionReset()
    mCompleteSubmissionArmed = False
    CheckYieldReset
End Sub
'@)
}

function Test-ProductionCompleteSubmission($Fixture,$Book,$Decoy,[string]$Canary) {
    $auth=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $authBytes=[IO.File]::ReadAllBytes($auth);$authHash=(Get-FileHash -LiteralPath $auth).Hash
    try {
        foreach($interruption in @('SignedOut','Permission')) {
            try {
                [void](Probe 'CompleteSubmissionReset')
                SelectTarget $Fixture 'config-producer'
                [void](Probe 'RunLocalReopen' @($Book.Name))
                if(-not [bool](Probe 'CompleteBaselinePrepare' @($true))){throw 'Complete submission fixture unavailable; not product RED.'}
                $sourceSelection=Owner 'CompleteSubmissionInputForTest'
                if($sourceSelection -isnot [array] -or $sourceSelection.Count -ne 2 -or [string]$sourceSelection[0] -ceq '' -or [double]$sourceSelection[1] -le 0){throw 'Exact staged consume input unavailable; not product RED.'}
                [void](Probe 'RunLocalShowAndCapture' @($Book.Name,'CHECK_IN'));$Decoy.Activate()
                [void](Probe 'CompleteSubmissionArm' @($interruption,$auth))
                $returned=[bool](Probe 'CompleteBaselineAct')
                $facts=([string](Probe 'CheckYieldEvidence')).Split('|')
                $eventId=[string](Probe 'CompleteSubmissionEvent')
                if($facts.Count -ne 4 -or $facts[0] -cne 'True' -or $facts[1] -cne 'True' -or $facts[3] -cne 'True' -or $eventId -ceq ''){throw 'Actual consume/interruption was not established; not product RED.'}
                $label='CompleteSubmission.'+$interruption
                Check ($label+'.ActualHandlerReturned') $returned
                Check ($label+'.NoLaterSubmissionAttempt') ([int](Probe 'CompleteSubmissionQueues') -eq 1)
                Check ($label+'.NoLaterReads') ([int]$facts[2] -eq 0)
                Check ($label+'.PartialOwnerStatePreserved') ([bool](Probe 'CheckYieldOwnerPreserved'))
                Check ($label+'.ProjectionAtBoundaryPreserved') ([bool](Probe 'CheckYieldProjectionPreserved'))
                $refused=if($interruption -ceq 'Permission'){[bool](Probe 'CheckBaselinePermissionRefused')}else{[bool](Probe 'CheckBaselineContextRefused')}
                Check ($label+'.VisibleInterruptionRefusal') $refused
                CaptureOwnedFormByCaptionEvidence 'Production' ('complete-submission-'+$interruption.ToLowerInvariant()+'.png')
            } finally {
                [void](Probe 'CompleteSubmissionReset')
                [IO.File]::WriteAllBytes($auth,$authBytes)
                SelectTarget $Fixture 'config-producer'
            }
            Check ($label+'.ExactInputsRemainConsumedOnce') ([bool](Owner 'CompleteBaselineBalancesForTest' @($true)))
            Check ($label+'.NoCompletedOutputOrSuccessState') ([bool](Owner 'CompleteSubmissionOutputsAbsentForTest'))
            $authority=$excel.Workbooks.Open((Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')),0,$true)
            try {
                $applied=Table $authority 'tblAppliedEvents';$audit=Table $authority 'tblInventoryLog'
                $appliedCount=@($applied.ListRows|Where-Object {$_.Range.Cells.Item(1,$applied.ListColumns.Item('EventID').Index).Value2 -ceq $eventId}).Count
                $auditCount=@($audit.ListRows|Where-Object {$_.Range.Cells.Item(1,$audit.ListColumns.Item('EventID').Index).Value2 -ceq $eventId -and $_.Range.Cells.Item(1,$audit.ListColumns.Item('System_Key').Index).Value2 -ceq $sourceSelection[0] -and $_.Range.Cells.Item(1,$audit.ListColumns.Item('QtyDelta').Index).Value2 -eq (-[double]$sourceSelection[1])}).Count
                Check ($label+'.ExactConsumeAppliedAndAudited') ($appliedCount -eq 1 -and $auditCount -eq 1)
            } finally {$authority.Close($false)}
            Check ($label+'.CustomValueAndFormulaPreserved') ($Book.Worksheets.Item(1).Cells.Item(2,1).Value2 -ceq $Canary -and $Book.Worksheets.Item(1).Cells.Item(2,2).Formula -ceq '=1+2')
            Check ($label+'.DecoyPreserved') ($Decoy.Worksheets.Count -eq 1 -and $Decoy.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq $Canary)
        }
        Check 'CompleteSubmission.AuthorizationFixtureRestored' ((Get-FileHash -LiteralPath $auth).Hash -ceq $authHash)
    } finally {[void](Probe 'CompleteSubmissionReset')}
}
