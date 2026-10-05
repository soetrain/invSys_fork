# Interrupt after real consume/output processing; retain each exact event.
# No source records, outcomes or business results are manufactured by this probe.
function Install-ProductionCompleteSubmissionProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $owner=$project.VBComponents.Item('modProductionReusableRun').CodeModule
    $start=$owner.ProcStartLine('CompleteReusableProcess',0)
    $end=$start+$owner.ProcCountLines('CompleteReusableProcess',0)
    $returns=@();$queues=@();$processors=@()
    for($line=$start;$line -lt $end;$line++){
        $text=$owner.Lines($line,1).Trim()
        if($text -imatch '^(modProductionCompleteActions\.)?AppendProcessorReport mProcessorReports, processorReport$'){$returns+=$line+1}
        if($text.StartsWith('If Not modRoleEventWriter.QueuePayloadEventCurrent(',[StringComparison]::OrdinalIgnoreCase)){$queues+=$line}
        if($text.StartsWith('processedNow = modProcessor.RunBatch(',[StringComparison]::OrdinalIgnoreCase)){$processors+=$line+1}
    }
    if($returns.Count -ne 2 -or $queues.Count -ne 2 -or $processors.Count -ne 2){throw 'Complete submission anchors changed; not product RED.'}
    $edits=@(
        @{Line=$returns[0];Code='        TestProductionDesigner.CompleteSubmissionReturned "ConsumeProcessed", eventId, processedNow'},
        @{Line=$returns[1];Code='    TestProductionDesigner.CompleteSubmissionReturned "OutputProcessed", eventId, processedNow'},
        @{Line=$processors[1];Code='    TestProductionDesigner.CompleteSubmissionOutputProcessed processedNow'}
    )
    foreach($line in $queues){$edits+=@{Line=$line;Code='    TestProductionDesigner.CompleteSubmissionBeforeQueue'}}
    foreach($edit in $edits){$edit.Preceding=$owner.Lines($edit.Line-1,1)}
    foreach($edit in @($edits|Sort-Object { [int]$_.Line } -Descending)){$owner.InsertLines($edit.Line,$edit.Code)}
    $readback=$owner.Lines($owner.ProcStartLine('CompleteReusableProcess',0),$owner.ProcCountLines('CompleteReusableProcess',0))
    foreach($edit in $edits){
        if($readback.IndexOf(($edit.Preceding+"`r`n"+$edit.Code),[StringComparison]::OrdinalIgnoreCase) -lt 0){
            throw 'Complete submission hook misplaced; not product RED.'
        }
    }
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
Public Function CompleteSubmissionOutputSelectionForTest() As Variant
    Dim output As Object, entity As String
    If mOutputs.Count <> 1 Or mOutputKeys.Count <> 1 Then Exit Function
    Set output = mOutputs(1)
    entity = ReusableRunOutputSystemKey(RunRecordText(output, "ProcessNodeId"), RunRecordText(output, "OutputId"))
    If entity = "" Or mCompleteBaselineBalances.Exists(entity) Then Exit Function
    CompleteSubmissionOutputSelectionForTest = Array(entity, ActualOutputQty(output))
End Function
'@)
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,'Private mCompleteSubmissionArmed As Boolean, mCompleteSubmissionQueues As Long, mCompleteSubmissionEvent As String, mCompleteSubmissionOutputEvent As String')
    $adapter.InsertLines(1,'Private mCompleteSubmissionOutputProcessed As Long')
    $adapter.AddFromString(@'
Public Sub CompleteSubmissionArm(ByVal interruption As String, ByVal authPath As String, Optional ByVal boundary As String = "ConsumeProcessed")
    mCompleteSubmissionQueues = 0: mCompleteSubmissionEvent = "": mCompleteSubmissionOutputEvent = ""
    mCompleteSubmissionOutputProcessed = -1
    CheckYieldArm boundary, 1, interruption, authPath
    mCompleteSubmissionArmed = True
End Sub
Public Sub CompleteSubmissionBeforeQueue()
    If mCompleteSubmissionArmed Then mCompleteSubmissionQueues = mCompleteSubmissionQueues + 1
End Sub
Public Sub CompleteSubmissionReturned(ByVal boundary As String, ByVal eventId As String, ByVal processed As Long)
    If Not mCompleteSubmissionArmed Then Exit Sub
    If boundary = "ConsumeProcessed" Then mCompleteSubmissionEvent = eventId
    If boundary = "OutputProcessed" Then mCompleteSubmissionOutputEvent = eventId
    CheckYieldReturned boundary, (processed > 0 And eventId <> "")
End Sub
Public Function CompleteSubmissionEvent() As String
    CompleteSubmissionEvent = mCompleteSubmissionEvent
End Function
Public Function CompleteSubmissionOutputEvent() As String
    CompleteSubmissionOutputEvent = mCompleteSubmissionOutputEvent
End Function
Public Sub CompleteSubmissionOutputProcessed(ByVal count As Long)
    If mCompleteSubmissionArmed Then mCompleteSubmissionOutputProcessed = count
End Sub
Public Function CompleteSubmissionOutputProcessorCount() As Long
    CompleteSubmissionOutputProcessorCount = mCompleteSubmissionOutputProcessed
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

function Test-ProductionCompleteSubmission($Fixture,$Book,$Decoy,[string]$Canary,[switch]$AfterOutput) {
    $auth=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $authBytes=[IO.File]::ReadAllBytes($auth);$authHash=(Get-FileHash -LiteralPath $auth).Hash
    $group=if($AfterOutput){'CompleteOutputReturn'}else{'CompleteSubmission'}
    $boundary=if($AfterOutput){'OutputProcessed'}else{'ConsumeProcessed'}
    $capturePrefix=if($AfterOutput){'complete-output-return'}else{'complete-submission'}
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
                [void](Probe 'CompleteSubmissionArm' @($interruption,$auth,$boundary))
                $returned=[bool](Probe 'CompleteBaselineAct')
                $facts=([string](Probe 'CheckYieldEvidence')).Split('|')
                $eventId=[string](Probe 'CompleteSubmissionEvent')
                if($facts.Count -ne 4 -or $facts[0] -cne 'True' -or $facts[1] -cne 'True' -or $facts[3] -cne 'True' -or $eventId -ceq ''){
                    [pscustomobject]@{Boundary=$boundary;Interruption=$interruption;HandlerReturned=$returned;Reached=($facts.Count -eq 4 -and $facts[0] -ceq 'True');Available=($facts.Count -eq 4 -and $facts[1] -ceq 'True');InterruptionValid=($facts.Count -eq 4 -and $facts[3] -ceq 'True');ConsumeEventPresent=($eventId -cne '');OutputEventPresent=([string](Probe 'CompleteSubmissionOutputEvent') -cne '');QueueAttempts=[int](Probe 'CompleteSubmissionQueues');OutputProcessorCount=[int](Probe 'CompleteSubmissionOutputProcessorCount')}|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'complete-submission-prerequisite.json')
                    throw 'Actual consume/interruption was not established; not product RED.'
                }
                if($AfterOutput){
                    $outputEventId=[string](Probe 'CompleteSubmissionOutputEvent')
                    $outputSelection=Owner 'CompleteSubmissionOutputSelectionForTest'
                    if($outputEventId -ceq '' -or $outputEventId -ceq $eventId -or $outputSelection -isnot [array] -or $outputSelection.Count -ne 2 -or [string]$outputSelection[0] -ceq '' -or [double]$outputSelection[1] -le 0){throw 'Actual output-processing boundary unavailable; not product RED.'}
                }
                $label=$group+'.'+$interruption
                Check ($label+'.ActualHandlerReturned') $returned
                Check ($label+'.NoLaterSubmissionAttempt') ([int](Probe 'CompleteSubmissionQueues') -eq $(if($AfterOutput){2}else{1}))
                Check ($label+'.NoLaterReads') ([int]$facts[2] -eq 0)
                Check ($label+'.PartialOwnerStatePreserved') ([bool](Probe 'CheckYieldOwnerPreserved'))
                Check ($label+'.ProjectionAtBoundaryPreserved') ([bool](Probe 'CheckYieldProjectionPreserved'))
                $refused=if($interruption -ceq 'Permission'){[bool](Probe 'CheckBaselinePermissionRefused')}else{[bool](Probe 'CheckBaselineContextRefused')}
                Check ($label+'.VisibleInterruptionRefusal') $refused
                CaptureOwnedFormByCaptionEvidence 'Production' ($capturePrefix+'-'+$interruption.ToLowerInvariant()+'.png')
            } finally {
                [void](Probe 'CompleteSubmissionReset')
                [IO.File]::WriteAllBytes($auth,$authBytes)
                SelectTarget $Fixture 'config-producer'
            }
            Check ($label+'.ExactInputsRemainConsumedOnce') ([bool](Owner 'CompleteBaselineBalancesForTest' @($true)))
            if($AfterOutput){Check ($label+'.ExactOutputRemainsApplied') ([bool](Owner 'CompleteBaselineOutputForTest'))}
            else{Check ($label+'.NoCompletedOutputOrSuccessState') ([bool](Owner 'CompleteSubmissionOutputsAbsentForTest'))}
            $authorityPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
            $beforeOpen=@($excel.Workbooks|Where-Object {[string]::Equals($_.FullName,$authorityPath,[StringComparison]::OrdinalIgnoreCase)})
            if($beforeOpen.Count -gt 1){throw 'Duplicate authority workbook during inspection.'}
            $authorityHash=Hash $authorityPath
            $openedForAudit=($beforeOpen.Count -eq 0)
            if($openedForAudit){$authority=$excel.Workbooks.Open($authorityPath,0,$true)}else{$authority=$beforeOpen[0]}
            try {
                $applied=Table $authority 'tblAppliedEvents';$audit=Table $authority 'tblInventoryLog'
                $appliedCount=@($applied.ListRows|Where-Object {$_.Range.Cells.Item(1,$applied.ListColumns.Item('EventID').Index).Value2 -ceq $eventId}).Count
                $auditCount=@($audit.ListRows|Where-Object {$_.Range.Cells.Item(1,$audit.ListColumns.Item('EventID').Index).Value2 -ceq $eventId -and $_.Range.Cells.Item(1,$audit.ListColumns.Item('System_Key').Index).Value2 -ceq $sourceSelection[0] -and $_.Range.Cells.Item(1,$audit.ListColumns.Item('QtyDelta').Index).Value2 -eq (-[double]$sourceSelection[1])}).Count
                Check ($label+'.ExactConsumeAppliedAndAudited') ($appliedCount -eq 1 -and $auditCount -eq 1)
                if($AfterOutput){
                    $outputAppliedCount=@($applied.ListRows|Where-Object {$_.Range.Cells.Item(1,$applied.ListColumns.Item('EventID').Index).Value2 -ceq $outputEventId}).Count
                    $outputAuditCount=@($audit.ListRows|Where-Object {$_.Range.Cells.Item(1,$audit.ListColumns.Item('EventID').Index).Value2 -ceq $outputEventId -and $_.Range.Cells.Item(1,$audit.ListColumns.Item('System_Key').Index).Value2 -ceq $outputSelection[0] -and $_.Range.Cells.Item(1,$audit.ListColumns.Item('QtyDelta').Index).Value2 -eq [double]$outputSelection[1]}).Count
                    Check ($label+'.ExactOutputAppliedAndAudited') ($outputAppliedCount -eq 1 -and $outputAuditCount -eq 1)
                }
            } finally {if($openedForAudit){$authority.Close($false)}}
            $afterOpen=@($excel.Workbooks|Where-Object {[string]::Equals($_.FullName,$authorityPath,[StringComparison]::OrdinalIgnoreCase)})
            $lifetime=($beforeOpen.Count -eq $afterOpen.Count)
            Check ($label+'.AuditWorkbookLifetimePreserved') $lifetime
            Check ($label+'.AuditWorkbookBytesPreserved') ((Hash $authorityPath) -ceq $authorityHash)
            if(-not $lifetime){throw 'Audit inspection changed the authority workbook lifetime; harness failure, not product RED.'}
            Check ($label+'.CustomValueAndFormulaPreserved') ($Book.Worksheets.Item(1).Cells.Item(2,1).Value2 -ceq $Canary -and $Book.Worksheets.Item(1).Cells.Item(2,2).Formula -ceq '=1+2')
            Check ($label+'.DecoyPreserved') ($Decoy.Worksheets.Count -eq 1 -and $Decoy.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq $Canary)
        }
        Check ($group+'.AuthorizationFixtureRestored') ((Get-FileHash -LiteralPath $auth).Hash -ceq $authHash)
    } finally {[void](Probe 'CompleteSubmissionReset')}
}
