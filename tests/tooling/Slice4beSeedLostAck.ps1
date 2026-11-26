# Instrument only the disposable, pre-compile copies. No runtime switch arms the fault.
function Install-SeedLostAckProbe {
    function Replace-ProbeLine($module,[string]$match,[string[]]$lines){
        $hits=@(for($i=1;$i -le $module.CountOfLines;$i++){
            if($module.Lines($i,1).Trim() -ceq $match){$i}
        })
        if($hits.Count -ne 1){throw 'Seed acknowledgment probe anchor changed.'}
        $module.DeleteLines($hits[0],1)
        $module.InsertLines($hits[0],($lines -join [Environment]::NewLine))
    }
    $evidence=New-Item -ItemType Directory -Path (Join-Path $runRoot 'lost-ack-boundary')
    [void](New-Item -ItemType Directory -Path (Join-Path $evidence.FullName 'Runs'))
    $ops=$packages['invSys.Operations.xlam'].VBProject
    $runner=$ops.VBComponents.Item('frmActionPathRun').CodeModule
    $counts=$ops.VBComponents.Add(1)
    $counts.Name='TestAckDispatch'
    $counts.CodeModule.AddFromString(@'
Option Explicit
Public Calls As Long, Lost As Long, Accepted As Long, FailedReturns As Long
Public Function Counts() As String
    Counts = Calls & "|" & Lost & "|" & Accepted & "|" & FailedReturns
End Function
'@)
    $target='"''invSys.Admin.xlam''!modAdminSeedReplay.ApplyBuiltIn"'
    $call="Application.Run($target, mToken, stepId, inputs)"
    Replace-ProbeLine $runner "delivered = $call" @(
        'TestAckDispatch.Calls = TestAckDispatch.Calls + 1',
        "delivered = LoseOwnerReply($call)",
        'If delivered Then TestAckDispatch.Accepted = TestAckDispatch.Accepted + 1')
    $failure='completed = modExecutionRun.FinishStep(mContext, mToken, False, notice)'
    Replace-ProbeLine $runner $failure @('TestAckDispatch.FailedReturns = TestAckDispatch.FailedReturns + 1',$failure)
    $fault=@'
Private Function LoseOwnerReply(ByVal result As Boolean) As Boolean
    Dim root As String, warehouse As String, journal As String, fs As Object
    If Not result Then Exit Function
    root = modNasConnection.GetCurrentTargetRuntimeRoot()
    warehouse = modNasConnection.GetCurrentTargetWarehouseId()
    journal = root & "\Training\ActionPaths\" & warehouse
    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.CopyFile root & "\" & warehouse & ".invSys.Data.Inventory.xlsb", "@BOUNDARY@\inventory.xlsb", False
    fs.CopyFile journal & "\*.json", "@BOUNDARY@\", False
    fs.CopyFile journal & "\Runs\*.json", "@BOUNDARY@\Runs\", False
    TestAckDispatch.Lost = TestAckDispatch.Lost + 1
    Err.Raise vbObjectError + 721, "Seed replay fixture", "Owner acknowledgment lost"
End Function
'@
    $runner.AddFromString($fault.Replace('@BOUNDARY@',$evidence.FullName))
    $admin=$packages['invSys.Admin.xlam'].VBProject
    $owner=$admin.VBComponents.Item('modAdminSeedReplay').CodeModule
    $counts=$admin.VBComponents.Add(1)
    $counts.Name='TestAckOwner'
    $counts.CodeModule.AddFromString(@'
Option Explicit
Public Entries As Long, Claims As Long, Applications As Long
Public Function Counts() As String
    Counts = Entries & "|" & Claims & "|" & Applications
End Function
'@)
    $entry='If kit <> modAdminInventorySeed.R1_DATA_SET_ID Then Exit Function'
    Replace-ProbeLine $owner $entry @('TestAckOwner.Entries = TestAckOwner.Entries + 1',$entry)
    $claim='If Not modExecutionRun.PendingSeed(token, stepId, kit, True) Then GoTo Done'
    Replace-ProbeLine $owner $claim @($claim,'TestAckOwner.Claims = TestAckOwner.Claims + 1')
    $observation=$admin.VBComponents.Item('modAdminSeedObservation').CodeModule
    $finish='FinishSeed action, outcome, eventId, notice'
    Replace-ProbeLine $observation $finish @($finish,'If replayContext <> "" And outcome = "APPLIED" Then TestAckOwner.Applications = TestAckOwner.Applications + 1')
}

function Test-SeedLostAck {
    function DispatchCounts { Run 'invSys.Operations.xlam' 'TestAckDispatch.Counts' }
    function OwnerCounts { Run 'invSys.Admin.xlam' 'TestAckOwner.Counts' }
    function Records($folder){
        Get-ChildItem -LiteralPath $folder -File -Filter '*.json'|ForEach-Object { Get-Content $_.FullName -Raw|ConvertFrom-Json }
    }
    function JournalPins {
        $result=@{}
        foreach($f in Get-ChildItem -LiteralPath $destRoot -File -Filter '*.json'){
            $result[$f.FullName]=(Get-FileHash $f.FullName).Hash
        }
        return $result
    }
    if((DispatchCounts) -cne '0|0|0|0' -or (OwnerCounts) -cne '0|0|0'){throw 'Expected a fresh instrumented destination.'}
    Deliver $runner 'btnStartRun' 'Click'
    Check 'SeedLostAck.OneInvocationOneLostReplyNoAcceptedSuccessOneFailureReturn' ((DispatchCounts) -ceq '1|1|0|1')
    Check 'SeedLostAck.OneAdminEntryClaimAndAppliedOwner' ((OwnerCounts) -ceq '1|1|1')
    $boundary=Join-Path $runRoot 'lost-ack-boundary'
    $runs=Join-Path $destRoot 'Runs'
    $prior=@(Records (Join-Path $boundary 'Runs')|Sort-Object Revision)
    $revisions=@(Records $runs|Sort-Object Revision)
    $run=$revisions[-1]
    $step=$run.Steps[0]
    $event=$step.SourceEventRefs[0]
    $oneRun=@($revisions.RunId|Sort-Object -Unique).Count -eq 1
    Check 'SeedLostAck.BoundRunStillPendingAtLoss' ($oneRun -and $prior[-1].RunId -ceq $run.RunId -and
        $prior[-1].State -ceq 'Running' -and $prior[-1].Steps.Count -eq 0 -and $prior[-1].Revision -lt $run.Revision -and
        $run.Guide.RecordId -ceq $guide.RecordId -and $run.Guide.ContentSha256 -ceq $guide.ContentSha256 -and
        $run.Profile.RecordId -ceq $saved.RecordId -and $run.Profile.ContentSha256 -ceq $saved.ContentSha256 -and
        $run.WarehouseId -ceq $Destination.Warehouse -and $run.PackageBuilds.Count -eq 5 -and
        @($run.PackageBuilds|Where-Object {$_.BuildIdentity -cne $expectedBuilds[$_.PackageId]}).Count -eq 0)
    Check 'SeedLostAck.FailedDispatchRetainsAppliedStep' ($run.State -ceq 'Failed' -and $run.ReasonCode -ceq 'OWNER_NOT_COMPLETED' -and
        $run.Steps.Count -eq 1 -and $step.StepId -ceq $guide.Steps[0].StepId -and $step.ControlId -ceq 'ADMIN_DEMO_SEED' -and
        $step.State -ceq 'Failed' -and $step.ReasonCode -ceq 'APPLIED' -and $step.SourceEventRefs.Count -eq 1 -and
        $event.SubmissionState -ceq 'Submitted' -and $event.WarehouseId -ceq $Destination.Warehouse)
    # REQUESTED and APPLIED share the same action ordinal. Journal versions order them.
    $journalAtLoss=@(Records $boundary|Where-Object SequenceId -ceq $run.Recording.SequenceId|Sort-Object Version)
    $observations=@(foreach($record in $journalAtLoss){if($record.RecordType -ceq 'Observation'){$record.Observations[0]}})
    $closed=@(Records $destRoot|Where-Object {$_.RecordType -ceq 'Close' -and $_.SequenceId -ceq $run.Recording.SequenceId})
    $pair=$closed[0].Observations
    $fresh=$event.EventId -cne $outcome[0].SourceEventRefs[0].EventId -and $step.ActivityId -cne $outcome[0].ActivityId
    Check 'SeedLostAck.PairSavedBeforeReplyLoss' ($journalAtLoss.Count -eq 3 -and $observations.Count -eq 2 -and
        $observations[0].OutcomeCode -ceq 'REQUESTED' -and $observations[1].OutcomeCode -ceq 'APPLIED' -and
        $observations[0].ActivityId -ceq $step.ActivityId -and $observations[1].ActivityId -ceq $step.ActivityId -and
        $observations[1].SourceEventRefs.Count -eq 1 -and $observations[1].SourceEventRefs[0].EventId -ceq $event.EventId)
    Check 'SeedLostAck.ClosedJournalRetainsCurrentOwnerPair' ($fresh -and $closed.Count -eq 1 -and
        $closed[0].PreviousSha256 -ceq $journalAtLoss[-1].ContentSha256 -and $pair.Count -eq 2 -and
        $pair[0].RecordId -ceq $observations[0].RecordId -and $pair[1].RecordId -ceq $observations[1].RecordId -and
        $pair[0].OutcomeCode -ceq 'REQUESTED' -and $pair[1].OutcomeCode -ceq 'APPLIED' -and
        @($pair|Where-Object {$_.ActivityId -cne $step.ActivityId -or $_.ControlId -cne $step.ControlId -or $_.SequenceId -cne $run.Recording.SequenceId}).Count -eq 0 -and
        $pair[1].SourceEventRefs.Count -eq 1 -and $pair[1].SourceEventRefs[0].EventId -ceq $event.EventId)
    $snapshot=Join-Path $boundary 'inventory.xlsb'
    $atLoss=Read-SeedReplayInventory $snapshot
    $inventoryHash=Get-AuthReadFixtureHash $inventory
    Check 'SeedLostAck.OneReceiptAndExactSavedKitBeforeLoss' ((Get-FileHash $snapshot).Hash -ceq $inventoryHash -and
        $atLoss.tblAppliedEvents.Count -eq ($baseline.tblAppliedEvents.Count + 1) -and
        $atLoss.tblInventoryLog.Count -eq ($baseline.tblInventoryLog.Count + 21) -and
        (Test-SeedReplayInventory $reference $baseline $atLoss $event.EventId $Destination))
    $runPins=BoundPins $runs
    $journalPins=JournalPins
    Deliver 'frmInventoryViewer' 'btnRefresh' 'Click'
    $evaluation=Read-SeedRunEvaluation $run $guide $destRoot
    $refs=@($evaluation.TerminalSources)
    $keys=@($atLoss.tblInventoryLog|Where-Object EventID -ceq $event.EventId|ForEach-Object System_Key|Sort-Object)
    Check 'SeedLostAck.VerifyConcludesFromExactCurrentEventAndKeys' ($evaluation.ResultState -ceq 'Concluded' -and
        $refs.Count -eq 1 -and $refs[0].EventId -ceq $event.EventId -and $refs[0].OwnerStatus -ceq 'Applied' -and
        $refs[0].LineCount -eq 21 -and (($refs[0].SystemKeys|Sort-Object) -join '|') -ceq ($keys -join '|'))
    Check 'SeedLostAck.VerifyPreservesDispatchAndRecording' ((BoundSame $runPins (BoundPins $runs)) -and (BoundSame $journalPins (JournalPins)))
    $artifacts=BoundPins $destRoot
    Deliver $runner 'btnCloseRun' 'Click'
    BoundSelect $guide
    Deliver $library 'btnRunHowTo' 'Click'
    Check 'SeedLostAck.ReopenReadyWithoutExecutionOrMutation' ((BoundControl 'btnStartRun' 'State' '' $runner) -ceq 'True|True' -and
        (BoundControl 'lblRunProfile' 'Label' '' $runner).Contains($saved.ContentSha256) -and
        (DispatchCounts) -ceq '1|1|0|1' -and (OwnerCounts) -ceq '1|1|1' -and
        (BoundSame $artifacts (BoundPins $destRoot)) -and (Get-AuthReadFixtureHash $inventory) -ceq $inventoryHash)
    Deliver $runner 'btnCloseRun' 'Click'
}
