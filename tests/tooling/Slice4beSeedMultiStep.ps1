# Disposable counters at the caller and the ordinary Receiving owner, armed after source authoring.
function Install-SeedReceivingProbe($ops,$runner) {
    $counter=$ops.VBComponents.Add(1)
    $counter.Name='TestSeedReceiving'
    $counter.CodeModule.AddFromString(@'
Option Explicit
Public Armed As Boolean, Dispatches As Long, OpenCalls As Long
Public Sub Arm()
    Armed = True
End Sub
Public Function Counts() As String
    Counts = Dispatches & "|" & OpenCalls
End Function
'@)
    $dispatch='delivered = modReceivingExecution.Dispatch(mContext, controlId, inputs, mCapturedWorkbook, mWorkbook, notice)'
    Replace-ProbeLine $runner $dispatch @('If TestSeedReceiving.Armed Then TestSeedReceiving.Dispatches = TestSeedReceiving.Dispatches + 1',$dispatch)
    $launcher=$ops.VBComponents.Item('modTS_Received').CodeModule
    $entry='If Not modReceivingActivityAction.BeginOpen(userControlAction, activityId, notice) Then Exit Sub'
    Replace-ProbeLine $launcher $entry @('If TestSeedReceiving.Armed Then TestSeedReceiving.OpenCalls = TestSeedReceiving.OpenCalls + 1',$entry)
}

function Test-SeedMultiStep {
    function Counters {
        $a=Run 'invSys.Operations.xlam' 'TestAckDispatch.Counts'
        $b=Run 'invSys.Admin.xlam' 'TestAckOwner.Counts'
        $c=Run 'invSys.Operations.xlam' 'TestSeedReceiving.Counts'
        return "$a/$b/$c"
    }
    function JournalPins {
        $p=@{}
        foreach($f in Get-ChildItem -LiteralPath $destRoot -File -Filter '*.json'){$p[$f.FullName]=(Get-FileHash $f.FullName).Hash}
        return $p
    }
    function VerifyOutcome($result) {
        $matches=@($result.Matches)
        $refs=@($result.TerminalSources)
        if($matches.Count -ne $count -or $result.ResultState -cne $verdict -or $result.FailedSteps.Count -or $result.UnavailableSteps.Count -or $result.ExtraActivityIds.Count){return $false}
        if($refs.Count -ne 1 -or $refs[0].EventId -cne $event.EventId -or $refs[0].OwnerStatus -cne 'Applied' -or $refs[0].LineCount -ne 21 -or (($refs[0].SystemKeys|Sort-Object) -join '|') -cne ($keys -join '|')){return $false}
        foreach($i in 0..($count-1)){
            if($matches[$i].StepId -cne $guide.ExpectedConclusion.Steps[$i].StepId -or $matches[$i].ActivityId -cne $attempt.Steps[$i].ActivityId -or $matches[$i].OutcomeCode -cne $attempt.Steps[$i].ReasonCode){return $false}
        }
        if($lost){return ($result.MissingSteps.Count -eq 1 -and $result.MissingSteps[0] -ceq $guide.ExpectedConclusion.Steps[1].StepId -and 'REQUIRED_STEP_MISSING' -cin $result.ReasonCodes)}
        return ($result.MissingSteps.Count -eq 0)
    }
    $lost=$AdminSeedMultiStep -ceq 'Lost'
    $count=2; $state='Completed'; $reason=''; $verdict='Concluded'; $counts='1|0|1|0/1|1|1/1|1'
    if($lost){$count=1; $state='Failed'; $reason='OWNER_NOT_COMPLETED'; $verdict='Failed'; $counts='1|1|0|1/1|1|1/0|0'}
    $case="SeedMulti.$AdminSeedMultiStep"
    $name=$Destination.Warehouse+'.Receiving.Operator.xlsm'
    $operatorPath=Join-Path $operatorRoot ($Destination.Warehouse+'/S1/'+$name)
    $choices=@((BoundControl 'lstRunSourceEntities' 'Values' '' $runner) -split "`n")
    if((Test-Path -LiteralPath $operatorPath) -or (Counters) -cne '0|0|0|0/0|0|0/0|0'){throw 'Fresh destination counts and absent Receiving workbook required.'}
    if((BoundControl 'btnStartRun' 'State' '' $runner) -cne 'True|False' -or $choices[0] -cnotin @($baseline.tblInventoryEntities.System_Key)){throw 'Mixed setup must require a real destination bootstrap entity.'}
    if((BoundControl 'lstRunSourceEntities' 'Select' '0' $runner) -cne 'SELECTED' -or (BoundControl 'lstRunSourceEntities' 'Selected' '' $runner) -cne $choices[0] -or (BoundControl 'btnStartRun' 'State' '' $runner) -cne 'True|True'){throw 'Exact selected destination key did not enable Start.'}
    Check "$case.EligibleBoundMixedRun" $true
    [void](Run 'invSys.Operations.xlam' 'TestSeedReceiving.Arm')
    Deliver $runner 'btnStartRun' 'Click'
    Check "$case.SeedDispatchReplyClaimApplicationAndReceivingDispatchOwnerCounts" ((Counters) -ceq $counts)
    $runs=Join-Path $destRoot 'Runs'
    $revisions=@(Get-ChildItem $runs -File -Filter '*.json'|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Sort-Object Revision)
    $attempt=$revisions[-1]; $first=$attempt.Steps[0]; $event=$first.SourceEventRefs[0]
    Check "$case.OneBoundRunAndSeedDispatchHistory" (@($revisions.RunId|Sort-Object -Unique).Count -eq 1 -and $attempt.State -ceq $state -and $attempt.ReasonCode -ceq $reason -and $attempt.Steps.Count -eq $count -and
        $attempt.Guide.RecordId -ceq $guide.RecordId -and $attempt.Profile.RecordId -ceq $saved.RecordId -and $attempt.WarehouseId -ceq $Destination.Warehouse -and
        $attempt.PackageBuilds.Count -eq 5 -and @($attempt.PackageBuilds|Where-Object {$_.BuildIdentity -cne $expectedBuilds[$_.PackageId]}).Count -eq 0 -and
        $first.StepId -ceq $saved.Steps[0].StepId -and $first.ControlId -ceq 'ADMIN_DEMO_SEED' -and $first.State -ceq $state -and $first.ReasonCode -ceq 'APPLIED' -and
        $first.SourceEventRefs.Count -eq 1 -and $event.SubmissionState -ceq 'Submitted' -and $event.WarehouseId -ceq $Destination.Warehouse -and
        $event.EventId -cne $outcome[0].SourceEventRefs[0].EventId -and $first.ActivityId -cne $outcome[0].ActivityId)
    $closed=@(Get-ChildItem $destRoot -File -Filter '*.json'|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -ceq 'Close')
    $obs=$closed[0].Observations
    Check "$case.ClosedCurrentSeedActivityAndEvent" ($closed.Count -eq 1 -and $closed[0].SequenceId -ceq $attempt.Recording.SequenceId -and $obs.Count -eq (2*$count) -and
        $obs[0].OutcomeCode -ceq 'REQUESTED' -and $obs[1].OutcomeCode -ceq 'APPLIED' -and $obs[0].ActivityId -ceq $first.ActivityId -and $obs[1].ActivityId -ceq $first.ActivityId -and
        $obs[1].SourceEventRefs[0].EventId -ceq $event.EventId)
    $owner=@($obs|Where-Object ControlId -ceq 'RECEIVING_OPEN')
    if($lost){
        Check "$case.NoLaterStepActivityOrWorkbook" (@($revisions|Where-Object {$_.Steps.Count -gt 1}).Count -eq 0 -and $owner.Count -eq 0 -and -not (Test-Path -LiteralPath $operatorPath))
    }else{
        $later=$attempt.Steps[1]
        $book=$excel.Workbooks.Item($name)
        Check "$case.ReceivingOwnerOpenedBoundWorkbook" ($later.StepId -ceq $saved.Steps[1].StepId -and $later.ControlId -ceq 'RECEIVING_OPEN' -and $later.State -ceq 'Completed' -and $later.ReasonCode -ceq 'OPENED' -and
            $owner.Count -eq 2 -and $owner[0].OutcomeCode -ceq 'REQUESTED' -and $owner[1].OutcomeCode -ceq 'OPENED' -and $owner[0].ActivityId -ceq $later.ActivityId -and
            $owner[1].ActivityId -ceq $later.ActivityId -and $later.ActivityId -cne $first.ActivityId -and $book.FullName -ieq (Get-Item -LiteralPath $operatorPath).FullName -and (Table $book 'ReceivedTally').ListRows.Count -eq 0)
    }
    $after=Read-SeedReplayInventory $inventory
    $keys=@($after.tblInventoryLog|Where-Object EventID -ceq $event.EventId|ForEach-Object System_Key|Sort-Object)
    $inventoryHash=Get-AuthReadFixtureHash $inventory
    Check "$case.SavedReceiptAndExactKit" ($after.tblAppliedEvents.Count -eq $baseline.tblAppliedEvents.Count+1 -and $after.tblInventoryLog.Count -eq $baseline.tblInventoryLog.Count+21 -and (Test-SeedReplayInventory $reference $baseline $after $event.EventId $Destination))
    if($lost){Check "$case.InventoryDurableBeforeReplyLost" ((Get-FileHash (Join-Path $runRoot 'seed-return-boundary/inventory.xlsb')).Hash -ceq $inventoryHash)}
    $runPins=BoundPins $runs; $journal=JournalPins
    Deliver 'frmInventoryViewer' 'btnRefresh' 'Click'
    $result=Read-SeedRunEvaluation $attempt $guide $destRoot
    Check "$case.WholeGuideVerdictAndAttributedSeedKeys" (VerifyOutcome $result)
    if($lost){
        $again=Read-SeedRunEvaluation $attempt $guide $destRoot
        Check "$case.RepeatedVerify" ((VerifyOutcome $again) -and $again.EvaluationId -cne $result.EvaluationId)
        Deliver $runner 'btnCloseRun' 'Click'
        BoundSelect $guide
        $artifacts=BoundPins $destRoot
        Deliver $library 'btnRunHowTo' 'Click'
        Check "$case.FreshSetupRequiresSelectionAndHasNoAttempt" ((BoundControl 'btnVerifyRun' 'State' '' $runner) -ceq 'True|False' -and (BoundControl 'lstRunSourceEntities' 'Select' '0' $runner) -ceq 'SELECTED' -and
            (BoundControl 'btnStartRun' 'State' '' $runner) -ceq 'True|True' -and (BoundSame $artifacts (BoundPins $destRoot)))
        Deliver $runner 'btnCloseRun' 'Click'
        Deliver $library 'btnCloseGuides' 'Click'
        Deliver 'frmActionPaths' 'btnPathRefresh' 'Click'
        if((BoundLibrary 'Select' $attempt.Recording.ActionPathId) -cne 'SELECTED'){throw 'Historical recording could not be selected.'}
        BoundOpen; BoundSelect $guide
        Deliver $library 'btnUseGuideForRun' 'Click'
        $historical=Read-SeedRunEvaluation $attempt $guide $destRoot -SelectedRecording
        Check "$case.ReevaluateHistoricalRunAfterReopen" (VerifyOutcome $historical)
    }else{Deliver $runner 'btnCloseRun' 'Click'}
    Check "$case.HistoryAndBusinessPreservedWithoutDispatch" ((Counters) -ceq $counts -and (BoundSame $runPins (BoundPins $runs)) -and (BoundSame $journal (JournalPins)) -and
        (Get-AuthReadFixtureHash $inventory) -ceq $inventoryHash -and (Test-Path -LiteralPath $operatorPath) -eq (-not $lost))
}
