# D18: real Shipping handlers participate in the Viewer-started recording.
# Existing owner/source assertions run alongside these checks. No activity is
# fabricated; submitted IDs come from the existing unsaved boundary observer.
function Test-ShippingRecordingAction($Fixture,$Before,[string]$Action,[string]$ControlId,$SubmittedIds,$State) {
    $State.Ordinal++
    $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $Before}|ForEach-Object {
        [IO.File]::ReadAllText($_)|ConvertFrom-Json
    })
    $attempt=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED')
    $outcome=@($records|Where-Object OutcomeCode -CNE 'REQUESTED')
    $pair=$records.Count -eq 2 -and $attempt.Count -eq 1 -and $outcome.Count -eq 1
    $label='ShippingRecording.'+$Action
    Check ($label+'.ExactlyOneAttemptAndOutcome') $pair
    if(-not $pair){return}
    if($State.Ordinal -eq 1){$State.Sequence=[string]$attempt[0].SequenceId}
    $actionPair=[pscustomobject]@{Attempt=$attempt[0];Outcome=$outcome[0]}
    Check ($label+'.SameSequenceOrderedOccurrence') ((HasSequence $actionPair $State.Ordinal) -and
        $attempt[0].SequenceId -ceq $State.Sequence -and $attempt[0].ActivityId -ceq $outcome[0].ActivityId -and
        $attempt[0].RecordId -cne $outcome[0].RecordId)
    Check ($label+'.CapturedActorWarehouseAndControl') (@($records|Where-Object {
        $_.UserId -cne 'config-reader' -or $_.WarehouseId -cne $Fixture.Warehouse -or
        $_.StationId -cne 'S1' -or $_.ControlId -cne $ControlId -or $_.OwnerId -cne 'SHIPPING_WORKFLOW'
    }).Count -eq 0)
    $refs=@($outcome[0].SourceEventRefs)
    $exact=$refs.Count -eq @($SubmittedIds).Count -and @($attempt[0].SourceEventRefs).Count -eq 0
    foreach($id in $SubmittedIds){
        $exact=$exact -and @($refs|Where-Object {$_.EventId -ceq $id -and $_.WarehouseId -ceq $Fixture.Warehouse -and
            $_.SourceKind -ceq 'Inventory' -and $_.SubmissionState -ceq 'Submitted'}).Count -eq 1
    }
    Check ($label+'.EverySubmittedSourceRetained') $exact
    $State.Records+=@($attempt[0],$outcome[0])
    $State.Sources+=@($SubmittedIds)
    $State.Activities+=@([string]$attempt[0].ActivityId)
}

function Test-ShippingRecordingClose($Fixture,$State) {
    Check 'ShippingRecording.ActualStopControl' ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')
    $closed=@(RecordingJournal $State.Sequence|Where-Object RecordType -CEQ 'Close')
    $valid=$closed.Count -eq 1
    if($valid){
        $entry=$closed[0]
        $valid=$entry.RecordKind -ceq 'Recording' -and $entry.SchemaVersion -eq 2 -and
            $entry.Lifecycle -ceq 'Stopped' -and $entry.ActionCount -eq 8 -and
            $entry.CreatedByUserId -ceq 'config-reader' -and $entry.WarehouseId -ceq $Fixture.Warehouse -and
            $entry.ExpectedConclusion.TerminalKind -ceq 'None' -and @($entry.Observations).Count -eq 16
    }
    Check 'ShippingRecording.StoppedContainsEightActionsSixteenObservations' $valid
    Check 'ShippingRecording.ImmutableJournalChain' (JournalChain $State.Sequence 18)
    Check 'ShippingRecording.RepeatedAddRemainsDistinct' ($State.Activities.Count -eq 8 -and
        @($State.Activities|Select-Object -Unique).Count -eq 8 -and $State.Activities[0] -cne $State.Activities[5])
    $exact=$valid -and $State.Records.Count -eq 16
    if($exact){
        for($index=0;$index -lt 16;$index++){
            # Activity-store integrity metadata wraps the observation body.
            # The journal has its own verified hash; compare every body field.
            $body=$State.Records[$index]|Select-Object -Property * -ExcludeProperty ContentSha256
            $exact=$exact -and ($entry.Observations[$index]|ConvertTo-Json -Depth 12 -Compress) -ceq
                ($body|ConvertTo-Json -Depth 12 -Compress)
        }
    }
    Check 'ShippingRecording.JournalRetainsExactOrderedActivityEnvelopes' $exact
    Check 'ShippingRecording.FourDistinctSubmissionIdentities' ($State.Sources.Count -eq 4 -and
        @($State.Sources|Select-Object -Unique).Count -eq 4)
    Check 'ShippingRecording.PriorActivityBytesPreserved' (PinsRetained $State.Pins)
}
