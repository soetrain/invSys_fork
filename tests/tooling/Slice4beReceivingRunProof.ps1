# Independent exact owner evidence and the existing evaluator protect actual replay.
function Test-ReceivingRunProof($Guide,$SavedProfile,$Original,$Latest,$Files,$PriorEvents,$PriorKeys,[string]$InventoryPath,$Staging,[bool]$Dispatched,[ref]$Proof) {
    $intact=$Files.Count -gt 0;$previous=$null;$revision=0
    $fields='SchemaVersion|RecordKind|RunId|Revision|RecordId|PreviousRecordId|PreviousSha256|WarehouseId|StationId|CreatedByUserId|CreatedAtUTC|ContextSha256|Guide|Profile|PackageSetVersion|PackageBuilds|Mode|State|ReasonCode|Recording|Steps|ContentSha256'
    foreach($file in @($Files|Sort-Object { [int]($_.BaseName -split '\.')[-1] })) {
        $text=[IO.File]::ReadAllText($file.FullName);$row=$text|ConvertFrom-Json;$revision++
        $match=[regex]::Match($text,',"ContentSha256":"([0-9a-f]{64})"\}$')
        if(-not $match.Success -or $text -match '[^\x00-\x7f]' -or $text.Length -gt 1048576){$intact=$false;continue}
        $body=$text.Substring(0,$match.Index)+'}';$sha=[Security.Cryptography.SHA256]::Create()
        try{$hash=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::ASCII.GetBytes($body))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
        $intact=$intact -and $row.SchemaVersion -eq 1 -and $row.RecordKind -ceq 'ExecutionRun' -and $row.Revision -eq $revision -and $row.ContentSha256 -ceq $hash -and (@($row.PSObject.Properties.Name|Sort-Object) -join '|') -ceq (@($fields -split '\|'|Sort-Object) -join '|')
        if($null -eq $previous){$intact=$intact -and $row.PreviousRecordId -ceq '' -and $row.PreviousSha256 -ceq ''}
        else{$intact=$intact -and $row.RunId -ceq $previous.RunId -and $row.PreviousRecordId -ceq $previous.RecordId -and $row.PreviousSha256 -ceq $previous.ContentSha256 -and $row.RecordId -cne $previous.RecordId}
        $previous=$row
    }
    Check 'ReceivingRun.ImmutableRunChain' ($Dispatched -and $intact)
    $exact=$Dispatched -and $Latest.Guide.ContentSha256 -ceq $Guide.ContentSha256 -and $Latest.Profile.ContentSha256 -ceq $SavedProfile.ContentSha256 -and $Latest.Profile.RecordId -ceq $SavedProfile.RecordId -and $Latest.WarehouseId -ceq $Fixture.Warehouse -and $Latest.Mode -ceq 'RunAll' -and $Latest.ContextSha256 -cmatch '^[0-9a-f]{64}$'
    Check 'ReceivingRun.RecordBindsCapturedGuideProfileTarget' $exact
    $fresh=$false;$closed=$null;$eventId='';$newKey='';$owner=$false
    if($Dispatched -and @($Latest.Recording.PSObject.Properties).Count -gt 0) {
        $journal=@(RecordingJournal $Latest.Recording.SequenceId|Sort-Object Version)
        if($journal.Count){$closed=$journal[-1]}
        $fresh=$null -ne $closed -and $closed.RecordType -ceq 'Close' -and $closed.ActionPathId -ceq $Latest.Recording.ActionPathId -and $closed.ActionPathId -cne $Original.ActionPathId -and $closed.SequenceId -cne $Original.SequenceId -and (JournalChain $closed.SequenceId $journal.Count)
    }
    Check 'ReceivingRun.FreshClosedRecording' $fresh
    $ordered=$false;$identity=$false;$source=$null
    if($fresh){
        $observed=@($closed.Observations|Where-Object OutcomeCode -CNE 'REQUESTED')
        $ordered=(@($observed|ForEach-Object ControlId) -join '|') -ceq (@($Guide.Steps|ForEach-Object ControlId) -join '|')
        $oldIds=@($Original.Observations|ForEach-Object ActivityId)
        $newIds=@($observed|ForEach-Object ActivityId)
        $identity=$newIds.Count -eq $Guide.Steps.Count -and @($newIds|Sort-Object -Unique).Count -eq $newIds.Count -and @($newIds|Where-Object {$_ -ceq '' -or $_ -cin $oldIds}).Count -eq 0 -and (@($Latest.Steps|ForEach-Object ActivityId) -join '|') -ceq ($newIds -join '|')
        $confirm=@($observed|Where-Object {$_.ControlId -ceq 'RECEIVING_CONFIRM_WRITES' -and $_.OutcomeCode -ceq 'CONFIRMED'})
        if($confirm.Count -eq 1 -and @($confirm[0].SourceEventRefs).Count -eq 1){$source=$confirm[0].SourceEventRefs[0];$eventId=[string]$source.EventId}
    }
    Check 'ReceivingRun.FreshOrdinaryHandlerOrderAndIdentities' ($fresh -and $ordered -and $identity)
    $newEvent=$null -ne $source -and $eventId -cne '' -and $eventId -cnotin $PriorEvents -and $source.WarehouseId -ceq $Fixture.Warehouse -and $source.SourceKind -ceq 'Inventory'
    Check 'ReceivingRun.FreshExactSourceEvent' $newEvent
    if($newEvent) {
        $authority=Open-ReceivingEvidenceBook $InventoryPath
        try {
            $applied=@(Get-ReceivingFixtureRows (Table $authority 'tblAppliedEvents')|Where-Object EventID -CEQ $eventId)
            $logged=@(Get-ReceivingFixtureRows (Table $authority 'tblInventoryLog')|Where-Object EventID -CEQ $eventId)
            if($logged.Count -eq 1){$newKey=[string]$logged[0].System_Key}
            $owner=$applied.Count -eq 1 -and $logged.Count -eq 1 -and $logged[0].QtyDelta -eq 2.5 -and $newKey -cne '' -and $newKey -cnotin $PriorKeys
        } finally {$authority.Close($false)}
    }
    Check 'ReceivingRun.NewEntityExactQuantityApplied' $owner
    Check 'ReceivingRun.OwnerClearsStagingPreservesCustomColumn' ($owner -and $Staging.ListRows.Count -eq 0 -and @($Staging.ListColumns|Where-Object Name -CEQ 'B0 Custom Display').Count -eq 1)
    # Refresh is an explicit separate operator action; Verify must never publish.
    [void](BoundControl 'btnRefresh' 'Click' '' 'frmInventoryViewer')
    $evaluationRoot=Join-Path $journalRoot 'Evaluations';$before=BoundPins $evaluationRoot
    $businessBefore=BusinessPins;$activityBefore=ActivityPins;$runBefore=BoundPins $runRoot
    $publishedBefore=Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest'
    if($publishedBefore -isnot [int]){throw 'Existing publication counter unavailable; not replay RED.'}
    $verified=(RunnerControl 'btnVerifyRun' 'Click') -ceq 'DELIVERED'
    $created=@()
    if(Test-Path -LiteralPath $evaluationRoot){$created=@(Get-ChildItem -LiteralPath $evaluationRoot -File -Filter '*.json'|Where-Object {-not $before.ContainsKey($_.FullName)})}
    $evaluation=$null
    if($created.Count -eq 1){$evaluation=Get-Content -LiteralPath $created[0].FullName -Raw|ConvertFrom-Json}
    $evaluated=$verified -and $fresh -and $owner -and $null -ne $evaluation -and $evaluation.ResultState -ceq 'Concluded' -and $evaluation.Guide.ContentSha256 -ceq $Guide.ContentSha256 -and $evaluation.ActionPathId -ceq $closed.ActionPathId -and $evaluation.JournalSha256 -ceq $closed.ContentSha256 -and @($evaluation.TerminalSources).Count -eq 1 -and $evaluation.TerminalSources[0].EventId -ceq $eventId -and @($evaluation.TerminalSources[0].SystemKeys).Count -eq 1 -and $evaluation.TerminalSources[0].SystemKeys[0] -ceq $newKey
    Check 'ReceivingRun.VerifyUsesFreshRecordingAndExactOwnerProof' $evaluated
    $visible=RunnerControl 'txtRunVerification' 'Text'
    Check 'ReceivingRun.DisplayMatchesFreshAppliedProof' ($evaluated -and $visible.StartsWith('Conclusion observed') -and $visible.Contains([string]$evaluation.EvaluationId) -and $visible.Contains($eventId))
    if($CaptureEvidence){
        [void](RunnerControl 'txtRunVerification' 'ViewportTop')
        CaptureOwnedFormByCaptionEvidence 'Run How-To' 'b0-run-concluded.png'
        [void](RunnerControl 'txtRunVerification' 'ViewportBottom')
        CaptureOwnedFormByCaptionEvidence 'Run How-To' 'b0-run-applied-source.png'
    }
    Check 'ReceivingRun.VerifyReadOnlyBusinessAndDispatch' ($verified -and (BoundSame $businessBefore (BusinessPins)) -and (BoundSame $runBefore (BoundPins $runRoot)) -and (Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest') -eq $publishedBefore -and (PinsRetained $activityBefore))
    $Proof.Value=($Dispatched -and $intact -and $exact -and $fresh -and $ordered -and $identity -and $newEvent -and $owner -and $evaluated)
}
