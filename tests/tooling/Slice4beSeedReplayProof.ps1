# Read persisted bytes, including when the publisher has a separate live projection.
function Read-SeedReplayInventory($path) {
    $copy=Join-Path $runRoot ('b1-read-'+[guid]::NewGuid().ToString('N')+'.xlsb')
    $book=$null
    try{
        $inputFile=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
        try{
            $outputFile=[IO.File]::Open($copy,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write)
            try{$inputFile.CopyTo($outputFile)}finally{$outputFile.Dispose()}
        }finally{$inputFile.Dispose()}
        $book=$excel.Workbooks.Open($copy,0,$true)
        $result=@{}
        foreach($name in 'tblInventoryEntities','tblInventoryLog','tblAppliedEvents','tblSkuCatalog'){
            $table=Table $book $name
            $values=$table.Range.Value2
            $rows=@(for($r=2;$r -le $values.GetLength(0);$r++){
                $row=[ordered]@{}
                for($c=1;$c -le $values.GetLength(1);$c++){$row[[string]$values[1,$c]]=$values[$r,$c]}
                [pscustomobject]$row
            })
            $result[$name]=$rows
        }
        return $result
    }finally{
        if($book){$book.Close($false)}
        if(Test-Path -LiteralPath $copy){Remove-Item -LiteralPath $copy}
    }
}

function Test-SeedReplayInventory($reference,$baseline,$after,$eventId,$target) {
    $entities=@($after.tblInventoryEntities)
    $prior=@($baseline.tblInventoryEntities)
    $logs=@($after.tblInventoryLog|Where-Object EventID -CEQ $eventId)
    $receipt=@($after.tblAppliedEvents|Where-Object EventID -CEQ $eventId)
    if(-not $eventId -or $eventId -cin @($baseline.tblAppliedEvents.EventID) -or $receipt.Count -ne 1 -or $receipt[0].Status -ine 'APPLIED'){return $false}
    $new=@($entities|Where-Object {$_.System_Key -cnotin @($prior.System_Key)})
    if($reference.tblInventoryEntities.Count -ne 24 -or $prior.Count -ne 3 -or $entities.Count -ne 24 -or $new.Count -ne 21 -or $logs.Count -ne 21){return $false}
    if(@($entities.System_Key|Sort-Object -Unique).Count -ne 24){return $false}
    foreach($old in $prior){
        $found=@($entities|Where-Object System_Key -CEQ $old.System_Key)
        if($found.Count -ne 1 -or ($old|ConvertTo-Json -Compress) -cne ($found[0]|ConvertTo-Json -Compress)){return $false}
    }
    foreach($entity in $new){
        $key=$entity.System_Key
        if(-not $key -or $key -cin @($reference.tblInventoryEntities.System_Key)){return $false}
        $expected=@($reference.tblInventoryEntities|Where-Object SKU -CEQ $entity.SKU)
        $line=@($logs|Where-Object System_Key -CEQ $key)
        if($expected.Count -ne 1 -or $line.Count -ne 1){return $false}
        foreach($field in 'SKU','QtyOnHand','Location','Condition','InventoryState'){
            if($expected[0].$field -cne $entity.$field){return $false}
        }
        if($line[0].QtyDelta -ne $entity.QtyOnHand -or $line[0].SKU -cne $entity.SKU -or $line[0].Location -cne $entity.Location -or $line[0].Condition -cne $entity.Condition -or $line[0].WarehouseId -cne $target.Warehouse -or $line[0].StationId -cne 'S1'){return $false}
        $item=@($after.tblSkuCatalog|Where-Object SKU -CEQ $entity.SKU)
        $master=@($reference.tblSkuCatalog|Where-Object SKU -CEQ $entity.SKU)
        if($item.Count -ne 1 -or $master.Count -ne 1){return $false}
        foreach($field in 'ITEM_CODE','ITEM','UOM','LOCATION','CATEGORY','DESCRIPTION'){
            if($item[0].$field -cne $master[0].$field){return $false}
        }
    }
    return (@($entities.SKU|Sort-Object -Unique).Count -eq 24)
}

function Read-SeedRunEvaluation($attempt,$guide,$destRoot) {
    $evalRoot=Join-Path $destRoot 'Evaluations'
    $before=BoundPins $evalRoot
    Deliver 'frmActionPathRun' 'btnVerifyRun' 'Click'
    $new=@(Get-ChildItem -LiteralPath $evalRoot -File -Filter '*.json'|Where-Object {-not $before.ContainsKey($_.FullName)})
    if($new.Count -ne 1){throw 'Verify did not save exactly one fresh evaluation.'}
    $evaluation=Get-Content $new[0].FullName -Raw|ConvertFrom-Json
    $journal=@(Get-ChildItem -LiteralPath $destRoot -File -Filter '*.json'|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object {$_.RecordType -ceq 'Close' -and $_.SequenceId -ceq $attempt.Recording.SequenceId})
    if($journal.Count -ne 1 -or $evaluation.Guide.ActionPathId -cne $guide.ActionPathId -or $evaluation.Guide.Version -ne $guide.Version -or $evaluation.Guide.ContentSha256 -cne $guide.ContentSha256 -or $evaluation.JournalSha256 -cne $journal[0].ContentSha256 -or $evaluation.ActionPathId -cne $attempt.Recording.ActionPathId){throw 'Verify consulted a different guide or recording.'}
    return $evaluation
}

function Test-SeedReplayStoppedWithStock($guide,$saved,$attempt,$destRoot,$inventory,$destination,$reference,$baseline,$applied) {
    $runner='frmActionPathRun'
    Deliver $runner 'btnCloseRun' 'Click'
    if(-not $applied){
        # Only after the positive attempt has closed, prepare stock for this separate case.
        if(-not [bool](Run 'invSys.Admin.xlam' 'TestSeedOutcome.Prepare' @($destination.Warehouse,$false))){throw 'Negative setup inbox unavailable.'}
        try{
            if((Run 'invSys.Admin.xlam' 'TestSeedOutcome.InvokeWorkflow' @('SEED','',$true)) -cne 'OK'){throw 'Negative setup Seed failed.'}
        }finally{[void](Run 'invSys.Admin.xlam' 'TestSeedOutcome.Release')}
    }
    $seeded=Read-SeedReplayInventory $inventory
    $event=@($seeded.tblAppliedEvents|Where-Object {$_.EventID -cnotin @($baseline.tblAppliedEvents.EventID)})
    if($event.Count -ne 1 -or -not (Test-SeedReplayInventory $reference $baseline $seeded $event[0].EventID $destination)){throw 'Expected exact pre-existing kit stock and its earlier receipt.'}
    if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Could not publish the earlier Seed event for the attribution check.'}
    Deliver 'frmInventoryViewer' 'btnRefresh' 'Click'
    $publication=Get-Content (Join-Path $destination.Root ($destination.Warehouse+'.invSys.Snapshot.Events.json')) -Raw|ConvertFrom-Json
    $group=@($publication.Groups|Where-Object {$_.Source -ceq 'Inventory' -and $_.SourceId -ceq $event[0].EventID})
    if($group.Count -ne 1 -or $group[0].Lines.Count -ne 21){throw 'Earlier Seed event and entity lines unavailable in the published evidence.'}
    $preserved=Get-AuthReadFixtureHash $inventory
    BoundSelect $guide
    Deliver 'frmActionPathLibrary' 'btnRunHowTo' 'Click'
    Deliver $runner 'cboRunMode' 'Write' 'Step through'
    $runs=Join-Path $destRoot 'Runs'
    $before=BoundPins $runs
    Deliver $runner 'btnStartRun' 'Click'
    Deliver $runner 'btnStopRun' 'Click'
    $revisions=@(Get-ChildItem -LiteralPath $runs -File -Filter '*.json'|Where-Object {-not $before.ContainsKey($_.FullName)}|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Sort-Object Revision)
    $stopped=$revisions[-1]
    if($stopped.RunId -ceq $attempt.RunId -or $stopped.Profile.RecordId -cne $saved.RecordId -or $stopped.State -cne 'Stopped' -or $stopped.ReasonCode -cne 'STOP_REQUESTED' -or $stopped.Steps.Count -ne 0){throw 'Explicit Stop did not prevent current-run Seed.'}
    $evaluation=Read-SeedRunEvaluation $stopped $guide $destRoot
    Check 'SeedReplay.Negative.MatchingStockCannotConcludeStoppedRun' ($evaluation.Publication.Availability -ceq 'Loaded' -and $evaluation.Publication.PublicationId -ceq $publication.PublicationId -and $evaluation.ResultState -ceq 'Failed' -and $evaluation.MissingSteps.Count -eq 1 -and $evaluation.TerminalSources.Count -eq 0 -and (Get-AuthReadFixtureHash $inventory) -ceq $preserved)
    Deliver $runner 'btnCloseRun' 'Click'
    $activityRoot=Join-Path $destination.Root ('Training/Activity/'+$destination.Warehouse)
    $old=BoundPins $activityRoot
    if(-not [bool](Run 'invSys.Admin.xlam' 'TestSeedOutcome.Prepare' @($destination.Warehouse,$false))){throw 'Repeated Seed inbox unavailable.'}
    try{
        if((Run 'invSys.Admin.xlam' 'TestSeedOutcome.InvokeWorkflow' @('SEED','',$true)) -cne 'OK'){throw 'Ordinary repeated Seed failed.'}
    }finally{[void](Run 'invSys.Admin.xlam' 'TestSeedOutcome.Release')}
    $terminal=@(Get-ChildItem $activityRoot -File -Filter '*.json'|Where-Object {-not $old.ContainsKey($_.FullName)}|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object {$_.ControlId -ceq 'ADMIN_DEMO_SEED' -and $_.OutcomeCode -cne 'REQUESTED'})
    $repeated=Read-SeedReplayInventory $inventory
    Check 'SeedReplay.OrdinaryRepeat.UnchangedWithoutEventOrDuplicate' ($terminal.Count -eq 1 -and $terminal[0].OutcomeCode -ceq 'UNCHANGED' -and $terminal[0].SourceEventRefs.Count -eq 0 -and $repeated.tblAppliedEvents.Count -eq $seeded.tblAppliedEvents.Count -and (Test-SeedReplayInventory $reference $baseline $repeated $event[0].EventID $destination))
    BoundSelect $guide
    Deliver 'frmActionPathLibrary' 'btnRunHowTo' 'Click'
    $before=BoundPins $runs
    Deliver $runner 'btnStartRun' 'Click'
    $revisions=@(Get-ChildItem $runs -File -Filter '*.json'|Where-Object {-not $before.ContainsKey($_.FullName)}|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Sort-Object Revision)
    $noop=$revisions[-1]
    if($noop.Steps.Count -ne 1){throw 'Repeat replay did not return its ordinary outcome.'}
    Deliver 'frmInventoryViewer' 'btnRefresh' 'Click'
    $evaluation=Read-SeedRunEvaluation $noop $guide $destRoot
    $afterRepeat=Read-SeedReplayInventory $inventory
    Check 'SeedReplay.ReplayRepeat.UnchangedNoFreshEffect' ($noop.RunId -cne $attempt.RunId -and $noop.State -ceq 'Completed' -and $noop.Steps[0].ReasonCode -ceq 'UNCHANGED' -and $noop.Steps[0].ActivityId -cne $attempt.Steps[0].ActivityId -and $noop.Steps[0].SourceEventRefs.Count -eq 0 -and $evaluation.ResultState -cne 'Concluded' -and $evaluation.TerminalSources.Count -eq 0 -and $afterRepeat.tblAppliedEvents.Count -eq $seeded.tblAppliedEvents.Count -and (Test-SeedReplayInventory $reference $baseline $afterRepeat $event[0].EventID $destination))
    Deliver $runner 'btnCloseRun' 'Click'
    foreach($change in 'Authority','Destination'){
        BoundSelect $guide
        Deliver 'frmActionPathLibrary' 'btnRunHowTo' 'Click'
        $priorRuns=BoundPins $runs
        $inventoryPin=Get-AuthReadFixtureHash $inventory
        if($change -ceq 'Authority'){
            SelectTarget $destination 'config-reader'
            if([bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ADMIN_MAINT','config-reader',$destination.Warehouse,'S1'))){throw 'Expected ordinary user to lack ADMIN_MAINT.'}
        }else{SelectTarget $Fixture}
        $click=BoundControl 'btnStartRun' 'Click' '' $runner
        Check ("SeedReplay.Stale."+$change+"CannotStartOrWrite") ($click -cin @('DELIVERED','DISABLED') -and (BoundSame $priorRuns (BoundPins $runs)) -and (Get-AuthReadFixtureHash $inventory) -ceq $inventoryPin)
        CloseRecordingViewer
        SelectTarget $destination
        OpenRecordingViewer
        if((BoundLibrary 'Open') -cne 'DELIVERED'){throw 'Could not reopen destination guide library.'}
        BoundOpen
    }
    BoundSelect $guide
    $admin=$packages['invSys.Admin.xlam']
    $before=BoundPins $runs
    $inventoryPin=Get-AuthReadFixtureHash $inventory
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
    $admin.Close($false)
    Deliver 'frmActionPathLibrary' 'btnRunHowTo' 'Click'
    Check 'SeedReplay.MissingAdmin.RefusesSetupWithoutWrite' ((BoundControl '' 'Count' '' $runner) -ceq '0' -and (BoundControl 'lblPublishedGuideStatus' 'Label').Contains('Admin') -and (BoundSame $before (BoundPins $runs)) -and (Get-AuthReadFixtureHash $inventory) -ceq $inventoryPin)
}
