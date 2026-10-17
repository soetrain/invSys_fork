# Independent real aggregations; authored intent and comparisons use packaged controls.
function Test-AggregationGuide($Fixture,$Other,$HostBook,$Expected){
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGeneralSettingsGuidePaths.ps1')
    $steps=@(@('Aggregate selected sources','ADMIN_AGGREGATION_AGGREGATE','COMPLETED'),@('Aggregate selected sources','ADMIN_AGGREGATION_AGGREGATE','COMPLETED'))
    $runs=@();$inventory=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb');$inventoryHash=Hash $inventory
    try{
        SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-admin',$Fixture.Warehouse,'S1'))){throw 'Explicit Aggregation guide permission unavailable.'}
        if(-not [bool](Activity 'SetPolicy' @('AggregationOnly'))){throw 'Aggregation capture policy unavailable.'}
        $HostBook.Windows.Item(1).Visible=$true;$HostBook.Activate();$excel.Visible=$true
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Initial Aggregation Viewer publication unavailable.'}
        foreach($kind in @('Source','Observed','Interrupted','Unconfirmed')){
            CloseRecordingViewer;SelectTarget $Fixture;OpenRecordingViewer
            $prior=BoundPins $journalRoot;$oldActivity=ActivityPins;$before=@(Get-Slice4beActivityFiles $Fixture)
            Check ('Aggregation.Recorded.'+$kind+'.ActualStart') ((RecordingControl 'Start Recording' 'Click') -ceq 'DELIVERED')
            $starts=@(Get-ChildItem $journalRoot -File -Filter '*.json'|Where-Object {-not $prior.ContainsKey($_.FullName)}|ForEach-Object {Get-Content -LiteralPath $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
            if($starts.Count -ne 1){throw 'One durable Aggregation recording Start required.'}
            $sequence=$starts[0].SequenceId;$incomplete=$kind -ceq 'Interrupted';$success=$kind -cin @('Source','Observed')
            for($ordinal=1;$ordinal -le 2;$ordinal++){
                [void](Invoke 'OpenForm' @($HostBook.Name))
                $refusal=$kind -ceq 'Unconfirmed' -and $ordinal -eq 2
                [void](Invoke 'SelectSource' @($(if($refusal){''}else{$Fixture.Warehouse})))
                [void](Invoke 'ResetTrackingNotices')
                $fault=$incomplete -and $ordinal -eq 2
                [void](Activity 'ArmBoundary' @($(if($fault){'OutcomeWriteFailure'}else{''}),'','',''))
                $ownPins=SourcePins $Fixture;$otherPins=SourcePins $Other
                try{
                    [void](Invoke 'Aggregate')
                    Check ('Aggregation.Recorded.'+$kind+'.SourcesPreserved.'+$ordinal) ((Preserved $ownPins) -and (Preserved $otherPins))
                    Check ('Aggregation.Recorded.'+$kind+'.OwnerTruth.'+$ordinal) ([bool](Invoke 'OwnerSuccess') -eq (-not $refusal))
                    Check ('Aggregation.Recorded.'+$kind+'.OwnerOnceOrRefused.'+$ordinal) ([int](Invoke 'OwnerCalls') -eq $(if($refusal){0}else{1}))
                    if(-not $refusal){Check ('Aggregation.Recorded.'+$kind+'.ExactPersistedOutput.'+$ordinal) (ExactOutput)}
                    Check ('Aggregation.Recorded.'+$kind+'.NoticeTruth.'+$ordinal) ([int](Invoke 'TrackingNotices') -eq $(if($fault){1}else{0}))
                    if($fault){Check 'Aggregation.Recorded.Interrupted.ActualOutcomeWriteFault' ([int](Activity 'WriteFailures') -eq 1)}
                }finally{[void](Activity 'ArmBoundary' @('','','',''));[void](Invoke 'CloseForm')}
            }
            $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ 'ADMIN_AGGREGATION_AGGREGATE')
            $count=if($incomplete){3}else{4};$valid=$records.Count -eq $count
            foreach($record in $records){$valid=$valid -and $record.SequenceId -ceq $sequence -and $record.CatalogVersion -eq 43 -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.UserId -ceq 'config-admin' -and $record.DataEffect -ceq 'Unknown' -and @($record.SourceEventRefs).Count -eq 0}
            foreach($ordinal in 1,2){$expectedCount=if($incomplete -and $ordinal -eq 2){1}else{2};$outcome=if($kind -ceq 'Unconfirmed' -and $ordinal -eq 2){'REJECTED'}else{'COMPLETED'};$part=@($records|Where-Object Ordinal -EQ $ordinal);$valid=$valid -and $part.Count -eq $expectedCount -and @($part|Where-Object OutcomeCode -CEQ 'REQUESTED').Count -eq 1 -and @($part|Where-Object OutcomeCode -CEQ $outcome).Count -eq ($expectedCount-1)}
            Check ('Aggregation.Recorded.'+$kind+'.ExactActionsAndContext') $valid
            if(-not $incomplete){Check ('Aggregation.Recorded.'+$kind+'.ActualStop') ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')}
            $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close');$valid=$closed.Count -eq 1
            if($valid){$valid=$closed[0].ActionCount -eq 2 -and @($closed[0].Observations).Count -eq $count -and $closed[0].Lifecycle -ceq $(if($incomplete){'Incomplete'}else{'Stopped'})}
            if($valid -and $incomplete){$valid=$closed[0].ReasonCode -ceq 'TRACKING_UNAVAILABLE'}
            Check ('Aggregation.Recorded.'+$kind+'.ExactJournalLifecycle') $valid
            Check ('Aggregation.Recorded.'+$kind+'.ImmutableJournalChain') (JournalChain $sequence ($count+2))
            Check ('Aggregation.Recorded.'+$kind+'.PriorActivityImmutable') (PinsRetained $oldActivity)
            if($closed.Count -ne 1){return}
            $runs+=@{Kind=$kind;Journal=$closed[0];Records=$records;Success=$success;Incomplete=$incomplete}
        }
        CloseRecordingViewer;SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Recorded Aggregation publication failed.'}
        $publication=Get-Content (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')) -Raw|ConvertFrom-Json
        $lines=@($publication.Groups|Where-Object Source -CEQ 'Activity'|ForEach-Object {$_.Lines});$exact=$true
        foreach($run in $runs){foreach($record in $run.Records){$matches=@($lines|Where-Object RecordId -CEQ $record.RecordId);$exact=$exact -and $matches.Count -eq 1;if($matches.Count -eq 1){$exact=$exact -and $matches[0].ActivityId -ceq $record.ActivityId -and $matches[0].OutcomeCode -ceq $record.OutcomeCode}}}
        Check 'Aggregation.Recorded.ExactPublishedRecords' $exact
        Test-GeneralSettingsGuidePaths $Fixture $runs $steps -AggregationActions
        Check 'Aggregation.Recorded.SourceInventoryPreserved' ((Hash $inventory) -ceq $inventoryHash)
    }finally{[void](Activity 'ArmBoundary' @('','','',''));[void](Invoke 'CloseForm');CloseRecordingViewer;SelectTarget $Fixture}
}
