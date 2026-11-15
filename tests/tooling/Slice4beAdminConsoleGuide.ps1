# Independent real Console journeys feed the existing guide/evaluation handlers.
function Test-AdminConsoleGuide($Fixture,$HostBook,[int]$CatalogVersion) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGeneralSettingsGuidePaths.ps1')
    function Hash([string]$Path){FileHash $Path}
    $steps=@(@('Admin Console','ADMIN_CONSOLE_OPEN','OPENED'),@('Admin Console','ADMIN_CONSOLE_OPEN','OPENED'))
    $runs=@();$inventory=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb');$inventoryHash=Hash $inventory
    try{
        SelectTarget $Fixture
        $authorized=[bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-admin',$Fixture.Warehouse,'S1'))
        if(-not $authorized){throw 'Explicit guide-maintenance fixture grant unavailable; not product RED.'}
        Check 'AdminConsole.Recorded.ExplicitGuidePermission' $authorized
        $HostBook.Windows.Item(1).Visible=$true;$excel.Visible=$true
        Check 'AdminConsole.Recorded.VisibleHost' ([bool]$excel.Visible -and [bool]$HostBook.Windows.Item(1).Visible)
        if(-not [bool](Probe 'SetPolicy' @('Capture'))){throw 'Console capture policy prerequisite failed.'}
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Initial Viewer publication unavailable.'}
        foreach($kind in @('Source','Observed','Interrupted')){
            CloseRecordingViewer;SelectTarget $Fixture;OpenRecordingViewer
            $prior=BoundPins $journalRoot;$oldActivity=ActivityPins;$before=@(Get-Slice4beActivityFiles $Fixture)
            Check ('AdminConsole.Recorded.'+$kind+'.ActualStart') ((RecordingControl 'Start Recording' 'Click') -ceq 'DELIVERED')
            $starts=@(Get-ChildItem $journalRoot -File -Filter '*.json'|Where-Object {-not $prior.ContainsKey($_.FullName)}|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
            if($starts.Count -ne 1){throw 'One durable Console Start required.'}
            $sequence=$starts[0].SequenceId;$success=$kind -cne 'Interrupted'
            for($ordinal=1;$ordinal -le 2;$ordinal++){
                $fault=-not $success -and $ordinal -eq 2;$mode=if($fault){'OutcomeWriteFailure'}else{''}
                [void](Probe 'ArmBoundary' @($mode,'','',''))
                try{
                    $values=([string](Run 'invSys.Admin.xlam' 'TestAdminConsole.Invoke' @($HostBook.Name))) -split '\|'
                    $faults=[int](Probe 'WriteFailures')
                }finally{[void](Probe 'ArmBoundary' @('','','',''))}
                if($fault -and $faults -ne 1){throw 'Real Console outcome-write fault was not reached.'}
                $notices=if($fault){'1'}else{'0'}
                Check ('AdminConsole.Recorded.'+$kind+'.OwnerOpensOnce.'+$ordinal) ($values.Count -eq 7 -and $values[0] -ceq '1' -and $values[1] -ceq 'True' -and $values[2] -ceq 'True' -and $values[3] -ceq $notices -and $values[4] -ceq '0' -and $values[5] -ceq 'True')
            }
            $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ 'ADMIN_CONSOLE_OPEN')
            $count=if($success){4}else{3};$valid=$records.Count -eq $count
            foreach($record in $records){$valid=$valid -and $record.SequenceId -ceq $sequence -and $record.CatalogVersion -eq $CatalogVersion -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.UserId -ceq 'config-admin' -and $record.DataEffect -ceq 'Unknown' -and @($record.SourceEventRefs).Count -eq 0}
            foreach($ordinal in 1,2){$expected=if(-not $success -and $ordinal -eq 2){1}else{2};$part=@($records|Where-Object Ordinal -EQ $ordinal);$valid=$valid -and $part.Count -eq $expected -and @($part|Where-Object OutcomeCode -CEQ 'REQUESTED').Count -eq 1 -and @($part|Where-Object OutcomeCode -CEQ 'OPENED').Count -eq ($expected-1)}
            Check ('AdminConsole.Recorded.'+$kind+'.ExactActionsAndContext') $valid
            if($success){Check ('AdminConsole.Recorded.'+$kind+'.ActualStop') ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')}
            $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close');$valid=$closed.Count -eq 1
            if($valid){$valid=$closed[0].ActionCount -eq 2 -and @($closed[0].Observations).Count -eq $count -and $closed[0].Lifecycle -ceq $(if($success){'Stopped'}else{'Incomplete'})}
            if($valid -and -not $success){$valid=$closed[0].ReasonCode -ceq 'TRACKING_UNAVAILABLE'}
            Check ('AdminConsole.Recorded.'+$kind+'.ExactJournalLifecycle') $valid
            Check ('AdminConsole.Recorded.'+$kind+'.ImmutableJournalChain') (JournalChain $sequence ($count+2))
            Check ('AdminConsole.Recorded.'+$kind+'.PriorActivityImmutable') (PinsRetained $oldActivity)
            if($closed.Count -ne 1){throw 'Closed Console recording unavailable.'}
            $runs+=@{Kind=$kind;Journal=$closed[0];Records=$records;Success=$success;Incomplete=(-not $success)}
        }
        CloseRecordingViewer;SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Recorded Console publication failed.'}
        $publication=Get-Content (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')) -Raw|ConvertFrom-Json
        $lines=@($publication.Groups|Where-Object Source -CEQ 'Activity'|ForEach-Object {$_.Lines});$exact=$true
        foreach($run in $runs){foreach($record in $run.Records){$matches=@($lines|Where-Object RecordId -CEQ $record.RecordId);$exact=$exact -and $matches.Count -eq 1;if($matches.Count -eq 1){$exact=$exact -and $matches[0].ActivityId -ceq $record.ActivityId -and $matches[0].OutcomeCode -ceq $record.OutcomeCode}}}
        Check 'AdminConsole.Recorded.ExactPublishedRecords' $exact
        Test-GeneralSettingsGuidePaths $Fixture $runs $steps -AdminConsole
        Check 'AdminConsole.Recorded.InventoryAndCustomValuePreserved' ((Hash $inventory) -ceq $inventoryHash -and [bool](Run 'invSys.Admin.xlam' 'TestAdminConsole.HostPreserved' @($HostBook.Name)))
    }finally{[void](Probe 'ArmBoundary' @('','','',''));CloseRecordingViewer;SelectTarget $Fixture}
}
