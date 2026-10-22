# Independent navigation recordings feed actual guide-authoring and evaluation controls.
function Test-InventoryGuide($Fixture,[switch]$Selections,[switch]$Uom) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGeneralSettingsGuidePaths.ps1')
    function Hash([string]$Path){Get-AdminInventoryEntryHash $Path}
    $steps=@(@('Add/Edit Inventory','ADMIN_INVENTORY_OPEN','OPENED'),@('Edit Item','ADMIN_INVENTORY_EDIT_MODE','SELECTED'),@('Add Item Mode','ADMIN_INVENTORY_ADD_MODE','SELECTED'),@('Edit Item','ADMIN_INVENTORY_EDIT_MODE','SELECTED'),@('Close','ADMIN_INVENTORY_CLOSE','CLOSED'))
    if($Selections){
        $steps=@(@('Add/Edit Inventory','ADMIN_INVENTORY_OPEN','OPENED'),@('Edit Item','ADMIN_INVENTORY_EDIT_MODE','SELECTED'))
        foreach($suffix in @('SELECT_ITEM','SELECT_ITEM','SELECT_UOM','SELECT_QUANTITY_MODE','SELECT_LOCATION','SELECT_CATEGORY')){
            $steps+=,@($suffix.Replace('SELECT_','Select ').Replace('_',' '),('ADMIN_INVENTORY_'+$suffix),'SELECTED')
        }
        foreach($outcome in @('REJECTED','STAGED','STAGED')){$steps+=,@('Add Additional Field','ADMIN_INVENTORY_CUSTOM_FIELD_ADD',$outcome)}
        $steps+=,@('Select Additional Field','ADMIN_INVENTORY_SELECT_CUSTOM_FIELD','SELECTED')
        foreach($outcome in @('STAGED','UNCHANGED')){$steps+=,@('Remove Additional Field','ADMIN_INVENTORY_CUSTOM_FIELD_REMOVE',$outcome)}
        $steps+=,@('Close','ADMIN_INVENTORY_CLOSE','CLOSED')
    }
    if($Uom){
        $steps=@(,@('Add/Edit Inventory','ADMIN_INVENTORY_OPEN','OPENED'))
        foreach($outcome in @('STAGED','STAGED','CANCELLED','CANCELLED')){$steps+=,@('Add UOM','ADMIN_INVENTORY_UOM_ADD',$outcome)}
        $steps+=,@('Select UOM','ADMIN_INVENTORY_SELECT_UOM','SELECTED')
        $steps+=,@('Close','ADMIN_INVENTORY_CLOSE','CLOSED')
    }
    $runs=@();$sourcePins=@{}
    function RecordCanonicalPins([string]$Stage){
        foreach($kind in @('Config','Auth','Data.Inventory')){
            $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.'+$kind+'.xlsb')
            [pscustomobject]@{Stage=$Stage;Store=$kind;Sha256=(Hash $path)}|ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'inventory-guide-canonical-pins.jsonl')
        }
    }
    RecordCanonicalPins 'Entry'
    $hostBook=$excel.Workbooks.Add();$hostBook.SaveAs((Join-Path $runRoot 'inventory-guide-host.xlsx'),51)
    try {
        SelectTarget $Fixture
        $authorized=[bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-admin',$Fixture.Warehouse,'S1'))
        if(-not $authorized){throw 'Explicit Inventory guide-maintenance fixture grant unavailable; not product RED.'}
        Check 'InventoryGuide.Recorded.ExplicitGuidePermission' $authorized
        $hostBook.Windows.Item(1).Visible=$true;$hostBook.Activate();$excel.Visible=$true
        Check 'InventoryGuide.Recorded.VisibleHost' ([bool]$excel.Visible -and [bool]$hostBook.Windows.Item(1).Visible)
        if(-not [bool](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.SetPolicy' @('Capture'))){throw 'Inventory capture policy fixture unavailable.'}
        # Saving the fixture's capture policy deliberately writes Config. Pin
        # after that setup; recording, publication and guide actions must preserve it.
        foreach($kind in @('Config','Auth','Data.Inventory')){
            $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.'+$kind+'.xlsb')
            $sourcePins[$path]=Hash $path
        }
        RecordCanonicalPins 'BeforeInitialPublication'
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Initial Inventory Viewer publication unavailable.'}
        RecordCanonicalPins 'AfterInitialPublication'
        foreach($kind in @('Source','Observed','Interrupted')){
            CloseRecordingViewer;SelectTarget $Fixture;OpenRecordingViewer
            $prior=BoundPins $journalRoot;$oldActivity=ActivityPins;$before=@(Get-Slice4beActivityFiles $Fixture)
            Check ('InventoryGuide.Recorded.'+$kind+'.ActualStart') ((RecordingControl 'Start Recording' 'Click') -ceq 'DELIVERED')
            $starts=@(Get-ChildItem $journalRoot -File -Filter '*.json'|Where-Object {-not $prior.ContainsKey($_.FullName)}|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
            if($starts.Count -ne 1){throw 'One durable Inventory Start required.'}
            $sequence=$starts[0].SequenceId;$success=$kind -cne 'Interrupted'
            $mode=if($success){'Enabled'}elseif($Uom){'InterruptAfterSecondPrompt'}elseif($Selections){'InterruptBeforeLastRemove'}else{'InterruptBeforeClose'}
            if($Uom){[void](Run 'invSys.Admin.xlam' 'TestInventoryUom.Arm' @($mode,$Fixture.Secret))}
            [void](Run 'invSys.Admin.xlam' 'TestInventoryRecording.Execute' @($hostBook.Name,$mode))
            Check ('InventoryGuide.Recorded.'+$kind+'.ActualHandlersOnce') ([long](Run 'invSys.Admin.xlam' 'TestInventoryRecording.Gestures') -eq 1 -and [bool](Run 'invSys.Admin.xlam' 'TestInventoryRecording.Finished'))
            $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -Like 'ADMIN_INVENTORY_*')
            $actions=if($success){$steps.Count}elseif($Uom){3}elseif($Selections){13}else{4};$count=$actions*2;$valid=$records.Count -eq $count
            foreach($record in $records){$valid=$valid -and $record.SequenceId -ceq $sequence -and $record.CatalogVersion -eq 47 -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.UserId -ceq 'config-admin' -and $record.OwnerId -ceq 'ADMIN_INVENTORY' -and @($record.SourceEventRefs).Count -eq 0}
            foreach($ordinal in 1..$actions){
                $part=@($records|Where-Object Ordinal -EQ $ordinal);$attempt=@($part|Where-Object OutcomeCode -CEQ 'REQUESTED');$outcome=@($part|Where-Object OutcomeCode -CEQ $steps[$ordinal-1][2])
                $valid=$valid -and $part.Count -eq 2 -and $attempt.Count -eq 1 -and $outcome.Count -eq 1
                if($attempt.Count -eq 1 -and $outcome.Count -eq 1){$valid=$valid -and $attempt[0].ActivityId -ceq $outcome[0].ActivityId -and $attempt[0].ControlId -ceq $steps[$ordinal-1][1] -and $attempt[0].DataEffect -ceq 'Unknown' -and $outcome[0].DataEffect -ceq 'Unchanged'}
            }
            Check ('InventoryGuide.Recorded.'+$kind+'.ExactActionsAndContext') $valid
            if($success){Check ('InventoryGuide.Recorded.'+$kind+'.ActualStop') ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')}
            $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close');$valid=$closed.Count -eq 1
            if($valid){$valid=$closed[0].ActionCount -eq $actions -and @($closed[0].Observations).Count -eq $count -and $closed[0].Lifecycle -ceq $(if($success){'Stopped'}else{'Incomplete'})}
            if($valid -and -not $success){$valid=$closed[0].ReasonCode -ceq 'SESSION_CHANGED'}
            Check ('InventoryGuide.Recorded.'+$kind+'.ExactJournalLifecycle') $valid
            Check ('InventoryGuide.Recorded.'+$kind+'.ImmutableJournalChain') (JournalChain $sequence ($count+2))
            Check ('InventoryGuide.Recorded.'+$kind+'.PriorActivityImmutable') (PinsRetained $oldActivity)
            if($closed.Count -ne 1){throw 'Closed Inventory recording unavailable.'}
            $runs+=@{Kind=$kind;Journal=$closed[0];Records=$records;Success=$success;Incomplete=(-not $success)}
        }
        RecordCanonicalPins 'AfterRecordings'
        CloseRecordingViewer;SelectTarget $Fixture
        RecordCanonicalPins 'BeforeFinalPublication'
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Recorded Inventory publication unavailable.'}
        RecordCanonicalPins 'AfterFinalPublication'
        $publication=Get-Content (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')) -Raw|ConvertFrom-Json
        $lines=@($publication.Groups|Where-Object Source -CEQ 'Activity'|ForEach-Object {$_.Lines});$exact=$true
        foreach($run in $runs){foreach($record in $run.Records){$matches=@($lines|Where-Object RecordId -CEQ $record.RecordId);$exact=$exact -and $matches.Count -eq 1;if($matches.Count -eq 1){$exact=$exact -and $matches[0].ActivityId -ceq $record.ActivityId -and $matches[0].OutcomeCode -ceq $record.OutcomeCode}}}
        Check 'InventoryGuide.Recorded.ExactPublishedRecords' $exact
        Test-GeneralSettingsGuidePaths $Fixture $runs $steps -InventoryActions -InventorySelections:$Selections -InventoryUom:$Uom
        RecordCanonicalPins 'AfterGuide'
        $same=$true;foreach($path in $sourcePins.Keys){$same=$same -and (Hash $path) -ceq $sourcePins[$path]}
        Check 'InventoryGuide.Recorded.CanonicalStoresPreserved' $same
    }finally{CloseRecordingViewer;SelectTarget $Fixture;$hostBook.Close($false)}
}
