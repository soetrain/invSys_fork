# Continuous real Users writes, independent recordings, and packaged guide controls.
# Only owned generated Auth fixtures are restored between completed recordings.
function Install-AdminUsersGuideProbe {
    . (Join-Path $PSScriptRoot 'Slice4beViewerPublishedRead.ps1')
    Test-Slice4beViewerPublishedRead $null $null $true
    . (Join-Path $PSScriptRoot 'Slice4beActionRecording.ps1')
    Test-Slice4beActionRecording $null $true
    . (Join-Path $PSScriptRoot 'Slice4beRecordingReader.ps1')
    Test-Slice4beRecordingReader $null $null $true
    . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
    Install-RecordingEvaluationProbe
    . (Join-Path $PSScriptRoot 'Slice4beGuideDraft.ps1')
    Install-GuideDraftProbe
}

function Test-AdminUsersGuide($Fixture,$Other) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGeneralSettingsGuidePaths.ps1')
    function Hash([string]$Path){Get-AuthReadFixtureHash $Path}
    $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $inventory=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $otherPath=Join-Path $Other.Root ($Other.Warehouse+'.invSys.Auth.xlsb')
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    if(-not [IO.Path]::GetFullPath($authPath).StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Generated Auth fixture required.'}
    $original=[IO.File]::ReadAllBytes($authPath);$originalHash=Hash $authPath
    $inventoryHash=Hash $inventory;$otherHash=Hash $otherPath
    $hostPath=Join-Path $runRoot 'users-guide-host.xlsb'
    if(Test-Path -LiteralPath $hostPath){throw 'Preserve existing guide host.'}
    $steps=@(@('Create / Update','ADMIN_USER_SAVE','SAVED'),@('Deactivate','ADMIN_USER_DEACTIVATE','SAVED'),@('Delete','ADMIN_USER_DELETE','SAVED'))
    $runs=@();$anchor=$null;$auth=$null
    try {
        $anchor=$excel.Workbooks.Add();$anchor.SaveAs($hostPath,50);$anchor.Close($false);$anchor=$null
        $hostHash=Hash $hostPath;$anchor=$excel.Workbooks.Open($hostPath,0,$false)
        $anchor.Windows.Item(1).Visible=$true;$excel.Visible=$true
        SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-admin',$Fixture.Warehouse,'S1'))){throw 'Explicit guide permission fixture unavailable.'}
        Check 'UsersRecorded.ExplicitGuidePermission' $true
        $catalogVersion=[int](Run 'invSys.Core.xlam' 'TestUsersRecording.CurrentCatalogVersion')
        if(-not [bool](Run 'invSys.Core.xlam' 'TestUsersRecording.Policy' @('Capture'))){throw 'Capture policy unavailable.'}
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Initial publication unavailable.'}
        foreach($kind in @('Source','Observed','Interrupted')){
            CloseRecordingViewer
            [IO.File]::WriteAllBytes($authPath,$original);SelectTarget $Fixture
            $anchor.Activate();OpenRecordingViewer
            if(-not [bool](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.LoadForTest' @($authPath))){throw 'Captured Users form unavailable.'}
            $prior=BoundPins $journalRoot;$oldActivity=ActivityPins;$before=@(Get-Slice4beActivityFiles $Fixture)
            $prefix='UsersRecorded.'+$kind
            Check ($prefix+'.ActualStart') ((RecordingControl 'Start Recording' 'Click') -ceq 'DELIVERED')
            $starts=@(Get-ChildItem $journalRoot -File -Filter '*.json'|Where-Object {-not $prior.ContainsKey($_.FullName)}|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
            if($starts.Count -ne 1){throw 'One durable Start required.'}
            $sequence=$starts[0].SequenceId;$success=$kind -cne 'Interrupted';$ordinal=0
            foreach($action in @('Save','Deactivate','Delete')){
                $ordinal++;$id=$steps[$ordinal-1][1];$fault=-not $success -and $ordinal -eq 3
                [void](Run 'invSys.Core.xlam' 'TestUsersRecordingStorage.Arm' @($id,$(if($fault){'SAVED'}else{''})))
                try{
                    [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.WriteReaderForTest' @($action))
                    $observations=([string](Run 'invSys.Core.xlam' 'TestUsersRecordingStorage.Observations')) -split '\|'
                }finally{[void](Run 'invSys.Core.xlam' 'TestUsersRecordingStorage.Arm' @('',''))}
                $expected=if($fault){'1|1|True|1|1'}else{'0|0|False|1|1'}
                Check ($prefix+'.'+$action+'.ExactFaultAndOnePersistence') (($observations -join '|') -ceq $expected)
                if($fault -and ($observations[0] -cne '1' -or $observations[1] -cne '1')){throw 'Real outcome Open failure not reached; not product RED.'}
                Check ($prefix+'.'+$action+'.FixedNoticeCount') ([int](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.TrackingNotices') -eq $(if($fault){1}else{0}))
                Check ($prefix+'.'+$action+'.OwnedAuthClosed') (@($excel.Workbooks|Where-Object {$_.FullName -ieq $authPath}).Count -eq 0)
                $auth=$excel.Workbooks.Open($authPath,0,$true)
                $users=Table $auth 'tblUsers';$caps=Table $auth 'tblCapabilities'
                $rows=@($users.ListRows|Where-Object {$_.Range.Cells.Item(1,$users.ListColumns.Item('UserId').Index).Value2 -ceq 'config-reader'})
                $state=$auth.ReadOnly -and $auth.Saved -and (Hash $authPath) -cne $originalHash
                if($action -ceq 'Delete'){
                    $state=$state -and $rows.Count -eq 0 -and @($caps.ListRows|Where-Object {$_.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-reader'}).Count -eq 0
                }else{
                    $status=if($action -ceq 'Save'){'ACTIVE'}else{'INACTIVE'}
                    $state=$state -and $rows.Count -eq 1 -and $rows[0].Range.Cells.Item(1,$users.ListColumns.Item('Status').Index).Value2 -ceq $status -and $rows[0].Range.Cells.Item(1,$users.ListColumns.Item('DisplayName').Index).Value2 -ceq 'Updated fixture reader'
                }
                Check ($prefix+'.'+$action+'.FreshSavedBusinessState') $state
                $auth.Close($false);$auth=$null
            }
            [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.CloseForTest')
            $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json})
            $count=if($success){6}else{5};$valid=$records.Count -eq $count;$safe=$true
            foreach($record in $records){
                $valid=$valid -and $record.SequenceId -ceq $sequence -and $record.CatalogVersion -eq $catalogVersion -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.UserId -ceq 'config-admin' -and $record.DataEffect -ceq 'Unknown' -and @($record.SourceEventRefs).Count -eq 0
                $wire=$record|ConvertTo-Json -Depth 20 -Compress
                foreach($value in @($Fixture.Root,$Fixture.Secret,(CredentialHash $Fixture.Secret),'config-reader','Updated fixture reader','PinHash','unavailable.pending')){$safe=$safe -and -not $wire.Contains($value)}
            }
            foreach($ordinal in 1,2,3){
                $expected=if(-not $success -and $ordinal -eq 3){1}else{2}
                $part=@($records|Where-Object Ordinal -EQ $ordinal)
                $valid=$valid -and $part.Count -eq $expected -and @($part|Where-Object ControlId -CEQ $steps[$ordinal-1][1]).Count -eq $expected -and @($part|Where-Object OutcomeCode -CEQ 'REQUESTED').Count -eq 1 -and @($part|Where-Object OutcomeCode -CEQ 'SAVED').Count -eq ($expected-1)
            }
            Check ($prefix+'.ExactActionsAndContext') $valid
            Check ($prefix+'.NoPrivatePayload') $safe
            if($success){Check ($prefix+'.ActualStop') ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')}
            $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close');$valid=$closed.Count -eq 1
            if($valid){$valid=$closed[0].ActionCount -eq 3 -and @($closed[0].Observations).Count -eq $count -and $closed[0].Lifecycle -ceq $(if($success){'Stopped'}else{'Incomplete'})}
            if($valid -and -not $success){$valid=$closed[0].ReasonCode -ceq 'TRACKING_UNAVAILABLE'}
            Check ($prefix+'.ExactJournalLifecycle') $valid
            Check ($prefix+'.ImmutableJournalChain') (JournalChain $sequence ($count+2))
            Check ($prefix+'.PriorActivityImmutable') (PinsRetained $oldActivity)
            if($closed.Count -ne 1){throw 'Closed recording unavailable.'}
            $runs+=@{Kind=$kind;Journal=$closed[0];Records=$records;Success=$success;Incomplete=(-not $success)}
        }
        CloseRecordingViewer
        [IO.File]::WriteAllBytes($authPath,$original);SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Recorded publication unavailable.'}
        $publication=Get-Content (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')) -Raw|ConvertFrom-Json
        $lines=@($publication.Groups|Where-Object Source -CEQ 'Activity'|ForEach-Object {$_.Lines});$exact=$true
        foreach($run in $runs){foreach($record in $run.Records){$matches=@($lines|Where-Object RecordId -CEQ $record.RecordId);$exact=$exact -and $matches.Count -eq 1;if($matches.Count -eq 1){$exact=$exact -and $matches[0].ActivityId -ceq $record.ActivityId -and $matches[0].OutcomeCode -ceq $record.OutcomeCode}}}
        Check 'UsersRecorded.ExactPublishedRecords' $exact
        $anchor.Activate()
        Test-GeneralSettingsGuidePaths $Fixture $runs $steps -Users
        Check 'UsersRecorded.OtherAuthorityInventoryAndHostPreserved' ((Hash $otherPath) -ceq $otherHash -and (Hash $inventory) -ceq $inventoryHash -and (Hash $hostPath) -ceq $hostHash)
    }finally{
        [void](Run 'invSys.Core.xlam' 'TestUsersRecordingStorage.Arm' @('',''))
        [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.CloseForTest')
        CloseRecordingViewer
        if($null -ne $auth){$auth.Close($false)}
        if($null -ne $anchor){$anchor.Close($false)}
        [IO.File]::WriteAllBytes($authPath,$original);SelectTarget $Fixture
    }
    Check 'UsersRecorded.AuthFixtureRestored' ((Hash $authPath) -ceq $originalHash)
}
