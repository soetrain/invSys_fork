# Actual built-in/uploaded Seed recordings and independently authored guide controls.
function Install-AdminSeedGuideProbe {
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
    if($AdminSeedGuideTimingOnly){
        . (Join-Path $PSScriptRoot 'Slice4beAdminSeedGuideTiming.ps1')
        Install-AdminSeedGuideTiming
    }
}

function Test-AdminSeedGuide($Fixture,$Other) {
    $catalogVersion=[int](Run 'invSys.Core.xlam' 'TestSeedRecording.CurrentCatalogVersion')
    if($catalogVersion -lt 37){throw 'Packaged activity catalog predates Demo Seed; not behavioral RED.'}
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGeneralSettingsGuidePaths.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beAdminSeedAppliedGuide.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beAuthReadOnly.ps1')
    # The saved host remains open, as in the qualified Users guide fixture.
    function Hash([string]$Path){Get-AuthReadFixtureHash $Path}
    function PrepareSeed {if(-not [bool](Run 'invSys.Admin.xlam' 'TestSeedOutcome.Prepare' @($Fixture.Warehouse,$false))){throw 'Seed inbox fixture unavailable.'}}
    function ReleaseSeed {[void](Run 'invSys.Admin.xlam' 'TestSeedOutcome.Release')}
    function CloseSeedProjection {if(-not [bool](Run 'invSys.Admin.xlam' 'TestSeedOutcome.CloseBoundaryProjection')){throw 'Generated projection did not close.'}}
    $inventory=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $auth=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $otherAuth=Join-Path $Other.Root ($Other.Warehouse+'.invSys.Auth.xlsb')
    $authHash=Hash $auth;$otherHash=Hash $otherAuth
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    if(-not [IO.Path]::GetFullPath($Fixture.Root).StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Generated Seed guide fixture required.'}
    $hostPath=Join-Path $runRoot 'seed-guide-host.xlsb'
    if(Test-Path -LiteralPath $hostPath){throw 'Preserve existing guide host.'}
    $steps=@(@('Seed built-in demo inventory','ADMIN_DEMO_SEED','APPLIED'),@('Seed the selected training CSV','ADMIN_DEMO_SEED','APPLIED'))
    $runs=@();$anchor=$null
    try{
        $anchor=$excel.Workbooks.Add();$anchor.Worksheets.Item(1).Cells.Item(1,1).Value2='Seed guide canary'
        $anchor.SaveAs($hostPath,50);$anchor.Close($false);$anchor=$null
        $hostHash=Hash $hostPath;$anchor=$excel.Workbooks.Open($hostPath,0,$false)
        $anchor.Windows.Item(1).Visible=$true;$excel.Visible=$true
        SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-admin',$Fixture.Warehouse,'S1'))){throw 'Explicit guide permission fixture unavailable.'}
        Check 'SeedRecorded.ExplicitGuidePermission' $true
        if(-not [bool](Run 'invSys.Core.xlam' 'TestSeedRecording.EnablePolicy' @('Capture'))){throw 'Capture policy unavailable.'}
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Initial publication unavailable.'}
        $kinds=if($AdminSeedGuideTimingOnly){@('Source')}else{@('Source','Observed','Interrupted')}
        foreach($kind in $kinds){
            CloseRecordingViewer;SelectTarget $Fixture
            PrepareSeed
            try{
                $deleted=[string](Run 'invSys.Admin.xlam' 'TestSeedOutcome.InvokeWorkflow' @('DELETE','',$true))
                $state=([string](Run 'invSys.Admin.xlam' 'TestSeedOutcome.State')) -split '\|'
                if($deleted -cne 'OK' -or $state[7] -cne '0'){throw 'Actual demo depletion prerequisite unavailable.'}
            }finally{ReleaseSeed}
            $anchor.Activate();OpenRecordingViewer
            $prior=BoundPins $journalRoot;$oldActivity=ActivityPins;$before=@(Get-Slice4beActivityFiles $Fixture)
            $prefix='SeedRecorded.'+$kind
            Check ($prefix+'.ActualStart') ((RecordingControl 'Start Recording' 'Click') -ceq 'DELIVERED')
            $starts=@(Get-ChildItem $journalRoot -File -Filter '*.json'|Where-Object {-not $prior.ContainsKey($_.FullName)}|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
            if($starts.Count -ne 1){throw 'One durable Start required.'}
            $sequence=$starts[0].SequenceId;$success=$kind -cne 'Interrupted';$eventIds=@()
            for($ordinal=1;$ordinal -le 2;$ordinal++){
                PrepareSeed
                $path='';$sku='DEMO-GUIDE-'+$kind.ToUpperInvariant()
                if($ordinal -eq 2){$path=Join-Path $Fixture.Root ('guide-'+$kind+'.csv');[IO.File]::WriteAllText($path,("ITEM_CODE,ITEM,QTY,UOM,LOCATION`r`n"+$sku+",Guide fixture,13,each,TRAINING`r`n"),[Text.UTF8Encoding]::new($false))}
                $fault=-not $success -and $ordinal -eq 2
                [void](Run 'invSys.Admin.xlam' 'TestSeedOutcome.ResetBoundary')
                [void](Run 'invSys.Core.xlam' 'TestSeedBoundary.Arm' @($(if($fault){'Storage'}else{''}),'APPLIED'))
                try{
                    $ack=[string](Run 'invSys.Admin.xlam' 'TestSeedOutcome.InvokeWorkflow' @('SEED',$path,$true))
                    $counts=([string](Run 'invSys.Admin.xlam' 'TestSeedOutcome.BoundaryCounts')) -split '\|'
                    $faults=([string](Run 'invSys.Core.xlam' 'TestSeedBoundary.Observations')) -split '\|'
                }finally{[void](Run 'invSys.Core.xlam' 'TestSeedBoundary.Arm' @('',''))}
                $state=([string](Run 'invSys.Admin.xlam' 'TestSeedOutcome.State')) -split '\|'
                $eventIds+=([string](Run 'invSys.Admin.xlam' 'TestSeedOutcome.ExactSeedEventForTest'))
                Check ($prefix+'.OwnerAppliedOnce.'+$ordinal) ($ack -ceq 'OK' -and $counts[0] -ceq '1' -and $state[0] -ceq '1' -and $state[2] -ceq 'True' -and $state[6] -ceq 'True' -and $state[7] -ceq $(if($ordinal -eq 1){'24'}else{'25'}))
                $expected=if($fault){'1|1|True|False|False'}else{'0|0|False|False|False'}
                Check ($prefix+'.ExactFaultAndNotice.'+$ordinal) (($faults -join '|') -ceq $expected -and $counts[1] -ceq $(if($fault){'1'}else{'0'}) -and $counts[2] -ceq 'True')
                if($fault -and ($faults[0] -cne '1' -or $faults[1] -cne '1')){throw 'Real final outcome failure not reached.'}
                if($ordinal -eq 2){
                    CloseSeedProjection;$savedHash=Hash $inventory
                    Check ($prefix+'.FreshUploadedEntityAndReceipt') ([bool](Run 'invSys.Admin.xlam' 'TestSeedOutcome.ExactBoundaryEntity' @($sku,13)))
                    Check ($prefix+'.FreshReadPreservesBytes') ((Hash $inventory) -ceq $savedHash)
                }
                ReleaseSeed
            }
            $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json})
            $count=if($success){4}else{3};$valid=$records.Count -eq $count;$safe=$true
            foreach($record in $records){
                $terminal=$record.OutcomeCode -ceq 'APPLIED';$refs=@($record.SourceEventRefs)
                $valid=$valid -and $record.ControlId -ceq 'ADMIN_DEMO_SEED' -and $record.SequenceId -ceq $sequence -and $record.CatalogVersion -eq $catalogVersion -and $record.UserId -ceq 'config-admin' -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.DataEffect -ceq $(if($terminal){'Changed'}else{'Unknown'})
                if($terminal){$valid=$valid -and $refs.Count -eq 1;if($refs.Count -eq 1){$valid=$valid -and $refs[0].EventId -ceq $eventIds[$record.Ordinal-1] -and $refs[0].WarehouseId -ceq $Fixture.Warehouse -and $refs[0].SubmissionState -ceq 'Submitted'}}else{$valid=$valid -and $refs.Count -eq 0}
                $wire=$record|ConvertTo-Json -Depth 20 -Compress
                foreach($value in @($Fixture.Root,$Fixture.Secret,(CredentialHash $Fixture.Secret),$sku,'Guide fixture','unavailable.pending','PinHash')){$safe=$safe -and -not $wire.Contains($value)}
            }
            foreach($ordinal in 1,2){$expected=if(-not $success -and $ordinal -eq 2){1}else{2};$part=@($records|Where-Object Ordinal -EQ $ordinal);$valid=$valid -and $part.Count -eq $expected -and @($part|Where-Object OutcomeCode -CEQ 'REQUESTED').Count -eq 1 -and @($part|Where-Object OutcomeCode -CEQ 'APPLIED').Count -eq ($expected-1)}
            Check ($prefix+'.ExactActionsContextAndSources') $valid
            Check ($prefix+'.NoPrivatePayload') $safe
            if($success){Check ($prefix+'.ActualStop') ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')}
            $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close');$valid=$closed.Count -eq 1
            if($valid){$valid=$closed[0].ActionCount -eq 2 -and @($closed[0].Observations).Count -eq $count -and $closed[0].Lifecycle -ceq $(if($success){'Stopped'}else{'Incomplete'})}
            if($valid -and -not $success){$valid=$closed[0].ReasonCode -ceq 'TRACKING_UNAVAILABLE'}
            Check ($prefix+'.ExactJournalLifecycle') $valid
            Check ($prefix+'.ImmutableJournalChain') (JournalChain $sequence ($count+2))
            Check ($prefix+'.PriorActivityImmutable') (PinsRetained $oldActivity)
            if($closed.Count -ne 1){throw 'Closed recording unavailable.'}
            $runs+=@{Kind=$kind;Journal=$closed[0];Records=$records;Success=$success;Incomplete=(-not $success);EventIds=$eventIds}
        }
        CloseRecordingViewer;SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Recorded publication unavailable.'}
        CloseSeedProjection;$inventoryHash=Hash $inventory
        $publication=Get-Content (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')) -Raw|ConvertFrom-Json
        $lines=@($publication.Groups|Where-Object Source -CEQ 'Activity'|ForEach-Object {$_.Lines});$exact=$true
        foreach($run in $runs){foreach($record in $run.Records){$matches=@($lines|Where-Object RecordId -CEQ $record.RecordId);$exact=$exact -and $matches.Count -eq 1;if($matches.Count -eq 1){$exact=$exact -and $matches[0].ActivityId -ceq $record.ActivityId -and $matches[0].OutcomeCode -ceq $record.OutcomeCode}}}
        Check 'SeedRecorded.ExactPublishedActivity' $exact
        $domain=$true;foreach($run in $runs){foreach($id in $run.EventIds){$domain=$domain -and @($publication.Groups|Where-Object {$_.Source -ceq 'Inventory' -and $_.SourceId -ceq $id}).Count -eq 1}}
        Check $(if($AdminSeedGuideTimingOnly){'SeedTiming.ExactDomainEventsPublished'}else{'SeedRecorded.AllSixExactDomainEventsPublished'}) $domain
        $anchor.Activate()
        if($AdminSeedGuideTimingOnly){
            . (Join-Path $PSScriptRoot 'Slice4beAdminSeedGuideTiming.ps1')
            Test-AdminSeedGuideTiming $Fixture $runs[0]
        }else{
            Test-GeneralSettingsGuidePaths $Fixture $runs $steps -Seed
            Test-AdminSeedAppliedGuide $Fixture $runs $inventory
        }
        Check 'SeedRecorded.AuthoritiesInventoryAndHostPreserved' ((Hash $auth) -ceq $authHash -and (Hash $otherAuth) -ceq $otherHash -and (Hash $inventory) -ceq $inventoryHash -and (Hash $hostPath) -ceq $hostHash -and [string]$anchor.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'Seed guide canary')
    }finally{
        [void](Run 'invSys.Core.xlam' 'TestSeedBoundary.Arm' @('',''))
        ReleaseSeed;CloseRecordingViewer
        if($null -ne $anchor){$anchor.Close($false)}
        SelectTarget $Fixture
    }
}
