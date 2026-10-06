# Record ordinary transfer handlers, preserve partial evidence, then evaluate
# explicit intent against the exact fresh journal and published observations.
function Test-GuideTransferRecorded($Fixture,$Other,$Guide) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beEvaluationContracts.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTransferWire.ps1')
    # Reuse the already compiled guide-action probe for expectation controls.
    function ExpectationControl([string]$Control,[string]$Action,[string]$Value='', [string]$Form='frmActionPathExpectation') {
        switch($Action){
            Count {return (BoundControl $Control 'Rows' '' $Form)}
            Boolean {return (BoundControl $Control 'Check' $Value $Form)}
            Index {if((BoundControl $Control 'Select' $Value $Form) -ceq 'SELECTED'){return 'DELIVERED'};return 'CHOICE_UNAVAILABLE'}
            Select {
                $choices=@((BoundControl $Control 'Values' '' $Form) -split "`n")
                if($Value -cnotin $choices){return 'CHOICE_UNAVAILABLE'}
                return (BoundControl $Control 'Write' $Value $Form)
            }
            default {return (BoundControl $Control $Action $Value $Form)}
        }
    }
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Core.xlam' ('TestGuideTransferRecorded.'+$Method) $Values}
    function Hash([string]$Path){(Get-FileHash -LiteralPath $Path).Hash}
    function RestoreConfig([byte[]]$Bytes){
        if(@($excel.Workbooks|Where-Object {$_.FullName -ieq $Fixture.Config}).Count){throw 'Close owned Config before restoring it.'}
        [IO.File]::WriteAllBytes($Fixture.Config,$Bytes)
    }
    function RestoreStore {
        if(Test-Path -LiteralPath $held -PathType Container){
            if(Test-Path -LiteralPath $activityRoot -PathType Leaf){Remove-Item -LiteralPath $activityRoot -Force}
            if(Test-Path -LiteralPath $activityRoot){throw 'Preserve held activity: destination unexpectedly occupied.'}
            Move-Item -LiteralPath $held -Destination $activityRoot
        }
    }
    function AuthorityPins {
        $pins=@{};foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File){
            if($file.FullName.StartsWith($training,[StringComparison]::OrdinalIgnoreCase)){continue}
            $pins[$file.FullName]=Hash $file.FullName
        };$pins
    }
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    $root=[IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\'
    $training=Join-Path $root 'Training\';$held=$activityRoot+'-transfer-recorded-held'
    $auth=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $candidate=Join-Path $Fixture.Root 'transfer-recorded-policy.xlsb'
    $invalidPolicy=Join-Path $Fixture.Root 'transfer-recorded-invalid-policy.xlsb'
    if(-not $root.StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Generated transfer recording fixture required.'}
    foreach($path in @($Fixture.Config,$auth,$activityRoot,$held,$candidate,$invalidPolicy)){
        if(-not [IO.Path]::GetFullPath($path).StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw 'Recording fault path escaped its fixture.'}
    }
    foreach($path in @($held,$candidate,$invalidPolicy)){if(Test-Path -LiteralPath $path){throw 'Preserve existing recording fixture artifacts.'}}
    $files=Join-Path $runRoot 'transfer-recorded';New-Item -ItemType Directory -Path $files|Out-Null
    $package=Join-Path $runRoot 'transfer-activity/export.json';$packageHash=Hash $package
    $invalid=Join-Path $files 'invalid.json';[IO.File]::WriteAllText($invalid,'{}',[Text.Encoding]::ASCII)
    $guideRoot=Join-Path $journalRoot 'Guides'
    $original=[IO.File]::ReadAllBytes($Fixture.Config);$originalHash=Hash $Fixture.Config
    $authBytes=[IO.File]::ReadAllBytes($auth);$authHash=Hash $auth
    $otherPins=BoundPins $Other.Root;$runs=@()
    try {
        SelectTarget $Fixture 'config-admin';SetRecordingPolicy $true
        $enabled=[IO.File]::ReadAllBytes($Fixture.Config);$enabledHash=Hash $Fixture.Config
        $version=Get-TrackingPolicyVersion $Fixture
        SetRecordingPolicy $true
        [IO.File]::WriteAllBytes($candidate,[IO.File]::ReadAllBytes($Fixture.Config))
        $cfg=$excel.Workbooks.Open($candidate,0,$false)
        try{(Table $cfg 'tblEventTrackingPolicies').ListColumns.Item('SchemaVersion').DataBodyRange.Value2=999.0;$cfg.SaveCopyAs($invalidPolicy)}finally{$cfg.Close($false)}
        RestoreConfig $enabled;$authority=AuthorityPins
        foreach($mode in @('Export','Import')){foreach($kind in @('PolicyChanged','StoreUnavailable','PolicyUnreadable','SignedOut','Permission','Cancel','Rejected','PickerFailure','Recovery','GuideRun')){
            $label='GuideTransferRecorded.'+$mode+'.'+$kind;$control='VIEWER_GUIDE_'+$mode.ToUpperInvariant()
            $terminal=$kind -cin @('PolicyChanged','StoreUnavailable','PolicyUnreadable')
            $incomplete=$terminal -or $kind -ceq 'SignedOut'
            $committed=$terminal -or $kind -cin @('Recovery','GuideRun')
            RestoreConfig $enabled;[IO.File]::WriteAllBytes($auth,$authBytes)
            CloseRecordingViewer;SelectTarget $Fixture 'config-admin';OpenRecordingViewer
            if([string](Run 'invSys.Core.xlam' 'TestGuideTransferActivity.Policy' @($control)) -cne 'True|True'){throw 'Transfer capture policy prerequisite failed.'}
            $prior=BoundPins $journalRoot
            Check ($label+'.ActualStartRecording') ((RecordingControl 'Start Recording' 'Click') -ceq 'DELIVERED')
            $starts=@(Get-ChildItem -LiteralPath $journalRoot -File -Filter '*.json'|Where-Object {-not $prior.ContainsKey($_.FullName)}|ForEach-Object {[IO.File]::ReadAllText($_.FullName)|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
            if($starts.Count -ne 1){throw 'Exactly one durable Start required.'}
            if((BoundLibrary 'Open') -cne 'DELIVERED'){throw 'Actual Action Paths entry unavailable.'}
            BoundOpen;BoundSelect $Guide
            $before=ActivityPins;$guides=BoundPins $guideRoot
            $path=if($mode -ceq 'Export'){Join-Path $files ($kind+'-export.json')}else{$package}
            if($kind -ceq 'Cancel'){$path=''}
            if($kind -ceq 'Rejected'){$path=if($mode -ceq 'Export'){$package}else{$invalid}}
            if($kind -ceq 'PickerFailure'){$path='TRANSFER_PICKER_FAULT'}
            $fault=if($terminal -or $kind -ceq 'Permission'){$kind}else{''}
            $next=if($kind -ceq 'PolicyUnreadable'){$invalidPolicy}else{$candidate}
            [void](Probe 'Arm' @($fault,$Fixture.Config,$next,$activityRoot,$auth))
            TransferSetFile $path ($kind -ceq 'SignedOut')
            try {
                Check ($label+'.ActualTransferHandler') ((BoundControl ('btn'+$mode+'Guide') 'Click') -ceq 'DELIVERED')
                if($fault){
                    $injected=[int](Probe 'Hits') -eq 1 -and [bool](Probe 'SameContext')
                    if(-not $injected){throw 'Transfer fault seam did not reach its captured boundary; not product RED.'}
                    Check ($label+'.ExactlyOneFaultAtCapturedContext') $injected
                }
                $after=BoundPins $guideRoot;$added=@($after.Keys|Where-Object {-not $guides.ContainsKey($_)})
                $owner=(TransferRetained $guides $after) -and $added.Count -eq $(if($committed -and $mode -ceq 'Import'){1}else{0})
                if($mode -ceq 'Export' -and $committed){
                    $wire=TransferRead $path 'GuideTransfer' 1 'SchemaVersion|RecordKind|TransferId|ExportedAtUTC|ExportedByUserId|SourceWarehouseId|Guide|ContentSha256'
                    $owner=$owner -and $null -ne $wire -and $wire.Guide.ContentSha256 -ceq $Guide.ContentSha256
                }elseif($mode -ceq 'Export' -and $kind -notin @('Cancel','Rejected','PickerFailure')){$owner=$owner -and -not(Test-Path -LiteralPath $path)}
                Check ($label+'.ExactOwnerEffect') $owner
                $message=if($mode -ceq 'Export'){'Guide exported.'}else{'Guide imported. Imported origin evidence; not locally observed.'}
                switch($kind){
                    PolicyChanged {$message+="`r`nTracking unavailable: the tracking policy changed during this action."}
                    PolicyUnreadable {$message+="`r`nTracking unavailable: the saved tracking policy is invalid."}
                    StoreUnavailable {$message+="`r`nTracking unavailable: the training record could not be saved."}
                    Permission {$message='Unavailable: guide transfer requires ACTION_PATH_MAINT.'}
                    Cancel {$message=$mode+' cancelled.'}
                    Rejected {$message=if($mode -ceq 'Export'){'Export destination already exists. Choose a new file.'}else{'Unavailable: the transfer file has invalid integrity, schema, bounds or provenance.'}}
                    PickerFailure {$message='Unavailable: the guide file could not be selected for '+$mode.ToLowerInvariant()+'.'}
                    SignedOut {$message='Unavailable: the invSys session or warehouse changed. Reopen Viewer.'}
                }
                Check ($label+'.ExactOwnerResultAndTrackingNotice') ((BoundControl 'lblPublishedGuideStatus' 'Label') -ceq $message)
                if($CaptureEvidence -and $kind -cin @('StoreUnavailable','Recovery')){CaptureOwnedFormByCaptionEvidence 'Published guides' ($label.ToLowerInvariant()+'.png')}
            }finally{[void](Probe 'Arm' @('','','','',''));TransferSetFile '';RestoreStore}
            $records=@(Get-ChildItem -LiteralPath $activityRoot -File -Filter '*.json'|Where-Object {-not $before.ContainsKey($_.Name)}|ForEach-Object {[IO.File]::ReadAllText($_.FullName)|ConvertFrom-Json})
            $count=if($incomplete){1}else{2}
            $outcome=switch($kind){Permission {'DENIED'} Cancel {'CANCELLED'} Rejected {'REJECTED'} PickerFailure {'FAILED'} default {'COMPLETED'}}
            $attempts=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$ends=@($records|Where-Object OutcomeCode -CEQ $outcome)
            $valid=$records.Count -eq $count -and $attempts.Count -eq 1
            if(-not $incomplete){$valid=$valid -and $ends.Count -eq 1;if($valid){$valid=$attempts[0].ActivityId -ceq $ends[0].ActivityId}}
            foreach($record in $records){$valid=$valid -and $record.ControlId -ceq $control -and $record.OwnerId -ceq 'CORE_GUIDE_TRANSFER' -and $record.UserId -ceq 'config-admin' -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.SequenceId -ceq $starts[0].SequenceId -and $record.Ordinal -eq 1 -and $record.PolicyVersion -eq $version -and $record.CatalogVersion -eq 30 -and @($record.SourceEventRefs).Count -eq 0}
            Check ($label+'.ExactRecordsWithoutInventedSuccess') $valid
            Check ($label+'.PriorActivityImmutable') (PinsRetained $before)
            [void](BoundControl 'btnCloseGuides' 'Click')
            if($kind -ceq 'PolicyUnreadable'){
                Check ($label+'.PendingJournalIntact') (JournalChain $starts[0].SequenceId 2)
                [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('Refresh',''))
                Check ($label+'.UnavailablePolicyDisablesStop') ((RecordingStatus) -ceq 'Tracking unavailable: the saved tracking policy is invalid.' -and (RecordingControl 'Stop Recording') -ceq 'True|False')
                if($CaptureEvidence){CaptureOwnedFormByCaptionEvidence ('Viewer - '+$Fixture.Warehouse) ($label.ToLowerInvariant()+'-viewer.png')}
                RestoreConfig $enabled
                [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('Refresh',''))
            }
            if(-not $incomplete -or $kind -ceq 'PolicyUnreadable'){
                Check ($label+'.ActualStopRecording') ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')
                Check ($label+'.VisibleCaptureStatus') ((RecordingStatus).StartsWith($(if($incomplete){'Incomplete evidence:'}else{'Stopped.'})))
            }
            $closed=@(RecordingJournal $starts[0].SequenceId|Where-Object RecordType -CEQ 'Close')
            $reason=switch($kind){PolicyChanged {'POLICY_CHANGED'} StoreUnavailable {'TRACKING_UNAVAILABLE'} PolicyUnreadable {'UNFINISHED_ACTIONS'} SignedOut {'SESSION_CHANGED'} default {''}}
            $valid=$closed.Count -eq 1
            if($valid){$j=$closed[0];$valid=$j.Lifecycle -ceq $(if($incomplete){'Incomplete'}else{'Stopped'}) -and $j.ReasonCode -ceq $reason -and $j.ActionCount -eq 1 -and @($j.Observations).Count -eq $count}
            Check ($label+'.ExactCaptureLifecycle') $valid
            Check ($label+'.ImmutableJournalChain') (JournalChain $starts[0].SequenceId ($count+2))
            if($closed.Count -eq 1){$runs+=@{Label=$label;Mode=$mode;Kind=$kind;Control=$control;Journal=$closed[0];Original=$records;Incomplete=$incomplete;Success=($kind -cin @('Recovery','GuideRun'))}}
            CloseRecordingViewer;RestoreConfig $enabled;[IO.File]::WriteAllBytes($auth,$authBytes)
            Check ($label+'.AuthorityAndConfigRestored') ((BoundSame $authority (AuthorityPins)) -and (Hash $Fixture.Config) -ceq $enabledHash)
            Check ($label+'.SourcePackageAndOtherWarehousePreserved') ((Hash $package) -ceq $packageHash -and (BoundSame $otherPins (BoundPins $Other.Root)))
        }}
        SelectTarget $Fixture 'config-admin'
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Actual transfer publication prerequisite failed.'}
        SelectTarget $Fixture 'config-reader';OpenRecordingViewer
        $trainingPins=BoundPins $journalRoot
        foreach($run in $runs){
            $label=$run.Label;$journal=$run.Journal
            if((Select-EvaluationRun $journal.ActionPathId) -cne 'SELECTED'){throw 'Exact transfer recording selection unavailable.'}
            $ready=Set-EvaluationDraft @(,@($run.Control,'COMPLETED','True')) 0 'CommandCompleted' -StopAtMissingChoice
            Check ($label+'.ActualExpectationEditor') $ready
            if(-not $ready){continue}
            $before=@(EvaluationFiles|ForEach-Object FullName)
            $clicked=ExpectationControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths'
            $fresh=@(EvaluationFiles|Where-Object {$_.FullName -cnotin $before})
            if($clicked -cne 'DELIVERED' -or $fresh.Count -ne 1){throw 'Actual Evaluate must append one result.'}
            $result=Get-Content $fresh[0].FullName -Raw|ConvertFrom-Json
            $state=if($run.Incomplete){'Incomplete'}elseif($run.Success){'Concluded'}else{'Failed'}
            $reason=if($run.Incomplete){'CAPTURE_INCOMPLETE'}elseif($run.Success){'COMMAND_COMPLETED'}else{'REQUIRED_OUTCOME_MISMATCH'}
            Check ($label+'.ExactDiagnosticState') ($result.ResultState -ceq $state -and $reason -cin @($result.ReasonCodes))
            Check ($label+'.ExactRunAndAuthoredIntent') ($result.JournalRecordId -ceq $journal.RecordId -and $result.JournalSha256 -ceq $journal.ContentSha256 -and $result.ExpectedConclusion.TerminalKind -ceq 'CommandCompleted' -and @($result.ExpectedConclusion.Steps).Count -eq 1 -and $result.ExpectedConclusion.Steps[0].ControlId -ceq $run.Control -and $result.ExpectedConclusion.Steps[0].RequiredOutcome -ceq 'COMPLETED')
            Check ($label+'.NoInventedCompletionOrDomainSources') (@($result.Matches).Count -eq $(if($run.Success){1}else{0}) -and @($result.TerminalSources).Count -eq 0)
            if($CaptureEvidence -and $run.Kind -cin @('Recovery','Permission','PolicyUnreadable')){CaptureOwnedFormByCaptionEvidence 'Action Paths' ($label.ToLowerInvariant()+'-diagnostic.png')}
        }
        Check 'GuideTransferRecorded.AllTwentyRunsCaptured' ($runs.Count -eq 20)
        . (Join-Path $PSScriptRoot 'Slice4beGuideTransferPaths.ps1')
        Test-GuideTransferPaths $Fixture $runs
        Check 'GuideTransferRecorded.PriorTrainingImmutableThroughEvaluation' (TransferRetained $trainingPins (BoundPins $journalRoot))
        Check 'GuideTransferRecorded.OtherWarehousePreservedThroughPublication' (BoundSame $otherPins (BoundPins $Other.Root))
    }finally{
        [void](Probe 'Arm' @('','','','',''));TransferSetFile '';RestoreStore;CloseRecordingViewer
        RestoreConfig $original;[IO.File]::WriteAllBytes($auth,$authBytes)
        foreach($path in @($candidate,$invalidPolicy)){if(Test-Path -LiteralPath $path){Remove-Item -LiteralPath $path -Force}}
        SelectTarget $Fixture 'config-admin'
    }
    Check 'GuideTransferRecorded.OriginalConfigAndAuthBytesRestored' ((Hash $Fixture.Config) -ceq $originalHash -and (Hash $auth) -ceq $authHash)
}
