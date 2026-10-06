# Actual Print, recording, publication and diagnostic handlers on disposable fixtures.
function Test-ProductionPrintRecorded($Fixture,$Other,[switch]$GuideDiagnostic){
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beEvaluationContracts.ps1')
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Hash([string]$Path){$s=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $s).Hash}finally{$s.Dispose()}}
    function Fingerprint($Sheet){ConvertTo-Json -Compress -Depth 10 -InputObject @($Sheet.UsedRange.Formula)}
    function AuthorityPins {
        $pins=@{};foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File|Where-Object {$_.Name -match '\.invSys\.(Data\.|Auth\.)' -and $_.Name -notlike '*.Snapshot.*'}){$pins[$file.FullName]=Hash $file.FullName};$pins
    }
    function Same($Before,$After){if($Before.Count -ne $After.Count){return $false};foreach($path in $Before.Keys){if(-not $After.ContainsKey($path) -or $After[$path] -cne $Before[$path]){return $false}};return $true}
    function RestoreConfig([byte[]]$Bytes){
        if(@($excel.Workbooks|Where-Object {$_.FullName -ieq $Fixture.Config}).Count){throw 'Close owned Config before restoring bytes.'}
        [IO.File]::WriteAllBytes($Fixture.Config,$Bytes)
    }
    function RestoreStore {
        if(Test-Path -LiteralPath $held -PathType Container){
            if(Test-Path -LiteralPath $activityRoot -PathType Leaf){Remove-Item -LiteralPath $activityRoot -Force}
            if(Test-Path -LiteralPath $activityRoot){throw 'Preserve held evidence: store unexpectedly occupied.'}
            Move-Item -LiteralPath $held -Destination $activityRoot
        }
    }
    $root=[IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\'
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    if(-not $root.StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Print recording fixture escaped owned root.'}
    $auth=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $candidate=Join-Path $Fixture.Root 'print-terminal-policy.xlsb'
    $invalid=Join-Path $Fixture.Root 'print-terminal-invalid.xlsb';$held=$activityRoot+'-terminal-held'
    foreach($path in @($auth,$Fixture.Config,$candidate,$invalid,$activityRoot,$held)){
        if(-not [IO.Path]::GetFullPath($path).StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw 'Print recording path escaped disposable fixture.'}
    }
    foreach($path in @($candidate,$invalid,$held)){if(Test-Path -LiteralPath $path){throw 'Preserve existing Print recording artifacts.'}}
    $original=[IO.File]::ReadAllBytes($Fixture.Config);$originalPin=Hash $Fixture.Config
    $authBytes=[IO.File]::ReadAllBytes($auth);$authPin=Hash $auth;$otherPins=RestartPins $Other.Root
    $book=$null;$work=$null;$decoy=$null;$runs=@()
    try{
        SelectTarget $Fixture
        if([string](Run 'invSys.Operations.xlam' 'modInventoryViewer.GuideDraftControlForTest' @('frmActionPathGuide','','Count','')) -cne '0' -or -not [bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-admin',$Fixture.Warehouse,'S1'))){throw 'Print guide probe/author fixture unavailable; not product RED.'}
        $seed=[string](Run 'invSys.Admin.xlam' 'modAdmin.RunDemoInventoryActionCallbackForAutomation' @($Fixture.Warehouse,'S1','config-admin','SEED'))
        Check 'PrintRecorded.SeedCallback' ($seed.StartsWith('OK|'))
        if(-not $seed.StartsWith('OK|')){throw 'Seed prerequisite unavailable; not product RED.'}
        $pins=AuthorityPins
        $inventory=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
        $authority=$excel.Workbooks.Open($inventory,0,$true)
        try{
            $entities=Table $authority 'tblInventoryEntities'
            $key=[string]$entities.ListColumns.Item('System_Key').DataBodyRange.Cells.Item(1,1).Value2
            $location=[string]$entities.ListColumns.Item('Location').DataBodyRange.Cells.Item(1,1).Value2
            if(-not $key -or -not $location){throw 'Seed identity/location unavailable; not product RED.'}
        }finally{$authority.Close($false)}
        Check 'PrintRecorded.ReadOnlySeedInspection' (Same $pins (AuthorityPins))
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1);$sheet.Name='Production'
        $sheet.Range('A1').Value2='PRINT-RECORDED';$sheet.Range('A2').Formula='=1+2'
        $sheet.Range('A5:E5').Value2=@('PROCESS','OUTPUT','RECALL CODE','LOCAL_NOTE','System_Key')
        $sheet.Range('A6').Value2='PRINT-PROCESS';$sheet.Range('B6').Value2='PRINT-OUTPUT'
        $sheet.Range('C6').Value2='PRINT-RECALL';$sheet.Range('D6').Formula='=2+3';$sheet.Range('E6').Value2=$key
        $table=$sheet.ListObjects.Add(1,$sheet.Range('A5:E6'),$null,1);$table.Name='ProductionOutput'
        $local=$book.Worksheets.Add();$local.Name='InventoryManagement'
        $local.Range('A1:D1').Value2=@('System_Key','ITEM','LOCATION','LOCAL_NOTE')
        $local.Range('A2').Value2=$key;$local.Range('B2').Value2='PRINT-ITEM';$local.Range('C2').Value2=$location;$local.Range('D2').Formula='=3+4'
        $table=$local.ListObjects.Add(1,$local.Range('A1:D2'),$null,1);$table.Name='invSys'
        $savedBook=Join-Path $runRoot 'print-recorded-source.xlsb'
        if(Test-Path -LiteralPath $savedBook){throw 'Preserve existing Print source fixture.'}
        $book.SaveAs($savedBook,50);$bookPin=Hash $savedBook
        $decoy=$excel.Workbooks.Add();$decoy.Worksheets.Item(1).Name='Production'
        $decoy.Worksheets.Item(1).Range('A1').Value2='PRINT-DECOY';$decoy.Worksheets.Item(1).Range('A2').Formula='=5+6'
        $foreign=Fingerprint $decoy.Worksheets.Item(1)
        SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadPolicyForTest' @($true))){throw 'Print recording policy unavailable.'}
        SetRecordingPolicy $true
        $enabled=[IO.File]::ReadAllBytes($Fixture.Config);$enabledPin=Hash $Fixture.Config
        $policy=([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @('PRODUCTION_RUN_PRINT'))).Split('|');$version=[int]$policy[3]
        if($policy[0] -cne 'True' -or $policy[1] -cne 'True' -or $version -lt 1){throw 'Saved Print capture policy prerequisite missing.'}
        SetRecordingPolicy $true
        [IO.File]::WriteAllBytes($candidate,[IO.File]::ReadAllBytes($Fixture.Config))
        $cfg=$excel.Workbooks.Open($candidate,0,$false)
        try{(Table $cfg 'tblEventTrackingPolicies').ListColumns.Item('SchemaVersion').DataBodyRange.Value2=999.0;$cfg.SaveCopyAs($invalid)}finally{$cfg.Close($false)}
        RestoreConfig $enabled
        $catalog=[int](Run 'invSys.Core.xlam' 'TestShippingCatalog.DeclaredCatalogVersionForTest')
        $cases=if($GuideDiagnostic){@('PolicyChanged','Permission','Recovery','GuideRun')}else{@('PolicyChanged','StoreUnavailable','PolicyUnreadable','SignedOut','Target','Permission','ClosedWorkbook','Recovery','GuideRun')}
        foreach($case in $cases){
            $kind=if($case -ceq 'GuideRun'){'Recovery'}else{$case}
            $terminal=$kind -cin @('PolicyChanged','StoreUnavailable','PolicyUnreadable')
            $incomplete=$terminal -or $kind -cin @('SignedOut','Target')
            $label='PrintRecorded.'+$case;$workPath=Join-Path $runRoot ($label+'.xlsb')
            if(Test-Path -LiteralPath $workPath){throw 'Preserve existing Print case workbook.'}
            RestoreConfig $enabled;SelectTarget $Fixture 'config-producer'
            $book.SaveCopyAs($workPath);$saved=Hash $workPath;$work=$excel.Workbooks.Open($workPath,0,$false)
            $name=$work.Name;$source=Fingerprint $work.Worksheets.Item('Production');$projection=Fingerprint $work.Worksheets.Item('InventoryManagement')
            [void](Probe 'OpenDesigner' @($name))
            OpenRecordingViewer
            $prior=@(if(Test-Path -LiteralPath $journalRoot){Get-ChildItem -LiteralPath $journalRoot -Filter '*.json' -File|ForEach-Object FullName})
            if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'Actual recording Start unavailable.'}
            $starts=@(Get-ChildItem -LiteralPath $journalRoot -Filter '*.json' -File|Where-Object {$_.FullName -cnotin $prior}|ForEach-Object {[IO.File]::ReadAllText($_.FullName)|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
            if($starts.Count -ne 1){throw 'One durable Start required.'}
            $before=ActivityPins
            [void](Probe 'RunLocalShowAndCapture' @($name,'PRINT'));[void](Probe 'ResetPrintPreviewForTest')
            [void](Probe 'PrintYieldArm' @($(if($terminal -or $kind -ceq 'Recovery'){''}else{$kind}),$auth))
            if($terminal){[void](Run 'invSys.Core.xlam' 'TestShippingCatalog.CheckTerminalArm' @($kind,$Fixture.Config,$(if($kind -ceq 'PolicyUnreadable'){$invalid}else{$candidate}),$activityRoot))}
            $decoy.Activate()
            try{
                Check ($label+'.ActualHandlerReturned') ([bool](Probe 'PrintYieldAct'))
                if($terminal){
                    Check ($label+'.ExactlyOneTerminalFault') ([int](Run 'invSys.Core.xlam' 'TestShippingCatalog.CheckTerminalHits') -eq 1)
                    Check ($label+'.CapturedContextAtFault') ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.CheckTerminalSameContext'))
                    Check ($label+'.OwnerSuccessAndTrackingNotice') ([bool](Probe 'CheckTerminalMessage' @($kind)))
                }elseif($kind -cne 'Recovery'){
                    if(-not [bool](Probe 'PrintYieldFact' @('Reached')) -or -not [bool](Probe 'PrintYieldFact' @('Valid'))){throw 'Actual preview interruption prerequisite failed; not product RED.'}
                }
                foreach($fact in @('PrintOwnerEntries','PrintReportReads','PrintPreviewCountForTest')){Check ($label+'.'+$fact) ([int](Probe $fact) -eq 1)}
                Check ($label+'.NoReinitialization') ([bool](Probe 'PrintYieldFact' @('NoReinitialization')))
                if($kind -ceq 'ClosedWorkbook'){
                    $work=$null
                    Check ($label+'.NativeWorkbookRemoved') ($name -cnotin @($excel.Workbooks|ForEach-Object Name))
                    Check ($label+'.BindingRejected') (-not [bool](Probe 'RunLocalClosedBindingCurrent'))
                }else{
                    Check ($label+'.ExactKeyAndCustomSourcePreserved') ((Fingerprint $work.Worksheets.Item('Production')) -ceq $source)
                    Check ($label+'.LocalProjectionPreserved') ((Fingerprint $work.Worksheets.Item('InventoryManagement')) -ceq $projection)
                }
                if([bool](Probe 'PrintYieldFact' @('Loaded'))){
                    Check ($label+'.GuardsRestored') ([bool](Probe 'PrintYieldFact' @('GuardsRestored')))
                    if(-not $terminal){
                        $message=if($kind -ceq 'Recovery'){'Print preview closed.'}elseif($kind -ceq 'Permission'){'Production permission changed. Reopen Production before continuing.'}else{'Session, warehouse, or captured workbook changed. Reopen Production before editing the draft.'}
                        if($kind -cin @('SignedOut','Target')){$message+=' Tracking unavailable: the action context is no longer current.'}
                        Check ($label+'.ExactVisibleResult') ([string](Probe 'PrintStatus') -ceq $message)
                    }
                    CaptureOwnedFormByCaptionEvidence 'Production' ($label.ToLowerInvariant()+'.png')
                }else{Check ($label+'.DismissedFormNotRecreated') ($kind -ceq 'ClosedWorkbook')}
            }finally{
                [void](Run 'invSys.Core.xlam' 'TestShippingCatalog.CheckTerminalArm' @('','','',''))
                [void](Probe 'PrintYieldArm' @('',''));RestoreStore
            }
            $records=@(Get-ChildItem -LiteralPath $activityRoot -Filter '*.json' -File|Where-Object {-not $before.ContainsKey($_.Name)}|ForEach-Object {[IO.File]::ReadAllText($_.FullName)|ConvertFrom-Json})
            $expectedCount=if($incomplete){1}else{2};$outcome=if($kind -ceq 'Recovery'){'PREVIEW_RETURNED'}else{'FAILED'}
            $valid=$records.Count -eq $expectedCount -and @($records|Where-Object OutcomeCode -CEQ 'REQUESTED').Count -eq 1
            if(-not $incomplete){$valid=$valid -and @($records|Where-Object OutcomeCode -CEQ $outcome).Count -eq 1 -and $records[0].ActivityId -ceq $records[1].ActivityId}
            foreach($record in $records){$valid=$valid -and $record.ControlId -ceq 'PRODUCTION_RUN_PRINT' -and $record.OwnerId -ceq 'PRODUCTION_RECALL_REPORT' -and $record.UserId -ceq 'config-producer' -and $record.SequenceId -ceq $starts[0].SequenceId -and $record.Ordinal -eq 1 -and $record.PolicyVersion -eq $version -and $record.CatalogVersion -eq $catalog -and @($record.SourceEventRefs).Count -eq 0}
            Check ($label+'.ExactOriginalRecordsWithoutInventedSuccess') $valid
            Check ($label+'.PriorActivityImmutable') (PinsRetained $before)
            if($kind -ceq 'PolicyUnreadable'){
                Check ($label+'.PendingJournalIntact') ((JournalChain $starts[0].SequenceId 2) -and @(RecordingJournal $starts[0].SequenceId|Where-Object Lifecycle -CNE 'Recording').Count -eq 0)
                Check ($label+'.ActualUnavailableViewerRefresh') ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('Refresh','')))
                Check ($label+'.VisibleUnavailablePolicyAndDisabledStop') ((RecordingStatus) -ceq 'Tracking unavailable: the saved tracking policy is invalid.' -and (RecordingControl 'Stop Recording') -ceq 'True|False')
                CaptureOwnedFormByCaptionEvidence ('Viewer - '+$Fixture.Warehouse) 'print-recorded-unreadable-viewer.png'
                RestoreConfig $enabled
                Check ($label+'.ActualRecoveredViewerRefresh') ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('Refresh','')))
            }
            if(-not $incomplete -or $kind -ceq 'PolicyUnreadable'){
                Check ($label+'.ActualStopDelivered') ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')
                Check ($label+'.ExactCaptureStatus') ((RecordingStatus).StartsWith($(if($incomplete){'Incomplete evidence:'}else{'Stopped.'})))
            }
            $closed=@(RecordingJournal $starts[0].SequenceId|Where-Object RecordType -CEQ 'Close')
            $reason=switch($kind){'PolicyChanged'{'POLICY_CHANGED'} 'StoreUnavailable'{'TRACKING_UNAVAILABLE'} 'PolicyUnreadable'{'UNFINISHED_ACTIONS'} 'SignedOut'{'SESSION_CHANGED'} 'Target'{'SESSION_CHANGED'} default{''}}
            $life=if($incomplete){'Incomplete'}else{'Stopped'}
            $valid=$closed.Count -eq 1
            if($valid){$entry=$closed[0];$valid=$entry.Lifecycle -ceq $life -and $entry.ReasonCode -ceq $reason -and $entry.ActionCount -eq 1 -and $entry.CreatedByUserId -ceq 'config-producer' -and @($entry.Observations).Count -eq $expectedCount}
            Check ($label+'.ExactCaptureLifecycle') $valid
            Check ($label+'.ExactImmutableJournalChain') (JournalChain $starts[0].SequenceId ($expectedCount+2))
            if($closed.Count -eq 1){$runs+=@{Label=$label;Journal=$closed[0];Original=$records;Incomplete=$incomplete;Success=($kind -ceq 'Recovery')}}
            CloseRecordingViewer;[void](Probe 'RunLocalSafeClose')
            if($null -ne $work){$work.Close($false);$work=$null}
            [IO.File]::WriteAllBytes($auth,$authBytes);RestoreConfig $enabled
            Check ($label+'.SavedOperatorBytesPreserved') ((Hash $workPath) -ceq $saved)
            Check ($label+'.AuthorityPreservedAfterRestoration') (Same $pins (AuthorityPins))
            Check ($label+'.ConfigBytesRestored') ((Hash $Fixture.Config) -ceq $enabledPin)
            Check ($label+'.DecoyPreserved') ((Fingerprint $decoy.Worksheets.Item(1)) -ceq $foreign)
            Check ($label+'.OtherWarehousePreserved') (RestartPinsEqual $otherPins $Other.Root)
        }
        SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Actual Admin publication unavailable.'}
        SelectTarget $Fixture 'config-reader';OpenRecordingViewer
        $trainingPins=RestartPins (Join-Path $Fixture.Root 'Training')
        $diagnostics=if($GuideDiagnostic){@()}else{$runs}
        foreach($run in $diagnostics){
            $label=$run.Label;$journal=$run.Journal
            if((Select-EvaluationRun $journal.ActionPathId) -cne 'SELECTED'){throw 'Actual recording selection unavailable.'}
            $ready=Set-EvaluationDraft @(,@('PRODUCTION_RUN_PRINT','PREVIEW_RETURNED','True')) 0 'CommandCompleted' -StopAtMissingChoice
            Check ($label+'.ActualExpectationEditor') $ready
            if(-not $ready){continue}
            $before=@(EvaluationFiles|ForEach-Object FullName)
            $clicked=ExpectationControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths'
            $fresh=@(EvaluationFiles|Where-Object {$_.FullName -cnotin $before})
            if($clicked -cne 'DELIVERED' -or $fresh.Count -ne 1){throw 'Actual Evaluate must append one result.'}
            $result=[IO.File]::ReadAllText($fresh[0].FullName)|ConvertFrom-Json
            $state=if($run.Incomplete){'Incomplete'}elseif($run.Success){'Concluded'}else{'Failed'}
            $reason=if($run.Incomplete){'CAPTURE_INCOMPLETE'}elseif($run.Success){'COMMAND_COMPLETED'}else{'REQUIRED_OUTCOME_MISMATCH'}
            Check ($label+'.ExactDiagnosticState') ($result.ResultState -ceq $state -and $reason -cin @($result.ReasonCodes))
            Check ($label+'.ExactRunAndAuthoredIntent') ($result.JournalRecordId -ceq $journal.RecordId -and $result.JournalSha256 -ceq $journal.ContentSha256 -and $result.ExpectedConclusion.TerminalKind -ceq 'CommandCompleted' -and @($result.ExpectedConclusion.Steps).Count -eq 1 -and $result.ExpectedConclusion.Steps[0].RequiredOutcome -ceq 'PREVIEW_RETURNED')
            Check ($label+'.NoInventedCompletionOrBusinessSources') (@($result.Matches).Count -eq $(if($run.Success){1}else{0}) -and @($result.TerminalSources).Count -eq 0)
            CaptureOwnedFormByCaptionEvidence 'Action Paths' ($label.ToLowerInvariant()+'-diagnostic.png')
        }
        if($GuideDiagnostic){Check 'PrintPathsDiagnostic.FourRequiredRunsCaptured' ($runs.Count -eq 4)}
        else{Check 'PrintRecorded.AllEightRunsCaptured' (@($runs|Where-Object Label -CNE 'PrintRecorded.GuideRun').Count -eq 8)}
        . (Join-Path $PSScriptRoot 'Slice4beProductionPrintPaths.ps1')
        Test-ProductionPrintPaths $Fixture $runs
        $retained=$true;foreach($file in $trainingPins.Keys){$retained=$retained -and (Hash $file) -ceq $trainingPins[$file]}
        Check 'PrintRecorded.PriorTrainingImmutableThroughEvaluation' $retained
        Check 'PrintRecorded.AuthorityPreservedThroughPublication' (Same $pins (AuthorityPins))
        Check 'PrintRecorded.SavedSourceBytesPreserved' ((Hash $savedBook) -ceq $bookPin)
        Check 'PrintRecorded.OtherWarehousePreservedThroughPublication' (RestartPinsEqual $otherPins $Other.Root)
    }finally{
        [void](Run 'invSys.Core.xlam' 'TestShippingCatalog.CheckTerminalArm' @('','','',''));[void](Probe 'PrintYieldArm' @('',''))
        RestoreStore;CloseRecordingViewer;[void](Probe 'RunLocalSafeClose')
        if($null -ne $work){$work.Close($false)};if($null -ne $book){$book.Close($false)};if($null -ne $decoy){$decoy.Close($false)}
        [IO.File]::WriteAllBytes($auth,$authBytes);RestoreConfig $original
        foreach($path in @($candidate,$invalid)){if(Test-Path -LiteralPath $path){Remove-Item -LiteralPath $path -Force}}
        SelectTarget $Fixture 'config-producer'
    }
    Check 'PrintRecorded.OriginalConfigBytesRestored' ((Hash $Fixture.Config) -ceq $originalPin)
    Check 'PrintRecorded.AuthBytesRestored' ((Hash $auth) -ceq $authPin)
}
