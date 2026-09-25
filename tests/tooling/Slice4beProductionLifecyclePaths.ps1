# D18 original handler observations and owning publication, before any evaluator edit.
function Test-ProductionLifecyclePaths($Fixture) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beEvaluationContracts.ps1')
    function Probe([string]$Method,[object[]]$Values=@()) {
        if($Method -cnotmatch '^[A-Za-z]+$'){throw 'Invalid fixture method.'}
        [pscustomobject]@{UTC=[DateTimeOffset]::UtcNow.ToString('o');Method=$Method}|ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'lifecycle-boundaries.jsonl')
        Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values
    }
    function SavedHash([string]$Path) {
        $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
        try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}
    }
    $book=$null;$decoy=$null;$cases=@();$original=@();$canary='PATH'+[guid]::NewGuid().ToString('N')
    try {
        SelectTarget $Fixture
        try {
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
            if(-not [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('DesignsEnabled','TRUE'))){throw 'Designs prerequisite unavailable; not RED.'}
        } finally {[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')}
        SetRecordingPolicy $true;SelectTarget $Fixture 'config-producer';OpenRecordingViewer
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$canary
        $book.SaveAs((Join-Path $runRoot 'lifecycle-path-operator.xlsb'),50)
        $decoy=$excel.Workbooks.Add();[void](Probe 'OpenDesigner' @($book.Name));[void](Probe 'RefreshDesigners')
        $ready=[string](Probe 'LifecyclePrepare' @($canary))
        if($ready -cnotmatch '^READY\|([^|]+)\|([^|]+)$'){throw 'Process preparation unavailable; not RED.'}
        $processId=$Matches[1];$processVersion=$Matches[2]
        $before=@(if(Test-Path $journalRoot){Get-ChildItem $journalRoot -File -Filter '*.json'|ForEach-Object FullName})
        Check 'LifecyclePaths.ActualStart' ((RecordingControl 'Start Recording' 'Click') -ceq 'DELIVERED')
        $starts=@(Get-ChildItem $journalRoot -File -Filter '*.json'|Where-Object {$_.FullName -cnotin $before}|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
        if($starts.Count -ne 1){throw 'Actual recording start unavailable; not RED.'}
        $sequence=[string]$starts[0].SequenceId;$ordinal=0
        $actions=@(@('Process','Save','DRAFT'),@('Process','Release','RELEASED'),@('Recipe','Save','DRAFT'),@('Recipe','Release','RELEASED'),@('Recipe','Obsolete','OBSOLETE'),@('Process','Obsolete','OBSOLETE'))
        foreach($action in $actions){
            $designer=$action[0];$verb=$action[1];$id='PRODUCTION_'+$designer.ToUpperInvariant()+'_'+$verb.ToUpperInvariant()
            if($designer -ceq 'Recipe' -and $verb -ceq 'Save'){
                if(-not [bool](Probe 'ReleasedRecipe' @($processId,$processVersion,$canary))){throw 'Released Process prerequisite unavailable; not RED.'}
            }
            foreach($mode in @('BeforeAppend','AfterAppend','Pending','Confirmed')){
                $ordinal++;$name=$designer+'.'+$verb+'.'+$mode;$label='LifecyclePaths.'+$name
                $queue=Join-Path $runRoot ('path-queue-'+$ordinal)
                $writerMode=if($mode -ceq 'Confirmed'){'Observe'}else{$mode}
                $ownerMode=if($mode -ceq 'Confirmed'){''}else{$mode}
                [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterFixtureForTest' @($queue,$writerMode,$Fixture.Warehouse))
                [void](Probe 'LifecycleFaultMode' @($ownerMode))
                $before=@(Get-Slice4beActivityFiles $Fixture);$decoy.Activate()
                $policyBefore=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @($id));$configBefore=SavedHash $Fixture.Config
                $returned=[string](Probe 'LifecycleAct' @($designer,$verb))
                $writer=([string](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterEvidenceForTest')).Split('|')
                $owner=([string](Probe 'LifecycleFaultEvidence')).Split('|')
                $reached=$returned -ceq 'RETURNED' -and $writer.Count -eq 5 -and $writer[1] -ceq '1'
                if($reached){$reached=switch($mode){
                    'BeforeAppend'{$writer[2] -ceq 'False' -and $writer[3] -ceq 'True' -and $writer[4] -ceq 'False'}
                    'AfterAppend'{$writer[2] -ceq 'True' -and $writer[3] -ceq 'True' -and $writer[4] -ceq 'True'}
                    'Pending'{$writer[2] -ceq 'True' -and $writer[4] -ceq 'True' -and $owner[1] -ceq 'True'}
                    'Confirmed'{$writer[2] -ceq 'True' -and [string](Probe 'LifecycleStatus' @($designer)) -ceq $action[2]}
                }}
                Check ($label+'.ActualOwningBoundary') $reached
                if(-not $reached){throw 'Actual owning boundary unavailable; not evaluator RED.'}
                Check ($label+'.ConfigReadPreservesPolicyAndBytes') ((SavedHash $Fixture.Config) -ceq $configBefore -and [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @($id)) -ceq $policyBefore)
                $outcome=if($mode -ceq 'Confirmed'){'CONFIRMED'}elseif($mode -ceq 'Pending'){'PENDING'}else{'FAILED'}
                $rows=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json})
                $attempt=@($rows|Where-Object OutcomeCode -CEQ 'REQUESTED');$result=@($rows|Where-Object OutcomeCode -CEQ $outcome)
                $valid=$rows.Count -eq 2 -and $attempt.Count -eq 1 -and $result.Count -eq 1
                if($valid){$valid=$attempt[0].ActivityId -ceq $result[0].ActivityId -and $attempt[0].ControlId -ceq $id -and $result[0].ControlId -ceq $id -and $attempt[0].SequenceId -ceq $sequence -and $result[0].SequenceId -ceq $sequence -and $attempt[0].Ordinal -eq $ordinal -and $result[0].Ordinal -eq $ordinal}
                Check ($label+'.ExactOriginalOccurrence') $valid
                if(-not $valid){
                    @($rows|ForEach-Object {[pscustomobject]@{ControlId=$_.ControlId;OutcomeCode=$_.OutcomeCode;PolicyVersion=$_.PolicyVersion;Ordinal=$_.Ordinal;SequenceMatches=($_.SequenceId -ceq $sequence)}})|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'unexpected-observation-shape.json')
                    $notice=[string](Probe 'LifecycleNotice');$classification=@{}
                    foreach($fragment in @('context','source references are invalid','tracking policy changed','saved tracking policy is invalid','policy does not include','not authorized','conflicting completion','result could not be recorded','Tracking unavailable','Recording','projection is not visible')){$classification[$fragment]=$notice.Contains($fragment)}
                    $classification.PolicyBefore=$policyBefore
                    $classification.PolicyAfter=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @($id))
                    $classification.ConfigBytesPreserved=(SavedHash $Fixture.Config) -ceq $configBefore
                    $classification|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'tracking-failure-classification.json')
                    throw 'Original observation pair unavailable; not evaluator RED.'
                }
                $refs=@($result[0].SourceEventRefs);$state=if($mode -ceq 'AfterAppend'){'Unknown'}else{'Submitted'}
                $exact=if($mode -ceq 'BeforeAppend'){$refs.Count -eq 0}else{$refs.Count -eq 1 -and $refs[0].EventId -ceq $writer[0] -and $refs[0].SourceKind -ceq 'Designs' -and $refs[0].WarehouseId -ceq $Fixture.Warehouse -and $refs[0].SubmissionState -ceq $state}
                Check ($label+'.ExactOwningReference') ($exact -and @($attempt[0].SourceEventRefs).Count -eq 0)
                $cases+=@{Name=$name;Id=$id;Mode=$mode;Outcome=$outcome;Record=$result[0];Attempt=$attempt[0];WriterId=$writer[0]}
                $original+=@($attempt[0],$result[0])
            }
        }
        [void](Probe 'LifecycleFaultMode' @(''))
        Check 'LifecyclePaths.ActualStop' ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')
        $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close')
        $intact=$closed.Count -eq 1 -and $closed[0].Lifecycle -ceq 'Stopped' -and $closed[0].ActionCount -eq 24 -and (JournalChain $sequence 50)
        if($intact){$intact=($closed[0].Observations.RecordId -join '|') -ceq ($original.RecordId -join '|')}
        Check 'LifecyclePaths.TwentyFourActionsInOriginalOrder' $intact
        if(-not $intact){throw 'Original journal unavailable; not evaluator RED.'}
        Check 'LifecyclePaths.UnknownOperatorColumnPreserved' ($sheet.Cells.Item(1,1).Value2 -ceq 'Operator Extra' -and $sheet.Cells.Item(2,1).Value2 -ceq $canary)
        $activityPins=ActivityPins;$journalPins=RestartPins $journalRoot
        [void](Probe 'CloseDesigner');CloseRecordingViewer;SelectTarget $Fixture
        $published=[bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')
        Check 'LifecyclePaths.ActualAdminPublication' $published
        if(-not $published){throw 'Owning publication unavailable; not evaluator RED.'}
        $eventsPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')
        $publication=Get-Content $eventsPath -Raw|ConvertFrom-Json
        $authorityPins=@{};foreach($file in Get-ChildItem $Fixture.Root -File -Filter '*.xlsb'|Where-Object {$_.Name -notlike '*.Snapshot.*' -and -not $_.Name.StartsWith('~$')}){$authorityPins[$file.FullName]=SavedHash $file.FullName}
        SelectTarget $Fixture 'config-producer';OpenRecordingViewer
        foreach($case in $cases){
            $label='LifecyclePaths.'+$case.Name;$activityId=[string]$case.Record.ActivityId
            $group=@($publication.Groups|Where-Object {$_.Source -ceq 'Activity' -and $_.SourceId -ceq $activityId})
            $same=$group.Count -eq 1 -and @($group[0].Lines).Count -eq 2
            if($same){foreach($record in @($case.Attempt,$case.Record)){
                $line=@($group[0].Lines|Where-Object RecordId -CEQ $record.RecordId)
                $same=$same -and $line.Count -eq 1
                if($line.Count -eq 1){$same=$same -and ($line[0]|ConvertTo-Json -Depth 20 -Compress) -ceq ($record|ConvertTo-Json -Depth 20 -Compress)}
            }}
            Check ($label+'.PublishedOriginalPair') $same
            Check ($label+'.ActualViewerSelectsOriginal') ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('SelectSource',$activityId)))
            Check ($label+'.DetailPreservesIdentity') ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadDetailForTest' @('Source event / activity ID',$activityId)))
            $source=@($publication.Groups|Where-Object {$_.Source -ceq 'Designs' -and $_.SourceId -ceq $case.WriterId})
            $exact=if($case.Mode -ceq 'Confirmed'){$source.Count -eq 1 -and @($source[0].Lines).Count -gt 0 -and @($source[0].Outcomes).Count -eq @($source[0].Lines).Count}else{$source.Count -eq 0}
            if($exact -and $source.Count -eq 1){foreach($line in $source[0].Lines){$exact=$exact -and $line.EventID -ceq $case.WriterId -and $line.WarehouseId -ceq $Fixture.Warehouse -and $line.AppliedAtUTC -cne '' -and $line.AppliedSeq -cmatch '^[1-9][0-9]*$'}}
            Check ($label+'.ExactPublishedDesignsState') $exact
        }
        . (Join-Path $PSScriptRoot 'Slice4beProductionLifecycleDiagnostics.ps1')
        Test-ProductionLifecycleDiagnostics $cases $closed[0] $publication $eventsPath
        Check 'LifecyclePaths.OriginalActivityImmutable' (PinsRetained $activityPins)
        $same=$true;foreach($file in $journalPins.Keys){$same=$same -and (Get-FileHash $file).Hash -ceq $journalPins[$file]}
        Check 'LifecyclePaths.OriginalJournalImmutable' $same
        $same=$true;foreach($file in $authorityPins.Keys){$same=$same -and (SavedHash $file) -ceq $authorityPins[$file]}
        Check 'LifecyclePaths.ReadAndEvaluatePreserveSavedAuthority' $same
    } finally {
        [void](Probe 'LifecycleFaultMode' @(''));[void](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterFixtureForTest' @('','',''))
        [void](Probe 'CloseDesigner');CloseRecordingViewer
        if($null -ne $book){$book.Close($false)}
        if($null -ne $decoy){$decoy.Close($false)}
        SelectTarget $Fixture
    }
}
