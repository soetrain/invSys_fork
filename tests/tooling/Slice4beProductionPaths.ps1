# Exact original Production actions -> recording -> publication -> actual Viewer/editor.
function Test-ProductionPaths($Fixture,[string]$WorkbookName,[string]$ProcessId,[string]$ProcessVersion,[string]$Canary) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beEvaluationContracts.ps1')
    $cases=@(
        @{Designer='Process';Action='New';Setup='Invalid';Outcome='STAGED'},
        @{Designer='Process';Action='Clear';Setup='Invalid';Outcome='STAGED'},
        @{Designer='Process';Action='Validate';Setup='Valid';Outcome='VALIDATED'},
        @{Designer='Recipe';Action='New';Setup='Invalid';Outcome='STAGED'},
        @{Designer='Recipe';Action='Clear';Setup='Invalid';Outcome='STAGED'},
        @{Designer='Recipe';Action='Validate';Setup='Valid';Outcome='VALIDATED'},
        @{Designer='Process';Action='Validate';Setup='Invalid';Outcome='REJECTED'},
        @{Designer='Recipe';Action='Validate';Setup='Invalid';Outcome='REJECTED'}
    )
    try {
        [void](Probe 'Close');SelectTarget $Fixture
        SetRecordingPolicy $true
        SelectTarget $Fixture 'config-producer'
        OpenRecordingViewer
        [void](Probe 'Open' @($WorkbookName))
        # This adapter creates a bound form without the launcher's InitializeFromProduction.
        # Use the existing Refresh handler to load its real released definitions before Start.
        [void](Probe 'RefreshDesigners')
        $ready=[bool](Probe 'ReleasedRecipe' @($ProcessId,$ProcessVersion,$Canary))
        Check 'ProductionPaths.ReopenedDesignerLoadsReleasedProcess' $ready
        if(-not $ready){throw ('Reopened designer fixture unavailable; stage='+[string](Probe 'RecipeFixtureStatus'))}
        $authorityPins=@{}
        foreach($file in @(Get-ChildItem $Fixture.Root -File -Filter '*.xlsb'|Where-Object {$_.Name -notlike '*.Snapshot.*' -and -not $_.Name.StartsWith('~$')})){$authorityPins[$file.FullName]=SavedHash $file.FullName}
        $startsBefore=@(if(Test-Path $journalRoot){Get-ChildItem $journalRoot -File -Filter '*.json'|Select-Object -ExpandProperty FullName})
        Check 'ProductionPaths.ActualStart' ((RecordingControl 'Start Recording' 'Click') -ceq 'DELIVERED')
        $starts=@(Get-ChildItem $journalRoot -File -Filter '*.json'|Where-Object {$_.FullName -cnotin $startsBefore}|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
        if($starts.Count -ne 1){throw 'Original recording Start unavailable; not terminal-map RED.'}
        $sequence=[string]$starts[0].SequenceId;$original=@();$ordinal=0
        foreach($case in $cases){
            $ordinal++;$case.Id='PRODUCTION_'+$case.Designer.ToUpperInvariant()+'_'+$case.Action.ToUpperInvariant()
            $case.Name=$case.Designer+'.'+$case.Action+'.'+$case.Outcome
            if($case.Setup -ceq 'Invalid'){[void](Probe 'Stage' @($case.Designer,$Canary))}
            elseif($case.Designer -ceq 'Process'){[void](Probe 'ValidProcess')}
            elseif(-not [bool](Probe 'ReleasedRecipe' @($ProcessId,$ProcessVersion,$Canary))){throw ('Actual released-Process recipe unavailable; stage='+[string](Probe 'RecipeFixtureStatus'))}
            $before=@(Get-Slice4beActivityFiles $Fixture)
            $report=[string](Probe 'Act' @($case.Designer,$case.Action))
            Pair $before $case.Id $case.Outcome ('ProductionPaths.'+$case.Name) -ObservedFixture $Fixture
            $rows=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json})
            $attempts=@($rows|Where-Object OutcomeCode -CEQ 'REQUESTED');$outcomes=@($rows|Where-Object OutcomeCode -CEQ $case.Outcome)
            $valid=$rows.Count -eq 2 -and $attempts.Count -eq 1 -and $outcomes.Count -eq 1
            if($valid){$valid=$attempts[0].SequenceId -ceq $sequence -and $outcomes[0].SequenceId -ceq $sequence -and $attempts[0].Ordinal -eq $ordinal -and $outcomes[0].Ordinal -eq $ordinal}
            Check ('ProductionPaths.'+$case.Name+'.ExactRecordingOccurrence') $valid
            if(-not $valid){throw 'Original Production recording incomplete; not terminal-map RED.'}
            $case.Record=$outcomes[0];$original+=@($attempts[0],$outcomes[0])
            if($CaptureProductionDesignerPaths -and $case.Outcome -ceq 'VALIDATED'){
                $page=if($case.Designer -ceq 'Process'){0}else{1}
                $window=[long](Probe 'Present' @($page))
                CaptureOwnedFormEvidence 'Production' ('production-'+$case.Designer.ToLowerInvariant()+'-validated.png') $window
            }
        }
        Check 'ProductionPaths.ActualStop' ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')
        $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close')
        $intact=$closed.Count -eq 1 -and $closed[0].Lifecycle -ceq 'Stopped' -and $closed[0].ActionCount -eq 8 -and (JournalChain $sequence 18)
        if($intact){$intact=($closed[0].Observations.RecordId -join '|') -ceq ($original.RecordId -join '|')}
        Check 'ProductionPaths.EightActionsPreserveExactOriginalOrder' $intact
        if(-not $intact){throw 'Stopped original recording unavailable; not terminal-map RED.'}
        $pathId=[string]$closed[0].ActionPathId;$activityPins=ActivityPins;$journalPins=RestartPins $journalRoot
        [void](Probe 'Close');CloseRecordingViewer;SelectTarget $Fixture
        $published=[bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')
        Check 'ProductionPaths.ActualAdminPublication' $published
        if(-not $published){throw 'Owning publication failed; not terminal-map RED.'}
        $eventsPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')
        $publication=Get-Content $eventsPath -Raw|ConvertFrom-Json
        SelectTarget $Fixture 'config-producer';OpenRecordingViewer
        foreach($case in $cases){
            $id=[string]$case.Record.ActivityId;$group=@($publication.Groups|Where-Object {$_.Source -ceq 'Activity' -and $_.SourceId -ceq $id})
            $retained=$group.Count -eq 1 -and @($group[0].Lines).Count -eq 2
            if($retained){
                # Publication preserves exact records; the journal owns action order.
                $expected=@($original|Where-Object ActivityId -CEQ $id)
                foreach($record in $expected){
                    $line=@($group[0].Lines|Where-Object RecordId -CEQ $record.RecordId)
                    $retained=$retained -and $line.Count -eq 1
                    if($line.Count -eq 1){$retained=$retained -and ($line[0]|ConvertTo-Json -Depth 20 -Compress) -ceq ($record|ConvertTo-Json -Depth 20 -Compress)}
                }
            }
            Check ('ProductionPaths.'+$case.Name+'.PublishedOriginalPair') $retained
            Check ('ProductionPaths.'+$case.Name+'.ActualViewerFindsSource') ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('ContainsSource',$id)))
            Check ('ProductionPaths.'+$case.Name+'.ActualViewerSelectsSource') ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('SelectSource',$id)))
            Check ('ProductionPaths.'+$case.Name+'.DetailPreservesIdentity') ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadDetailForTest' @('Source event / activity ID',$id)))
            if($CaptureProductionDesignerPaths -and $case.Outcome -ceq 'VALIDATED'){
                $outcomeIndex=0
                while($group[0].Lines[$outcomeIndex].OutcomeCode -cne 'VALIDATED'){$outcomeIndex++}
                [void](ExpectationControl 'lstEventLines' 'Index' ([string]$outcomeIndex) 'frmEventDetail')
                CaptureOwnedFormEvidence 'Event Detail' ('production-'+$case.Designer.ToLowerInvariant()+'-detail.png') ([long](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedDetailLabelForTest' @('Window','')))
            }
        }
        . (Join-Path $PSScriptRoot 'Slice4beProductionDiagnostics.ps1')
        Test-ProductionDiagnostics $cases $closed[0] $eventsPath
        Check 'ProductionPaths.OriginalActivityImmutable' (PinsRetained $activityPins)
        $same=$true;foreach($file in $journalPins.Keys){$same=$same -and (Get-FileHash $file).Hash -ceq $journalPins[$file]}
        Check 'ProductionPaths.OriginalJournalImmutable' $same
        $same=$true;foreach($file in $authorityPins.Keys){$same=$same -and (SavedHash $file) -ceq $authorityPins[$file]}
        Check 'ProductionPaths.SavedAuthorityPreserved' $same
    } finally {
        [void](Probe 'Close');CloseRecordingViewer
    }
}
