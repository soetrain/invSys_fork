# Companion to the unchanged 349-check baseline. Real writer/processor boundaries
# are observed through the existing disposable lifecycle failure seams.
function Test-ProductionAssignmentSafety($Fixture,$Other) {
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    function Hash([string]$Path){$s=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $s).Hash}finally{$s.Dispose()}}
    function SourceRows {
        $wire=[string](Run 'invSys.Designs.Domain.xlam' 'modDesignsBridgeApi.ReadDesignsQueryBridgeResult' @('PUBLICATION_EVENTS',$Fixture.Warehouse,$Fixture.Root))
        $lines=@($wire -split '\r?\n')
        if($lines.Count -lt 2 -or $lines[0] -notmatch '^EVTSRC1\tDesigns\tAvailable\t'){throw 'Owning Designs source unavailable; not behavioral RED.'}
        $headers=$lines[1].Split("`t")
        @(foreach($line in $lines|Select-Object -Skip 2){if($line -ne ''){$fields=$line.Split("`t");if($fields.Count -ne $headers.Count){throw 'Designs source field count mismatch.'};$row=@{};for($i=0;$i -lt $headers.Count;$i++){$row[$headers[$i]]=$fields[$i]};[pscustomobject]$row}})
    }
    $canary='ASSIGNFAULT'+[guid]::NewGuid().ToString('N');$book=$null;$decoy=$null;$recordPins=@{}
    $faultRoot=Join-Path $runRoot 'assignment-fault-queues'
    New-Item -ItemType Directory -Path $faultRoot|Out-Null
    try{
        SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.AssignmentPolicyForTest' @($true,$true))){throw 'Assignment safety policy unavailable; not product RED.'}
        SelectTarget $Fixture 'config-producer'
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$canary;$sheet.Cells.Item(2,2).Formula='=1+2'
        $path=Join-Path $runRoot 'assignment-safety-operator.xlsb';$book.SaveAs($path,50);$book.Close($false)
        $bookPin=Hash $path;$book=$excel.Workbooks.Open($path,0,$false);$sheet=$book.Worksheets.Item(1)
        $decoy=$excel.Workbooks.Add();$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name))
        if(-not [bool](Probe 'AssignmentPrepare' @($canary))){throw 'Released Assignment safety fixture unavailable; not product RED.'}
        foreach($file in Files){$recordPins[$file]=Hash $file}
        foreach($mode in @('Observe','BeforeAppend','AfterAppend','Pending','Refresh')){
            [void](Probe 'AssignmentStage' @('Normal'))
            $queueRoot=Join-Path $faultRoot ([guid]::NewGuid().ToString('N'))
            [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterFixtureForTest' @($queueRoot,$mode,$Fixture.Warehouse))
            [void](Probe 'LifecycleFaultMode' @($mode))
            $before=@(Files);$sourceBefore=@(SourceRows)
            $returned=[string](Probe 'AssignmentAct' @('SAVE'))
            $writer=([string](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterEvidenceForTest')).Split('|')
            $owner=([string](Probe 'LifecycleFaultEvidence')).Split('|')
            $case='AssignmentSafety.Save.'+$mode
            $reached=$writer.Count -eq 5 -and $writer[0] -cmatch '^[A-Za-z0-9_-]+$' -and $writer[1] -ceq '1' -and $owner.Count -eq 3
            if($reached){$reached=switch($mode){
                'BeforeAppend'{$writer[2] -ceq 'False' -and $writer[3] -ceq 'True'}
                'AfterAppend'{$writer[2] -ceq 'True' -and $writer[3] -ceq 'True'}
                'Pending'{$writer[2] -ceq 'True' -and $owner[0] -ceq $writer[0] -and $owner[1] -ceq 'True'}
                'Refresh'{$writer[2] -ceq 'True' -and $owner[0] -ceq $writer[0] -and $owner[2] -ceq 'True'}
                'Observe'{$writer[2] -ceq 'True' -and $owner[0] -ceq $writer[0] -and $owner[1] -ceq 'False' -and $owner[2] -ceq 'False'}
            }}
            if(-not $reached){throw ('Assignment owner boundary not reached: '+$mode+'; not product RED.')}
            Check ($case+'.RealBoundaryReached') (-not $returned.StartsWith('HANDLER_ERROR|'))
            $source=@(SourceRows|Where-Object{$_.EventID -cnotin @($sourceBefore|ForEach-Object EventID)})
            $stored=switch($mode){
                'BeforeAppend'{$writer[4] -ceq 'False' -and $source.Count -eq 0}
                {$_ -in @('Observe','Refresh')}{$source.Count -eq 1 -and $source[0].EventID -ceq $writer[0] -and $source[0].WarehouseId -ceq $Fixture.Warehouse -and $source[0].EventType -ceq 'PROCESS_SAVE' -and $source[0].AppliedAtUTC -ne ''}
                default{$writer[4] -ceq 'True' -and $source.Count -eq 0}
            }
            Check ($case+'.ActualWriteAndApplicationBoundary') $stored
            $expected=if($mode -ceq 'Observe'){'CONFIRMED'}elseif($mode -in @('Pending','Refresh')){'PENDING'}else{'FAILED'}
            $state=if($mode -ceq 'BeforeAppend'){''}elseif($mode -ceq 'AfterAppend'){'Unknown'}else{'Submitted'}
            $raw=@(Files|Where-Object{$_ -cnotin $before}|ForEach-Object{[IO.File]::ReadAllText($_)})
            $rows=@($raw|ForEach-Object{$_|ConvertFrom-Json});$attempt=@($rows|Where-Object OutcomeCode -CEQ 'REQUESTED');$outcome=@($rows|Where-Object OutcomeCode -CEQ $expected)
            $pair=$rows.Count -eq 2 -and $attempt.Count -eq 1 -and $outcome.Count -eq 1
            $exact=$pair;$linked=$pair;$effect=$pair;$safe=$pair;$terminal=$false
            if($pair){
                $refs=@($outcome[0].SourceEventRefs)
                $exact=if($state -ceq ''){$refs.Count -eq 0}else{$refs.Count -eq 1 -and $refs[0].EventId -ceq $writer[0] -and $refs[0].WarehouseId -ceq $Fixture.Warehouse -and $refs[0].SourceKind -ceq 'Designs' -and $refs[0].SubmissionState -ceq $state}
                $exact=$exact -and @($attempt[0].SourceEventRefs).Count -eq 0
                $linked=$attempt[0].ActivityId -cne '' -and $attempt[0].ActivityId -ceq $outcome[0].ActivityId -and $attempt[0].RecordId -cne $outcome[0].RecordId
                $severity=if($expected -ceq 'CONFIRMED'){'Info'}elseif($expected -ceq 'PENDING'){'Notice'}else{'Error'}
                $effect=$outcome[0].DataEffect -ceq 'Unknown' -and $outcome[0].Severity -ceq $severity -and $outcome[0].ControlId -ceq 'PRODUCTION_ASSIGNMENT_SAVE' -and $outcome[0].OwnerId -ceq 'PRODUCTION_ASSIGNMENT' -and $outcome[0].EventCode -ceq ('PRODUCTION_ASSIGNMENT_SAVE_'+$expected)
                $completed=[bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @(($outcome[0]|ConvertTo-Json -Depth 20 -Compress)))
                $terminal=$completed -eq ($mode -ceq 'Observe')
            }
            foreach($text in $raw){foreach($value in @($canary,$Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,$queueRoot,'Injected lifecycle','mBtn','ASSIGN-OLD')){if($text.Contains($value)){$safe=$false}}}
            Check ($case+'.ExactAttemptAndOutcome') $pair
            Check ($case+'.ExactOwnerReferenceOrNoUnusedId') $exact
            Check ($case+'.DistinctCorrelatedRecords') $linked
            Check ($case+'.UnknownEffectWithoutRollbackClaim') $effect
            Check ($case+'.NoEnteredDataOrRawFailure') $safe
            Check ($case+'.OnlyConfirmedCommandCompletes') $terminal
            Check ($case+'.UnknownValuesPreserved') ($sheet.Cells.Item(1,1).Value2 -ceq 'Operator Extra' -and $sheet.Cells.Item(2,1).Value2 -ceq $canary -and $sheet.Cells.Item(2,2).Formula -ceq '=1+2')
            [void](Probe 'LifecycleFaultMode' @(''))
        }
        Test-ProductionAssignmentPolicy $Fixture $Other $book $sheet $canary
        $same=$true;foreach($file in $recordPins.Keys){$same=$same -and (Hash $file) -ceq $recordPins[$file]};Check 'AssignmentSafety.OlderRecordsImmutable' $same
        [void](Probe 'CloseDesigner');$book.Close($false);$book=$null
        Check 'AssignmentSafety.OperatorBytesPreserved' ((Hash $path) -ceq $bookPin)
    }finally{
        [void](Probe 'LifecycleFaultMode' @(''))
        [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterFixtureForTest' @('','',''))
        [void](Probe 'CloseDesigner');if($null -ne $book){$book.Close($false)};if($null -ne $decoy){$decoy.Close($false)}
        SelectTarget $Fixture
    }
}
