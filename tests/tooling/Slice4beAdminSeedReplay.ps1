# B1 uses a real source recording and a destination-local execution profile.
function Test-AdminSeedReplayProfile($Fixture,$Destination) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTransferWire.ps1')
    function Deliver($form,$control,$action,$value='') {
        if((BoundControl $control $action $value $form) -cne 'DELIVERED'){throw "Fixture control unavailable: $form.$control"}
    }
    function OpenTarget($target) {
        SelectTarget $target
        foreach($capability in @('ADMIN_MAINT','ACTION_PATH_MAINT')){
            if(-not [bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @($capability,'config-admin',$target.Warehouse,'S1'))){throw 'Required fixture capability unavailable.'}
        }
        if(-not [bool](Run 'invSys.Core.xlam' 'TestSeedRecording.EnablePolicy' @('Capture'))){throw 'Training capture policy unavailable.'}
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Initial publication unavailable.'}
        $anchor.Activate()
        OpenRecordingViewer
    }
    function OnlyGuide($root) {
        $files=@(Get-ChildItem -LiteralPath $root -Filter '*.json' -File)
        if($files.Count -ne 1){throw 'One saved guide required.'}
        $guide=ReadGuideExpectationRecord $files[0].FullName
        if($null -eq $guide -or $guide.Steps.Count -ne 1 -or $guide.Steps[0].ControlId -cne 'ADMIN_DEMO_SEED' -or $guide.ExpectedConclusion.TerminalKind -cne 'SourceEventsApplied'){throw 'Expected a sealed single-step Seed guide.'}
        return $guide
    }
    $editor='frmActionPathGuide'
    $library='frmActionPathLibrary'
    $profileEditor='frmActionPathExecution'
    $export=Join-Path $runRoot 'b1-source-guide.json'
    $inventory=Join-Path $Destination.Root ($Destination.Warehouse+'.invSys.Data.Inventory.xlsb')
    if($Fixture.Warehouse -ceq $Destination.Warehouse -or $Fixture.Root -eq $Destination.Root){throw 'Source and destination must be distinct.'}
    $anchor=$excel.Workbooks.Add()
    try{
        $anchor.SaveAs((Join-Path $runRoot 'b1-host.xlsb'),50)
        $excel.Visible=$true
        OpenTarget $Fixture
        if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'Source recording unavailable.'}
        if(-not [bool](Run 'invSys.Admin.xlam' 'TestSeedOutcome.Prepare' @($Fixture.Warehouse,$false))){throw 'Source inbox unavailable.'}
        try{
            if((Run 'invSys.Admin.xlam' 'TestSeedOutcome.InvokeWorkflow' @('SEED','',$true)) -cne 'OK'){throw 'Ordinary built-in source Seed failed.'}
        }finally{[void](Run 'invSys.Admin.xlam' 'TestSeedOutcome.Release')}
        if((RecordingControl 'Stop Recording' 'Click') -cne 'DELIVERED'){throw 'Source Stop unavailable.'}
        $closed=@(Get-ChildItem -LiteralPath $journalRoot -File -Filter '*.json'|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Close')
        if($closed.Count -ne 1 -or $closed[0].ActionCount -ne 1 -or $closed[0].Lifecycle -cne 'Stopped'){throw 'Single-step source recording unavailable.'}
        $outcome=@($closed[0].Observations|Where-Object {$_.ControlId -ceq 'ADMIN_DEMO_SEED' -and $_.OutcomeCode -ceq 'APPLIED'})
        if($outcome.Count -ne 1 -or $outcome[0].SourceEventRefs.Count -ne 1){throw 'Source recording lacks owner evidence.'}
        if((BoundLibrary 'Open') -cne 'DELIVERED' -or (BoundLibrary 'Select' $closed[0].ActionPathId) -cne 'SELECTED'){throw 'Source recording selection unavailable.'}
        Deliver 'frmActionPaths' 'btnCreateGuide' 'Click'
        Deliver $editor 'txtGuideSubject' 'Write' 'Training inventory'
        Deliver $editor 'txtGuideName' 'Write' 'Seed the built-in R1 Workflow Kit'
        Deliver $editor 'btnGuideExpectedConclusion' 'Click'
        $expect='frmActionPathExpectation'
        Deliver $expect 'cboExpectedControl' 'Write' 'ADMIN_DEMO_SEED'
        Deliver $expect 'cboExpectedOutcome' 'Write' 'APPLIED'
        Deliver $expect 'chkExpectedRetry' 'Check' 'False'
        Deliver $expect 'btnAddExpectedStep' 'Click'
        if((BoundControl 'cboTerminalStep' 'Select' '0' $expect) -cne 'SELECTED'){throw 'Expected terminal step unavailable.'}
        Deliver $expect 'cboTerminalKind' 'Write' 'SourceEventsApplied'
        Deliver $expect 'btnUseExpectation' 'Click'
        Deliver $editor 'btnSaveGuide' 'Click'
        $sourceRoot=Join-Path $journalRoot 'Guides'
        $source=OnlyGuide $sourceRoot
        if($source.SourceRun.ContentSha256 -cne $closed[0].ContentSha256){throw 'Source guide bound the wrong recording.'}
        Deliver $editor 'btnCancelGuide' 'Click'
        BoundOpen
        BoundSelect $source
        TransferSetFile $export
        Deliver $library 'btnExportGuide' 'Click'
        if(-not (Test-Path -LiteralPath $export -PathType Leaf)){throw 'Source export unavailable.'}
        CloseRecordingViewer
        $pins=BoundPins $sourceRoot
        OpenTarget $Destination
        $before=(Get-FileHash -LiteralPath $inventory).Hash
        $authority=$excel.Workbooks.Open($inventory,0,$true)
        try{
            $entities=Table $authority 'tblInventoryEntities'
            if($entities.ListRows.Count -ne 3){throw 'Expected only the generated bootstrap inventory before replay.'}
        }finally{$authority.Close($false)}
        if((BoundLibrary 'Open') -cne 'DELIVERED'){throw 'Destination Action Paths unavailable.'}
        BoundOpen
        Deliver $library 'btnImportGuide' 'Click'
        $destRoot=Join-Path $Destination.Root ('Training/ActionPaths/'+$Destination.Warehouse)
        $guide=OnlyGuide (Join-Path $destRoot 'Guides')
        if($guide.WarehouseId -cne $Destination.Warehouse -or $guide.TransferOrigin.ContentSha256 -cne $source.ContentSha256 -or $guide.TransferOrigin.RecordId -cne $source.RecordId -or $guide.ActionPathId -ceq $source.ActionPathId){throw 'Local guide provenance mismatch.'}
        BoundSelect $guide
        Check 'SeedReplay.Prerequisite.SourceGuideAndFreshSeparateDestination' $true
        Deliver $library 'btnConfigureExecution' 'Click'
        $opened=(BoundControl '' 'Count' '' $profileEditor) -ceq '1'
        if(-not $opened){
            if((BoundControl 'lblPublishedGuideStatus' 'Label') -cne 'Execution is not configured for one or more guide controls. No actions were skipped or run.'){throw 'Unexpected Configure refusal; investigate before classifying RED.'}
            Deliver $library 'btnRunHowTo' 'Click'
            if((BoundControl 'lblPublishedGuideStatus' 'Label') -cne 'Execution not configured for this exact guide version.'){throw 'Unexpected Run setup refusal; no profile RED qualified.'}
        }
        Check 'SeedReplay.Profile.ConfigureOpensForExactSeedGuide' $opened
        if($opened){
            if((BoundControl 'lstExecutionSteps' 'Select' '0' $profileEditor) -cne 'SELECTED'){throw 'Seed profile step unavailable.'}
            Deliver $profileEditor 'cboExecutionInput' 'Write' 'DataSet'
            Deliver $profileEditor 'cboExecutionBinding' 'Write' 'Literal'
            Deliver $profileEditor 'txtExecutionValue' 'Write' 'R1_WORKFLOW_KIT'
            Deliver $profileEditor 'btnApplyExecutionInput' 'Click'
            Deliver $profileEditor 'btnSaveExecutionProfile' 'Click'
            $profiles=@(Get-ChildItem -LiteralPath (Join-Path $destRoot 'ExecutionProfiles') -File -Filter '*.json')
            $exact=$false
            if($profiles.Count -eq 1){
                $saved=Get-Content $profiles[0].FullName -Raw|ConvertFrom-Json
                $step=$saved.Steps[0]
                $exact=$saved.Guide.RecordId -ceq $guide.RecordId -and $saved.Guide.ContentSha256 -ceq $guide.ContentSha256 -and $saved.WarehouseId -ceq $Destination.Warehouse -and $saved.Steps.Count -eq 1 -and $step.StepId -ceq $guide.Steps[0].StepId -and $step.AdapterId -ceq 'ADMIN_DEMO_SEED' -and $step.AdapterVersion -eq 1 -and $step.Inputs.Count -eq 1 -and $step.Inputs[0].Name -ceq 'DataSet' -and $step.Inputs[0].Binding.Kind -ceq 'Literal' -and $step.Inputs[0].Binding.Value -ceq 'R1_WORKFLOW_KIT'
            }
            Check 'SeedReplay.Profile.ExactGuideAndExplicitBuiltInKitSaved' $exact
        }
        Check 'SeedReplay.Profile.PreservesInventoryAndSourceGuide' ((Get-FileHash -LiteralPath $inventory).Hash -ceq $before -and (BoundSame $pins (BoundPins $sourceRoot)))
    }finally{
        TransferSetFile ''
        CloseRecordingViewer
        $anchor.Close($false)
    }
}
