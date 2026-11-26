# B1 uses a real source recording and a destination-local execution profile.
function Install-SeedReplayGuardProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $module=$project.VBComponents.Item('frmActionPathRun').CodeModule
    $start=$module.ProcStartLine('DispatchSteps',0)
    $count=$module.ProcCountLines('DispatchSteps',0)
    $code=$module.Lines($start,$count)
    $owner='"''invSys.Admin.xlam''!modAdminSeedReplay.ApplyBuiltIn"'
    $call="delivered = Application.Run($owner, mToken, stepId, inputs)"
    if(-not $code.Contains($call)){return}
    $probe=$project.VBComponents.Add(1)
    $probe.Name='TestSeedBridgeGuard'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Public Refused As Long
Public Function CountRefused() As Long
    CountRefused = Refused
End Function
'@)
    $checks=@(
        "If Not Application.Run($owner, mToken, stepId, inputs & ""-wrong"") Then TestSeedBridgeGuard.Refused = TestSeedBridgeGuard.Refused + 1",
        "If Not Application.Run($owner, mToken, stepId & ""-wrong"", inputs) Then TestSeedBridgeGuard.Refused = TestSeedBridgeGuard.Refused + 1",
        "If Not Application.Run($owner, mToken & ""-wrong"", stepId, inputs) Then TestSeedBridgeGuard.Refused = TestSeedBridgeGuard.Refused + 1",
        $call,
        "If Not Application.Run($owner, mToken, stepId, inputs) Then TestSeedBridgeGuard.Refused = TestSeedBridgeGuard.Refused + 1"
    )
    $module.DeleteLines($start,$count)
    $module.InsertLines($start,$code.Replace($call,($checks -join "`r`n            ")))
}

function Test-AdminSeedReplayProfile($Fixture,$Destination) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTransferWire.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beSeedReplayProof.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beAuthReadOnly.ps1')
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
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $expectedBuilds=@{}
        foreach($package in $packages.Keys){
            $expectedPath=Join-Path $deploy $package
            if($packages[$package].FullName -ine $expectedPath){throw 'Wrong loaded package path.'}
            $archive=[IO.Compression.ZipFile]::OpenRead((Join-Path $inputDeploy $package))
            try{
                $reader=[IO.StreamReader]::new($archive.GetEntry('docProps/custom.xml').Open())
                try{$properties=([xml]$reader.ReadToEnd()).DocumentElement.ChildNodes}finally{$reader.Dispose()}
                $expectedBuilds[$package]=($properties|Where-Object {$_.GetAttribute('name') -ceq 'invSysBuildIdentity'}).InnerText
            }finally{$archive.Dispose()}
        }
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
        $sourceInventory=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
        $sourceHash=Get-AuthReadFixtureHash $sourceInventory
        if($AdminSeedReplayRun){$reference=Read-SeedReplayInventory $sourceInventory}
        $pins=BoundPins $sourceRoot
        OpenTarget $Destination
        $before=(Get-FileHash -LiteralPath $inventory).Hash
        if($AdminSeedReplayRun){$baseline=Read-SeedReplayInventory $inventory}
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
        if($AdminSeedReplayRun){
            if(-not $opened -or -not $exact){throw 'A saved bound profile is required for the replay RED.'}
            Deliver $profileEditor 'btnCloseExecution' 'Click'
            Deliver $library 'btnRunHowTo' 'Click'
            $runner='frmActionPathRun'
            if((BoundControl '' 'Count' '' $runner) -cne '1' -or -not (BoundControl 'lblRunGuide' 'Label' '' $runner).Contains($guide.ContentSha256) -or -not (BoundControl 'lblRunProfile' 'Label' '' $runner).Contains($saved.ContentSha256) -or -not (BoundControl 'lblRunTarget' 'Label' '' $runner).Contains($Destination.Warehouse)){throw 'Actual runner did not bind the exact destination, guide and profile.'}
            $ready=(BoundControl 'btnStartRun' 'State' '' $runner) -ceq 'True|True'
            Check 'SeedReplay.Run.StartNeedsNoReceivingEntity' $ready
            if(-not $ready){return}
            Deliver $runner 'btnStartRun' 'Click'
            $attempts=@(Get-ChildItem (Join-Path $destRoot 'Runs') -File -Filter '*.json'|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Sort-Object Revision)
            if(-not $attempts.Count){throw 'No durable replay attempt; dispatch RED unavailable.'}
            $attempt=$attempts[-1]
            if($attempt.Guide.RecordId -cne $guide.RecordId -or $attempt.Profile.RecordId -cne $saved.RecordId -or $attempt.WarehouseId -cne $Destination.Warehouse -or $attempt.Steps.Count -ne 1 -or $attempt.Steps[0].StepId -cne $guide.Steps[0].StepId -or $attempt.Steps[0].ControlId -cne 'ADMIN_DEMO_SEED'){throw 'Durable run/step identity did not match.'}
            $applied=$attempt.Steps[0].ReasonCode -ceq 'APPLIED' -and $attempt.Steps[0].ActivityId -cne '' -and $attempt.Steps[0].SourceEventRefs.Count -eq 1
            Check 'SeedReplay.Run.OrdinaryOwnerAppliedWithFreshEvent' $applied
            if($applied){Check 'SeedReplay.Bridge.RefusesAlteredKitStepTokenAndDuplicate' ((Run 'invSys.Operations.xlam' 'TestSeedBridgeGuard.CountRefused') -eq 4)}
            if(-not $applied -and ($attempt.State -cne 'Unknown' -or $attempt.Steps[0].ReasonCode -cne 'CAPTURE_UNAVAILABLE' -or $attempt.Steps[0].SourceEventRefs.Count -ne 0 -or (Get-FileHash $inventory).Hash -cne $before)){throw 'Unexpected replay failure; inspect before classifying RED.'}
            $effect=$false
            if($applied){
                $event=$attempt.Steps[0].SourceEventRefs[0]
                $journal=@(Get-ChildItem $destRoot -File -Filter '*.json'|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object {$_.RecordType -ceq 'Close' -and $_.SequenceId -ceq $attempt.Recording.SequenceId})
                $pair=@($journal[0].Observations|Where-Object {$_.ActivityId -ceq $attempt.Steps[0].ActivityId -and $_.WarehouseId -ceq $Destination.Warehouse -and $_.SequenceId -ceq $attempt.Recording.SequenceId -and $_.ControlId -ceq 'ADMIN_DEMO_SEED'})
                Check 'SeedReplay.Proof.FiveExactBuildsAndCurrentOwnerPair' ($attempt.PackageBuilds.Count -eq 5 -and @($attempt.PackageBuilds|Where-Object {$_.BuildIdentity -cne $expectedBuilds[$_.PackageId]}).Count -eq 0 -and $journal.Count -eq 1 -and $journal[0].Observations.Count -eq 2 -and $pair.Count -eq 2 -and $pair[0].OutcomeCode -ceq 'REQUESTED' -and $pair[1].OutcomeCode -ceq 'APPLIED' -and $pair[1].SourceEventRefs.Count -eq 1 -and $pair[1].SourceEventRefs[0].EventId -ceq $event.EventId -and $event.SubmissionState -ceq 'Submitted')
                $after=Read-SeedReplayInventory $inventory
                $effect=$attempt.State -ceq 'Completed' -and $event.WarehouseId -ceq $Destination.Warehouse -and $event.EventId -cne $outcome[0].SourceEventRefs[0].EventId -and $attempt.Steps[0].ActivityId -cne $outcome[0].ActivityId -and (Test-SeedReplayInventory $reference $baseline $after $event.EventId $Destination)
                Deliver 'frmInventoryViewer' 'btnRefresh' 'Click'
                $evaluation=Read-SeedRunEvaluation $attempt $guide $destRoot
                $refs=@($evaluation.TerminalSources)
                $keys=@($after.tblInventoryLog|Where-Object EventID -CEQ $event.EventId|ForEach-Object System_Key|Sort-Object)
                $effect=$effect -and $evaluation.ResultState -ceq 'Concluded' -and $refs.Count -eq 1 -and $refs[0].OwnerStatus -ceq 'Applied' -and $refs[0].EventId -ceq $event.EventId -and $refs[0].LineCount -eq 21 -and (($refs[0].SystemKeys|Sort-Object) -join '|') -ceq ($keys -join '|')
            }
            Check 'SeedReplay.Run.ExactSavedInventoryAndVerifiedCurrentEvent' $effect
            Test-SeedReplayStoppedWithStock $guide $saved $attempt $destRoot $inventory $Destination $reference $baseline $applied
            Check 'SeedReplay.Run.PreservesSourceGuideAndWarehouse' ((BoundSame $pins (BoundPins $sourceRoot)) -and (Get-AuthReadFixtureHash $sourceInventory) -ceq $sourceHash)
        }
    }finally{
        TransferSetFile ''
        CloseRecordingViewer
        $anchor.Close($false)
    }
}
