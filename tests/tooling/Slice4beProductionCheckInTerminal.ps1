# Faults occur after the real owner and REQUESTED, before Core's terminal read/write.
# Instrumentation is unsaved and restricted to the controller's disposable paths.
function Install-ProductionCheckInTerminalProbe {
    param([ValidateSet('PRODUCTION_RUN_CHECK_IN','PRODUCTION_RUN_NEXT_BATCH')][string]$ControlId='PRODUCTION_RUN_CHECK_IN')
    . (Join-Path $PSScriptRoot 'Slice4beProductionPathsProbe.ps1')
    Install-ProductionPathsProbe
    $core=$packages['invSys.Core.xlam'].VBProject
    $adapter=$core.VBComponents.Item('TestShippingCatalog').CodeModule
    $adapter.InsertLines(1,'Private mTerminalKind As String, mTerminalConfig As String, mTerminalCandidate As String, mTerminalStore As String, mTerminalHits As Long, mTerminalSameContext As Boolean')
    $adapter.AddFromString(@'
Public Sub CheckTerminalArm(ByVal kind As String, ByVal config As String, ByVal candidate As String, ByVal store As String)
    mTerminalKind = kind: mTerminalConfig = config: mTerminalCandidate = candidate: mTerminalStore = store
    mTerminalHits = 0: mTerminalSameContext = False
End Sub
Public Sub CheckTerminalBoundary()
    Dim fs As Object, stream As Object, kind As String, context As String
    If mTerminalKind = "" Then Exit Sub
    kind = mTerminalKind: mTerminalKind = "": mTerminalHits = mTerminalHits + 1
    context = modActivity.CaptureContext()
    Set fs = CreateObject("Scripting.FileSystemObject")
    If kind = "PolicyChanged" Or kind = "PolicyUnreadable" Then
        fs.CopyFile mTerminalCandidate, mTerminalConfig, True
    ElseIf kind = "StoreUnavailable" Then
        fs.MoveFolder mTerminalStore, mTerminalStore & "-terminal-held"
        Set stream = fs.CreateTextFile(mTerminalStore, False)
        stream.Write "Disposable terminal-write fault": stream.Close
    Else
        Err.Raise 5
    End If
    mTerminalSameContext = (context <> "" And context = modActivity.CaptureContext())
End Sub
Public Function CheckTerminalHits() As Long
    CheckTerminalHits = mTerminalHits
End Function
Public Function CheckTerminalSameContext() As Boolean
    CheckTerminalSameContext = mTerminalSameContext
End Function
'@)
    $module=$core.VBComponents.Item('modActivity').CodeModule
    $start=$module.ProcStartLine('FinishAction',0);$end=$start+$module.ProcCountLines('FinishAction',0)
    $lines=@(for($i=$start;$i -lt $end;$i++){
        if($module.Lines($i,1).Trim() -ieq 'If Not modActivityPolicy.ReadPolicy(target, action("ControlId"), version, collect, visible, notice) Then GoTo CleanExit'){$i}
    })
    if($lines.Count -ne 1){throw 'Terminal observation boundary changed; not product RED.'}
    $module.InsertLines($lines[0],('    If action("ControlId") = "'+$ControlId+'" Then TestShippingCatalog.CheckTerminalBoundary'))
    $messageProbe=@'
Public Function CheckTerminalMessageForTest(ByVal kind As String) As Boolean
    Dim notice As String
    If kind = "PolicyChanged" Then
        notice = "Tracking unavailable: the tracking policy changed during this action."
    ElseIf kind = "PolicyUnreadable" Then
        notice = "Tracking unavailable: the saved tracking policy is invalid."
    ElseIf kind = "StoreUnavailable" Then
        notice = "Tracking unavailable: the training record could not be saved."
    Else
        Exit Function
    End If
    CheckTerminalMessageForTest = InStr(1, mTxtStatus.Text, "Checked in ", vbBinaryCompare) = 1 And _
        Right$(mTxtStatus.Text, Len(notice) + 1) = " " & notice
End Function
'@
    if($ControlId -ceq 'PRODUCTION_RUN_NEXT_BATCH'){$messageProbe=$messageProbe.Replace('"Checked in "','"Next Batch "')}
    $packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmProduction').CodeModule.AddFromString($messageProbe)
    $packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function CheckTerminalMessage(ByVal kind As String) As Boolean
    CheckTerminalMessage = mForm.CheckTerminalMessageForTest(kind)
End Function
'@)
}

function Test-ProductionCheckInTerminal($Fixture,$Other,$Book,$Decoy,[string]$SelectedKey,[string]$Canary,[string[]]$Keys,[switch]$NextBatch){
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beEvaluationContracts.ps1')
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Hash([string]$Path){$s=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $s).Hash}finally{$s.Dispose()}}
    $controlId='PRODUCTION_RUN_CHECK_IN';$prefix='CheckInTerminal'
    $catalog=[int](Run 'invSys.Core.xlam' 'TestShippingCatalog.DeclaredCatalogVersionForTest')
    if($NextBatch){
        $controlId='PRODUCTION_RUN_NEXT_BATCH';$prefix='NextTerminal'
        $catalog=[int](Run 'invSys.Core.xlam' 'TestShippingCatalog.NextPolicyCatalogVersion')
        . (Join-Path $PSScriptRoot 'Slice4beProductionNextPaths.ps1')
        function SavedHash([string]$Path){Hash $Path}
        function PathCheck([string]$Name,[bool]$Passed){Check ($Name.Replace('InstructionPaths.','')) $Passed}
    }
    function RestoreStore {
        if(Test-Path -LiteralPath $held -PathType Container){
            if(Test-Path -LiteralPath $activityRoot -PathType Leaf){Remove-Item -LiteralPath $activityRoot -Force}
            if(Test-Path -LiteralPath $activityRoot){throw 'Unexpected occupied store; preserve the held evidence.'}
            Move-Item -LiteralPath $held -Destination $activityRoot
        }
    }
    function RestoreConfig([byte[]]$Bytes){
        if(@($excel.Workbooks|Where-Object{$_.FullName -ieq $Fixture.Config}).Count){throw 'Refuse to overwrite an open Config fixture.'}
        [IO.File]::WriteAllBytes($Fixture.Config,$Bytes)
    }
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    $root=[IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\'
    $held=$activityRoot+'-terminal-held';$candidate=Join-Path $Fixture.Root 'check-in-terminal-policy.xlsb'
    $invalidCandidate=Join-Path $Fixture.Root 'check-in-terminal-invalid-policy.xlsb'
    if(-not $root.StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Terminal fixture escaped the owned controller root.'}
    foreach($file in @($Fixture.Config,$activityRoot,$held,$candidate,$invalidCandidate)){
        if(-not [IO.Path]::GetFullPath($file).StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw 'Terminal path escaped the disposable fixture.'}
    }
    if((Test-Path -LiteralPath $held) -or (Test-Path -LiteralPath $candidate) -or (Test-Path -LiteralPath $invalidCandidate)){throw 'Preserve existing terminal fixture artifacts.'}
    $original=[IO.File]::ReadAllBytes($Fixture.Config);$originalHash=Hash $Fixture.Config
    $otherPins=RestartPins $Other.Root;$runs=@();$nextWork=$null
    try{
        # CheckBaselineReopen first copies the source fixture from the live form.
        # Preserve it until that adapter has rebound it to the next actor.
        SelectTarget $Fixture
        if($NextBatch -and -not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadPolicyForTest' @($true))){throw 'Next Batch command policy unavailable; not product RED.'}
        SetRecordingPolicy $true
        $enabled=[IO.File]::ReadAllBytes($Fixture.Config);$enabledHash=Hash $Fixture.Config
        # Build a valid external policy version ahead of time, preserving every
        # historical row. The terminal hook installs it without switching actor.
        $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
        try{
            $headers=Table $cfg 'tblEventTrackingPolicies';$controls=Table $cfg 'tblEventTrackingControls'
            $pv=$headers.ListColumns.Item('PolicyVersion').Index;$cv=$controls.ListColumns.Item('PolicyVersion').Index
            $latest=0;$source=0
            for($i=1;$i -le $headers.ListRows.Count;$i++){
                $v=[int]$headers.ListRows.Item($i).Range.Cells.Item(1,$pv).Value2
                if($v -gt $latest){$latest=$v;$source=$i}
            }
            if($latest -lt 1){throw 'Saved capture policy required before fault setup.'}
            $row=$headers.ListRows.Add()
            [void]$headers.ListRows.Item($source).Range.Copy($row.Range)
            $row.Range.Cells.Item(1,$pv).Value2=[double]($latest+1)
            $newPolicyRow=$headers.ListRows.Count
            $count=$controls.ListRows.Count
            for($i=1;$i -le $count;$i++){
                if([int]$controls.ListRows.Item($i).Range.Cells.Item(1,$cv).Value2 -eq $latest){
                    $row=$controls.ListRows.Add()
                    [void]$controls.ListRows.Item($i).Range.Copy($row.Range)
                    $row.Range.Cells.Item(1,$cv).Value2=[double]($latest+1)
                }
            }
            if($NextBatch){
                # Current schema-2 policies retain their per-user rows in the new version.
                $users=Table $cfg 'tblEventTrackingUsers';$uv=$users.ListColumns.Item('PolicyVersion').Index
                $count=$users.ListRows.Count
                for($i=1;$i -le $count;$i++){
                    if([int]$users.ListRows.Item($i).Range.Cells.Item(1,$uv).Value2 -eq $latest){
                        $row=$users.ListRows.Add();[void]$users.ListRows.Item($i).Range.Copy($row.Range)
                        $row.Range.Cells.Item(1,$uv).Value2=[double]($latest+1)
                    }
                }
            }
            $cfg.SaveCopyAs($candidate)
            $headers.ListRows.Item($newPolicyRow).Range.Cells.Item(1,$headers.ListColumns.Item('SchemaVersion').Index).Value2=999.0
            $cfg.SaveCopyAs($invalidCandidate)
        }finally{$cfg.Close($false)}
        Write-Output ($prefix+': terminal fault policy fixture prepared.')
        foreach($kind in @('PolicyChanged','StoreUnavailable','PolicyUnreadable')){
            foreach($mode in @('Reusable','Worksheet')){
                RestoreConfig $enabled
                SelectTarget $Fixture 'config-producer'
                if($NextBatch){
                    $nextWork=$excel.Workbooks.Add();$nextSheet=$nextWork.Worksheets.Item(1)
                    $nextSheet.Cells.Item(2,1).Value2=$Canary;$nextSheet.Cells.Item(2,2).Formula='=1+2'
                    [void](Probe 'RunLocalReopen' @($nextWork.Name));Prepare-NextPath $nextWork $Canary $mode
                    $ownerPins=Get-NextPathAuthorityPins $Fixture
                }else{
                    [void](Probe 'CheckBaselineReopen' @($Book.Name))
                    $ready=if($mode -ceq 'Reusable'){[bool](Probe 'CheckBaselineReusableStage' @('Selected'))}else{[bool](Probe 'CheckBaselineWorksheetStage' @($SelectedKey,$Canary))}
                    if(-not $ready){throw 'Terminal owner prerequisites unavailable; not product RED.'}
                }
                OpenRecordingViewer
                $prior=@(if(Test-Path -LiteralPath $journalRoot){Get-ChildItem -LiteralPath $journalRoot -File -Filter '*.json'|ForEach-Object FullName})
                if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'Actual recording Start unavailable.'}
                $starts=@(Get-ChildItem -LiteralPath $journalRoot -File -Filter '*.json'|Where-Object{$_.FullName -cnotin $prior}|ForEach-Object{[IO.File]::ReadAllText($_.FullName)|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
                if($starts.Count -ne 1){throw 'Exactly one real recording Start required.'}
                $before=ActivityPins;$label=$prefix+'.'+$kind+'.'+$mode
                Write-Output ($label+': invoking actual handler after recording Start.')
                [void](Probe 'CheckBaselineResetOwnerEntries')
                $activeBook=if($NextBatch){$nextWork}else{$Book}
                [void](Probe 'RunLocalShowAndCapture' @($activeBook.Name,'CHECK_IN'));$Decoy.Activate()
                $selectedPolicy=if($kind -ceq 'PolicyUnreadable'){$invalidCandidate}else{$candidate}
                [void](Run 'invSys.Core.xlam' 'TestShippingCatalog.CheckTerminalArm' @($kind,$Fixture.Config,$selectedPolicy,$activityRoot))
                try{
                    if($NextBatch){Invoke-NextPath $Fixture $Canary $mode $label $ownerPins -ChangingPolicy $Fixture.Config}
                    else{Check ($label+'.ActualHandlerReturned') ([bool](Probe 'CheckBaselineAct' @('')))}
                    Check ($label+'.ExactlyOneTerminalFault') ([int](Run 'invSys.Core.xlam' 'TestShippingCatalog.CheckTerminalHits') -eq 1)
                    Check ($label+'.CapturedContextPreservedAtFault') ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.CheckTerminalSameContext'))
                    if(-not $NextBatch){
                      Check ($label+'.OwnerEnteredOnce') ([int](Probe 'CheckBaselineOwnerEntries') -eq 1)
                      if($mode -ceq 'Reusable'){
                        Check ($label+'.CheckedFrozenWithoutCompletion') ([bool](Probe 'CheckBaselineReusableResult' @($true)))
                        Check ($label+'.ExactIdentityDisplay') ([bool](Probe 'CheckBaselineReusableDisplay' @($Keys[0],$Keys[1])))
                      }else{
                        foreach($fact in @('ReachedCheckRows','ExactSelectedKey','CustomValue','CustomFormula','PalettePreserved','HeadersPreserved','DisplayColumns')){Check ($label+'.'+$fact) ([bool](Probe 'CheckBaselineWorksheetFact' @($fact)))}
                      }
                    }
                    Check ($label+'.ExactSuccessAndFailureNotice') ([bool](Probe 'CheckTerminalMessage' @($kind)))
                    # NextActivityAct checks its own guards before its adapter resets them.
                    if(-not $NextBatch){Check ($label+'.GuardsRestored') ([bool](Probe 'CheckBaselineGuards'))}
                    if(-not $NextBatch){Check ($label+'.CanonicalEntitiesPreserved') ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.StockSourcePreservedForTest'))}
                    $capturePrefix=if($NextBatch){'next-terminal-'}else{'check-in-terminal-'}
                    CaptureOwnedFormByCaptionEvidence 'Production' ($capturePrefix+$kind.ToLowerInvariant()+'-'+$mode.ToLowerInvariant()+'.png')
                }finally{
                    [void](Run 'invSys.Core.xlam' 'TestShippingCatalog.CheckTerminalArm' @('','','',''))
                    RestoreStore
                }
                $policy=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @($controlId))
                if($kind -ceq 'PolicyUnreadable'){
                    Check ($label+'.UnreadablePolicyObserved') ($policy -ceq 'False|False|False|0')
                }else{Check ($label+'.ExpectedValidPolicyVersion') ($policy -ceq ('True|True|True|'+$(if($kind -ceq 'PolicyChanged'){$latest+1}else{$latest})))}
                $records=@(Get-ChildItem -LiteralPath $activityRoot -File -Filter '*.json'|Where-Object{-not $before.ContainsKey($_.Name)}|ForEach-Object{[IO.File]::ReadAllText($_.FullName)|ConvertFrom-Json})
                $valid=$records.Count -eq 1
                if($valid){$r=$records[0];$valid=$r.ControlId -ceq $controlId -and $r.OwnerId -ceq 'PRODUCTION_RUN_LOCAL' -and $r.OutcomeCode -ceq 'REQUESTED' -and $r.UserId -ceq 'config-producer' -and $r.SequenceId -ceq $starts[0].SequenceId -and $r.Ordinal -eq 1 -and $r.PolicyVersion -eq $latest -and $r.CatalogVersion -eq $catalog -and @($r.SourceEventRefs).Count -eq 0}
                Check ($label+'.OnlyOriginalRequestedNoFalseTerminal') $valid
                Check ($label+'.PreviousActivityImmutable') (PinsRetained $before)
                if($kind -ceq 'PolicyUnreadable'){
                    $pending=@(RecordingJournal $starts[0].SequenceId)
                    Check ($label+'.UnfinishedJournalCannotClaimCompletion') ($pending.Count -eq 2 -and @($pending|Where-Object Lifecycle -CNE 'Recording').Count -eq 0 -and (JournalChain $starts[0].SequenceId 2))
                    Check ($label+'.ActualViewerRefreshWhileUnreadable') ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('Refresh','')))
                    Check ($label+'.UnavailablePolicyVisibleAndStopDisabled') ((RecordingStatus) -ceq 'Tracking unavailable: the saved tracking policy is invalid.' -and (RecordingControl 'Stop Recording') -ceq 'True|False')
                    $viewerPrefix=if($NextBatch){'next-unreadable-policy-viewer-'}else{'check-in-unreadable-policy-viewer-'}
                    CaptureOwnedFormByCaptionEvidence ('Viewer - '+$Fixture.Warehouse) ($viewerPrefix+$mode.ToLowerInvariant()+'.png')
                    # Availability recovery cannot supply the missing terminal.
                    RestoreConfig $enabled
                    Check ($label+'.ActualViewerRefreshAfterRecovery') ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('Refresh','')))
                    Check ($label+'.RecoveredStopDelivered') ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')
                    Check ($label+'.StopReportsIncomplete') ((RecordingStatus).StartsWith('Incomplete evidence:'))
                }
                $closed=@(RecordingJournal $starts[0].SequenceId|Where-Object RecordType -CEQ 'Close')
                $reason=if($kind -ceq 'PolicyChanged'){'POLICY_CHANGED'}elseif($kind -ceq 'PolicyUnreadable'){'UNFINISHED_ACTIONS'}else{'TRACKING_UNAVAILABLE'}
                $valid=$closed.Count -eq 1
                if($valid){$c=$closed[0];$valid=$c.Lifecycle -ceq 'Incomplete' -and $c.ReasonCode -ceq $reason -and $c.ActionCount -eq 1 -and $c.CreatedByUserId -ceq 'config-producer' -and @($c.Observations).Count -eq 1 -and $c.Observations[0].ActivityId -ceq $records[0].ActivityId -and $c.Observations[0].OutcomeCode -ceq 'REQUESTED'}
                $closeCheck=if($kind -ceq 'PolicyUnreadable'){'.ExactIncompleteJournalAfterStop'}else{'.ExactIncompleteJournalWithoutStop'}
                Check ($label+$closeCheck) $valid
                Check ($label+'.IncrementalJournalChain') (JournalChain $starts[0].SequenceId 3)
                if(-not $valid){throw 'Recording interruption contract failed; inspect focused behavioral checks.'}
                $runs+=@{Label=$label;Journal=$closed[0]}
                CloseRecordingViewer
                RestoreConfig $enabled
                Check ($label+'.ConfigBytesRestored') ((Hash $Fixture.Config) -ceq $enabledHash)
                if($NextBatch){
                    $after=Get-NextPathAuthorityPins $Fixture;$same=$after.Count -eq $ownerPins.Count
                    foreach($file in $ownerPins.Keys){$same=$same -and $after.ContainsKey($file) -and $after[$file] -ceq $ownerPins[$file]}
                    Check ($label+'.AllAuthorityBytesPreservedAfterPolicyRestoration') $same
                    Check ($label+'.OperatorCustomValueAndFormulaPreserved') ($nextSheet.Cells.Item(2,1).Value2 -ceq $Canary -and $nextSheet.Cells.Item(2,2).Formula -ceq '=1+2')
                    Check ($label+'.DecoyPreserved') ($Decoy.Worksheets.Count -eq 1 -and $Decoy.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq $Canary)
                    [void](Probe 'CloseDesigner');$nextWork.Close($false);$nextWork=$null
                }
            }
        }
        $authority=@{}
        foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File|Where-Object{$_.Extension -in '.xlsb','.xlsm' -and $_.Name -notlike '*.Snapshot.*' -and $_.FullName -ine $candidate -and $_.FullName -ine $invalidCandidate -and $_.FullName -ine $Fixture.Config -and $_.Name -notlike '~$*'}){$authority[$file.FullName]=Hash $file.FullName}
        [void](Probe 'CloseDesigner');SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Owning publication unavailable.'}
        SelectTarget $Fixture 'config-reader';OpenRecordingViewer
        $trainingPins=RestartPins (Join-Path $Fixture.Root 'Training')
        foreach($run in $runs){
            $label=$run.Label;$journal=$run.Journal
            if((Select-EvaluationRun $journal.ActionPathId) -cne 'SELECTED'){throw 'Actual interrupted recording selection unavailable.'}
            $ready=Set-EvaluationDraft @(,@($controlId,'STAGED','True')) 0 'CommandCompleted'
            Check ($label+'.ActualExpectationEditor') $ready
            if(-not $ready){throw 'Existing Check In expectation unavailable.'}
            $before=@(EvaluationFiles|ForEach-Object FullName)
            $invoked=ExpectationControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths'
            $fresh=@(EvaluationFiles|Where-Object{$_.FullName -cnotin $before})
            if($invoked -cne 'DELIVERED' -or $fresh.Count -ne 1){throw 'Actual Evaluate must append one result.'}
            $result=[IO.File]::ReadAllText($fresh[0].FullName)|ConvertFrom-Json
            $status=ExpectationControl 'lblEvaluationStatus' 'Caption' '' 'frmActionPaths'
            Check ($label+'.DiagnosticCannotConclude') ($result.ResultState -ceq 'Incomplete' -and 'CAPTURE_INCOMPLETE' -cin @($result.ReasonCodes) -and $status.StartsWith('Incomplete evidence'))
            Check ($label+'.ExactRunAndIntent') ($result.JournalRecordId -ceq $journal.RecordId -and $result.JournalSha256 -ceq $journal.ContentSha256 -and $result.ExpectedConclusion.TerminalKind -ceq 'CommandCompleted' -and @($result.ExpectedConclusion.Steps).Count -eq 1 -and $result.ExpectedConclusion.Steps[0].RequiredOutcome -ceq 'STAGED')
            Check ($label+'.NoInventedCompletionOrDomainSources') (@($result.Matches).Count -eq 0 -and @($result.TerminalSources).Count -eq 0)
            CaptureOwnedFormByCaptionEvidence 'Action Paths' ($label.ToLowerInvariant()+'.png')
        }
        $retained=$true;foreach($file in $trainingPins.Keys){$retained=$retained -and (Hash $file) -ceq $trainingPins[$file]}
        Check ($prefix+'.PriorTrainingImmutableThroughEvaluation') $retained
        $retained=$true;foreach($file in $authority.Keys){$retained=$retained -and (Hash $file) -ceq $authority[$file]}
        Check ($prefix+'.AuthorityPreservedThroughPublicationAndEvaluation') $retained
        Check ($prefix+'.OtherWarehouseUnchanged') (RestartPinsEqual $otherPins $Other.Root)
    }finally{
        [void](Run 'invSys.Core.xlam' 'TestShippingCatalog.CheckTerminalArm' @('','','',''))
        RestoreStore;[void](Probe 'CloseDesigner');CloseRecordingViewer
        if($null -ne $nextWork){$nextWork.Close($false);$nextWork=$null}
        RestoreConfig $original
        if(Test-Path -LiteralPath $candidate){Remove-Item -LiteralPath $candidate -Force}
        if(Test-Path -LiteralPath $invalidCandidate){Remove-Item -LiteralPath $invalidCandidate -Force}
        SelectTarget $Fixture 'config-producer'
    }
    Check ($prefix+'.OriginalConfigBytesRestored') ((Hash $Fixture.Config) -ceq $originalHash)
}
