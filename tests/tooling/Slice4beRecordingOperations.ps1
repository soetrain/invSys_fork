# D18: one real Admin/Operations recording, repeated multi-event submissions,
# and deferred owning application. No fabricated activity or source identities.
function Test-Slice4beRecordingOperations([bool]$InstallOnly=$false) {
    if(-not $InstallOnly){CloseRecordingViewer}
    if($CheckRecordingEvaluation){
        . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
        if(-not $script:RecordingOperationsProbeInstalled){Install-RecordingEvaluationProbe}
        if($CheckEvaluationContracts){
            . (Join-Path $PSScriptRoot 'Slice4beEvaluationContracts.ps1')
            . (Join-Path $PSScriptRoot 'Slice4beEvaluationBinding.ps1')
            if(-not $script:RecordingOperationsProbeInstalled){Install-EvaluationBindingProbe}
        }
        if($CheckExpectationCompatibility){. (Join-Path $PSScriptRoot 'Slice4beExpectationCompatibility.ps1')}
    }
    if(-not $script:RecordingOperationsProbeInstalled){
    $operations=$packages['invSys.Operations.xlam'].VBProject
    $control=$operations.VBComponents.Add(2);$control.Name='TestRecordingRibbonControl'
    $control.CodeModule.AddFromString(@'
Option Explicit
Implements Office.IRibbonControl
Private Property Get IRibbonControl_Id() As String
    IRibbonControl_Id = "btnOperationsReceivingForm"
End Property
Private Property Get IRibbonControl_Tag() As String
    IRibbonControl_Tag = ""
End Property
Private Property Get IRibbonControl_Context() As Object
    Set IRibbonControl_Context = Application.ActiveWindow
End Property
'@)
    $operations.VBComponents.Item('frmReceiving').CodeModule.AddFromString(@'
Public Function RecordingPrepareForTest() As String
    Dim before As Long
    before = mLstReceiveItems.ListCount
    mBtnRefresh_Click
    mTabs.Value = 0
    ApplyReceivingTab
    RecordingPrepareForTest = CStr(before) & "|" & CStr(mLstReceiveItems.ListCount)
End Function
Public Function RecordingAddForTest() As String
    Dim before As Long, lo As ListObject
    Set lo = mOperatorWorkbook.Worksheets("ReceivedTally").ListObjects("ReceivedTally")
    before = lo.ListRows.Count
    RecordingAddForTest = "NO_ITEMS"
    If mLstReceiveItems.ListCount = 0 Then Exit Function
    mLstReceiveItems.ListIndex = 0
    mLstReceiveItems_Click
    mTxtRef.Value = "RECORDING-PRIVATE-REFERENCE-" & CStr(before + 1)
    mTxtQty.Value = "1"
    mTxtReceiveLocation.Value = "RECORDING-PRIVATE-LOCATION"
    mCboCondition.Value = "GOOD"
    mBtnAdd_Click
    If lo.ListRows.Count = before + 1 Then
        RecordingAddForTest = "ADDED"
    Else
        RecordingAddForTest = "ROWS_" & CStr(before) & "_" & CStr(lo.ListRows.Count)
    End If
End Function
Public Function RecordingConfirmForTest() As Boolean
    mBtnConfirm_Click
    RecordingConfirmForTest = modTS_Received.LastConfirmWritesSucceeded()
End Function
'@)
    $operations.VBComponents.Item('modTS_Received').CodeModule.AddFromString(@'
Public Function RecordingReceivingForTest(ByVal action As String) As String
    Dim control As New TestRecordingRibbonControl
    Select Case action
        Case "Launch"
            modRibbonGenerated.RibbonOnActionOperations control
        Case "Add"
            RecordingReceivingForTest = mReceivingLauncherForm.RecordingAddForTest(): Exit Function
        Case "Prepare"
            RecordingReceivingForTest = mReceivingLauncherForm.RecordingPrepareForTest(): Exit Function
        Case "Confirm"
            RecordingReceivingForTest = CStr(mReceivingLauncherForm.RecordingConfirmForTest()): Exit Function
        Case "Close"
            If Not mReceivingLauncherForm Is Nothing Then Unload mReceivingLauncherForm
            Set mReceivingLauncherForm = Nothing
            Exit Function
    End Select
    If mReceivingLauncherForm Is Nothing Then Err.Raise 5, , "Receiving launcher fixture missing."
    RecordingReceivingForTest = mReceivingLauncherWorkbookName
End Function
'@)
    $operations.VBComponents.Item('modInventoryViewer').CodeModule.AddFromString(@'
Public Function RecordingEvaluateForTest(ByVal invoke As Boolean) As String
    Dim form As Object, control As Object
    For Each form In VBA.UserForms
        If form.Name = "frmActionPaths" Then
            For Each control In form.Controls
                If TypeName(control) = "CommandButton" Then
                    If control.Caption = "Evaluate" Then
                        If Not control.Visible Or Not control.Enabled Then RecordingEvaluateForTest = "DISABLED": Exit Function
                        If invoke Then control.Value = True
                        RecordingEvaluateForTest = "AVAILABLE": Exit Function
                    End If
                End If
            Next control
        End If
    Next form
    RecordingEvaluateForTest = "MISSING"
End Function
'@)
    $gate=$packages['invSys.Core.xlam'].VBProject.VBComponents.Add(1);$gate.Name='TestRecordingDeferred'
    $gate.CodeModule.AddFromString(@'
Option Explicit
Public Withhold As Boolean
Public Sub SetWithhold(ByVal value As Boolean)
    Withhold = value
End Sub
'@)
    $bridge=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modOperationsPrimitiveBridge').CodeModule
    $start=$bridge.ProcStartLine('RunBatchAndRefreshOperatorWorkbook',0)
    $count=$bridge.ProcCountLines('RunBatchAndRefreshOperatorWorkbook',0)
    $original=$bridge.Lines($start,$count)
    $changed=$original -replace '(?m)^(\s*)Dim wb As Workbook', ('$1Dim wb As Workbook'+[Environment]::NewLine+'    If TestRecordingDeferred.Withhold Then report = "Fixture application withheld.": Exit Function')
    if($changed -ceq $original){throw 'Deferred processor fixture seam missing.'}
    $bridge.DeleteLines($start,$count);$bridge.InsertLines($start,$changed)
    $packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('modAdminConsole').CodeModule.AddFromString(@'
Public Function RecordingProcessForTest(ByVal adminWorkbookName As String) As Long
    Dim report As String
    RecordingProcessForTest = RunProcessorFromConsole("config-admin", "", Application.Workbooks(adminWorkbookName), report)
End Function
Public Function RecordingPublishForTest(ByVal adminWorkbookName As String) As Boolean
    Dim report As String
    RecordingPublishForTest = GenerateInventorySnapshot("config-admin", "", Nothing, "", Application.Workbooks(adminWorkbookName), report)
End Function
'@)
    $script:RecordingOperationsProbeInstalled=$true
    }
    if($InstallOnly){return}
    # All project edits precede fixture creation; repeated calls only exercise actions.
    [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride' @((Join-Path $repo 'deploy/current/templates')))
    [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetLocalOperatorRootOverrideForAutomation' @((Join-Path $runRoot 'operators')))
    $Fixture=NewFixture 'recording-operations'
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beReceivingActivity.ps1')
    SelectTarget $Fixture 'config-admin'
    $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($Fixture.Warehouse,'S1','config-admin'))
    if(-not $seed.StartsWith('OK|')){throw 'Recording Operations seed fixture failed.'}
    SetRecordingPolicy $true
    $adminBook=$excel.Workbooks.Add()
    $adminBook.SaveAs((Join-Path $runRoot 'recording-admin.xlsm'),52)
    $other=$excel.Workbooks.Add()
    $other.Worksheets.Item(1).Cells.Item(1,1).Value2='unrelated recording fixture'
    $other.SaveAs((Join-Path $runRoot 'recording-unrelated.xlsm'),52)
    $otherHash=Get-ReceivingFixtureHash $other.FullName
    $otherObservations=[Collections.Generic.List[object]]::new()
    function ObserveRecordingOther([string]$Stage){
        $otherObservations.Add([pscustomobject]@{Stage=$Stage;Saved=[bool]$other.Saved;
            FileBytesPreserved=((Get-ReceivingFixtureHash $other.FullName) -ceq $otherHash);
            SheetCount=[int]$other.Worksheets.Count;
            UsedRows=[int]$other.Worksheets.Item(1).UsedRange.Rows.Count;
            UsedColumns=[int]$other.Worksheets.Item(1).UsedRange.Columns.Count;
            SentinelPreserved=($other.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'unrelated recording fixture')})
    }
    ObserveRecordingOther 'Created'
    $operatorName=[string](Run 'invSys.Operations.xlam' 'modTS_Received.RecordingReceivingForTest' @('Launch'))
    $operator=$excel.Workbooks.Item($operatorName)
    ObserveRecordingOther 'InitialLauncher'
    $staging=Table $operator 'ReceivedTally'
    if($staging.ListRows.Count -ne 0){throw 'Recording requires an empty generated Receiving staging table.'}
    $unknown=$staging.ListColumns.Add();$unknown.Name='Operator Recording Extra'
    $prepared=[string](Run 'invSys.Operations.xlam' 'modTS_Received.RecordingReceivingForTest' @('Prepare'))
    $ready=$prepared -match '^[0-9]+\|[1-9][0-9]*$'
    Check 'RecordingOperations.ActualRefreshPreparesManagedItems' $ready
    if(-not $ready){throw 'Actual Receiving Refresh did not prepare managed item choices.'}
    ObserveRecordingOther 'Prepared'
    OpenRecordingViewer
    [void](Run 'invSys.Core.xlam' 'TestRecordingDeferred.SetWithhold' @($true))
    try {
        if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'Actual Start control unavailable.'}
        $first=SaveRecordedSetting '651';$sequence=[string]$first.Attempt.SequenceId
        ObserveRecordingOther 'FirstAdminSave'
        Check 'RecordingOperations.AdminStartsSharedSequence' (HasSequence $first 1)
        $other.Activate()
        $reused=[string](Run 'invSys.Operations.xlam' 'modTS_Received.RecordingReceivingForTest' @('Launch'))
        Check 'RecordingOperations.RibbonReusesCapturedWorkbook' ($reused -ceq $operatorName)
        ObserveRecordingOther 'RepeatedLauncher'
        $submissionIds=@();$entityKeys=@()
        foreach($batch in 1..2){
            foreach($item in 1..2){
                $other.Activate()
                $added=[string](Run 'invSys.Operations.xlam' 'modTS_Received.RecordingReceivingForTest' @('Add'))
                if($added -cne 'ADDED'){
                    if($added -notmatch '^(NO_ITEMS|ROWS_[0-9]+_[0-9]+)$'){$added='UNAVAILABLE'}
                    throw ('Actual Receiving Add fixture failed: '+$added)
                }
            }
            $rows=@(Get-ReceivingFixtureRows $staging)
            $ids=@($rows|ForEach-Object {[string]$_.EventId})
            $keys=@($rows|ForEach-Object {[string]$_.System_Key})
            if($ids.Count -ne 2*$batch -or @($ids|Select-Object -Unique).Count -ne $ids.Count -or '' -in $ids -or '' -in $keys){throw 'Staging source identity fixture invalid.'}
            if($batch -eq 1){$unknown.DataBodyRange.Cells.Item(1,1).Value2='first extra';$unknown.DataBodyRange.Cells.Item(2,1).Value2='second extra'}
            $before=ActivityPins
            $other.Activate()
            $confirmed=[string](Run 'invSys.Operations.xlam' 'modTS_Received.RecordingReceivingForTest' @('Confirm'))
            ObserveRecordingOther ('Submission'+$batch)
            $records=@(foreach($file in Get-ChildItem -LiteralPath $activityRoot -Filter '*.json' -File){if(-not $before.ContainsKey($file.Name)){Get-Content -LiteralPath $file.FullName -Raw|ConvertFrom-Json}})
            $attempt=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$outcome=@($records|Where-Object OutcomeCode -CEQ 'PENDING')
            $pair=$records.Count -eq 2 -and $attempt.Count -eq 1 -and $outcome.Count -eq 1
            Check ('RecordingOperations.Submission'+$batch+'.PendingOwnerOutcome') ($confirmed -ceq 'False' -and $pair)
            if(-not $pair){throw 'Receiving fixture lacks its owning pending observation pair.'}
            $action=[pscustomobject]@{Attempt=$attempt[0];Outcome=$outcome[0]}
            Check ('RecordingOperations.Submission'+$batch+'.SameSequenceOrderedOccurrence') ((HasSequence $action (2+3*$batch)) -and $action.Attempt.SequenceId -ceq $sequence)
            $refs=@($outcome[0].SourceEventRefs)
            $exact=$refs.Count -eq $ids.Count -and @($attempt[0].SourceEventRefs).Count -eq 0 -and $outcome[0].DataEffect -ceq 'Unknown'
            foreach($id in $ids){$exact=$exact -and @($refs|Where-Object {$_.EventId -ceq $id -and $_.WarehouseId -ceq $Fixture.Warehouse -and $_.SourceKind -ceq 'Inventory' -and $_.SubmissionState -ceq 'Submitted'}).Count -eq 1}
            Check ('RecordingOperations.Submission'+$batch+'.EveryExactEventReference') $exact
            if($batch -eq 2){Check 'RecordingOperations.RepeatedSubmissionPreservesEarlierIds' (@($submissionIds|Where-Object {$_ -cnotin $ids}).Count -eq 0 -and @($entityKeys|Where-Object {$_ -cnotin $keys}).Count -eq 0)}
            $submissionIds=$ids;$entityKeys=$keys
        }
        $last=SaveRecordedSetting '652'
        ObserveRecordingOther 'LastAdminSave'
        Check 'RecordingOperations.AdminReturnsToSameSequence' ((HasSequence $last 9) -and $last.Attempt.SequenceId -ceq $sequence)
        [void](RecordingControl 'Stop Recording' 'Click')
        $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close')
        Check 'RecordingOperations.StopRetainsAllOrderedObservations' ($closed.Count -eq 1 -and (JournalFact $sequence 'Close' 9 'Stopped') -and (JournalChain $sequence 20) -and @($closed[0].Observations).Count -eq 18)
        if($closed.Count -ne 1){throw 'Recorded Operations fixture has no closing journal.'}
        $pins=RestartPins $journalRoot
        $pathId=[string]$closed[0].ActionPathId
        $eventsPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')
        $stages=if($CheckRecordingEvaluation){@('Pending','Partial','Applied')}else{@('Pending','Applied')}
        $totalProcessed=0
        foreach($stage in $stages){
            if($stage -ne 'Pending'){
                if($CheckRecordingEvaluation -and $stage -eq 'Partial'){
                    $batchSetting=SaveRecordedSetting '3'
                    Check 'RecordingEvaluation.BatchLimitUsesActualAdminHandler' ($batchSetting.Outcome.SequenceId -ceq '')
                }
                [void](Run 'invSys.Core.xlam' 'TestRecordingDeferred.SetWithhold' @($false))
                $processed=[long](Run 'invSys.Admin.xlam' 'modAdminConsole.RecordingProcessForTest' @($adminBook.Name))
                $totalProcessed+=$processed
                if($stage -eq 'Partial'){Check 'RecordingEvaluation.OwnerAppliesOnlyThreeEvents' ($processed -eq 3)}
                if($stage -eq 'Applied'){Check 'RecordingOperations.OwnerAppliesFourDistinctEvents' ($totalProcessed -eq 4)}
                ObserveRecordingOther 'Processor'
            }
            if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.RecordingPublishForTest' @($adminBook.Name))){throw 'Owning publication failed in Operations fixture.'}
            ObserveRecordingOther ($stage+'Publication')
            $published=Get-Content -LiteralPath $eventsPath -Raw|ConvertFrom-Json
            $events=@($published.Groups|Where-Object {$_.Source -ceq 'Inventory' -and $_.SourceId -cin $submissionIds})
            $expectedCount=if($stage -eq 'Applied'){4}elseif($stage -eq 'Partial'){3}else{0}
            Check ('RecordingOperations.'+$stage+'.PublishedOwnerEventCount') ($events.Count -eq $expectedCount)
            if($stage -eq 'Applied'){
                $exact=$true
                foreach($i in 0..3){$group=@($events|Where-Object SourceId -CEQ $submissionIds[$i]);$exact=$exact -and $group.Count -eq 1 -and @($group[0].Lines|Where-Object {$_.System_Key -ceq $entityKeys[$i] -and $_.AppliedAtUTC -ne ''}).Count -eq 1}
                Check 'RecordingOperations.PublishedApplicationPreservesExactKeys' $exact
            }
            if($CheckEvaluationContracts -and $stage -ne 'Pending'){Test-EvaluationLoadedPublication $stage}
            [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('Refresh',''))
            [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Open',''))
            $selected=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Select',$pathId))
            Check ('RecordingOperations.'+$stage+'.RealSequenceSelectable') ($selected -ceq 'SELECTED')
            $evaluate=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingEvaluateForTest' @($true))
            Check ('RecordingOperations.'+$stage+'.EvaluateActionAvailable') ($evaluate -ceq 'AVAILABLE')
            $evidence=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Evidence',''))
            $referencesVisible=([regex]::Matches($evidence,'Source event: ').Count -eq 6)
            foreach($id in $submissionIds){$referencesVisible=$referencesVisible -and $evidence.Contains($id)}
            Check ('RecordingOperations.'+$stage+'.RepeatedSourceReferencesVisible') $referencesVisible
            Check ('RecordingOperations.'+$stage+'.NoConclusionWithoutExpectation') ($evidence -notmatch '(?i)conclusion observed')
            if($CheckRecordingEvaluation){Test-RecordingEvaluationStage $stage}
            ObserveRecordingOther ($stage+'Viewer')
            if($CaptureEvidence){CaptureFormEvidence 'Action Paths' ('recording-operations-'+$stage.ToLowerInvariant()+'.png')}
        }
        if($CheckEvaluationContracts){
            if($CheckExpectationCompatibility){Test-ExpectationCompatibility}
            Test-EvaluationContracts
            Test-EvaluationBinding
            if($CheckExpectationCompatibility){Test-ExpectationEditorBinding}
            ObserveRecordingOther 'EvaluationContracts'
        }
        # D18 permits new derived evaluation records; original journal bytes stay immutable.
        $preserved=$true
        foreach($path in $pins.Keys){
            if(-not (Test-Path -LiteralPath $path) -or (Get-FileHash -LiteralPath $path).Hash -cne $pins[$path]){$preserved=$false}
        }
        Check 'RecordingOperations.PublicationAndEvaluationPreserveJournal' $preserved
        Check 'RecordingOperations.UnknownStagingColumnPreserved' ($staging.ListColumns.Item('Operator Recording Extra').DataBodyRange.Cells.Item(1,1).Value2 -ceq 'first extra' -and $staging.ListColumns.Item('Operator Recording Extra').DataBodyRange.Cells.Item(2,1).Value2 -ceq 'second extra')
        $lastOther=$otherObservations[$otherObservations.Count-1]
        # Existing Receiving GREEN permits Saved=False while preserving the file.
        # Keep that flag diagnostic; never force it true or infer a data write from it.
        Check 'RecordingOperations.UnrelatedWorkbookPreserved' ($lastOther.FileBytesPreserved -and $lastOther.SheetCount -eq 1 -and $lastOther.UsedRows -eq 1 -and $lastOther.UsedColumns -eq 1 -and $lastOther.SentinelPreserved)
        Check 'RecordingOperations.UnrelatedFileBytesPreserved' $lastOther.FileBytesPreserved
        Check 'RecordingOperations.UnrelatedContentsPreserved' ($lastOther.SheetCount -eq 1 -and $lastOther.UsedRows -eq 1 -and $lastOther.UsedColumns -eq 1 -and $lastOther.SentinelPreserved)
    } finally {
        $otherObservations|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $reportRoot 'recording-operations-other-state.json')
        [void](Run 'invSys.Core.xlam' 'TestRecordingDeferred.SetWithhold' @($false))
        CloseRecordingViewer
        [void](Run 'invSys.Operations.xlam' 'modTS_Received.RecordingReceivingForTest' @('Close'))
    }
}
