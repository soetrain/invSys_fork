# D18: one-shot reentry at commit/return boundaries, through the real list handler.
# Probes exist only in disposable unsaved projects and emit no operational values.
function Install-EvaluationBindingProbe {
    $manager=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modInventoryViewer').CodeModule
    $manager.AddFromString(@'
Private EvaluationStageForTest As String
Private EvaluationPathForTest As String
Private EvaluationActionForTest As String
Private EvaluationTakenForTest As Boolean
Public Sub ArmEvaluationBoundaryForTest(ByVal stage As String, ByVal pathId As String, Optional ByVal action As String = "Select")
    EvaluationStageForTest = stage: EvaluationPathForTest = pathId: EvaluationActionForTest = action: EvaluationTakenForTest = False
End Sub
Public Sub EvaluationBoundaryForTest(ByVal stage As String)
    Dim pathId As String, ignored As String, count As Long, index As Long
    If stage <> EvaluationStageForTest Or EvaluationPathForTest = "" Then Exit Sub
    pathId = EvaluationPathForTest: EvaluationStageForTest = "": EvaluationPathForTest = ""
    Select Case EvaluationActionForTest
        Case "Select": EvaluationTakenForTest = (RecordingLibraryForTest("Select", pathId) = "SELECTED")
        Case "ClearIntent"
            ignored = RecordingExpectationForTest("frmActionPaths", "btnExpectedConclusion", "Click")
            If ignored <> "DELIVERED" Then Err.Raise 5, , "Expectation reentry fixture unavailable."
            count = CLng(RecordingExpectationForTest("frmActionPathExpectation", "lstExpectedSteps", "Count"))
            For index = count - 1 To 0 Step -1
                ignored = RecordingExpectationForTest("frmActionPathExpectation", "lstExpectedSteps", "Index", CStr(index))
                ignored = RecordingExpectationForTest("frmActionPathExpectation", "btnRemoveExpectedStep", "Click")
            Next index
            ignored = RecordingExpectationForTest("frmActionPathExpectation", "cboTerminalKind", "Select", "None")
            EvaluationTakenForTest = (RecordingExpectationForTest("frmActionPathExpectation", "btnUseExpectation", "Click") = "DELIVERED")
        Case "CloseAndRefresh"
            ignored = RecordingControlForTest("Stop Recording", "Click")
            If ignored <> "DELIVERED" Then Err.Raise 5, , "Stop reentry fixture unavailable."
            ignored = RecordingLibraryForTest("Refresh")
            EvaluationTakenForTest = (RecordingExpectationForTest("frmActionPaths", "lstActionPaths", "Value") = pathId)
        Case Else: Err.Raise 5, , "Unknown reentry fixture action."
    End Select
End Sub
Public Function EvaluationReentryTakenForTest() As Boolean
    EvaluationReentryTakenForTest = EvaluationTakenForTest
End Function
Public Function SelectedEvaluationIdForTest() As String
    Dim form As Object
    For Each form In VBA.UserForms
        If form.Name = "frmActionPaths" Then SelectedEvaluationIdForTest = form.SelectedEvaluationIdForTest(): Exit Function
    Next form
    Err.Raise 5, , "Evaluation library fixture missing."
End Function
'@)
    $form=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmActionPaths').CodeModule
    $form.AddFromString(@'
Public Function SelectedEvaluationIdForTest() As String
    SelectedEvaluationIdForTest = mEvaluationId
End Function
'@)
    $start=$form.ProcStartLine('mEvaluate_Click',0);$count=$form.ProcCountLines('mEvaluate_Click',0)
    $original=$form.Lines($start,$count)
    $pattern='(?im)^(\s*succeeded = modPathEvaluation\.Evaluate[^\r\n]*)'
    if(-not [regex]::IsMatch($original,$pattern)){throw 'Evaluate return boundary unavailable; not product RED.'}
    $changed=[regex]::Replace($original,$pattern,('$1'+[Environment]::NewLine+'    modInventoryViewer.EvaluationBoundaryForTest "AfterReturn"'))
    $form.DeleteLines($start,$count);$form.InsertLines($start,$changed)
    $core=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modPathEvaluation').CodeModule
    $start=$core.ProcStartLine('Evaluate',0);$count=$core.ProcCountLines('Evaluate',0)
    $original=$core.Lines($start,$count)
    $pattern='(?im)^(\s*If context <> modActivity\.CaptureContext\(\) Then GoTo Failed)'
    if(-not [regex]::IsMatch($original,$pattern)){throw 'Evaluate append boundary unavailable; not product RED.'}
    $call='    Application.Run "''invSys.Operations.xlam''!modInventoryViewer.EvaluationBoundaryForTest", "BeforeSave"'
    $changed=[regex]::Replace($original,$pattern,($call+[Environment]::NewLine+'$1'))
    $core.DeleteLines($start,$count);$core.InsertLines($start,$changed)
}

function Selected-EvaluationId {
    [string](Run 'invSys.Operations.xlam' 'modInventoryViewer.SelectedEvaluationIdForTest')
}

function Test-EvaluationBinding {
    $command=@(,@('ADMIN_SETTINGS_SAVE_VALUE','COMPLETED','True'))
    $originalPath=$pathId
    # Create another real stopped run; reentry selects it through ListIndex/Change.
    if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'Binding fixture recorder unavailable.'}
    $action=SaveRecordedSetting '691'
    [void](RecordingControl 'Stop Recording' 'Click')
    $closed=@(RecordingJournal $action.Attempt.SequenceId|Where-Object RecordType -CEQ 'Close')
    if($closed.Count -ne 1){throw 'Binding fixture has no real closed journal.'}
    $otherPath=[string]$closed[0].ActionPathId
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Refresh',''))
    foreach($stage in @('BeforeSave','AfterReturn')){
        if((Select-EvaluationRun $originalPath) -cne 'SELECTED' -or -not (Set-EvaluationDraft $command 0 'CommandCompleted')){throw 'Binding fixture selected expectation unavailable.'}
        $context=[string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
        $before=@(EvaluationFiles|ForEach-Object FullName)
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.ArmEvaluationBoundaryForTest' @($stage,$otherPath))
        $clicked=ExpectationControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths'
        $taken=[bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.EvaluationReentryTakenForTest')
        $selected=ExpectationControl 'lstActionPaths' 'Value' '' 'frmActionPaths'
        $sameContext=$context -ceq [string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
        Check ('EvaluationBinding.'+$stage+'.ActualSelectionChangedWithinSameContext') ($clicked -ceq 'DELIVERED' -and $taken -and $selected -ceq $otherPath -and $sameContext)
        if(-not $taken -or $selected -cne $otherPath -or -not $sameContext){throw 'Reentry fixture did not reach the selected boundary; not product RED.'}
        $created=@(EvaluationFiles|Where-Object {$_.FullName -cnotin $before})
        if($stage -ceq 'BeforeSave'){
            Check 'EvaluationBinding.SelectionBeforeSavePreventsStaleAppend' ($created.Count -eq 0)
        }else{
            $preserved=$false
            if($created.Count -eq 1){
                $saved=Get-Content -LiteralPath $created[0].FullName -Raw|ConvertFrom-Json
                $preserved=$saved.ActionPathId -ceq $originalPath -and $saved.ResultState -ceq 'Concluded'
            }
            Check 'EvaluationBinding.SelectionAfterSavePreservesOriginalBoundResult' $preserved
        }
        $text=ExpectationControl 'txtPathEvaluation' 'Value' '' 'frmActionPaths'
        Check ('EvaluationBinding.'+$stage+'.NewSelectionHasNoStaleResultAttachment') ($text -ceq '' -and (Selected-EvaluationId) -ceq '')
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Refresh',''))
        Check ('EvaluationBinding.'+$stage+'.RefreshDoesNotReattachStaleResult') ((ExpectationControl 'txtPathEvaluation' 'Value' '' 'frmActionPaths') -ceq '' -and (Selected-EvaluationId) -ceq '')
    }

    if((Select-EvaluationRun $originalPath) -cne 'SELECTED' -or -not (Set-EvaluationDraft $command 0 'CommandCompleted')){throw 'Intent reentry fixture unavailable.'}
    $before=@(EvaluationFiles).Count
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.ArmEvaluationBoundaryForTest' @('BeforeSave',$originalPath,'ClearIntent'))
    [void](ExpectationControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths')
    $taken=[bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.EvaluationReentryTakenForTest')
    $summary=ExpectationControl 'lblExpectationSummary' 'Caption' '' 'frmActionPaths'
    Check 'EvaluationBinding.IntentChangedByActualEditorBeforeSave' ($taken -and $summary.Contains('0 expected step'))
    if(-not $taken){throw 'Intent reentry fixture failed.'}
    Check 'EvaluationBinding.IntentChangePreventsStaleAppendAndAttachment' (@(EvaluationFiles).Count -eq $before -and (Selected-EvaluationId) -ceq '' -and (ExpectationControl 'txtPathEvaluation' 'Value' '' 'frmActionPaths') -ceq '')

    # Evaluate a real unfinished journal, then close it through the ordinary control.
    if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'Advancing journal fixture recorder unavailable.'}
    $action=SaveRecordedSetting '692'
    $entries=@(RecordingJournal $action.Attempt.SequenceId|Sort-Object Version)
    if($entries.Count -ne 3){throw 'Advancing journal fixture lacks incremental observations.'}
    $advancingPath=[string]$entries[0].ActionPathId
    $prefixPins=@{};foreach($entry in $entries){$file=Join-Path $journalRoot ($advancingPath+'.'+$entry.Version+'.json');$prefixPins[$file]=(Get-FileHash -LiteralPath $file).Hash}
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Refresh',''))
    if((Select-EvaluationRun $advancingPath) -cne 'SELECTED' -or -not (Set-EvaluationDraft $command 0 'CommandCompleted')){throw 'Unfinished journal selection fixture failed.'}
    [void](ExpectationControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths')
    $oldId=Selected-EvaluationId
    if($oldId -notmatch '^[0-9a-fA-F-]{36}$'){throw 'Existing evaluator did not save unfinished fixture result.'}
    $oldFile=Join-Path (Join-Path $journalRoot 'Evaluations') ($oldId+'.1.json')
    $oldHash=(Get-FileHash -LiteralPath $oldFile).Hash
    $oldResult=Get-Content -LiteralPath $oldFile -Raw|ConvertFrom-Json
    $oldText=ExpectationControl 'txtPathEvaluation' 'Value' '' 'frmActionPaths'
    Check 'EvaluationBinding.UnfinishedResultBindsExactObservedVersion' ($oldResult.JournalVersion -eq 3 -and $oldResult.ResultState -ceq 'Incomplete' -and $oldText.Length -gt 0)
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Refresh',''))
    Check 'EvaluationBinding.SameVersionRefreshRetainsSavedResult' ((Selected-EvaluationId) -ceq $oldId -and (ExpectationControl 'txtPathEvaluation' 'Value' '' 'frmActionPaths') -ceq $oldText)
    $before=@(EvaluationFiles).Count
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.ArmEvaluationBoundaryForTest' @('BeforeSave',$advancingPath,'CloseAndRefresh'))
    [void](ExpectationControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths')
    $taken=[bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.EvaluationReentryTakenForTest')
    $closed=@(RecordingJournal $action.Attempt.SequenceId|Where-Object RecordType -CEQ 'Close')
    if($closed.Count -ne 1 -or $closed[0].Version -ne 4){throw 'Ordinary Stop did not advance the same journal.'}
    Check 'EvaluationBinding.JournalAdvancedByActualStopDuringEvaluate' $taken
    if(-not $taken){throw 'Journal advance did not reach reentry boundary.'}
    Check 'EvaluationBinding.VersionAdvanceBeforeSavePreventsStaleAppend' (@(EvaluationFiles).Count -eq $before)
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Refresh',''))
    $selected=ExpectationControl 'lstActionPaths' 'Value' '' 'frmActionPaths'
    $text=ExpectationControl 'txtPathEvaluation' 'Value' '' 'frmActionPaths'
    $summary=ExpectationControl 'lblExpectationSummary' 'Caption' '' 'frmActionPaths'
    Check 'EvaluationBinding.NewJournalVersionClearsSavedResultSelection' ($selected -ceq $advancingPath -and $text -ceq '' -and (Selected-EvaluationId) -ceq '')
    Check 'EvaluationBinding.NewJournalVersionClearsAnalysisIntent' (-not $summary.Contains('This evaluation'))
    if(-not (Set-EvaluationDraft $command 0 'CommandCompleted')){throw 'Closed journal expectation fixture unavailable.'}
    [void](ExpectationControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths')
    $newId=Selected-EvaluationId
    $newFile=Join-Path (Join-Path $journalRoot 'Evaluations') ($newId+'.1.json')
    $newResult=Get-Content -LiteralPath $newFile -Raw|ConvertFrom-Json
    Check 'EvaluationBinding.NewEvaluationBindsNewClosingVersion' ($newId -cne $oldId -and $newResult.JournalVersion -eq 4 -and $newResult.JournalSha256 -ceq $closed[0].ContentSha256 -and $newResult.ResultState -ceq 'Concluded' -and $newResult.PreviousEvaluationId -ceq '')
    $unchanged=(Get-FileHash -LiteralPath $oldFile).Hash -ceq $oldHash
    foreach($file in $prefixPins.Keys){$unchanged=$unchanged -and (Get-FileHash -LiteralPath $file).Hash -ceq $prefixPins[$file]}
    Check 'EvaluationBinding.AdvancePreservesPriorResultAndJournalBytes' $unchanged
    [void](Select-EvaluationRun $originalPath)
}
