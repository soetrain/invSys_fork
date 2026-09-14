# D18 expectation/evaluation probes use actual UI events. Missing product controls
# return a fixed observation; unexpected COM/VBA errors remain harness failures.
function Install-RecordingEvaluationProbe {
    $packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modInventoryViewer').CodeModule.AddFromString(@'
Public Function RecordingExpectationForTest(ByVal formName As String, ByVal controlName As String, _
                                           ByVal action As String, Optional ByVal value As String = "") As String
    Dim form As Object, target As Object, control As Object, r As Long, c As Long
    For Each form In VBA.UserForms
        If form.Name = formName Then Set target = form: Exit For
    Next form
    RecordingExpectationForTest = "MISSING"
    If target Is Nothing Then Exit Function
    If action = "FormCaption" Then RecordingExpectationForTest = CStr(target.Caption): Exit Function
    If action = "WindowHandle" Then
        RecordingExpectationForTest = CStr(modUserFormResizeWin.GetUserFormWindowHandle(target)): Exit Function
    End If
    For Each control In target.Controls
        If control.Name = controlName Then
            If action <> "ValueAnyVisibility" And action <> "CountAnyVisibility" Then
                If Not target.Visible Or Not control.Visible Then RecordingExpectationForTest = "HIDDEN": Exit Function
            End If
            Select Case action
                Case "Click"
                    If Not control.Enabled Then RecordingExpectationForTest = "DISABLED": Exit Function
                    control.Value = True
                Case "Select"
                    For r = 0 To control.ListCount - 1
                        For c = 0 To control.ColumnCount - 1
                            If CStr(control.List(r, c)) = value Then
                                control.ListIndex = r: RecordingExpectationForTest = "DELIVERED": Exit Function
                            End If
                        Next c
                    Next r
                    RecordingExpectationForTest = "CHOICE_UNAVAILABLE": Exit Function
                Case "Index": control.ListIndex = CLng(value)
                Case "Boolean": control.Value = (value = "True")
                Case "Count", "CountAnyVisibility": RecordingExpectationForTest = CStr(control.ListCount): Exit Function
                Case "Value", "ValueAnyVisibility": RecordingExpectationForTest = CStr(control.Value): Exit Function
                Case "Caption": RecordingExpectationForTest = CStr(control.Caption): Exit Function
                Case "Locked": RecordingExpectationForTest = CStr(control.Locked): Exit Function
                Case "ViewportTop", "ViewportBottom"
                    If Not control.Locked Then RecordingExpectationForTest = "UNLOCKED": Exit Function
                    control.SetFocus
                    control.SelStart = IIf(action = "ViewportBottom", Len(CStr(control.Value)), 0)
                    control.SelLength = 0
                    target.Repaint: DoEvents
                Case Else: Err.Raise 5, , "Unsupported expectation probe action."
            End Select
            RecordingExpectationForTest = "DELIVERED": Exit Function
        End If
    Next control
End Function

Public Function RecordingExpectationLayoutForTest(ByVal formName As String, ByVal width As Single, ByVal height As Single) As String
    Dim form As Object, target As Object, control As Object, other As Object
    RecordingExpectationLayoutForTest = "MISSING"
    For Each form In VBA.UserForms
        If form.Name = formName Then Set target = form: Exit For
    Next form
    If target Is Nothing Then Exit Function
    If Not target.Visible Then RecordingExpectationLayoutForTest = "HIDDEN": Exit Function
    target.Width = width: target.Height = height
    target.Repaint: DoEvents
    For Each control In target.Controls
        If control.Visible Then
            If control.Left < 0 Or control.Top < 0 Or control.Width <= 0 Or control.Height <= 0 Or _
               control.Left + control.Width > target.InsideWidth Or control.Top + control.Height > target.InsideHeight Then
                RecordingExpectationLayoutForTest = "OUTSIDE:" & control.Name & "|Inside=" & CStr(target.InsideWidth) & "," & _
                    CStr(target.InsideHeight) & "|Bounds=" & CStr(control.Left) & "," & CStr(control.Top) & "," & _
                    CStr(control.Width) & "," & CStr(control.Height): Exit Function
            End If
            For Each other In target.Controls
                If other.Visible And other.Name <> control.Name Then
                    If control.Left < other.Left + other.Width And other.Left < control.Left + control.Width And _
                       control.Top < other.Top + other.Height And other.Top < control.Top + control.Height Then
                        RecordingExpectationLayoutForTest = "OVERLAP:" & control.Name & ":" & other.Name: Exit Function
                    End If
                End If
            Next other
        End If
    Next control
    RecordingExpectationLayoutForTest = "FITS"
End Function

'@)
}

function ExpectationControl([string]$Control,[string]$Action,[string]$Value='', [string]$Form='frmActionPathExpectation') {
    [string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingExpectationForTest' @($Form,$Control,$Action,$Value))
}

function EvaluationFiles {
    $directory=Join-Path $journalRoot 'Evaluations'
    if(Test-Path -LiteralPath $directory){Get-ChildItem -LiteralPath $directory -File -Filter '*.json'}
}

function Test-RecordingEvaluationStage([string]$Stage) {
    . (Join-Path $PSScriptRoot 'Slice4beEvaluationEvidence.ps1')
    $prefix='RecordingEvaluation.'+$Stage+'.'
    $oldActivity=ActivityPins
    $oldResults=@{}
    foreach($file in @(EvaluationFiles)){$oldResults[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash}
    $open=ExpectationControl 'btnExpectedConclusion' 'Click' '' 'frmActionPaths'
    Check ($prefix+'ExpectedConclusionOpensEditor') ($open -ceq 'DELIVERED')
    $empty=ExpectationControl 'lstExpectedSteps' 'Count'
    Check ($prefix+'AnalysisDefaultsToNone') ($empty -ceq '0')
    $cancel=ExpectationControl 'btnCancelExpectation' 'Click'
    Check ($prefix+'CancelDraftAvailable') ($cancel -ceq 'DELIVERED')
    $cancelledCount=@(EvaluationFiles).Count
    Check ($prefix+'CancelDoesNotSaveResult') ($cancelledCount -eq $oldResults.Count)
    [void](ExpectationControl 'btnExpectedConclusion' 'Click' '' 'frmActionPaths')
    $definitionReady=$true
    foreach($step in @(
        @('ADMIN_SETTINGS_SAVE_VALUE','COMPLETED'),
        @('RECEIVING_CONFIRM_WRITES','PENDING'),
        @('RECEIVING_CONFIRM_WRITES','PENDING'),
        @('ADMIN_SETTINGS_SAVE_VALUE','COMPLETED')
    )){
        # Fixed authored expectation; no inference from whatever happened to pass.
        $control=ExpectationControl 'cboExpectedControl' 'Select' $step[0]
        $outcome=ExpectationControl 'cboExpectedOutcome' 'Select' $step[1]
        $retry=ExpectationControl 'chkExpectedRetry' 'Boolean' 'True'
        $add=ExpectationControl 'btnAddExpectedStep' 'Click'
        $definitionReady=$definitionReady -and $control -ceq 'DELIVERED' -and $outcome -ceq 'DELIVERED' -and $retry -ceq 'DELIVERED' -and $add -ceq 'DELIVERED'
    }
    $count=ExpectationControl 'lstExpectedSteps' 'Count'
    Check ($prefix+'OrderedRepeatedStepsAuthored') ($definitionReady -and $count -ceq '4')
    # The third authored step is the second Confirm, which emits all four IDs.
    $terminal=ExpectationControl 'cboTerminalStep' 'Index' '2'
    $kind=ExpectationControl 'cboTerminalKind' 'Select' 'SourceEventsApplied'
    $use=ExpectationControl 'btnUseExpectation' 'Click'
    Check ($prefix+'ExplicitAllSourceEventsExpectationStaged') ($terminal -ceq 'DELIVERED' -and $kind -ceq 'DELIVERED' -and $use -ceq 'DELIVERED')
    $summary=ExpectationControl 'lblExpectationSummary' 'Caption' '' 'frmActionPaths'
    Check ($prefix+'AnalysisProvenanceVisible') ($summary.Contains('This evaluation'))
    $before=@(EvaluationFiles|ForEach-Object FullName)
    $clicked=ExpectationControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths'
    $status=ExpectationControl 'lblEvaluationStatus' 'Caption' '' 'frmActionPaths'
    $expected=if($Stage -eq 'Applied'){'Conclusion observed'}else{'Awaiting published result'}
    Check ($prefix+'OwningEvidenceDeterminesConclusion') ($clicked -ceq 'DELIVERED' -and $status.Contains($expected))
    $locked=ExpectationControl 'txtPathEvaluation' 'Locked' '' 'frmActionPaths'
    Check ($prefix+'ResultReadOnly') ($locked -ceq 'True')
    $created=@(EvaluationFiles|Where-Object {$_.FullName -cnotin $before})
    $valid=$created.Count -eq 1
    if($valid){
        $file=$created[0];$record=Get-Content -LiteralPath $file.FullName -Raw|ConvertFrom-Json
        $valid=$file.Length -le 1048576 -and $file.Name -cmatch '^[0-9a-fA-F-]{36}\.1\.json$' -and
            $record.SchemaVersion -eq 1 -and $record.RecordKind -ceq 'Evaluation' -and $record.Version -eq 1 -and
            $record.ActionPathId -ceq $pathId -and $record.SequenceId -ceq $sequence -and
            $record.WarehouseId -ceq $Fixture.Warehouse
    }
    Check ($prefix+'SeparateDerivedResultBoundToRun') $valid
    $savedFile=if($created.Count -eq 1){$created[0]}else{$null}
    Test-EvaluationEvidence $Stage $savedFile
    foreach($file in @(EvaluationFiles)){$oldResults[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash}
    [void](ExpectationControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths')
    Check ($prefix+'ReevaluationAppendsNewResult') (@(EvaluationFiles).Count -eq $oldResults.Count+1)
    $preserved=$true
    foreach($path in $oldResults.Keys){if(-not (Test-Path -LiteralPath $path) -or (Get-FileHash -LiteralPath $path).Hash -cne $oldResults[$path]){$preserved=$false}}
    Check ($prefix+'PriorDerivedResultsImmutable') $preserved
    Check ($prefix+'TrainingControlsDoNotCreateWorkflowActivity') ((PinsRetained $oldActivity) -and (ActivityPins).Count -eq $oldActivity.Count)
    # Exact original observation text remains after real analysis/evaluation.
    $after=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Evidence',''))
    Check ($prefix+'AnalysisPreservesOriginalObservations') ($after -ceq $evidence)
}
