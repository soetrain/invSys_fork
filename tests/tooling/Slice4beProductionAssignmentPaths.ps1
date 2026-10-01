# Disposable inputs select existing rows; every recorded action enters its real handler.
# The acceptable ITEM_CODE is a synthetic UI fixture, not an inventory allocation.
function Install-ProductionAssignmentPathProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $project.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Sub AssignmentPathInputForTest(ByVal action As String)
    Dim prior As Boolean, index As Long
    prior = mLoading: mLoading = True: mReadModeForTest = "Normal"
    Select Case action
        Case "PROCESS", "PROCESS_SELECT"
            For index = 0 To mLstAssignRecipes.ListCount - 1
                If NzStr(mLstAssignRecipes.List(index, 0)) = mReadProcessIdForTest And _
                    NzStr(mLstAssignRecipes.List(index, 1)) = mReadProcessVersionForTest Then
                    mLstAssignRecipes.ListIndex = index: Exit For
                End If
            Next index
        Case "REQUIREMENT", "REQUIREMENT_SELECT"
            mLstAssignIngredients.ListIndex = 0
        Case "ADD"
            mLstAssignInventory.Clear: mLstAssignInventory.AddItem "ASSIGN-PATH-KEY"
            mLstAssignInventory.List(0, 1) = mReadCanaryForTest
            mLstAssignInventory.List(0, 2) = "EA": mLstAssignInventory.List(0, 6) = "ASSIGN-PATH"
            mLstAssignInventory.ListIndex = 0
        Case "REMOVE"
            mLstAssignAllowed.ListIndex = 0
    End Select
    mLoading = prior
End Sub
Public Function AssignmentPathResultForTest() As Boolean
    Dim records As Collection, record As Variant, report As String, draft As Boolean, alternative As Boolean
    Set records = modProductionReusableDesigns.ParseReusableDefinitionRecords( _
        modOperationsPrimitiveBridge.GetProcessVersion(mTxtProcessId.Text, mTxtProcessVersion.Text), report)
    If records Is Nothing Then Exit Function
    For Each record In records
        If modProductionReusableDesigns.ReusableRecordText(record, "RecordType") = "PROCESS" Then _
            draft = (modProductionReusableDesigns.ReusableRecordText(record, "Status") = "DRAFT")
        If modProductionReusableDesigns.ReusableRecordText(record, "RecordType") = "ALTERNATIVE" Then _
            alternative = (modProductionReusableDesigns.ReusableRecordText(record, "RequirementId") = "A02" And _
                modProductionReusableDesigns.ReusableRecordText(record, "ITEM_CODE") = "ASSIGN-PATH")
    Next record
    AssignmentPathResultForTest = draft And alternative And mProcessAlternatives.Count = 1 And _
        mTxtProcessId.Text = mReadProcessIdForTest And mTxtProcessVersion.Text <> mReadProcessVersionForTest
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Sub AssignmentPathInput(ByVal action As String)
    mForm.AssignmentPathInputForTest action
End Sub
Public Function AssignmentPathResult() As Boolean
    AssignmentPathResult = mForm.AssignmentPathResultForTest()
End Function
'@)
}
function Get-AssignmentPathInstruction([string]$Action) {
    switch($Action){
        'REFRESH' {'Click Refresh in Ingredients Assignment and inspect the lists. Refresh completion does not prove source availability.'}
        'PROCESS_SELECT' {'Select the released Process in Processes. Inspect its requirements; selection replaces local assignment data.'}
        'PROCESS' {'Keep the intended Process selected and click Select Process to load its requirements and alternatives.'}
        'REQUIREMENT_SELECT' {'Select the ingredient requirement in Ingredient Requirements to view its acceptable item types.'}
        'REQUIREMENT' {'Keep that requirement selected and click Select Requirement to refresh its acceptable items.'}
        'ADD' {'Select the compatible inventory item type and click Add Acceptable. This stages an ITEM_CODE alternative; it does not allocate inventory.'}
        'REMOVE' {'Select the staged acceptable item and click Remove Row. Inspect the local alternatives.'}
        'CLEAR' {'Click Clear to empty the local requirement and acceptable-item views and the shared alternatives draft.'}
        'SAVE' {'Click Save Alternatives. Inspect the new DRAFT Process version; use the exact published Designs event to establish application separately.'}
    }
}
function Test-AssignmentPathAction($Terminal,$Attempt,$Fixture,[string]$Label,[int]$Ordinal) {
    $refs=@($Terminal.SourceEventRefs);$save=$Terminal.ControlId -ceq 'PRODUCTION_ASSIGNMENT_SAVE'
    $exact=@($Attempt.SourceEventRefs).Count -eq 0 -and $Attempt.ControlId -ceq $Terminal.ControlId -and $Attempt.SequenceId -ceq $Terminal.SequenceId -and $Attempt.Ordinal -eq $Ordinal
    if($save){
        $writer=([string](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterEvidenceForTest')).Split('|')
        $exact=$exact -and $writer.Count -eq 5 -and $writer[1] -ceq '1' -and $writer[2] -ceq 'True' -and $refs.Count -eq 1
        if($refs.Count -eq 1){$exact=$exact -and $refs[0].EventId -ceq $writer[0] -and $refs[0].SourceKind -ceq 'Designs' -and $refs[0].WarehouseId -ceq $Fixture.Warehouse -and $refs[0].SubmissionState -ceq 'Submitted' -and @($refs[0].PSObject.Properties).Count -eq 4}
    }else{$exact=$exact -and $refs.Count -eq 0}
    Check ('AssignmentPaths.'+$Label+'.'+$Ordinal+'.ExactOwnerFacts') $exact
}
function Test-AssignmentPathSources($Source,$Observed,$Publication,$Fixture) {
    foreach($run in @(@{Name='GuideSource';Value=$Source},@{Name='ObservedRun';Value=$Observed})){
        $refs=@($run.Value.Terminals[-1].SourceEventRefs);$exact=$refs.Count -eq 1
        foreach($ref in $refs){
            $group=@($Publication.Groups|Where-Object{$_.Source -ceq 'Designs' -and $_.SourceId -ceq $ref.EventId})
            $exact=$exact -and $group.Count -eq 1
            if($group.Count -eq 1){
                $exact=$exact -and @($group[0].Lines).Count -gt 0 -and @($group[0].Outcomes).Count -eq @($group[0].Lines).Count
                foreach($line in $group[0].Lines){$exact=$exact -and $line.EventID -ceq $ref.EventId -and $line.WarehouseId -ceq $Fixture.Warehouse -and $line.AppliedAtUTC -cne '' -and $line.AppliedSeq -cmatch '^[1-9][0-9]*$'}
            }
        }
        Check ('AssignmentPaths.'+$run.Name+'.ExactPublishedDesignsEvent') $exact
    }
    Check 'AssignmentPaths.IndependentOwningEventIdentities' ($Source.Terminals[-1].SourceEventRefs[0].EventId -cne $Observed.Terminals[-1].SourceEventRefs[0].EventId)
}
function Test-AssignmentPathTerminalSource($Result,$Terminal,$Publication) {
    . (Join-Path $PSScriptRoot 'Slice4beEvaluationEvidence.ps1')
    $refs=@($Result.TerminalSources);$exact=$refs.Count -eq 1
    if($exact){
        $ref=$refs[0];$expected=$Terminal.SourceEventRefs[0]
        $group=@($Publication.Groups|Where-Object{$_.Source -ceq 'Designs' -and $_.SourceId -ceq $expected.EventId})
        $exact=$group.Count -eq 1 -and $ref.EventId -ceq $expected.EventId -and $ref.SourceKind -ceq 'Designs' -and $ref.WarehouseId -ceq $Terminal.WarehouseId -and $ref.SubmissionState -ceq 'Submitted' -and $ref.OwnerStatus -ceq 'Applied' -and @($ref.SystemKeys).Count -eq 0
        if($group.Count -eq 1){
            $body=[ordered]@{Lines=@($group[0].Lines)}|ConvertTo-Json -Depth 16 -Compress
            $exact=$exact -and $ref.LineCount -eq @($group[0].Lines).Count -and $ref.LinesSha256 -ceq (EvaluationSha $body)
        }
    }
    Check 'AssignmentPaths.ExactAppliedSourceProof' $exact
}
function Test-AssignmentPathLocalTerminals($Observed) {
    foreach($record in $Observed.Terminals|Select-Object -First 8){foreach($kind in @('CommandCompleted','SourceEventsApplied')){
        $label='AssignmentPaths.LocalTerminal.'+$record.ControlId+'.'+$kind
        $ready=Set-EvaluationDraft @(,@($record.ControlId,$record.OutcomeCode,'True')) 0 $kind -StopAtMissingChoice
        Check ($label+'.ActualEditor') $ready
        if(-not $ready){continue}
        $before=@(EvaluationFiles|ForEach-Object FullName);Delivered (ExpectationControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths')
        $fresh=@(EvaluationFiles|Where-Object{$_.FullName -cnotin $before});if($fresh.Count -ne 1){throw 'Assignment local evaluation unavailable.'}
        $result=Get-Content $fresh[0].FullName -Raw|ConvertFrom-Json
        $valid=if($kind -ceq 'CommandCompleted'){$result.ResultState -ceq 'Concluded' -and 'COMMAND_COMPLETED' -cin @($result.ReasonCodes)}else{$result.ResultState -ceq 'Incomplete' -and 'SOURCE_UNAVAILABLE' -cin @($result.ReasonCodes)}
        Check ($label+'.ExactLocalOnlyConclusion') ($valid -and $result.JournalRecordId -ceq $Observed.Journal.RecordId -and $result.Matches[-1].ActivityId -ceq $record.ActivityId -and @($result.TerminalSources).Count -eq 0)
    }}
}
