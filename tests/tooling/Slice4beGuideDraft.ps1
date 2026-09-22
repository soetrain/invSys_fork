# D18 guide draft entry through actual packaged controls. The disposable probe
# observes controls or delivers their normal action; it never constructs a guide.
function Install-GuideDraftProbe {
    $manager=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modInventoryViewer').CodeModule
    $manager.AddFromString(@'
Public Function GuideDraftControlForTest(ByVal formName As String, ByVal name As String, ByVal action As String, Optional ByVal value As String = "") As String
    Dim instance As Object, owner As Object, control As Object, item As Object, count As Long, labels As String, index As Long
    For Each instance In VBA.UserForms
        If TypeName(instance) = formName Then Set owner = instance: count = count + 1
    Next instance
    If action = "Count" Then GuideDraftControlForTest = CStr(count): Exit Function
    GuideDraftControlForTest = "MISSING"
    If owner Is Nothing Then Exit Function
    If action = "Caption" Then GuideDraftControlForTest = owner.Caption: Exit Function
    If action = "Labels" Then
        For Each item In owner.Controls
            If TypeName(item) = "Label" Then labels = labels & item.Caption & vbLf
        Next item
        GuideDraftControlForTest = labels: Exit Function
    End If
    If action = "Fit" Then
        Select Case value
            Case "Minimum": owner.Width = 760: owner.Height = 600
            Case "Default", "Restored": owner.Width = 900: owner.Height = 650
            Case "Larger": owner.Width = 1000: owner.Height = 760
            Case "Current"
            Case Else: Err.Raise 5, , "Unknown guide layout fixture."
        End Select
        If formName = "frmActionPathExpectation" Then
            If value = "Minimum" Then owner.Height = 550
            If value = "Default" Or value = "Restored" Then owner.Height = 630
        End If
        If formName = "frmActionPathView" Then
            If value = "Minimum" Then owner.Width = 840: owner.Height = 600
            If value = "Default" Or value = "Restored" Then owner.Width = 960: owner.Height = 680
        End If
        owner.Repaint: DoEvents
        GuideDraftControlForTest = "False"
        For Each control In owner.Controls
            If formName = "frmActionPathView" And Not control.Visible Then GoTo NextFitControl
            If Not control.Visible Or control.Left < 0 Or control.Top < 0 Or control.Width <= 0 Or control.Height <= 0 Then Exit Function
            If control.Left + control.Width > owner.InsideWidth Or control.Top + control.Height > owner.InsideHeight Then Exit Function
            For Each item In owner.Controls
                If item.Name <> control.Name And (item.Visible Or formName <> "frmActionPathView") Then
                    If control.Left < item.Left + item.Width And item.Left < control.Left + control.Width And _
                       control.Top < item.Top + item.Height And item.Top < control.Top + control.Height Then Exit Function
                End If
            Next item
NextFitControl:
        Next control
        GuideDraftControlForTest = "True": Exit Function
    End If
    For Each item In owner.Controls
        If item.Name = name Then Set control = item: Exit For
    Next item
    If control Is Nothing And formName = "frmEventTrackingSettings" Then
        For Each item In owner.Controls("mpOperationsSettings").Pages(0).Controls
            If item.Name = name Then Set control = item: Exit For
        Next item
    End If
    If control Is Nothing Then Exit Function
    Select Case action
        Case "State": GuideDraftControlForTest = CStr(control.Visible) & "|" & CStr(control.Enabled)
        Case "Text": GuideDraftControlForTest = CStr(control.Value)
        Case "Label": GuideDraftControlForTest = CStr(control.Caption)
        Case "Rows": GuideDraftControlForTest = CStr(control.ListCount)
        Case "Locked": GuideDraftControlForTest = CStr(control.Locked)
        Case "Selected": GuideDraftControlForTest = CStr(control.Value)
        Case "ViewportTop", "ViewportBottom"
            If TypeName(control) <> "TextBox" Then GuideDraftControlForTest = "NOT TEXT": Exit Function
            If Not control.Locked Or Not control.Visible Or Not owner.Visible Then GuideDraftControlForTest = "UNAVAILABLE": Exit Function
            control.SetFocus
            control.SelStart = IIf(action = "ViewportBottom", Len(CStr(control.Value)), 0)
            control.SelLength = 0
            owner.Repaint: DoEvents
            GuideDraftControlForTest = "DELIVERED"
        Case "Values"
            For index = 0 To control.ListCount - 1
                labels = labels & CStr(control.List(index, 0)) & vbLf
            Next index
            GuideDraftControlForTest = labels
        Case "Select"
            index = CLng(value)
            If index < 0 Or index >= control.ListCount Then GuideDraftControlForTest = "OUT OF RANGE": Exit Function
            control.ListIndex = -1: control.ListIndex = index
            GuideDraftControlForTest = "SELECTED"
        Case "Click", "Write", "Check"
            If Not owner.Visible Or Not control.Visible Or Not control.Enabled Then
                GuideDraftControlForTest = "DISABLED": Exit Function
            End If
            If action = "Click" Then
                control.Value = True
            ElseIf action = "Check" Then
                control.Value = CBool(value)
            Else
                control.Value = value
            End If
            GuideDraftControlForTest = "DELIVERED"
    End Select
End Function
'@)
}

function Test-GuideDraftEntry($Fixture,$Other) {
    function CaptureGuide([string]$Stage) {
        if(-not $CaptureEvidence){return}
        Initialize-SettingsCapture
        for($attempt=1;$attempt -le 3;$attempt++){
            try {
                $handle=[InvSysSettingsCapture]::OwnedVisibleForm('Action Path guide',[IntPtr]$excel.Hwnd).ToInt64()
                if($handle -eq 0){throw 'Requested form window unavailable.'}
                $activation=New-Object -ComObject WScript.Shell
                try{[void]$activation.AppActivate('Action Path guide')}finally{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($activation)}
                CaptureFormEvidence 'Action Path guide' ('guide-draft-'+$Stage.ToLowerInvariant()+'.png') $handle
                Check ('GuideDraft.VisibleCapture.'+$Stage) $true
                return
            } catch {
                if($_.Exception.GetBaseException().Message -cnotin @('Requested form is not in the foreground.','Requested form window unavailable.')){throw}
                if($attempt -lt 3){Start-Sleep -Milliseconds 300}
            }
        }
        Check ('GuideDraft.VisibleCapture.'+$Stage) $false
    }
    function GuideControl([string]$Name,[string]$Action,[string]$Value='', [string]$Form='frmActionPathGuide') {
        [string](Run 'invSys.Operations.xlam' 'modInventoryViewer.GuideDraftControlForTest' @($Form,$Name,$Action,$Value))
    }
    function GuideLibrary([string]$Action,[string]$Value='') {
        [string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @($Action,$Value))
    }
    function CreateGuide { GuideControl 'btnCreateGuide' 'Click' '' 'frmActionPaths' }
    function TrainingPins {
        $pins=@{}
        foreach($file in Get-ChildItem -LiteralPath $journalRoot -Recurse -File){$pins[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash}
        return $pins
    }
    function TrainingUnchanged($Before) {
        $after=TrainingPins
        if($after.Count -ne $Before.Count){return $false}
        foreach($path in $Before.Keys){if(-not $after.ContainsKey($path) -or $after[$path] -cne $Before[$path]){return $false}}
        return $true
    }
    CloseRecordingViewer
    SelectTarget $Fixture 'config-admin'
    $allowed=Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-admin',$Fixture.Warehouse,'S1')
    if($allowed -isnot [bool] -or -not $allowed){throw 'Guide-author capability fixture unavailable; not product RED.'}
    SetRecordingPolicy $true
    OpenRecordingViewer
    if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'Existing recorder entry unavailable; not guide RED.'}
    $first=SaveRecordedSetting '630'
    $second=SaveRecordedSetting '631'
    if((RecordingControl 'Stop Recording' 'Click') -cne 'DELIVERED'){throw 'Existing recorder stop unavailable; not guide RED.'}
    $sequence=[string]$first.Attempt.SequenceId
    if(-not (HasSequence $first 1) -or -not (HasSequence $second 2) -or -not (JournalChain $sequence 6)){
        throw 'Real Admin sequence fixture unavailable; not guide RED.'
    }
    $entries=@(RecordingJournal $sequence|Sort-Object Version)
    $pathId=[string]$entries[0].ActionPathId
    if((GuideLibrary 'Open') -cne 'DELIVERED' -or (GuideLibrary 'Select' $pathId) -cne 'SELECTED'){
        throw 'Existing saved-run selection unavailable; not guide RED.'
    }
    $originalEvidence=GuideLibrary 'Evidence'
    if(-not $originalEvidence.Contains($first.Attempt.ActivityId) -or -not $originalEvidence.Contains($second.Attempt.ActivityId)){
        throw 'Existing source evidence lacks the actual actions; not guide RED.'
    }
    $training=TrainingPins; $activity=ActivityPins
    $configHash=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $publishBefore=[long](Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest')
    try {
        Check 'GuideDraft.EntryEnabledForPermittedSelectedRun' ((GuideControl 'btnCreateGuide' 'State' '' 'frmActionPaths') -ceq 'True|True')
        $delivered=CreateGuide
        $opened=(GuideControl '' 'Count') -ceq '1'
        Check 'GuideDraft.ActualCreateControlOpensEditor' ($delivered -ceq 'DELIVERED' -and $opened)
        Check 'GuideDraft.ApprovedEditorCaption' ($opened -and (GuideControl '' 'Caption') -ceq 'Action Path guide')
        $source=GuideControl 'lblGuideSource' 'Label'
        Check 'GuideDraft.ExactSourceRunAndVersionVisible' ($opened -and $source.Contains($pathId) -and $source.Contains($sequence) -and $source -match '(?i)version\s*:\s*6\b')
        Check 'GuideDraft.DistinctObservedActionsBecomeDistinctDraftSteps' ($opened -and (GuideControl 'lstGuideSteps' 'Rows') -ceq '2')
        $observed=GuideControl 'txtGuideEvidence' 'Text'
        $firstIndex=$observed.IndexOf($first.Attempt.ActivityId,[StringComparison]::Ordinal)
        $secondIndex=$observed.IndexOf($second.Attempt.ActivityId,[StringComparison]::Ordinal)
        Check 'GuideDraft.OriginalObservedEvidenceOrderedAndLabelled' ($opened -and $firstIndex -ge 0 -and $secondIndex -gt $firstIndex -and $observed.Contains('Observed control') -and $observed.Contains('Save Value') -and $observed.Contains('REQUESTED') -and $observed.Contains('COMPLETED'))
        Check 'GuideDraft.ObservedEvidenceIsLocked' ($opened -and (GuideControl 'txtGuideEvidence' 'Locked') -ceq 'True')
        Check 'GuideDraft.AuthoredInstructionsBeginEmpty' ($opened -and (GuideControl 'txtGuideInstructions' 'Text') -ceq '' -and (GuideControl 'txtGuideStepInstruction' 'Text') -ceq '')
        Check 'GuideDraft.AuthoredInstructionWordingDistinctFromObservedControl' ($opened -and (GuideControl '' 'Labels').Contains('Authored instruction') -and $observed.Contains('Observed control'))
        $nameSet=GuideControl 'txtGuideName' 'Write' 'Review a Settings change'
        $textSet=GuideControl 'txtGuideInstructions' 'Write' 'Review the setting, save it, then inspect its recorded outcome.'
        Check 'GuideDraft.AuthoredTextStagesThroughEditor' ($opened -and $nameSet -ceq 'DELIVERED' -and $textSet -ceq 'DELIVERED' -and (GuideControl 'txtGuideName' 'Text') -ceq 'Review a Settings change')
        [void](CreateGuide)
        Check 'GuideDraft.RepeatedEntryReusesStagedDraft' ($opened -and (GuideControl '' 'Count') -ceq '1' -and (GuideControl 'txtGuideName' 'Text') -ceq 'Review a Settings change')
        Check 'GuideDraft.AuthoredTextDoesNotRewriteObservationPane' ($opened -and (GuideControl 'txtGuideEvidence' 'Text') -ceq $observed)
        $tagsSet=GuideControl 'txtGuideTags' 'Write' 'settings, training'
        Check 'GuideDraft.TagsStageAsAuthoredContent' ($opened -and $tagsSet -ceq 'DELIVERED' -and (GuideControl 'txtGuideTags' 'Text') -ceq 'settings, training')
        [void](GuideControl 'lstGuideSteps' 'Select' '0')
        $firstStep=GuideControl 'lstGuideSteps' 'Selected'
        [void](GuideControl 'txtGuideStepInstruction' 'Write' 'First authored instruction')
        [void](GuideControl 'lstGuideSteps' 'Select' '1')
        $secondStep=GuideControl 'lstGuideSteps' 'Selected'
        [void](GuideControl 'txtGuideStepInstruction' 'Write' 'Second authored instruction')
        $firstGuid=[guid]::Empty;$secondGuid=[guid]::Empty
        Check 'GuideDraft.StableStepIdsDifferFromObservedActionIds' ($opened -and [guid]::TryParse($firstStep,[ref]$firstGuid) -and [guid]::TryParse($secondStep,[ref]$secondGuid) -and $firstGuid -ne [guid]::Empty -and $secondGuid -ne [guid]::Empty -and $firstGuid -ne $secondGuid -and $firstStep -cnotin @($first.Attempt.ActivityId,$second.Attempt.ActivityId) -and $secondStep -cnotin @($first.Attempt.ActivityId,$second.Attempt.ActivityId))
        $up=GuideControl 'btnGuideStepUp' 'Click'
        [void](GuideControl 'lstGuideSteps' 'Select' '0')
        Check 'GuideDraft.MoveUpKeepsStepIdentityAndInstruction' ($opened -and $up -ceq 'DELIVERED' -and (GuideControl 'lstGuideSteps' 'Selected') -ceq $secondStep -and (GuideControl 'txtGuideStepInstruction' 'Text') -ceq 'Second authored instruction')
        if($opened){CaptureGuide 'Reordered'}
        $down=GuideControl 'btnGuideStepDown' 'Click'
        [void](GuideControl 'lstGuideSteps' 'Select' '1')
        Check 'GuideDraft.MoveDownKeepsStepIdentityAndInstruction' ($opened -and $down -ceq 'DELIVERED' -and (GuideControl 'lstGuideSteps' 'Selected') -ceq $secondStep -and (GuideControl 'txtGuideStepInstruction' 'Text') -ceq 'Second authored instruction')
        $remove=GuideControl 'btnRemoveGuideStep' 'Click'
        [void](GuideControl 'lstGuideSteps' 'Select' '0')
        Check 'GuideDraft.RemovalRetainsRemainingAuthoredStep' ($opened -and $remove -ceq 'DELIVERED' -and (GuideControl 'lstGuideSteps' 'Rows') -ceq '1' -and (GuideControl 'lstGuideSteps' 'Selected') -ceq $firstStep -and (GuideControl 'txtGuideStepInstruction' 'Text') -ceq 'First authored instruction')
        Check 'GuideDraft.StepEditsPreserveOriginalObservationOrder' ($opened -and (GuideControl 'txtGuideEvidence' 'Text') -ceq $observed)
        foreach($layout in @('Minimum','Default','Larger','Restored')){
            Check ('GuideDraft.Layout.'+$layout) ($opened -and (GuideControl '' 'Fit' $layout) -ceq 'True')
            if($opened){CaptureGuide $layout}
        }
        $cancel=GuideControl 'btnCancelGuide' 'Click'
        Check 'GuideDraft.CancelDiscardsEditor' ($opened -and $cancel -ceq 'DELIVERED' -and (GuideControl '' 'Count') -ceq '0')
        Check 'GuideDraft.EntryEditingCancelPreserveTrainingActivityConfig' ((TrainingUnchanged $training) -and (PinsRetained $activity) -and (ActivityPins).Count -eq $activity.Count -and (Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configHash)
        Check 'GuideDraft.NoPublicationOrSourceReplacement' ([long](Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest') -eq $publishBefore -and (GuideLibrary 'Evidence') -ceq $originalEvidence)
        [void](CreateGuide)
        $wasOpen=(GuideControl '' 'Count') -ceq '1'
        SelectTarget $Other 'config-admin'
        [void](GuideControl 'txtGuideName' 'Write' 'Stale draft must not retarget')
        [void](CreateGuide)
        $status=GuideControl 'lblGuideStatus' 'Label'
        Check 'GuideDraft.ContextChangeInvalidatesRetainedDraft' ($wasOpen -and ((GuideControl '' 'Count') -ceq '0' -or $status -match '(?i)unavailable|changed|reopen'))
        $retainedEvidence=GuideControl 'txtGuideEvidence' 'Text'
        Check 'GuideDraft.ContextChangeClearsOriginalEvidence' ($wasOpen -and $retainedEvidence -cin @('','MISSING'))
        CloseRecordingViewer
        SelectTarget $Fixture 'config-reader'
        $denied=Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-reader',$Fixture.Warehouse,'S1')
        if($denied -isnot [bool] -or $denied){throw 'Guide-reader capability fixture unavailable; not product RED.'}
        OpenRecordingViewer
        if((GuideLibrary 'Open') -cne 'DELIVERED' -or (GuideLibrary 'Select' $pathId) -cne 'SELECTED'){throw 'Ordinary recording reader fixture unavailable.'}
        Check 'GuideDraft.OrdinaryReaderCannotCreateGuide' ((GuideControl 'btnCreateGuide' 'State' '' 'frmActionPaths') -ceq 'True|False' -and (CreateGuide) -ceq 'DISABLED' -and (GuideControl '' 'Count') -ceq '0')
        Check 'GuideDraft.DenialAndContextChangePreserveSourceTraining' (TrainingUnchanged $training)
    } finally {
        CloseRecordingViewer
        SelectTarget $Fixture 'config-admin'
    }
}
