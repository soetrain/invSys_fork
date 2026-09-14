# D18 recording lifecycle through actual Viewer controls and Admin Save Value.
# The probes are installed in disposable, unsaved package projects. Missing
# product controls return an explicit observation; missing test seams throw.
function Test-Slice4beActionRecording($Fixture) {
    $form=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmInventoryViewer').CodeModule
    $form.AddFromString(@'
Public Function RecordingControlForTest(ByVal caption As String, ByVal operation As String) As String
    Dim control As Object
    For Each control In Me.Controls
        If TypeName(control) = "CommandButton" Then
            If StrComp(control.Caption, caption, vbBinaryCompare) = 0 Then
                Select Case operation
                    Case "State"
                        RecordingControlForTest = CStr(control.Visible) & "|" & CStr(control.Enabled)
                    Case "Click"
                        If Not control.Visible Or Not control.Enabled Then
                            RecordingControlForTest = "DISABLED": Exit Function
                        End If
                        control.Value = True
                        RecordingControlForTest = "DELIVERED"
                End Select
                Exit Function
            End If
        End If
    Next control
    RecordingControlForTest = "MISSING"
End Function
Public Function RecordingStatusForTest() As String
    Dim control As Object
    For Each control In Me.Controls
        If TypeName(control) = "Label" Then
            If control.Name = "lblRecordingStatus" Then
                RecordingStatusForTest = control.Caption: Exit Function
            End If
        End If
    Next control
    RecordingStatusForTest = "MISSING"
End Function
'@)
    $manager=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modInventoryViewer').CodeModule
    $manager.AddFromString(@'
Public Function RecordingControlForTest(ByVal caption As String, ByVal operation As String) As String
    If mInventoryViewer Is Nothing Then Err.Raise 5, , "Recording fixture Viewer is not open."
    RecordingControlForTest = mInventoryViewer.RecordingControlForTest(caption, operation)
End Function
Public Function RecordingStatusForTest() As String
    If mInventoryViewer Is Nothing Then Err.Raise 5, , "Recording fixture Viewer is not open."
    RecordingStatusForTest = mInventoryViewer.RecordingStatusForTest()
End Function
'@)
    $policy=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('cAdminTrackingPolicy').CodeModule
    $policy.AddFromString(@'
Public Function RecordingCapturePolicyForTest(ByVal enabled As Boolean) As Boolean
    mCapture.Value = enabled
    mCapture_Click
    mSave_Click
    RecordingCapturePolicyForTest = (InStr(1, mStatus.Caption, " saved.", vbBinaryCompare) > 0)
End Function
'@)
    $settings=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('frmAdminSettings').CodeModule
    $settings.AddFromString(@'
Public Function RecordingCapturePolicyForTest(ByVal enabled As Boolean) As Boolean
    RecordingCapturePolicyForTest = mTracking.RecordingCapturePolicyForTest(enabled)
End Function
'@)
    $commands=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('TestD5Commands').CodeModule
    $commands.AddFromString(@'
Public Function RecordingCapturePolicyForTest(ByVal enabled As Boolean) As Boolean
    RecordingCapturePolicyForTest = mForm.RecordingCapturePolicyForTest(enabled)
End Function
'@)
    function RecordingControl([string]$Caption,[string]$Operation='State') {
        [string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingControlForTest' @($Caption,$Operation))
    }
    function RecordingStatus { [string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingStatusForTest') }
    function OpenRecordingViewer {
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OpenInventoryViewer')
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('Events',''))
    }
    function CloseRecordingViewer { [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.CloseInventoryViewerForTest') }
    function SetRecordingPolicy([bool]$Enabled) {
        try {
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
            if(-not [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.RecordingCapturePolicyForTest' @($Enabled))) {
                throw 'Actual recording policy fixture save failed; not behavioral RED.'
            }
        } finally { [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings') }
    }
    $activityRoot=Join-Path $Fixture.Root ('Training/Activity/'+$Fixture.Warehouse)
    function ActivityPins {
        $pins=@{}
        if(Test-Path -LiteralPath $activityRoot) {
            foreach($file in Get-ChildItem -LiteralPath $activityRoot -Filter '*.json' -File) {
                $pins[$file.Name]=(Get-FileHash -LiteralPath $file.FullName).Hash
            }
        }
        return $pins
    }
    function SaveRecordedSetting([string]$Value) {
        $before=ActivityPins
        try {
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
            if(-not [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('BatchSize',$Value))) {
                throw 'Actual Admin Save Value fixture failed; not behavioral RED.'
            }
        } finally { [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings') }
        $records=@(foreach($file in Get-ChildItem -LiteralPath $activityRoot -Filter '*.json' -File) {
            if(-not $before.ContainsKey($file.Name)) { [IO.File]::ReadAllText($file.FullName)|ConvertFrom-Json }
        })
        $attempts=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED')
        $outcomes=@($records|Where-Object OutcomeCode -CEQ 'COMPLETED')
        if($records.Count -ne 2 -or $attempts.Count -ne 1 -or $outcomes.Count -ne 1 -or
            $attempts[0].ControlId -cne 'ADMIN_SETTINGS_SAVE_VALUE' -or
            $attempts[0].ActivityId -cne $outcomes[0].ActivityId) {
            throw 'Existing Admin activity fixture is invalid; not recording RED.'
        }
        [pscustomobject]@{Attempt=$attempts[0];Outcome=$outcomes[0]}
    }
    function HasSequence($Action,[int]$Ordinal) {
        $id=[guid]::Empty
        return ([guid]::TryParse([string]$Action.Attempt.SequenceId,[ref]$id) -and
            $id -ne [guid]::Empty -and $Action.Attempt.SequenceId -ceq $Action.Outcome.SequenceId -and
            $Action.Attempt.Ordinal -eq $Ordinal -and $Action.Outcome.Ordinal -eq $Ordinal)
    }
    function PinsRetained($Pins) {
        foreach($name in $Pins.Keys) {
            if((Get-FileHash -LiteralPath (Join-Path $activityRoot $name)).Hash -cne $Pins[$name]) { return $false }
        }
        return $true
    }
    SelectTarget $Fixture 'config-admin'
    try {
        SetRecordingPolicy $false
        OpenRecordingViewer
        foreach($caption in @('Start Recording','Stop Recording','Cancel Recording')) {
            Check ('Recording.ControlPresent.'+$caption.Replace(' ','')) ((RecordingControl $caption) -match '^True\|')
        }
        Check 'Recording.DisabledPolicyExplained' ((RecordingStatus) -match '(?i)(capture|recording).*(off|disabled)')
        Check 'Recording.DisabledPolicyCannotStart' ((RecordingControl 'Start Recording') -ceq 'True|False')
        $ordinary=SaveRecordedSetting '611'
        Check 'Recording.OutsideSequenceRetainsOrdinaryActivity' ($ordinary.Attempt.SequenceId -ceq '' -and $ordinary.Outcome.SequenceId -ceq '' -and $ordinary.Attempt.Ordinal -eq 0 -and $ordinary.Outcome.Ordinal -eq 0)
        CloseRecordingViewer
        SetRecordingPolicy $true
        OpenRecordingViewer
        Check 'Recording.StartUsesActualControl' ((RecordingControl 'Start Recording' 'Click') -ceq 'DELIVERED')
        Check 'Recording.StartShowsZeroCounter' ((RecordingStatus) -match '(?i)recording.*0\s*/\s*256')
        $first=SaveRecordedSetting '612'
        Check 'Recording.FirstAttemptAndOutcomeShareSequenceAndOrdinal' (HasSequence $first 1)
        $firstPins=ActivityPins
        $second=SaveRecordedSetting '613'
        Check 'Recording.SecondAttemptAndOutcomeShareSequenceAndOrdinal' (HasSequence $second 2)
        Check 'Recording.RepeatedCommandRetainsDistinctOccurrence' ((HasSequence $first 1) -and (HasSequence $second 2) -and $first.Attempt.SequenceId -ceq $second.Attempt.SequenceId -and $first.Attempt.ActivityId -cne $second.Attempt.ActivityId)
        Check 'Recording.LaterActionPreservesEarlierBytes' (PinsRetained $firstPins)
        Check 'Recording.StopUsesActualControl' ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')
        Check 'Recording.StopFreezesWithoutAssertingConclusion' ((RecordingStatus) -match '(?i)stopped' -and (RecordingStatus) -notmatch '(?i)conclusion observed')
        $after=SaveRecordedSetting '614'
        Check 'Recording.AfterStopUsesOrdinaryActivity' ($after.Attempt.SequenceId -ceq '' -and $after.Outcome.SequenceId -ceq '' -and $after.Attempt.Ordinal -eq 0 -and $after.Outcome.Ordinal -eq 0)
        Check 'Recording.SecondStartUsesActualControl' ((RecordingControl 'Start Recording' 'Click') -ceq 'DELIVERED')
        $cancelled=SaveRecordedSetting '615'
        Check 'Recording.NewRunHasDistinctSequence' ((HasSequence $cancelled 1) -and (HasSequence $first 1) -and $cancelled.Attempt.SequenceId -cne $first.Attempt.SequenceId)
        $cancelPins=ActivityPins
        Check 'Recording.CancelUsesActualControl' ((RecordingControl 'Cancel Recording' 'Click') -ceq 'DELIVERED')
        Check 'Recording.CancelShowsCancelled' ((RecordingStatus) -match '(?i)cancelled')
        Check 'Recording.CancelPreservesObservedActivity' (PinsRetained $cancelPins)
    } finally { CloseRecordingViewer }
}
