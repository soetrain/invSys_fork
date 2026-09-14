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
Public Function RecordingGeometryForTest() As Boolean
    Dim name As Variant, control As Object, right As Single, filter As Object
    On Error GoTo Missing
    Set filter = Me.Controls("cboEventsOutcome")
    For Each name In Array("btnStartRecording", "btnStopRecording", "btnCancelRecording", "lblRecordingStatus")
        Set control = Me.Controls(CStr(name))
        If Not control.Visible Or control.Left < right Or control.Width <= 0 Or control.Height <= 0 Then Exit Function
        If control.Left + control.Width > Me.InsideWidth Or control.Top < filter.Top + filter.Height + 4 Then Exit Function
        If control.Top + control.Height > mLstInventory.Top - 22 Then Exit Function
        right = control.Left + control.Width
    Next name
    RecordingGeometryForTest = True
Missing:
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
Public Function RecordingGeometryForTest() As Boolean
    If mInventoryViewer Is Nothing Then Err.Raise 5, , "Recording fixture Viewer is not open."
    RecordingGeometryForTest = mInventoryViewer.RecordingGeometryForTest()
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
    $journalRoot=Join-Path $Fixture.Root ('Training/ActionPaths/'+$Fixture.Warehouse)
    function RecordingJournal([string]$Sequence) {
        if(-not $Sequence -or -not (Test-Path -LiteralPath $journalRoot)){return}
        foreach($file in Get-ChildItem -LiteralPath $journalRoot -Filter '*.json' -File) {
            if($file.Length -gt 1048576){throw 'Recording product wrote an oversized record.'}
            $text=[IO.File]::ReadAllText($file.FullName)
            $record=$text|ConvertFrom-Json
            if($record.SequenceId -ceq $Sequence){$record}
        }
    }
    function JournalFact([string]$Sequence,[string]$Type,[int]$Count,[string]$Lifecycle='Recording') {
        $entries=@(RecordingJournal $Sequence|Where-Object {$_.RecordType -ceq $Type}|Sort-Object Version)
        if(-not $entries.Count){return $false}
        $record=$entries[-1]
        return ($record.SchemaVersion -eq 1 -and $record.RecordKind -ceq 'Recording' -and
            $record.Lifecycle -ceq $Lifecycle -and $record.ActionCount -eq $Count -and
            $record.WarehouseId -ceq $Fixture.Warehouse -and $record.CreatedByUserId -ceq 'config-admin' -and
            $record.ActionPathId -cne $record.SequenceId -and $record.RecordId -cne $record.ActionPathId)
    }
    function JournalChain([string]$Sequence,[int]$ExpectedCount) {
        $entries=@(RecordingJournal $Sequence|Sort-Object Version)
        if($entries.Count -ne $ExpectedCount){return $false}
        $previousId='';$previousHash='';$version=0
        foreach($entry in $entries){
            $version++
            $path=Join-Path $journalRoot ($entry.ActionPathId+'.'+$version+'.json')
            if(-not (Test-Path -LiteralPath $path)){return $false}
            $text=[IO.File]::ReadAllText($path);$marker=$text.LastIndexOf(',"ContentSha256":"',[StringComparison]::Ordinal)
            if($marker -lt 0){return $false}
            $body=$text.Substring(0,$marker)+'}'
            $sha=[Security.Cryptography.SHA256]::Create()
            try{$hash=([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($body)))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
            if($entry.Version -ne $version -or $entry.PreviousRecordId -cne $previousId -or $entry.PreviousSha256 -cne $previousHash -or $entry.ContentSha256 -cne $hash){return $false}
            $previousId=$entry.RecordId;$previousHash=$hash
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
        foreach($layout in @('FitMinimum','FitDefault','FitLarger','FitDefault')) {
            $suffix=if($layout -eq 'FitDefault' -and @($results.Check) -contains 'Recording.Layout.FitDefault'){'.Restored'}else{''}
            $fits=[bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @($layout,''))
            Check ('Recording.Layout.'+$layout+$suffix) ($fits -and [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingGeometryForTest'))
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
        $atStart=@(if(Test-Path -LiteralPath $journalRoot){Get-ChildItem -LiteralPath $journalRoot -Filter '*.json' -File|ForEach-Object {[IO.File]::ReadAllText($_.FullName)|ConvertFrom-Json}})
        Check 'Recording.StartDurableBeforeAnyAction' ($atStart.Count -eq 1 -and $atStart[0].RecordType -ceq 'Start' -and $atStart[0].ActionCount -eq 0)
        $first=SaveRecordedSetting '612'
        Check 'Recording.FirstAttemptAndOutcomeShareSequenceAndOrdinal' (HasSequence $first 1)
        Check 'Recording.DurableStartUsesDistinctIdentities' (JournalFact $first.Attempt.SequenceId 'Start' 0)
        Check 'Recording.AttemptAndResultPersistBeforeStop' ((JournalFact $first.Attempt.SequenceId 'Observation' 1) -and (JournalChain $first.Attempt.SequenceId 3))
        $firstPins=ActivityPins
        $second=SaveRecordedSetting '613'
        Check 'Recording.SecondAttemptAndOutcomeShareSequenceAndOrdinal' (HasSequence $second 2)
        Check 'Recording.RepeatedCommandRetainsDistinctOccurrence' ((HasSequence $first 1) -and (HasSequence $second 2) -and $first.Attempt.SequenceId -ceq $second.Attempt.SequenceId -and $first.Attempt.ActivityId -cne $second.Attempt.ActivityId)
        Check 'Recording.LaterActionPreservesEarlierBytes' (PinsRetained $firstPins)
        Check 'Recording.StopUsesActualControl' ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')
        Check 'Recording.StopFreezesWithoutAssertingConclusion' ((RecordingStatus) -match '(?i)stopped' -and (RecordingStatus) -notmatch '(?i)conclusion observed')
        Check 'Recording.StopAppendsHashedClosingRecord' ((JournalFact $first.Attempt.SequenceId 'Close' 2 'Stopped') -and (JournalChain $first.Attempt.SequenceId 6))
        $closed=@(RecordingJournal $first.Attempt.SequenceId|Where-Object RecordType -CEQ 'Close')
        $savedIds=@(if($closed.Count -eq 1){$closed[0].Observations|ForEach-Object RecordId})
        $originalIds=@($first.Attempt.RecordId,$first.Outcome.RecordId,$second.Attempt.RecordId,$second.Outcome.RecordId)
        Check 'Recording.CloseRetainsEveryOriginalObservationInOrder' ($savedIds.Count -eq 4 -and ($savedIds -join '|') -ceq ($originalIds -join '|'))
        $after=SaveRecordedSetting '614'
        Check 'Recording.AfterStopUsesOrdinaryActivity' ($after.Attempt.SequenceId -ceq '' -and $after.Outcome.SequenceId -ceq '' -and $after.Attempt.Ordinal -eq 0 -and $after.Outcome.Ordinal -eq 0)
        Check 'Recording.SecondStartUsesActualControl' ((RecordingControl 'Start Recording' 'Click') -ceq 'DELIVERED')
        $cancelled=SaveRecordedSetting '615'
        Check 'Recording.NewRunHasDistinctSequence' ((HasSequence $cancelled 1) -and (HasSequence $first 1) -and $cancelled.Attempt.SequenceId -cne $first.Attempt.SequenceId)
        $cancelPins=ActivityPins
        Check 'Recording.CancelUsesActualControl' ((RecordingControl 'Cancel Recording' 'Click') -ceq 'DELIVERED')
        Check 'Recording.CancelShowsCancelled' ((RecordingStatus) -match '(?i)cancelled')
        Check 'Recording.CancelPreservesObservedActivity' (PinsRetained $cancelPins)
        Check 'Recording.CancelAppendsCancelledRecord' ((JournalFact $cancelled.Attempt.SequenceId 'Close' 1 'Cancelled') -and (JournalChain $cancelled.Attempt.SequenceId 4))
        [void](RecordingControl 'Start Recording' 'Click')
        $policyRun=SaveRecordedSetting '616'
        SetRecordingPolicy $false
        Check 'Recording.PolicyChangeClosesIncomplete' (JournalFact $policyRun.Attempt.SequenceId 'Close' 1 'Incomplete')
        CloseRecordingViewer
        SetRecordingPolicy $true
        OpenRecordingViewer
        [void](RecordingControl 'Start Recording' 'Click')
        $bindingRun=SaveRecordedSetting '617'
        CloseRecordingViewer
        Check 'Recording.ViewerClosureMarksIncomplete' (JournalFact $bindingRun.Attempt.SequenceId 'Close' 1 'Incomplete')
        OpenRecordingViewer
        Check 'Recording.ReopenDoesNotResumeOldSequence' ((RecordingControl 'Stop Recording') -ceq 'True|False')
        [void](RecordingControl 'Start Recording' 'Click')
        $sessionRun=SaveRecordedSetting '618'
        [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
        Check 'Recording.SignOutClosesIncomplete' (JournalFact $sessionRun.Attempt.SequenceId 'Close' 1 'Incomplete')
        CloseRecordingViewer
        SelectTarget $Fixture 'config-admin'
        # Fault only the generated fixture's training directory; restore in finally.
        $resolved=[IO.Path]::GetFullPath($journalRoot)
        $allowed=[IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\'
        if(-not $resolved.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)){throw 'Journal fault path escaped fixture.'}
        $backup=[IO.Path]::GetFullPath($resolved+'.test-backup')
        if(-not $backup.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)){throw 'Journal backup path escaped fixture.'}
        $hadJournal=Test-Path -LiteralPath $resolved -PathType Container
        if($hadJournal){Move-Item -LiteralPath $resolved -Destination $backup}
        New-Item -ItemType Directory -Path (Split-Path $resolved) -Force|Out-Null
        try {
            [IO.File]::WriteAllText($resolved,'Fixture storage obstruction.')
            OpenRecordingViewer
            [void](RecordingControl 'Start Recording' 'Click')
            Check 'Recording.StorageFailureNeverReportsActive' ((RecordingStatus) -match '(?i)(unavailable|could not|failed|incomplete)' -and (RecordingControl 'Stop Recording') -ceq 'True|False')
            $failedStart=SaveRecordedSetting '619'
            Check 'Recording.StorageFailurePreservesOrdinaryAction' ($failedStart.Attempt.SequenceId -ceq '' -and $failedStart.Outcome.SequenceId -ceq '' -and $failedStart.Attempt.Ordinal -eq 0)
        } finally {
            CloseRecordingViewer
            if(Test-Path -LiteralPath $resolved -PathType Leaf){Remove-Item -LiteralPath $resolved}
            if($hadJournal){Move-Item -LiteralPath $backup -Destination $resolved}
        }
        SelectTarget $Fixture 'config-reader'
        OpenRecordingViewer
        Check 'Recording.OrdinaryViewerUserCanStartOwnRun' ((RecordingControl 'Start Recording' 'Click') -ceq 'DELIVERED' -and (RecordingStatus) -match '(?i)recording.*0\s*/\s*256')
        Check 'Recording.OrdinaryViewerUserCanStopOwnRun' ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED' -and (RecordingStatus) -match '(?i)stopped')
        if($CheckRecordingLimits) {
            CloseRecordingViewer
            SelectTarget $Fixture 'config-admin'
            OpenRecordingViewer
            [void](RecordingControl 'Start Recording' 'Click')
            $beforeLimit=ActivityPins
            try {
                [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
                foreach($index in 1..256) {
                    if(-not [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('BatchSize',[string](700+$index)))){throw 'Action-limit owning Settings handler failed.'}
                    if($index % 32 -eq 0){Write-Output ('Recording boundary fixture: '+$index+' ordinary actions completed.')}
                }
            } finally {[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')}
            $limitRecords=@(foreach($file in Get-ChildItem -LiteralPath $activityRoot -Filter '*.json' -File) {
                if(-not $beforeLimit.ContainsKey($file.Name)){[IO.File]::ReadAllText($file.FullName)|ConvertFrom-Json}
            })
            $attempts=@($limitRecords|Where-Object OutcomeCode -CEQ 'REQUESTED'|Sort-Object Ordinal)
            if($limitRecords.Count -ne 512 -or $attempts.Count -ne 256){throw 'Action-limit activity fixture lacks complete owning outcomes.'}
            $sequence=[string]$attempts[0].SequenceId
            Check 'Recording.LimitRetainsExactly256DistinctOccurrences' ($sequence -ne '' -and @($attempts.SequenceId|Select-Object -Unique).Count -eq 1 -and @($attempts.ActivityId|Select-Object -Unique).Count -eq 256 -and ($attempts.Ordinal -join ',') -ceq ((1..256) -join ','))
            Check 'Recording.LimitClosesPartialWithAll512Observations' ((JournalFact $sequence 'Close' 256 'Incomplete') -and (JournalChain $sequence 514) -and @((RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close').Observations).Count -eq 512)
            [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('Refresh',''))
            Check 'Recording.LimitVisibleCounterAndReason' ((RecordingStatus) -match '(?i)partial.*action limit reached.*256\s*/\s*256' -and (RecordingControl 'Stop Recording') -ceq 'True|False')
            $afterLimit=SaveRecordedSetting '999'
            Check 'Recording.Action257ContinuesOutsideClosedSequence' ($afterLimit.Attempt.SequenceId -ceq '' -and $afterLimit.Outcome.SequenceId -ceq '' -and $afterLimit.Attempt.Ordinal -eq 0)
        }
        if($CheckRecordingReader) {
            . (Join-Path $PSScriptRoot 'Slice4beRecordingReader.ps1')
            Test-Slice4beRecordingReader $Fixture
        }
    } finally { CloseRecordingViewer }
}
