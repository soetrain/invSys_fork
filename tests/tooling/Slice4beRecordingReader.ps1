# D18 full-journal reads through the actual Viewer/library controls.
# Called inside the recording suite so its real handler fixture helpers remain
# in scope. No substitute reader, inferred business outcome or product seam.
function Test-Slice4beRecordingReader($Fixture) {
    CloseRecordingViewer
    $manager=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modInventoryViewer').CodeModule
    $manager.AddFromString(@'
Public Function RecordingLibraryForTest(ByVal action As String, Optional ByVal value As String = "") As String
    Dim item As Object, library As Object, control As Object, count As Long, r As Long, c As Long
    If action = "Open" Then
        If mInventoryViewer Is Nothing Then Err.Raise 5, , "Viewer fixture is missing."
        For Each control In mInventoryViewer.Controls
            If TypeName(control) = "CommandButton" Then
                If control.Caption = "Action Paths" Then
                    If Not control.Visible Or Not control.Enabled Then RecordingLibraryForTest = "DISABLED": Exit Function
                    control.Value = True
                    RecordingLibraryForTest = "DELIVERED": Exit Function
                End If
            End If
        Next control
        RecordingLibraryForTest = "MISSING": Exit Function
    End If
    For Each item In VBA.UserForms
        If item.Name = "frmActionPaths" Then Set library = item: count = count + 1
    Next item
    If action = "Count" Then RecordingLibraryForTest = CStr(count): Exit Function
    If library Is Nothing Then RecordingLibraryForTest = "MISSING": Exit Function
    Select Case action
        Case "Select"
            Set control = library.Controls("lstActionPaths")
            For r = 0 To control.ListCount - 1
                For c = 0 To control.ColumnCount - 1
                    If CStr(control.List(r, c)) = value Then
                        control.ListIndex = -1: control.ListIndex = r
                        RecordingLibraryForTest = "SELECTED": Exit Function
                    End If
                Next c
            Next r
            RecordingLibraryForTest = "NOT FOUND": Exit Function
        Case "Search": library.Controls("txtPathSearch").Value = value
        Case "Refresh": library.Controls("btnPathRefresh").Value = True
        Case "Rows": RecordingLibraryForTest = CStr(library.Controls("lstActionPaths").ListCount): Exit Function
        Case "Evidence": RecordingLibraryForTest = CStr(library.Controls("txtPathEvidence").Value): Exit Function
        Case "Status": RecordingLibraryForTest = CStr(library.Controls("lblPathStatus").Caption): Exit Function
        Case "ReadOnly": RecordingLibraryForTest = CStr(library.Controls("txtPathEvidence").Locked): Exit Function
    End Select
    RecordingLibraryForTest = "DELIVERED"
End Function
'@)
    function Library([string]$Action,[string]$Value='') {
        [string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @($Action,$Value))
    }
    function InspectRun {
        [void](Library 'Refresh')
        [void](Library 'Select' $pathId)
        (Library 'Status')+"`n"+(Library 'Evidence')
    }
    function SaveVisibility([bool]$Visible) {
        try {
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
            if(-not [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.PublishedReadVisibilityForTest' @($Visible))) {
                throw 'Actual Admin visibility policy save failed; not reader RED.'
            }
        } finally { [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings') }
    }
    SelectTarget $Fixture 'config-admin'
    SetRecordingPolicy $true
    OpenRecordingViewer
    if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'Durable recorder baseline is missing; not reader RED.'}
    $first=SaveRecordedSetting '620'
    $second=SaveRecordedSetting '621'
    [void](RecordingControl 'Stop Recording' 'Click')
    $sequence=$first.Attempt.SequenceId
    if(-not (HasSequence $first 1) -or -not (HasSequence $second 2) -or -not (JournalChain $sequence 6)) {
        throw 'Real owning actions did not produce the required journal fixture.'
    }
    $entries=@(RecordingJournal $sequence|Sort-Object Version)
    $pathId=[string]$entries[0].ActionPathId
    $bytes=@{}
    $trainingPins=@{}
    foreach($file in Get-ChildItem -LiteralPath (Join-Path $Fixture.Root 'Training') -Recurse -File) {
        $trainingPins[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash
    }
    foreach($entry in $entries){$file=Join-Path $journalRoot ($pathId+'.'+$entry.Version+'.json');$bytes[$file]=[IO.File]::ReadAllBytes($file)}
    $closePath=Join-Path $journalRoot ($pathId+'.6.json')
    $middlePath=Join-Path $journalRoot ($pathId+'.2.json')
    $utf8=[Text.UTF8Encoding]::new($false)
    $publishBefore=[long](Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest')
    $shippingBefore=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.PublishedReadAuthorityCallsForTest')
    try {
        Check 'RecordingRead.ViewerActionPathsOpensLibrary' ((Library 'Open') -ceq 'DELIVERED' -and (Library 'Count') -ceq '1')
        [void](Library 'Open')
        Check 'RecordingRead.RepeatedLaunchReusesLibrary' ((Library 'Count') -ceq '1')
        Check 'RecordingRead.RealRunSelectable' ((Library 'Select' $pathId) -ceq 'SELECTED')
        $text=(Library 'Status')+"`n"+(Library 'Evidence')
        Check 'RecordingRead.StoppedNeverAssertsConclusion' ($text -match '(?i)stopped' -and $text -notmatch '(?i)conclusion observed')
        $firstPosition=$text.IndexOf($first.Attempt.ActivityId,[StringComparison]::Ordinal)
        $secondPosition=$text.IndexOf($second.Attempt.ActivityId,[StringComparison]::Ordinal)
        Check 'RecordingRead.PreservesOrderedDistinctOccurrences' ($firstPosition -ge 0 -and $secondPosition -gt $firstPosition -and $text -match 'REQUESTED' -and $text -match 'COMPLETED')
        Check 'RecordingRead.EvidenceIsReadOnly' ((Library 'ReadOnly') -ceq 'True')
        [void](Library 'Search' 'NO-MATCH-RECORDING-FIXTURE')
        Check 'RecordingRead.SearchFiltersLibrary' ((Library 'Rows') -ceq '0')
        [void](Library 'Search' '')
        Check 'RecordingRead.SearchClearRestoresRun' ((Library 'Select' $pathId) -ceq 'SELECTED')
        foreach($fault in @('MissingClose','MissingMiddle','DamagedHash','BrokenPreviousLink','ClosingObservationOmitted','RepeatedOrdinal')) {
            try {
                if($fault -eq 'MissingClose'){Remove-Item -LiteralPath $closePath}
                elseif($fault -eq 'MissingMiddle'){Remove-Item -LiteralPath $middlePath}
                elseif($fault -eq 'DamagedHash') {
                    $body=$utf8.GetString($bytes[$closePath]).Replace('"ContentSha256":"','"ContentSha256":"0')
                    [IO.File]::WriteAllText($closePath,$body,$utf8)
                } else {
                    $previousHash=''
                    foreach($entry in $entries) {
                        $file=Join-Path $journalRoot ($pathId+'.'+$entry.Version+'.json')
                        $model=$utf8.GetString($bytes[$file])|ConvertFrom-Json
                        $model.PSObject.Properties.Remove('ContentSha256')
                        $model.PreviousSha256=$previousHash
                        if($fault -eq 'BrokenPreviousLink' -and $model.Version -eq 6){$model.PreviousSha256='0'*64}
                        if($fault -eq 'ClosingObservationOmitted' -and $model.Version -eq 6){$model.Observations=@($model.Observations|Select-Object -First 1)}
                        if($fault -eq 'RepeatedOrdinal'){foreach($observation in $model.Observations){if($observation.ActivityId -ceq $second.Attempt.ActivityId){$observation.Ordinal=1}}}
                        $body=$model|ConvertTo-Json -Depth 30 -Compress
                        $sha=[Security.Cryptography.SHA256]::Create()
                        try{$previousHash=([BitConverter]::ToString($sha.ComputeHash($utf8.GetBytes($body)))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
                        $content=$body.Substring(0,$body.Length-1)+',"ContentSha256":"'+$previousHash+'"}'
                        [IO.File]::WriteAllText($file,$content,$utf8)
                    }
                }
                $text=InspectRun
                $required=if($fault -eq 'MissingClose'){'(?i)interrupted'}else{'(?i)incomplete evidence'}
                Check ('RecordingRead.'+$fault+'IsExplicit') ($text -match $required -and $text -notmatch '(?i)conclusion observed')
            } finally {foreach($file in $bytes.Keys){[IO.File]::WriteAllBytes($file,$bytes[$file])}}
        }
        $text=InspectRun
        Check 'RecordingRead.RestoredJournalReadsAgain' ($text -match '(?i)stopped' -and $text.Contains($second.Attempt.ActivityId))
        SaveVisibility $false
        [void](Library 'Search' 'Recorded')
        [void](Library 'Select' $pathId)
        $text=(Library 'Status')+"`n"+(Library 'Evidence')
        Check 'RecordingRead.CurrentPolicyHidesLoadedObservations' ($text -match '(?i)incomplete evidence' -and -not $text.Contains($first.Attempt.ActivityId) -and -not $text.Contains($second.Attempt.ActivityId))
        SaveVisibility $true
        $text=InspectRun
        Check 'RecordingRead.CurrentPolicyCanRestorePermittedEvidence' ($text.Contains($first.Attempt.ActivityId) -and $text.Contains($second.Attempt.ActivityId))
        CloseRecordingViewer
        Check 'RecordingRead.ViewerCloseClosesLibrary' ((Library 'Count') -ceq '0')
        OpenRecordingViewer
        [void](Library 'Open')
        [void](Library 'Select' $pathId)
        [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
        [void](Library 'Refresh')
        $text=(Library 'Status')+"`n"+(Library 'Evidence')
        Check 'RecordingRead.SignOutClearsCapturedEvidence' ((Library 'Count') -ceq '1' -and $text -match '(?i)(session|sign.in|context|unavailable)' -and -not $text.Contains($first.Attempt.ActivityId))
        Check 'RecordingRead.ReadsNeverPublishOrReadShippingAuthority' ([long](Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest') -eq $publishBefore -and [long](Run 'invSys.Operations.xlam' 'modTS_Shipments.PublishedReadAuthorityCallsForTest') -eq $shippingBefore)
        $unchanged=$trainingPins.Count -eq @(Get-ChildItem -LiteralPath (Join-Path $Fixture.Root 'Training') -Recurse -File).Count
        foreach($file in $trainingPins.Keys){if((Get-FileHash -LiteralPath $file).Hash -cne $trainingPins[$file]){$unchanged=$false}}
        Check 'RecordingRead.JournalAndActivityBytesPreserved' $unchanged
    } finally {
        foreach($file in $bytes.Keys){[IO.File]::WriteAllBytes($file,$bytes[$file])}
        CloseRecordingViewer
    }
}
