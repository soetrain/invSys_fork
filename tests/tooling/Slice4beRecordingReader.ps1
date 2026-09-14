# D18 full-journal reads through the actual Viewer/library controls.
# Called inside the recording suite so its real handler fixture helpers remain
# in scope. No substitute reader, inferred business outcome or product seam.
function Test-Slice4beRecordingReader($Fixture,$OtherFixture,[bool]$InstallOnly=$false) {
    CloseRecordingViewer
    $manager=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modInventoryViewer').CodeModule
    $manager.AddFromString(@'
Public Function RecordingLibraryForTest(ByVal action As String, Optional ByVal value As String = "") As String
    Dim item As Object, library As Object, control As Object, count As Long, r As Long, c As Long
    Dim other As Object
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
        Case "Close": library.Controls("btnClose").Value = True
        Case "Visible": RecordingLibraryForTest = CStr(library.Visible): Exit Function
        Case "FitMinimum", "FitDefault", "FitLarger"
            Select Case action
                Case "FitMinimum": library.Width = 720: library.Height = 520
                Case "FitDefault": library.Width = 820: library.Height = 640
                Case "FitLarger": library.Width = 1000: library.Height = 760
            End Select
            library.Repaint: DoEvents
            RecordingLibraryForTest = "False"
            For Each control In library.Controls
                If Not control.Visible Or control.Left < 0 Or control.Top < 0 Or control.Width <= 0 Or control.Height <= 0 Then Exit Function
                If control.Left + control.Width > library.InsideWidth Or control.Top + control.Height > library.InsideHeight Then Exit Function
                For Each other In library.Controls
                    If control.Name <> other.Name Then
                        If control.Left < other.Left + other.Width And other.Left < control.Left + control.Width And _
                           control.Top < other.Top + other.Height And other.Top < control.Top + control.Height Then Exit Function
                    End If
                Next other
            Next control
            RecordingLibraryForTest = "True": Exit Function
    End Select
    RecordingLibraryForTest = "DELIVERED"
End Function
'@)
    if($InstallOnly){return}
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
        foreach($layout in @('FitMinimum','FitDefault','FitLarger','FitDefault')) {
            $suffix=if($layout -eq 'FitDefault' -and @($results.Check) -contains 'RecordingRead.Layout.FitDefault'){'FitDefault.Restored'}else{$layout}
            Check ('RecordingRead.Layout.'+$suffix) ((Library $layout) -ceq 'True')
        }
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
        foreach($fault in @('MissingClose','MissingMiddle','DamagedHash','BrokenPreviousLink','ClosingObservationOmitted','RepeatedOrdinal',
            'ObservationPackageMismatch','ObservationBuildMismatch','ObservationCatalogMismatch','ForeignWarehouse','UnsupportedSchema','OlderPackage','OlderBuild','OlderCatalog')) {
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
                        if($fault -eq 'ForeignWarehouse'){$model.WarehouseId='FOREIGN-RECORDING-FIXTURE'}
                        if($fault -eq 'UnsupportedSchema'){$model.SchemaVersion=99}
                        if($fault -eq 'OlderPackage'){$model.PackageSetVersion='prior-package-fixture'}
                        if($fault -eq 'OlderBuild'){$model.BuildIdentity='prior-build-fixture'}
                        if($fault -eq 'OlderCatalog'){$model.CatalogVersion=7}
                        foreach($observation in $model.Observations) {
                            if($fault -in @('ObservationPackageMismatch','OlderPackage')){$observation.PackageSetVersion='prior-package-fixture'}
                            if($fault -eq 'OlderBuild' -or ($fault -eq 'ObservationBuildMismatch' -and $observation.OutcomeCode -cne 'REQUESTED')){$observation.BuildIdentity='prior-build-fixture'}
                            if($fault -in @('ObservationCatalogMismatch','OlderCatalog')){$observation.CatalogVersion=7}
                        }
                        $body=$model|ConvertTo-Json -Depth 30 -Compress
                        $sha=[Security.Cryptography.SHA256]::Create()
                        try{$previousHash=([BitConverter]::ToString($sha.ComputeHash($utf8.GetBytes($body)))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
                        $content=$body.Substring(0,$body.Length-1)+',"ContentSha256":"'+$previousHash+'"}'
                        [IO.File]::WriteAllText($file,$content,$utf8)
                    }
                    if($fault -ne 'BrokenPreviousLink' -and -not (JournalChain $sequence 6)) {
                        throw 'Rehashed reader fault fixture lacks a complete six-entry chain.'
                    }
                }
                $text=InspectRun
                # Preserve established check IDs; these opaque fixture labels
                # cannot establish chronological age from package/build IDs.
                $required=if($fault -eq 'MissingClose'){'(?i)interrupted'}elseif($fault -in @('OlderPackage','OlderBuild')){'(?i)different release/build.*relative age unavailable'}elseif($fault -eq 'OlderCatalog'){'(?i)older release'}else{'(?i)incomplete evidence'}
                Check ('RecordingRead.'+$fault+'IsExplicit') ($text -match $required -and $text -notmatch '(?i)conclusion observed')
                if($fault -in @('OlderPackage','OlderBuild','OlderCatalog')) {
                    Check ('RecordingRead.'+$fault+'RetainsOriginalEvidence') ($text.Contains($first.Attempt.ActivityId) -and $text.Contains($second.Attempt.ActivityId) -and $text -match 'COMPLETED')
                }
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
        [void](Library 'Close')
        Check 'RecordingRead.ActualCloseHidesLibrary' ((Library 'Visible') -ceq 'False')
        [void](Library 'Open')
        Check 'RecordingRead.OpenAfterCloseReusesAndReadsLibrary' ((Library 'Visible') -ceq 'True' -and (Library 'Count') -ceq '1' -and (Library 'Select' $pathId) -ceq 'SELECTED' -and (Library 'Evidence').Contains($first.Attempt.ActivityId))
        CloseRecordingViewer
        Check 'RecordingRead.ViewerCloseClosesLibrary' ((Library 'Count') -ceq '0')
        OpenRecordingViewer
        [void](Library 'Open')
        [void](Library 'Select' $pathId)
        if($null -eq $OtherFixture -or $OtherFixture.Warehouse -ceq $Fixture.Warehouse){throw 'Distinct warehouse fixture is required.'}
        SelectTarget $OtherFixture 'config-reader'
        [void](Library 'Refresh')
        $text=(Library 'Status')+"`n"+(Library 'Evidence')
        Check 'RecordingRead.TargetChangeClearsCapturedEvidence' ((Library 'Count') -ceq '1' -and $text -match '(?i)(session|warehouse|context|unavailable)' -and -not $text.Contains($first.Attempt.ActivityId))
        CloseRecordingViewer
        SelectTarget $Fixture 'config-admin'
        OpenRecordingViewer
        [void](Library 'Open')
        Check 'RecordingRead.ReopenAfterTargetChangeReadsOriginalRun' ((Library 'Select' $pathId) -ceq 'SELECTED' -and (Library 'Evidence').Contains($first.Attempt.ActivityId))
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
