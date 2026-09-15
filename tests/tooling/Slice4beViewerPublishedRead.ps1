# D18: real Settings activity -> real Admin publication -> actual Viewer handlers.
# Faults affect only disposable publication bytes; ordering uses a synthetic wire.
function Test-Slice4beViewerPublishedRead($Fixture,$OtherFixture,[bool]$InstallOnly=$false) {
    if(-not $InstallOnly -and -not $RecordingEvaluationDiagnostic){
    SelectTarget $Fixture 'config-admin'
    try {
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
        if(-not [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('BatchSize','601'))) {throw 'Actual Settings activity fixture failed.'}
    } finally {[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')}
    }
    if(-not $script:PublishedReadProbeInstalled){
    $admin=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('modAdminConsole').CodeModule
    $admin.AddFromString(@'
Public Function PublishReadFixtureForTest() As Boolean
    Dim report As String
    PublishReadFixtureForTest = GenerateInventorySnapshot("config-admin", "", Nothing, "", Nothing, report)
End Function
'@)
    }
    if(-not $InstallOnly -and -not $RecordingEvaluationDiagnostic){
    if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')) {throw 'Actual Admin publication fixture failed.'}
    $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')
    if(-not (Test-Path -LiteralPath $path)){throw 'Published-read fixture requires the independently tested publisher candidate.'}
    $original=[IO.File]::ReadAllText($path)
    $model=$original|ConvertFrom-Json
    $activity=@($model.Groups|Where-Object {$_.Source -ceq 'Activity' -and @($_.Lines|Where-Object ControlId -CEQ 'ADMIN_SETTINGS_SAVE_VALUE').Count -gt 0})
    if($activity.Count -ne 1 -or $activity[0].Lines.Count -ne 2 -or $activity[0].Outcomes.Count -ne 1){throw 'Actual publication lacks the correlated Settings activity fixture.'}
    $activityId=[string]$activity[0].SourceId
    }
    if(-not $script:PublishedReadProbeInstalled){
    $core=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modInventoryViewerData').CodeModule
    $core.InsertLines($core.CountOfDeclarationLines+1,'Private PublishedOrderingPayloadForTest As String')
    $core.AddFromString(@'
Public Function PublishedReadFixtureValidForTest(ByVal path As String, ByVal warehouse As String) As Boolean
    Dim model As Object
    Set model = modEventsPublicationStore.Read(path, warehouse)
    PublishedReadFixtureValidForTest = Not model Is Nothing
End Function
Public Sub SetPublishedOrderingForTest(ByVal payload As String)
    PublishedOrderingPayloadForTest = payload
End Sub
'@)
    $entry=$core.ProcBodyLine('LoadCurrentInventoryEventViewerData',0)
    $core.InsertLines($entry+1,@'
    If PublishedOrderingPayloadForTest <> "" Then
        LoadCurrentInventoryEventViewerData = PublishedOrderingPayloadForTest
        Exit Function
    End If
'@)
    }
    if(-not $InstallOnly -and -not $RecordingEvaluationDiagnostic){
    $valid=[bool](Run 'invSys.Core.xlam' 'modInventoryViewerData.PublishedReadFixtureValidForTest' @($path,$Fixture.Warehouse))
    Check 'PublishedRead.RealPublicationFixtureValid' $valid
    if(-not $valid){throw 'Fixture schema or integrity is invalid; not a behavioral RED.'}
    Check 'PublishedRead.RealSettingsAttemptAndOutcomePublished' ($activity[0].Lines[0].RecordId -cne $activity[0].Lines[1].RecordId)
    }

    if(-not $script:PublishedReadProbeInstalled){
    $shipping=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modTS_Shipments').CodeModule
    $entry=$shipping.ProcBodyLine('LoadShippingViewerSupplementEvents',0)
    $shipping.InsertLines($entry+1,'    PublishedReadAuthorityCountForTest = PublishedReadAuthorityCountForTest + 1')
    $shipping.AddFromString(@'
Private PublishedReadAuthorityCountForTest As Long
Public Function PublishedReadAuthorityCallsForTest() As Long
    PublishedReadAuthorityCallsForTest = PublishedReadAuthorityCountForTest
End Function
'@)
    $publisher=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modWarehouseSync').CodeModule
    $entry=$publisher.ProcBodyLine('GenerateWarehouseSnapshot',0)
    # This signature spans lines; insert after its declaration is complete.
    while($publisher.Lines($entry,1).TrimEnd().EndsWith('_')){$entry++}
    $publisher.InsertLines($entry+1,'    PublishedReadPublishCountForTest = PublishedReadPublishCountForTest + 1')
    $publisher.AddFromString(@'
Private PublishedReadPublishCountForTest As Long
Public Function PublishedReadPublishCallsForTest() As Long
    PublishedReadPublishCallsForTest = PublishedReadPublishCountForTest
End Function
'@)
    $form=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmInventoryViewer').CodeModule
    $form.AddFromString(@'
Public Function PublishedReadActionForTest(ByVal action As String, Optional ByVal expected As String = "") As Boolean
    Dim r As Long, c As Long, source As Long
    On Error GoTo Failed
    Select Case action
        Case "Events": mCboEventRange.Value = "All": mTabs.Value = 1: mBtnRefresh_Click
        Case "Refresh": mBtnRefresh_Click
        Case "Day": mCboEventRange.Value = "Day": mBtnRefresh_Click
        Case "Search": mTxtSearch.Value = expected
        Case "MixedOrder"
            If mVisibleIndexes.Count <> 3 Then Exit Function
            PublishedReadActionForTest = (CStr(mRows(CLng(mVisibleIndexes(1)), 11)) = "MIX-A-INVENTORY" And _
                CStr(mRows(CLng(mVisibleIndexes(2)), 11)) = "MIX-Z-DESIGNS" And CStr(mRows(CLng(mVisibleIndexes(3)), 11)) = "MIX-A-ACTIVITY"): Exit Function
        Case "FitMinimum": Me.Width = 720: Me.Height = 430: Me.Repaint: DoEvents: PublishedReadActionForTest = PublishedReadPageFitsForTest(): Exit Function
        Case "FitDefault": Me.Width = 860: Me.Height = 535: Me.Repaint: DoEvents: PublishedReadActionForTest = PublishedReadPageFitsForTest(): Exit Function
        Case "FitLarger": Me.Width = 1000: Me.Height = 650: Me.Repaint: DoEvents: PublishedReadActionForTest = PublishedReadPageFitsForTest(): Exit Function
        Case "ContainsSource", "SelectSource"
            If IsEmpty(mRows) Then Exit Function
            For r = 1 To mVisibleIndexes.Count
                source = CLng(mVisibleIndexes(r))
                For c = LBound(mRows, 2) To UBound(mRows, 2)
                    If StrComp(CStr(mRows(source, c)), expected, vbBinaryCompare) = 0 Then
                        If action = "SelectSource" Then mLstInventory.ListIndex = r - 1
                        PublishedReadActionForTest = True: Exit Function
                    End If
                Next c
            Next r
            Exit Function
        Case "Unavailable"
            PublishedReadActionForTest = (InStr(1, mLblStatus.Caption, "Unavailable", vbTextCompare) > 0): Exit Function
        Case "Stale"
            PublishedReadActionForTest = (InStr(1, mLblStatus.Caption, "Stale", vbTextCompare) > 0): Exit Function
        Case "Fresh"
            PublishedReadActionForTest = (InStr(1, mLblStatus.Caption, "Stale", vbTextCompare) = 0 And InStr(1, mLblStatus.Caption, "Unavailable", vbTextCompare) = 0): Exit Function
        Case "Empty"
            PublishedReadActionForTest = (mLstInventory.ListCount = 0 And IsEmpty(mRows)): Exit Function
        Case Else: Exit Function
    End Select
    PublishedReadActionForTest = True
Failed:
End Function
Private Function PublishedReadPageFitsForTest() As Boolean
    Dim previous As Object, following As Object, page As Object, name As Variant, control As Object
    On Error GoTo Failed
    Set previous = Me.Controls("btnEventsPrevious"): Set following = Me.Controls("btnEventsNext")
    Set page = Me.Controls("lblEventPage")
    For Each name In Array("lstInventory", "btnEventsPrevious", "btnEventsNext", "lblEventPage", "lblStatus", "btnClose")
        Set control = Me.Controls(CStr(name))
        If Not control.Visible Or control.Left < 0 Or control.Top < 0 Or control.Width <= 0 Or control.Height <= 0 Then Exit Function
        If control.Left + control.Width > Me.InsideWidth + 1 Or control.Top + control.Height > Me.InsideHeight + 1 Then Exit Function
    Next name
    If mLstInventory.Top + mLstInventory.Height > previous.Top Then Exit Function
    If previous.Left + previous.Width > following.Left Or following.Left + following.Width > page.Left Then Exit Function
    If previous.Top + previous.Height > mLblStatus.Top Or following.Top + following.Height > mLblStatus.Top Or page.Top + page.Height > mLblStatus.Top Then Exit Function
    PublishedReadPageFitsForTest = (mLstInventory.ListCount > 0)
Failed:
End Function
Public Function PublishedReadDetailForTest(ByVal caption As String, ByVal expected As String) As Boolean
    Dim fields As Variant, r As Long
    On Error GoTo Failed
    If mDetail Is Nothing Then Exit Function
    fields = mDetail.Fields(1)
    If IsEmpty(fields) Then Exit Function
    For r = LBound(fields, 1) To UBound(fields, 1)
        If CStr(fields(r, 0)) = caption Then
            PublishedReadDetailForTest = (InStr(1, CStr(fields(r, 1)), expected, vbBinaryCompare) > 0): Exit Function
        End If
    Next r
Failed:
End Function
Public Function PublishedReadGeometryForTest() As String
    Dim name As Variant, control As Object
    For Each name In Array("lstInventory", "btnEventsPrevious", "btnEventsNext", "lblEventPage", "lblStatus", "btnClose")
        Set control = Me.Controls(CStr(name))
        PublishedReadGeometryForTest = PublishedReadGeometryForTest & CStr(name) & vbTab & CStr(control.Left) & vbTab & CStr(control.Top) & vbTab & _
            CStr(control.Width) & vbTab & CStr(control.Height) & vbTab & CStr(control.Visible) & vbTab & CStr(Me.InsideWidth) & vbTab & CStr(Me.InsideHeight) & vbLf
    Next name
End Function
'@)
    $manager=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modInventoryViewer').CodeModule
    $manager.AddFromString(@'
Public Function PublishedReadActionForTest(ByVal action As String, Optional ByVal expected As String = "") As Boolean
    If Not mInventoryViewer Is Nothing Then PublishedReadActionForTest = mInventoryViewer.PublishedReadActionForTest(action, expected)
End Function
Public Function PublishedReadDetailForTest(ByVal caption As String, ByVal expected As String) As Boolean
    If Not mInventoryViewer Is Nothing Then PublishedReadDetailForTest = mInventoryViewer.PublishedReadDetailForTest(caption, expected)
End Function
Public Function PublishedReadWindowForTest() As Double
    mInventoryViewer.Repaint: DoEvents
    PublishedReadWindowForTest = CDbl(modUserFormResizeWin.GetUserFormWindowHandle(mInventoryViewer))
End Function
Public Function PublishedReadGeometryForTest() As String
    PublishedReadGeometryForTest = mInventoryViewer.PublishedReadGeometryForTest()
End Function
Public Function PublishedDetailLabelForTest(ByVal action As String, Optional ByVal expected As String = "") As String
    Dim instance As Object, lines As Object, index As Long, labels As String
    For Each instance In VBA.UserForms
        If TypeName(instance) = "frmEventDetail" Then
            If action = "Prompt" Then
                PublishedDetailLabelForTest = CStr(instance.Controls("lblDetailLines").Caption = expected)
            ElseIf action = "Window" Then
                PublishedDetailLabelForTest = CStr(modUserFormResizeWin.GetUserFormWindowHandle(instance))
            ElseIf action = "Labels" Then
                Set lines = instance.Controls("lstEventLines")
                For index = 0 To lines.ListCount - 1
                    If index > 0 Then labels = labels & vbLf
                    labels = labels & CStr(lines.List(index, 0))
                Next index
                PublishedDetailLabelForTest = CStr(labels = expected)
            End If
            Exit Function
        End If
    Next instance
End Function
'@)
    $policy=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('cAdminTrackingPolicy').CodeModule
    $policy.AddFromString(@'
Public Function PublishedReadVisibilityForTest(ByVal enabled As Boolean) As Boolean
    mAdminVisible.Value = enabled
    mAdminVisible_Click
    mSave_Click
    PublishedReadVisibilityForTest = (InStr(1, mStatus.Caption, " saved.", vbBinaryCompare) > 0)
End Function
'@)
    $settings=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('frmAdminSettings').CodeModule
    $settings.AddFromString(@'
Public Function PublishedReadVisibilityForTest(ByVal enabled As Boolean) As Boolean
    PublishedReadVisibilityForTest = mTracking.PublishedReadVisibilityForTest(enabled)
End Function
'@)
    $commands=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('TestD5Commands').CodeModule
    $commands.AddFromString(@'
Public Function PublishedReadVisibilityForTest(ByVal enabled As Boolean) As Boolean
    PublishedReadVisibilityForTest = mForm.PublishedReadVisibilityForTest(enabled)
End Function
'@)
    $script:PublishedReadProbeInstalled=$true
    }
    if($InstallOnly -or $RecordingEvaluationDiagnostic){return}
    function ReadAct([string]$Action,[string]$Expected='') {[bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @($Action,$Expected))}
    function ReadDetail([string]$Caption,[string]$Expected) {[bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadDetailForTest' @($Caption,$Expected))}
    function WriteReadFixture([string]$Body) {
        $sha=[Security.Cryptography.SHA256]::Create()
        try{$hash=([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Body)))).Replace('-','')}finally{$sha.Dispose()}
        [IO.File]::WriteAllText($path,$Body.Substring(0,$Body.Length-1)+',"ContentSha256":"'+$hash+'"}',[Text.UTF8Encoding]::new($false))
    }
    $pins=@{}
    foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File){$pins[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash}
    $viewerAuthorityBefore=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.PublishedReadAuthorityCallsForTest')
    $viewerPublishBefore=[long](Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest')
    SelectTarget $Fixture 'config-reader'
    try {
        $excel.Visible=$true
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OpenInventoryViewer')
        [void](ReadAct 'Events')
        Check 'PublishedRead.EventsHandlerLoadsRealActivitySource' (ReadAct 'ContainsSource' $activityId)
        Check 'PublishedRead.SelectionUsesActualListHandler' (ReadAct 'SelectSource' $activityId)
        Check 'PublishedRead.DetailPreservesExactActivityIdentity' (ReadDetail 'Source event / activity ID' $activityId)
        Check 'PublishedRead.DetailShowsVerifiedPublicationTime' (ReadDetail 'Published' ([string]$model.PublishedAtUTC).Substring(0,10))
        Check 'PublishedRead.DetailLabelsActivityProvenance' (ReadDetail 'Source classification' 'User activity')
        $labels=@($activity[0].Lines|ForEach-Object {([string]$_.Caption)+' - '+([string]$_.OutcomeCode)}) -join "`n"
        Check 'PublishedRead.DetailPicker.DistinguishesActualAdminObservations' ([string](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedDetailLabelForTest' @('Labels',$labels)) -ceq 'True')
        Check 'PublishedRead.DetailPicker.SourceNeutralPrompt' ([string](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedDetailLabelForTest' @('Prompt','Contributing lines - select a line to inspect its fields')) -ceq 'True')
        if($CaptureEvidence){CaptureOwnedFormEvidence 'Event Detail' 'viewer-activity-detail-labels.png' ([long](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedDetailLabelForTest' @('Window','')))}
        [void](ReadAct 'Search' $activityId)
        Check 'PublishedRead.SearchRetainsExactActivityGroup' (ReadAct 'ContainsSource' $activityId)
        [void](ReadAct 'Search' '')
        Check 'PublishedRead.OpenRefreshSearchAvoidShippingAuthority' ([long](Run 'invSys.Operations.xlam' 'modTS_Shipments.PublishedReadAuthorityCallsForTest') -eq $viewerAuthorityBefore)
        [void](ReadAct 'Day')
        Check 'PublishedRead.DayFilterRetainsVerifiedUtcActivity' (ReadAct 'ContainsSource' $activityId)
        [void](ReadAct 'Events')
        Check 'PublishedRead.AllDatesRestoresActivity' (ReadAct 'ContainsSource' $activityId)
        foreach($layout in @('FitMinimum','FitDefault','FitLarger','FitDefault')) {
            $suffix=if($layout -eq 'FitDefault' -and @($results.Check) -contains 'PublishedRead.Layout.FitDefault'){'.Restored'}else{''}
            Check ('PublishedRead.Layout.'+$layout+$suffix) ((ReadAct $layout) -and (ReadAct 'ContainsSource' $activityId))
            [string](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadGeometryForTest') | Set-Content -LiteralPath (Join-Path $reportRoot ($layout+$suffix+'-geometry.tsv'))
        }
        # A synthetic serialized projection isolates the Operations comparator.
        # It never replaces a business owner or writes an authority/artifact.
        $wire=[string](Run 'invSys.Core.xlam' 'modInventoryViewerData.LoadCurrentInventoryEventViewerData')
        $header=($wire -split "`r`n")[0] -split "`t"
        if($header.Count -ne 12 -or $header[4] -cne 'EVENTS1'){throw 'Mixed ordering fixture requires the validated EVENTS1 reader.'}
        $fieldIds=$header[9] -split ','
        $header[3]='3';$header[6]=[guid]::NewGuid().ToString();$header[8]='Synthetic ordering fixture; no business effect.'
        $mixed=@($header -join "`t")
        foreach($record in @(@('Inventory','MIX-A-INVENTORY','Business event','Receiving','RECEIVE'),@('Activity','MIX-A-ACTIVITY','User activity','Admin','CONFIG_SAVE_REQUESTED'),@('Designs','MIX-Z-DESIGNS','Business event','Designs','PROCESS_SAVE'))) {
            $values=[string[]]::new(18+$fieldIds.Count)
            for($i=0;$i -lt $values.Length;$i++){$values[$i]=''}
            $values[0]=$header[2].Replace('T',' ').Replace('Z',' UTC');$values[1]=$record[4]
            $values[2]=$record[1];$values[9]='Synthetic ordering fixture.';$values[10]=$record[1]
            $values[12]=$record[4];$values[17]=$record[0]
            $details=@{SourceId=$record[1];SourceKind=$record[2];EventFamily=$record[3];EventCode=$record[4];EventType=$record[4];WarehouseId=$Fixture.Warehouse;TimeProvenance='Verified UTC';Coverage=$header[8];Explanation=$values[9]}
            foreach($field in $details.Keys){$column=[array]::IndexOf($fieldIds,$field);if($column -lt 0){throw 'Ordering detail field is missing.'};$values[18+$column]=$details[$field]}
            $mixed+=($values -join "`t")
        }
        try {
            [void](Run 'invSys.Core.xlam' 'modInventoryViewerData.SetPublishedOrderingForTest' @($mixed -join "`r`n"))
            [void](ReadAct 'Events')
            Check 'PublishedRead.MixedSourceEqualTimeOrder' (ReadAct 'MixedOrder')
        } finally {
            [void](Run 'invSys.Core.xlam' 'modInventoryViewerData.SetPublishedOrderingForTest' @(''))
            [void](ReadAct 'Events')
        }

        $utf8=[Text.UTF8Encoding]::new($false)
        [IO.File]::WriteAllText($path,$original.Replace('"ContentSha256":"','"ContentSha256":"0'),$utf8)
        [void](ReadAct 'Refresh')
        Check 'PublishedRead.DamagedHashRefreshIsStale' (ReadAct 'Stale')
        Check 'PublishedRead.DamagedHashRetainsLoadedActivity' (ReadAct 'ContainsSource' $activityId)
        [IO.File]::WriteAllText($path,$original,$utf8)
        [void](ReadAct 'Refresh')
        Check 'PublishedRead.RestoredPublicationRecoversActivity' ((ReadAct 'Fresh') -and (ReadAct 'ContainsSource' $activityId))
        if($CaptureEvidence -and -not $CheckViewerFilters){CaptureOwnedFormEvidence ('Viewer - '+$Fixture.Warehouse) 'viewer-published-read.png' ([long](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadWindowForTest'))}

        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.CloseInventoryViewerForTest')
        $body=$original.Substring(0,$original.LastIndexOf(',"ContentSha256":"',[StringComparison]::Ordinal))+'}'
        foreach($fault in @('Schema','Warehouse','Missing')) {
            if($fault -eq 'Schema'){WriteReadFixture ($body.Replace('"SchemaVersion":1','"SchemaVersion":99'))}
            elseif($fault -eq 'Warehouse'){WriteReadFixture ($body.Replace('"WarehouseId":"'+$Fixture.Warehouse+'"','"WarehouseId":"'+$OtherFixture.Warehouse+'"'))}
            else {Remove-Item -LiteralPath $path}
            if([bool](Run 'invSys.Core.xlam' 'modInventoryViewerData.PublishedReadFixtureValidForTest' @($path,$Fixture.Warehouse))) {throw 'Fault fixture was not rejected by the calibrated publication validator.'}
            [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OpenInventoryViewer')
            [void](ReadAct 'Events')
            Check ('PublishedRead.'+$fault+'ColdOpenUnavailable') ((ReadAct 'Unavailable') -and (ReadAct 'Empty'))
            [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.CloseInventoryViewerForTest')
            [IO.File]::WriteAllText($path,$original,$utf8)
        }
        # Change current visibility through the actual Admin controls after
        # publication. Reading must obey that policy without republishing history.
        foreach($visible in @($false,$true)) {
            SelectTarget $Fixture 'config-admin'
            try {
                [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
                if(-not [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.PublishedReadVisibilityForTest' @($visible))) {throw 'Actual Admin visibility handler did not save the fixture policy.'}
            } finally {[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')}
            # Only the intentional fixture policy change advances this pin.
            $pins[$Fixture.Config]=(Get-FileHash -LiteralPath $Fixture.Config).Hash
            SelectTarget $Fixture 'config-reader'
            [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OpenInventoryViewer')
            [void](ReadAct 'Events')
            $name=if($visible){'CurrentPolicyRestoresPublishedActivity'}else{'CurrentPolicyHidesPublishedAdminActivity'}
            Check ('PublishedRead.'+$name) ((ReadAct 'ContainsSource' $activityId) -eq $visible)
            Check ('PublishedRead.PolicyChangeDoesNotRepublish.'+$visible) ([IO.File]::ReadAllText($path) -ceq $original)
            [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.CloseInventoryViewerForTest')
        }
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OpenInventoryViewer')
        [void](ReadAct 'Events')
        SelectTarget $OtherFixture 'config-reader'
        [void](ReadAct 'Search' 'context changed')
        Check 'PublishedRead.TargetChangeClearsCapturedContent' (ReadAct 'Empty')
        [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
        [void](ReadAct 'Refresh')
        Check 'PublishedRead.SignOutCannotReloadProjection' (ReadAct 'Empty')
        Check 'PublishedRead.ViewerNeverPublishes' ([long](Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest') -eq $viewerPublishBefore)
    } finally {
        [IO.File]::WriteAllText($path,$original,[Text.UTF8Encoding]::new($false))
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.CloseInventoryViewerForTest')
    }
    $unchanged=$true
    foreach($pin in $pins.Keys){if((Get-FileHash -LiteralPath $pin).Hash -cne $pins[$pin]){$unchanged=$false}}
    $unchanged=$unchanged -and @((Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File)).Count -eq $pins.Count
    Check 'PublishedRead.AuthorityActivityAndRestoredProjectionBytesUnchanged' $unchanged
}
