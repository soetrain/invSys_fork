# D18: real Settings activity -> real Admin publication -> actual Viewer handlers.
# Only disposable published bytes are faulted. No runtime reader is replaced.
function Test-Slice4beViewerPublishedRead($Fixture,$OtherFixture) {
    SelectTarget $Fixture 'config-admin'
    try {
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
        if(-not [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('BatchSize','601'))) {throw 'Actual Settings activity fixture failed.'}
    } finally {[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')}
    $admin=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('modAdminConsole').CodeModule
    $admin.AddFromString(@'
Public Function PublishReadFixtureForTest() As Boolean
    Dim report As String
    PublishReadFixtureForTest = GenerateInventorySnapshot("config-admin", "", Nothing, "", Nothing, report)
End Function
'@)
    if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')) {throw 'Actual Admin publication fixture failed.'}
    $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')
    if(-not (Test-Path -LiteralPath $path)){throw 'Published-read fixture requires the independently tested publisher candidate.'}
    $original=[IO.File]::ReadAllText($path)
    $model=$original|ConvertFrom-Json
    $activity=@($model.Groups|Where-Object {$_.Source -ceq 'Activity' -and @($_.Lines|Where-Object ControlId -CEQ 'ADMIN_SETTINGS_SAVE_VALUE').Count -gt 0})
    if($activity.Count -ne 1 -or $activity[0].Lines.Count -ne 2 -or $activity[0].Outcomes.Count -ne 1){throw 'Actual publication lacks the correlated Settings activity fixture.'}
    $activityId=[string]$activity[0].SourceId
    $core=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modInventoryViewerData').CodeModule
    $core.AddFromString(@'
Public Function PublishedReadFixtureValidForTest(ByVal path As String, ByVal warehouse As String) As Boolean
    Dim model As Object
    Set model = modEventsPublicationStore.Read(path, warehouse)
    PublishedReadFixtureValidForTest = Not model Is Nothing
End Function
'@)
    $valid=[bool](Run 'invSys.Core.xlam' 'modInventoryViewerData.PublishedReadFixtureValidForTest' @($path,$Fixture.Warehouse))
    Check 'PublishedRead.RealPublicationFixtureValid' $valid
    if(-not $valid){throw 'Fixture schema or integrity is invalid; not a behavioral RED.'}
    Check 'PublishedRead.RealSettingsAttemptAndOutcomePublished' ($activity[0].Lines[0].RecordId -cne $activity[0].Lines[1].RecordId)

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
        Case "Search": mTxtSearch.Value = expected
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
    function ReadAct([string]$Action,[string]$Expected='') {[bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @($Action,$Expected))}
    function ReadDetail([string]$Caption,[string]$Expected) {[bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadDetailForTest' @($Caption,$Expected))}
    function WriteReadFixture([string]$Body) {
        $sha=[Security.Cryptography.SHA256]::Create()
        try{$hash=([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Body)))).Replace('-','')}finally{$sha.Dispose()}
        [IO.File]::WriteAllText($path,$Body.Substring(0,$Body.Length-1)+',"ContentSha256":"'+$hash+'"}',[Text.UTF8Encoding]::new($false))
    }
    $pins=@{}
    foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File){$pins[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash}
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
        [void](ReadAct 'Search' $activityId)
        Check 'PublishedRead.SearchRetainsExactActivityGroup' (ReadAct 'ContainsSource' $activityId)
        [void](ReadAct 'Search' '')
        Check 'PublishedRead.OpenRefreshSearchAvoidShippingAuthority' ([long](Run 'invSys.Operations.xlam' 'modTS_Shipments.PublishedReadAuthorityCallsForTest') -eq 0)

        $utf8=[Text.UTF8Encoding]::new($false)
        [IO.File]::WriteAllText($path,$original.Replace('"ContentSha256":"','"ContentSha256":"0'),$utf8)
        [void](ReadAct 'Refresh')
        Check 'PublishedRead.DamagedHashRefreshIsStale' (ReadAct 'Stale')
        Check 'PublishedRead.DamagedHashRetainsLoadedActivity' (ReadAct 'ContainsSource' $activityId)
        [IO.File]::WriteAllText($path,$original,$utf8)
        [void](ReadAct 'Refresh')
        Check 'PublishedRead.RestoredPublicationRecoversActivity' ((ReadAct 'Fresh') -and (ReadAct 'ContainsSource' $activityId))
        if($CaptureEvidence){CaptureFormEvidence '' 'viewer-published-read.png' ([long](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadWindowForTest'))}

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
        Check 'PublishedRead.ViewerNeverPublishes' ([long](Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest') -eq 0)
    } finally {
        [IO.File]::WriteAllText($path,$original,[Text.UTF8Encoding]::new($false))
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.CloseInventoryViewerForTest')
    }
    $unchanged=$true
    foreach($pin in $pins.Keys){if((Get-FileHash -LiteralPath $pin).Hash -cne $pins[$pin]){$unchanged=$false}}
    $unchanged=$unchanged -and @((Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File)).Count -eq $pins.Count
    Check 'PublishedRead.AuthorityActivityAndRestoredProjectionBytesUnchanged' $unchanged
}
