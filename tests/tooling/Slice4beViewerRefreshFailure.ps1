# D18 failure injection at the projection-read boundary. Actual Viewer launch,
# Refresh, Search and tab handlers remain intact. Synthetic rows never leave VBA.
function Test-Slice4beViewerRefreshFailure($Fixture,$OtherFixture) {
    $core = $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modInventoryViewerData').CodeModule
    $entry = $core.ProcBodyLine('LoadCurrentInventoryEventViewerData',0)
    $core.InsertLines($entry+1, @'
    If ViewerFailureEnabledForTest Then
        ViewerFailureReadsForTest = ViewerFailureReadsForTest + 1
        LoadCurrentInventoryEventViewerData = ViewerFailurePayloadForTest
        Exit Function
    End If
'@)
    $core.AddFromString(@'
Private ViewerFailureEnabledForTest As Boolean
Private ViewerFailurePayloadForTest As String
Private ViewerFailureReadsForTest As Long
Public Sub SetViewerFailureForTest(ByVal mode As Long)
    ViewerFailureEnabledForTest = (mode <> 0)
    ViewerFailureReadsForTest = 0
    ViewerFailurePayloadForTest = "FAIL" & vbTab & "Published Events are unavailable. Try Refresh after publication is restored."
    If mode = 1 Then
        ViewerFailurePayloadForTest = "OK" & vbTab & modNasConnection.GetCurrentTargetWarehouseId() & vbTab & "2026-09-13 12:00:00" & vbTab & "2" & vbCrLf & _
            "2026-09-13 11:00:00" & vbTab & "Receipt" & vbTab & "SYNTHETIC-A" & vbTab & "Synthetic item A" & vbTab & "1" & vbTab & "each" & vbTab & "DOCK" & vbTab & "GOOD" & vbTab & "fixture" & vbTab & "" & vbCrLf & _
            "2026-09-13 11:01:00" & vbTab & "Receipt" & vbTab & "SYNTHETIC-B" & vbTab & "Synthetic item B" & vbTab & "2" & vbTab & "each" & vbTab & "DOCK" & vbTab & "GOOD" & vbTab & "fixture" & vbTab & ""
    ElseIf mode = 3 Then
        ViewerFailurePayloadForTest = "OK" & vbTab & modNasConnection.GetCurrentTargetWarehouseId() & vbTab & "2026-09-13 12:00:00" & vbTab & "0"
    ElseIf mode = 4 Then
        ViewerFailurePayloadForTest = ""
    End If
End Sub
Public Function ViewerReadCountForTest() As Long
    ViewerReadCountForTest = ViewerFailureReadsForTest
End Function
'@)
    $shipping = $packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modTS_Shipments').CodeModule
    $entry = $shipping.ProcBodyLine('LoadShippingViewerSupplementEvents',0)
    $shipping.InsertLines($entry+1,'    ViewerSupplementReadsForTest = ViewerSupplementReadsForTest + 1')
    $shipping.AddFromString(@'
Private ViewerSupplementReadsForTest As Long
Public Function ViewerSupplementCountForTest(Optional ByVal reset As Boolean = False) As Long
    ViewerSupplementCountForTest = ViewerSupplementReadsForTest
    If reset Then ViewerSupplementReadsForTest = 0
End Function
'@)
    $form = $packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmInventoryViewer').CodeModule
    $form.AddFromString(@'
Private mRememberedEventsForTest As String
Private Function ViewerValuesForTest() As String
    Dim r As Long, c As Long, value As String
    ViewerValuesForTest = CStr(mLstInventory.ListCount) & "|" & CStr(mLstInventory.ListIndex) & "|" & CStr(mTxtSearch.Value)
    If IsEmpty(mRows) Then Exit Function
    For r = LBound(mRows, 1) To UBound(mRows, 1)
        For c = LBound(mRows, 2) To UBound(mRows, 2)
            value = CStr(mRows(r, c))
            ViewerValuesForTest = ViewerValuesForTest & "|" & CStr(Len(value)) & ":" & value
        Next c
    Next r
End Function
Public Function ViewerFailureActionForTest(ByVal action As String) As Boolean
    On Error GoTo Failed
    Select Case action
        Case "Events": mCboEventRange.Value = "All": mTabs.Value = 1
        Case "Remember"
            If mLstInventory.ListCount < 2 Then Exit Function
            mTxtSearch.Value = "Synthetic item": mLstInventory.ListIndex = 1
            mRememberedEventsForTest = ViewerValuesForTest()
        Case "Refresh": mBtnRefresh_Click
        Case "Preserved": ViewerFailureActionForTest = (mRememberedEventsForTest = ViewerValuesForTest()): Exit Function
        Case "Stale": ViewerFailureActionForTest = (InStr(1, mLblStatus.Caption, "Stale", vbTextCompare) > 0): Exit Function
        Case "Unavailable": ViewerFailureActionForTest = (InStr(1, mLblStatus.Caption, "Unavailable", vbTextCompare) > 0): Exit Function
        Case "Empty": ViewerFailureActionForTest = (mLstInventory.ListCount = 0 And IsEmpty(mRows)): Exit Function
        Case "Fresh": ViewerFailureActionForTest = (InStr(1, mLblStatus.Caption, "Stale", vbTextCompare) = 0 And InStr(1, mLblStatus.Caption, "Unavailable", vbTextCompare) = 0): Exit Function
        Case "Search": mTxtSearch.Value = "Synthetic item A"
        Case "Inventory": mTabs.Value = 0
        Case Else: Exit Function
    End Select
    ViewerFailureActionForTest = True
Failed:
End Function
'@)
    $manager = $packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modInventoryViewer').CodeModule
    $manager.AddFromString(@'
Public Function ViewerFailureActionForTest(ByVal action As String) As Boolean
    If Not mInventoryViewer Is Nothing Then ViewerFailureActionForTest = mInventoryViewer.ViewerFailureActionForTest(action)
End Function
Public Sub ViewerFailureCloseForTest()
    If Not mInventoryViewer Is Nothing Then Unload mInventoryViewer
End Sub
Public Function ViewerFailureWindowForTest() As Double
    mInventoryViewer.Repaint
    DoEvents
    ViewerFailureWindowForTest = CDbl(modUserFormResizeWin.GetUserFormWindowHandle(mInventoryViewer))
End Function
'@)
    SelectTarget $Fixture 'config-reader'
    $pins = @{}
    foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Filter '*.xlsb') { $pins[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash }
    function SetProjection([int]$Mode) { [void](Run 'invSys.Core.xlam' 'modInventoryViewerData.SetViewerFailureForTest' @($Mode)) }
    function Act([string]$Action) { [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.ViewerFailureActionForTest' @($Action)) }
    SetProjection 1
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OpenInventoryViewer')
    [void](Act 'Events')
    Check 'ViewerRefresh.PopulatedFixtureThroughActualEventsHandler' (Act 'Remember')
    SetProjection 2
    [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ViewerSupplementCountForTest' @($true))
    [void](Act 'Refresh')
    Check 'ViewerRefresh.FailedRefreshPreservesRowsSearchSelection' (Act 'Preserved')
    Check 'ViewerRefresh.FailedRefreshMarkedStale' (Act 'Stale')
    if($CaptureEvidence) {
        $window = [long](Run 'invSys.Operations.xlam' 'modInventoryViewer.ViewerFailureWindowForTest')
        CaptureFormEvidence '' 'viewer-stale-events.png' $window
    }
    Check 'ViewerRefresh.FailedProjectionDoesNotReadShippingSupplements' ([long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ViewerSupplementCountForTest') -eq 0)
    Check 'ViewerRefresh.OneProjectionReadPerRefresh' ([long](Run 'invSys.Core.xlam' 'modInventoryViewerData.ViewerReadCountForTest') -eq 1)
    SetProjection 4
    Check 'ViewerRefresh.EmptyFailureDoesNotRaise' (Act 'Refresh')
    Check 'ViewerRefresh.EmptyFailureReadOnce' ([long](Run 'invSys.Core.xlam' 'modInventoryViewerData.ViewerReadCountForTest') -eq 1)
    Check 'ViewerRefresh.EmptyFailureRetainsStaleContent' ((Act 'Preserved') -and (Act 'Stale'))
    [void](Act 'Search')
    Check 'ViewerRefresh.SearchRetainsStaleNotice' (Act 'Stale')
    SetProjection 1
    [void](Act 'Refresh')
    Check 'ViewerRefresh.SuccessClearsStaleNotice' (Act 'Fresh')
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.ViewerFailureCloseForTest')
    SetProjection 2
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OpenInventoryViewer')
    [void](Act 'Events')
    Check 'ViewerRefresh.FirstFailureUnavailable' (Act 'Unavailable')
    Check 'ViewerRefresh.FirstFailureDoesNotShowInventoryAsEvents' (Act 'Empty')
    SetProjection 3
    [void](Act 'Refresh')
    Check 'ViewerRefresh.ValidEmptyProjectionIsDistinctFromFailure' ((Act 'Empty') -and (Act 'Fresh'))
    SetProjection 1
    [void](Act 'Refresh')
    [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
    [void](Act 'Search')
    Check 'ViewerRefresh.SignOutInvalidatesLoadedContent' (Act 'Empty')
    SelectTarget $OtherFixture 'config-reader'
    [void](Act 'Refresh')
    Check 'ViewerRefresh.ChangedTargetCannotRefreshOldViewer' ((Act 'Empty') -and (Act 'Unavailable'))
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.ViewerFailureCloseForTest')
    SetProjection 0
    $unchanged = $true
    foreach($path in $pins.Keys) { if((Get-FileHash -LiteralPath $path).Hash -ne $pins[$path]) { $unchanged=$false } }
    Check 'ViewerRefresh.GeneratedAuthorityBytesUnchanged' $unchanged
}
