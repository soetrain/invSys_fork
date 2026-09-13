# D18 paged-group/read-boundary RED through real Viewer handlers. The volume
# fixture changes only an Admin-generated disposable published snapshot. It
# does not fabricate canonical inventory or claim publisher-bound acceptance.
function Test-Slice4beViewerEventGroups($Fixture,$OtherFixture) {
    SelectTarget $Fixture 'config-reader'
    $manager=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modInventoryViewer').CodeModule
    $manager.AddFromString(@'
Public Function PrepareViewerGroupsForTest(ByVal workbookName As String) As Boolean
    Dim wb As Workbook, ws As Worksheet, lo As ListObject, found As ListObject, column As ListColumn
    Dim headers As Variant, values() As Variant, item As Variant, i As Long, j As Long, group As Long, stamp As Date, key As String, unit As String
    Set wb = Application.Workbooks(workbookName)
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = "tblInventoryEvents" Then Set found = lo
        Next lo
    Next ws
    If found Is Nothing Then Exit Function
    Set column = found.ListColumns.Add(1): column.Name = "User group sentinel"
    If column.Name <> "User group sentinel" Then Exit Function
    found.ListColumns("System_Key").Name = "system_KEY"
    headers = Array("EventID", "EventType", "OccurredAtUTC", "AppliedAtUTC", "StationId", "UserId", "system_KEY", "SKU", "QtyDelta", "Location", "Condition", "Note", "User group sentinel")
    found.Resize found.HeaderRowRange.Resize(5004, found.ListColumns.Count)
    ReDim values(1 To 5003, 1 To found.ListColumns.Count)
    stamp = Now
    For i = 1 To 5003
        group = i: key = "SYS-GROUP-" & Format$(i, "00000"): unit = "EA"
        If i > 5001 Then group = 100
        If group = 100 Then
            key = "SYS-GROUP-A"
            If i = 5002 Then unit = "LB"
            If i = 5003 Then key = "SYS-GROUP-B"
        End If
        item = Array("EVT-GROUP-" & Format$(group, "00000"), "RECEIVE", CDbl(DateAdd("s", -group, stamp)), _
            CDbl(DateAdd("s", -group, stamp)), "S1", "group-fixture", key, "SKU-GROUP", 1, "DOCK", "GOOD", _
            "Reference=PAGE-GROUP-" & Format$(group, "00000") & ";Item=Group fixture;UOM=" & unit, "DO-NOT-DISPLAY")
        For j = 0 To UBound(headers)
            values(i, found.ListColumns(CStr(headers(j))).Index) = item(j)
        Next j
    Next i
    found.DataBodyRange.Value2 = values
    PrepareViewerGroupsForTest = (found.ListRows.Count = 5003)
End Function
Public Function ViewerGroupsActionForTest(ByVal action As String) As Boolean
    If Not mInventoryViewer Is Nothing Then ViewerGroupsActionForTest = mInventoryViewer.ViewerGroupsActionForTest(action)
End Function
Public Function ViewerGroupsDetailForTest() As Boolean
    Dim form As Object, list As Object, i As Long, firstKey As Long, secondKey As Long
    For Each form In VBA.UserForms
        If TypeName(form) = "frmEventDetail" Then
            Set list = form.Controls("lstEventLines")
            For i = 0 To list.ListCount - 1
                If CStr(list.List(i, 0)) = "SYS-GROUP-A" Then firstKey = firstKey + 1
                If CStr(list.List(i, 0)) = "SYS-GROUP-B" Then secondKey = secondKey + 1
            Next i
            ViewerGroupsDetailForTest = (list.ListCount = 3 And firstKey = 2 And secondKey = 1)
            Exit Function
        End If
    Next form
End Function
Public Sub CloseViewerGroupsForTest()
    If Not mInventoryViewer Is Nothing Then Unload mInventoryViewer
End Sub
Public Function ViewerGroupsWindowForTest() As Double
    mInventoryViewer.Repaint: DoEvents
    ViewerGroupsWindowForTest = CDbl(modUserFormResizeWin.GetUserFormWindowHandle(mInventoryViewer))
End Function
'@)
    $snapshot=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Inventory.xlsb')
    $book=$excel.Workbooks.Open($snapshot,0,$false)
    if(-not [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PrepareViewerGroupsForTest' @($book.Name))) {throw 'Published group fixture preparation failed.'}
    $book.Save();$book.Close($false)
    $pins=@{}
    foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Filter '*.xlsb') {$pins[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash}

    $shipping=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modTS_Shipments').CodeModule
    $entry=$shipping.ProcBodyLine('AppendBoxDesignViewerEvents',0)
    $shipping.InsertLines($entry+1,'    ViewerGroupsAuthorityReadsForTest = ViewerGroupsAuthorityReadsForTest + 1')
    $shipping.AddFromString(@'
Private ViewerGroupsAuthorityReadsForTest As Long
Public Function ViewerGroupsAuthorityCountForTest(Optional ByVal reset As Boolean = False) As Long
    ViewerGroupsAuthorityCountForTest = ViewerGroupsAuthorityReadsForTest
    If reset Then ViewerGroupsAuthorityReadsForTest = 0
End Function
'@)
    $form=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmInventoryViewer').CodeModule
    $form.AddFromString(@'
Public Function ViewerGroupsActionForTest(ByVal action As String) As Boolean
    Dim i As Long, label As String
    On Error GoTo Failed
    Select Case action
        Case "Events"
            mCboEventRange.Value = "All": mTabs.Value = 1: mBtnRefresh_Click
        Case "Refresh": mBtnRefresh_Click
        Case "Loaded"
            ViewerGroupsActionForTest = (mLstInventory.ListCount > 0 And Not IsEmpty(mRows)): Exit Function
        Case "FirstPage"
            If mLstInventory.ListCount <> 100 Then Exit Function
            ViewerGroupsActionForTest = (mLstInventory.List(0, 2) = "PAGE-GROUP-00001" And mLstInventory.List(99, 2) = "PAGE-GROUP-00100"): Exit Function
        Case "PageLabel"
            label = Me.Controls("lblEventPage").Caption
            ViewerGroupsActionForTest = (InStr(1, label, "Page 1 of", vbTextCompare) > 0 And InStr(1, label, "matching", vbTextCompare) > 0): Exit Function
        Case "PreviousDisabled"
            ViewerGroupsActionForTest = Not Me.Controls("btnEventsPrevious").Enabled: Exit Function
        Case "Next"
            Me.Controls("btnEventsNext").Value = True
            If mLstInventory.ListCount <> 100 Then Exit Function
            ViewerGroupsActionForTest = (mLstInventory.List(0, 2) = "PAGE-GROUP-00101" And mLstInventory.List(99, 2) = "PAGE-GROUP-00200"): Exit Function
        Case "Previous"
            Me.Controls("btnEventsPrevious").Value = True
            ViewerGroupsActionForTest = ViewerGroupsActionForTest("FirstPage"): Exit Function
        Case "SearchUnits"
            mTxtSearch.Value = "LB"
            ViewerGroupsActionForTest = (mLstInventory.ListCount = 1 And mLstInventory.List(0, 2) = "PAGE-GROUP-00100"): Exit Function
        Case "SummaryRetainsMultipleUnits"
            If mLstInventory.ListCount <> 1 Then Exit Function
            ViewerGroupsActionForTest = (Not IsNumeric(mLstInventory.List(0, 4)) And mLstInventory.List(0, 5) <> "LB" And mLstInventory.List(0, 5) <> "EA"): Exit Function
        Case "SelectBoundary"
            For i = 0 To mLstInventory.ListCount - 1
                If mLstInventory.List(i, 2) = "PAGE-GROUP-00100" Then mLstInventory.ListIndex = i: ViewerGroupsActionForTest = True: Exit Function
            Next i
            Exit Function
        Case "SearchOne"
            mTxtSearch.Value = "PAGE-GROUP-00001"
            ViewerGroupsActionForTest = (mLstInventory.ListCount = 1 And mLstInventory.List(0, 2) = "PAGE-GROUP-00001"): Exit Function
        Case "NextDisabled"
            ViewerGroupsActionForTest = Not Me.Controls("btnEventsNext").Enabled: Exit Function
        Case "ClearSearch": mTxtSearch.Value = ""
        Case "SignedOut"
            mTxtSearch.Value = "context check"
            ViewerGroupsActionForTest = (mLstInventory.ListCount = 0 And IsEmpty(mRows)): Exit Function
        Case Else: Exit Function
    End Select
    ViewerGroupsActionForTest = True
Failed:
End Function
'@)
    function GroupsAct([string]$Action) {[bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.ViewerGroupsActionForTest' @($Action))}
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OpenInventoryViewer')
    [void](GroupsAct 'Events')
    Check 'ViewerGroups.PublishedVolumeLoadedThroughEventsHandler' (GroupsAct 'Loaded')
    Check 'ViewerGroups.FirstPageHas100SortedCompleteGroups' (GroupsAct 'FirstPage')
    Check 'ViewerGroups.PageAndMatchingCountVisible' (GroupsAct 'PageLabel')
    Check 'ViewerGroups.PreviousDisabledAtStart' (GroupsAct 'PreviousDisabled')
    Check 'ViewerGroups.NextHandlerShowsNext100Groups' (GroupsAct 'Next')
    Check 'ViewerGroups.PreviousHandlerRestoresFirstPage' (GroupsAct 'Previous')
    [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ViewerGroupsAuthorityCountForTest' @($true))
    [void](GroupsAct 'Refresh')
    Check 'ViewerGroups.RefreshAvoidsShippingAuthorityRead' ([long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ViewerGroupsAuthorityCountForTest') -eq 0)
    Check 'ViewerGroups.SearchMatchesContributingUnit' (GroupsAct 'SearchUnits')
    Check 'ViewerGroups.GroupSummaryDoesNotChooseOrSumUnlikeUnits' (GroupsAct 'SummaryRetainsMultipleUnits')
    Check 'ViewerGroups.BoundaryGroupSelectedThroughListHandler' (GroupsAct 'SelectBoundary')
    Check 'ViewerGroups.FilteredBoundaryRetainsEveryExactKeyLine' ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.ViewerGroupsDetailForTest'))
    Check 'ViewerGroups.SearchFindsSingleGroup' (GroupsAct 'SearchOne')
    Check 'ViewerGroups.NextDisabledForSingleMatch' (GroupsAct 'NextDisabled')
    [void](GroupsAct 'ClearSearch')
    Check 'ViewerGroups.ClearingSearchStartsFirstPage' (GroupsAct 'FirstPage')
    if($CaptureEvidence) {
        $window=[long](Run 'invSys.Operations.xlam' 'modInventoryViewer.ViewerGroupsWindowForTest')
        CaptureFormEvidence '' 'viewer-groups.png' $window
    }
    [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
    Check 'ViewerGroups.SignOutInvalidatesLoadedContent' (GroupsAct 'SignedOut')
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.CloseViewerGroupsForTest')
    $unchanged=$true
    foreach($path in $pins.Keys){if((Get-FileHash -LiteralPath $path).Hash -cne $pins[$path]){$unchanged=$false}}
    Check 'ViewerGroups.GeneratedAuthorityAndUnknownColumnBytesUnchanged' $unchanged
}
