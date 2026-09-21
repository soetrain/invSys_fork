# D18 published-fixture selection through the packaged Viewer. No canonical
# inventory rows are fabricated: only the disposable published test projection
# receives synthetic edge cases after Admin Generate Warehouse created it.
function Test-Slice4beViewerEventDetail($Fixture,$OtherFixture) {
    SelectTarget $Fixture 'config-reader'
    $snapshot = Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Inventory.xlsb')
    if(-not (Test-Path -LiteralPath $snapshot)) { throw 'Admin-generated snapshot fixture is missing.' }
    $manager = $packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modInventoryViewer').CodeModule
    $manager.AddFromString(@'
Public Function PrepareDetailProjectionForTest(ByVal workbookName As String) As Boolean
    Dim wb As Workbook, ws As Worksheet, lo As ListObject, found As ListObject, column As ListColumn, record As ListRow
    Dim headers As Variant, values As Variant, keys As Variant, units As Variant, i As Long, j As Long, stamp As Date
    Set wb = Application.Workbooks(workbookName)
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = "tblInventoryEvents" Then Set found = lo
        Next lo
    Next ws
    If found Is Nothing Then Exit Function
    If Not found.DataBodyRange Is Nothing Then found.DataBodyRange.Delete
    Set column = found.ListColumns.Add(1)
    column.Name = "User detail sentinel"
    If column.Name <> "User detail sentinel" Then Exit Function
    found.ListColumns("System_Key").Name = "system_KEY"
    headers = Array("EventID", "EventType", "OccurredAtUTC", "AppliedAtUTC", "StationId", "UserId", "system_KEY", "SKU", "QtyDelta", "Location", "Condition", "Note", "User detail sentinel")
    keys = Array("SYS-DETAIL-A", "SYS-DETAIL-A", "SYS-DETAIL-B")
    units = Array("EA", "LB", "EA")
    stamp = DateAdd("n", -5, Now)
    For i = 0 To 2
        Set record = found.ListRows.Add
        values = Array("EVT-DETAIL-COMPLETE", "RECEIVE", CDbl(stamp), CDbl(stamp), "S1", "detail-fixture", keys(i), "SKU-DETAIL", i + 2, "DOCK", "GOOD", _
            "Reference=DETAIL-REF;Item=Detail fixture;UOM=" & units(i) & ";UnapprovedField=DO-NOT-DISPLAY", "DO-NOT-DISPLAY")
        For j = 0 To UBound(headers)
            record.Range.Cells(1, found.ListColumns(CStr(headers(j))).Index).Value2 = values(j)
        Next j
    Next i
    PrepareDetailProjectionForTest = (found.ListRows.Count = 3 And found.ListColumns("User detail sentinel").DataBodyRange.Cells(1, 1).Value2 = "DO-NOT-DISPLAY")
End Function
'@)
    $book = $excel.Workbooks.Open($snapshot,0,$false)
    if(-not [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PrepareDetailProjectionForTest' @($book.Name))) { throw 'Typed published fixture preparation failed.' }
    $book.Save(); $book.Close($false)
    . (Join-Path $PSScriptRoot 'Slice4bePublishedProjectionFixture.ps1')
    Publish-Slice4beProjectionFixture $Fixture $snapshot
    $pins = @{}
    foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Filter '*.xlsb') { $pins[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash }
    $form = $packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmInventoryViewer').CodeModule
    $selectBody = @'
Public Function SelectDetailFixtureForTest() As Boolean
    Dim index As Long
    mCboEventRange.Value = "All"
    mTabs.Value = 1
    mBtnRefresh_Click
    For index = 0 To mLstInventory.ListCount - 1
        If CStr(mLstInventory.List(index, 2)) = "DETAIL-REF" Then
            mLstInventory.ListIndex = index
            SelectDetailFixtureForTest = True
            Exit Function
        End If
    Next index
End Function
'@
    $form.AddFromString($selectBody)
    $form.AddFromString(@'
Public Function FilterDetailFixtureForTest() As Boolean
    mTxtSearch.Value = "LB"
    If mLstInventory.ListCount <> 1 Then Exit Function
    mLstInventory.ListIndex = 0
    FilterDetailFixtureForTest = True
End Function
Public Sub RefreshDetailFixtureForTest()
    mBtnRefresh_Click
End Sub
Public Function SelectStateFixtureForTest() As Boolean
    Dim i As Long
    mTxtSearch.Value = "": mCboEventRange.Value = "All"
    mBtnRefresh_Click
    For i = 0 To mLstInventory.ListCount - 1
        If mLstInventory.List(i, 2) = "DETAIL-STATE" Then mLstInventory.ListIndex = i: SelectStateFixtureForTest = True: Exit Function
    Next i
End Function
Public Function DetailListSafeForTest() As Boolean
    Dim i As Long, j As Long
    For i = 0 To mLstInventory.ListCount - 1
        For j = 0 To mLstInventory.ColumnCount - 1
            If InStr(1, CStr(mLstInventory.List(i, j)), "DO-NOT-DISPLAY", vbBinaryCompare) > 0 Then Exit Function
        Next j
    Next i
    DetailListSafeForTest = True
End Function
'@)
    foreach($hook in @(@('modEventDetailSettings','ReadViewer','DetailProfileReadCountForTest'),@('modInventoryViewerData','LoadCurrentInventoryEventViewerData','DetailProjectionReadCountForTest'))) {
        $code=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item($hook[0]).CodeModule
        $entry=$code.ProcBodyLine($hook[1],0)
        $code.InsertLines($entry+1,('    '+$hook[2]+' = '+$hook[2]+' + 1'))
        $code.AddFromString(('Private '+$hook[2]+" As Long`nPublic Function DetailReadCounterForTest(Optional ByVal reset As Boolean = False) As Long`n    DetailReadCounterForTest = "+$hook[2]+"`n    If reset Then "+$hook[2]+" = 0`nEnd Function"))
    }
    $code=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modInventoryViewerData').CodeModule
    $entry=$code.ProcBodyLine('LoadCurrentInventoryEventViewerData',0)
    $code.InsertLines($entry+1,@'
    If DetailStateTypeForTest <> "" Then
        LoadCurrentInventoryEventViewerData = DetailStatePayloadForTest()
        Exit Function
    End If
'@)
    $code.AddFromString(@'
Private DetailStateTypeForTest As String
Public Sub SetDetailStateTypeForTest(ByVal value As String)
    DetailStateTypeForTest = value
End Sub
Private Function DetailStatePayloadForTest() As String
    Dim row As String
    row = "2026-09-13 10:00:00" & vbTab & DetailStateTypeForTest & vbTab & "DETAIL-STATE" & vbTab & "State fixture" & vbTab & "1" & vbTab & "EA" & vbTab & "DOCK" & vbTab & "" & vbTab & "" & vbTab & ""
    DetailStatePayloadForTest = "OK" & vbTab & modNasConnection.GetCurrentTargetWarehouseId() & vbTab & "2026-09-13 10:00:00" & vbTab & "2" & vbCrLf & row & vbCrLf & Replace$(row, "DETAIL-STATE", "OTHER-STATE")
End Function
'@)
    $manager = $packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modInventoryViewer').CodeModule
    $manager.AddFromString(@'
Public Function SelectDetailFixtureForTest() As Boolean
    SelectDetailFixtureForTest = mInventoryViewer.SelectDetailFixtureForTest()
End Function
Private Function DetailFixtureFormForTest() As Object
    Dim instance As Object
    For Each instance In VBA.UserForms
        If TypeName(instance) = "frmEventDetail" Then Set DetailFixtureFormForTest = instance: Exit Function
    Next instance
End Function
Public Function DetailFixtureFactForTest(ByVal fact As String) As Boolean
    Dim instance As Object, fields As Object, lines As Object, r As Long, a As Long, b As Long, source As Boolean, required As Long
    On Error GoTo Failed
    Set instance = DetailFixtureFormForTest()
    If instance Is Nothing Then Exit Function
    If fact = "Visible" Then DetailFixtureFactForTest = instance.Visible: Exit Function
    Set fields = instance.Controls("lstEventFields")
    Set lines = instance.Controls("lstEventLines")
    Select Case fact
        Case "SourceIdentity"
            For r = 0 To fields.ListCount - 1
                If fields.List(r, 0) = "Source event / activity ID" And fields.List(r, 1) = "EVT-DETAIL-COMPLETE" Then source = True
            Next r
            DetailFixtureFactForTest = source
        Case "EveryKey"
            For r = 0 To lines.ListCount - 1
                If lines.List(r, 0) = "SYS-DETAIL-A" Then a = a + 1
                If lines.List(r, 0) = "SYS-DETAIL-B" Then b = b + 1
            Next r
            DetailFixtureFactForTest = (lines.ListCount = 3 And a = 2 And b = 1)
        Case "SafeFields"
            For r = 0 To fields.ListCount - 1
                If InStr(1, fields.List(r, 0) & " " & fields.List(r, 1), "DO-NOT-DISPLAY", vbBinaryCompare) > 0 Then Exit Function
            Next r
            DetailFixtureFactForTest = (fields.ListCount >= 19)
        Case "UnknownZone"
            For r = 0 To fields.ListCount - 1
                If fields.List(r, 0) = "Time provenance" And InStr(1, fields.List(r, 1), "zone unavailable", vbTextCompare) > 0 Then DetailFixtureFactForTest = True
            Next r
    End Select
Failed:
End Function
Public Sub CloseDetailFixtureForTest()
    If Not mInventoryViewer Is Nothing Then Unload mInventoryViewer
End Sub
Public Function DetailLineValuesForTest() As Boolean
    Dim instance As Object, lines As Object, fields As Object, i As Long, j As Long, key As String, qty As String, unit As String, seen As Object
    On Error GoTo Failed
    Set instance = DetailFixtureFormForTest()
    If instance Is Nothing Then Exit Function
    Set lines = instance.Controls("lstEventLines"): Set fields = instance.Controls("lstEventFields")
    Set seen = CreateObject("Scripting.Dictionary")
    For i = 0 To lines.ListCount - 1
        lines.ListIndex = i: key = "": qty = "": unit = ""
        For j = 0 To fields.ListCount - 1
            Select Case fields.List(j, 0)
                Case "Inventory identity (System_Key)": key = fields.List(j, 1)
                Case "Quantity": qty = fields.List(j, 1)
                Case "Unit of measure": unit = fields.List(j, 1)
            End Select
        Next j
        seen.Add key & "|" & qty & "|" & unit, True
    Next i
    DetailLineValuesForTest = (seen.Count = 3 And seen.Exists("SYS-DETAIL-A|2|EA") And seen.Exists("SYS-DETAIL-A|3|LB") And seen.Exists("SYS-DETAIL-B|4|EA"))
Failed:
End Function
Public Function DetailProfileStateForTest(ByVal expectedVersion As Long, ByVal quantityVisible As Boolean, Optional ByVal reordered As Boolean = False) As Boolean
    Dim instance As Object, fields As Object, i As Long, quantity As Boolean, itemCode As Long, itemName As Long
    On Error GoTo Failed
    Set instance = DetailFixtureFormForTest()
    If instance Is Nothing Then Exit Function
    If InStr(1, instance.Controls("lblDetailProfile").Caption, "Detail profile " & CStr(expectedVersion) & ".", vbBinaryCompare) <> 1 Then Exit Function
    Set fields = instance.Controls("lstEventFields")
    itemCode = -1: itemName = -1
    For i = 0 To fields.ListCount - 1
        Select Case fields.List(i, 0)
            Case "Quantity": quantity = True
            Case "Item code": itemCode = i
            Case "Item name": itemName = i
        End Select
    Next i
    DetailProfileStateForTest = (quantity = quantityVisible)
    If reordered Then DetailProfileStateForTest = DetailProfileStateForTest And itemName >= 0 And itemCode > itemName
Failed:
End Function
Public Function SaveDetailFixtureProfileForTest() As Boolean
    Dim request As String, report As String
    request = modEventDetailSettings.DefaultRequest()
    request = modEventDetailSettings.StageField(request, "Receiving", "ItemName", True, -1)
    request = modEventDetailSettings.StageField(request, "Receiving", "Quantity", False)
    SaveDetailFixtureProfileForTest = modEventDetailSettings.SaveProfile(modActivity.CaptureContext(), 0, request, report)
End Function
Public Function DetailActionForTest(ByVal action As String) As Boolean
    Dim instance As Object, control As Object, lines As Object
    On Error GoTo Failed
    Set instance = DetailFixtureFormForTest()
    Select Case action
        Case "Filter": DetailActionForTest = mInventoryViewer.FilterDetailFixtureForTest(): Exit Function
        Case "ListSafe": DetailActionForTest = mInventoryViewer.DetailListSafeForTest(): Exit Function
        Case "Refresh": mInventoryViewer.RefreshDetailFixtureForTest
        Case "Closed": DetailActionForTest = (instance Is Nothing): Exit Function
        Case "Invalidated"
            If instance Is Nothing Then Exit Function
            Set lines = instance.Controls("lstEventLines")
            If lines.ListCount > 1 Then lines.ListIndex = 1
            DetailActionForTest = (lines.ListCount = 0 And instance.Controls("lstEventFields").ListCount = 0 And InStr(1, instance.Controls("lblDetailStatus").Caption, "Unavailable", vbTextCompare) > 0): Exit Function
        Case "ResizeAfterSignOut", "ResizeAgainAfterSignOut"
            If instance Is Nothing Then Exit Function
            If action = "ResizeAfterSignOut" Then instance.Width = 960 Else instance.Width = 820
            instance.Repaint
            DoEvents
            DetailActionForTest = (instance.Controls("lstEventLines").ListCount = 0 And instance.Controls("lstEventFields").ListCount = 0 And InStr(1, instance.Controls("lblDetailStatus").Caption, "Unavailable", vbTextCompare) > 0): Exit Function
        Case "FitDefault", "FitLarger", "FitCurrent"
            If instance Is Nothing Then Exit Function
            If action = "FitLarger" Then instance.Width = 960: instance.Height = 720
            If action = "FitDefault" Then instance.Width = 820: instance.Height = 640
            instance.Repaint
            DoEvents
            For Each control In instance.Controls
                If control.Visible Then
                    If control.Left < 0 Or control.Top < 0 Or control.Left + control.Width > instance.InsideWidth + 1 Or control.Top + control.Height > instance.InsideHeight + 1 Then Exit Function
                End If
            Next control
        Case Else: Exit Function
    End Select
    DetailActionForTest = True
Failed:
End Function
Public Function DetailWindowForTest() As Double
    Dim instance As Object
    Set instance = DetailFixtureFormForTest()
    instance.Repaint
    DoEvents
    DetailWindowForTest = CDbl(modUserFormResizeWin.GetUserFormWindowHandle(instance))
End Function
Public Function DetailGeometryForTest() As String
    Dim instance As Object, control As Object
    Set instance = DetailFixtureFormForTest()
    For Each control In instance.Controls
        DetailGeometryForTest = DetailGeometryForTest & control.Name & vbTab & CStr(control.Left) & vbTab & CStr(control.Top) & vbTab & CStr(control.Width) & vbTab & CStr(control.Height) & vbTab & CStr(instance.InsideWidth) & vbTab & CStr(instance.InsideHeight) & vbLf
    Next control
End Function
Public Function DetailCompleteTextForTest() As Boolean
    Dim instance As Object, fields As Object, measure As Object, widths As Variant
    Dim row As Long, column As Long, complete As Boolean, capacity As Single
    Set instance = DetailFixtureFormForTest()
    If instance Is Nothing Then Err.Raise 5, , "Detail fixture was not opened."
    Set fields = instance.Controls("lstEventFields")
    If fields.ListCount = 0 Or fields.ColumnCount <> 2 Then Err.Raise 5, , "Detail fixture has no field rows."
    widths = Split(fields.ColumnWidths, ";")
    If UBound(widths) <> 1 Then Err.Raise 5, , "Unexpected field-column geometry."
    Set measure = instance.Controls.Add("Forms.Label.1", "DetailTextMeasureForTest", False)
    measure.Font.Name = fields.Font.Name: measure.Font.Size = fields.Font.Size
    measure.Font.Bold = fields.Font.Bold: measure.Font.Italic = fields.Font.Italic
    measure.WordWrap = False: measure.AutoSize = True
    complete = fields.Locked
    For row = 0 To fields.ListCount - 1
        For column = 0 To 1
            measure.Caption = CStr(fields.List(row, column))
            capacity = Val(widths(column))
            If column = 1 And fields.Width - Val(widths(0)) - 20 > capacity Then capacity = fields.Width - Val(widths(0)) - 20
            If measure.Width > capacity Then complete = False
        Next column
    Next row
    instance.Controls.Remove "DetailTextMeasureForTest"
    DetailCompleteTextForTest = complete
End Function
Public Function DetailOverflowPointForTest() As String
    Dim instance As Object, fields As Object, widths As Variant
    Set instance = DetailFixtureFormForTest()
    Set fields = instance.Controls("lstEventFields")
    widths = Split(fields.ColumnWidths, ";")
    If Val(widths(0)) + Val(widths(1)) <= fields.Width Then Exit Function
    DetailOverflowPointForTest = CStr(fields.Left + fields.Width - 36) & vbTab & _
        CStr(fields.Top + fields.Height - 6) & vbTab & CStr(instance.InsideWidth) & vbTab & CStr(instance.InsideHeight)
End Function
Public Function SelectStateFixtureForTest() As Boolean
    SelectStateFixtureForTest = mInventoryViewer.SelectStateFixtureForTest()
End Function
Public Function DetailStateFactForTest(ByVal expectedFamily As String, ByVal checkClassification As Boolean) As Boolean
    Dim instance As Object, fields As Object, i As Long, classification As String, family As String, id As String
    On Error GoTo Failed
    Set instance = DetailFixtureFormForTest()
    If instance Is Nothing Then Exit Function
    Set fields = instance.Controls("lstEventFields")
    For i = 0 To fields.ListCount - 1
        Select Case fields.List(i, 0)
            Case "Source classification": classification = fields.List(i, 1)
            Case "Event family": family = fields.List(i, 1)
            Case "Source event / activity ID": id = fields.List(i, 1)
        End Select
    Next i
    If checkClassification Then
        DetailStateFactForTest = (classification = "Current state" And family = expectedFamily)
    Else
        DetailStateFactForTest = (id = "Unavailable" And instance.Controls("lstEventLines").ListCount = 1 And instance.Controls("lstEventLines").List(0, 0) = "Unavailable")
    End If
Failed:
End Function
'@)
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OpenInventoryViewer')
    Check 'EventDetail.PublishedFixtureSelectedThroughViewer' ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.SelectDetailFixtureForTest'))
    foreach($fact in @('Visible','SourceIdentity','EveryKey','SafeFields','UnknownZone')) {
        Check ('EventDetail.'+$fact) ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailFixtureFactForTest' @($fact)))
    }
    Check 'EventDetail.DefaultProfileForNonAdmin' ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailProfileStateForTest' @(0,$true)))
    foreach($module in @('modEventDetailSettings','modInventoryViewerData')) { [void](Run 'invSys.Core.xlam' ($module+'.DetailReadCounterForTest') @($true)) }
    Check 'EventDetail.UnlikeUnitsAndRepeatedLineValuesRetained' ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailLineValuesForTest'))
    foreach($module in @('modEventDetailSettings','modInventoryViewerData')) { Check ('EventDetail.SelectionDoesNotRead.'+$module) ([long](Run 'invSys.Core.xlam' ($module+'.DetailReadCounterForTest')) -eq 0) }
    foreach($action in @('ListSafe','Filter','FitDefault','FitLarger','FitDefault')) {
        $suffix=if($action -eq 'FitDefault' -and $results.Check -contains 'EventDetail.Action.FitDefault'){'.Restored'}else{''}
        Check ('EventDetail.Action.'+$action+$suffix) ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailActionForTest' @($action)))
        if($action -like 'Fit*') {
            $complete=Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailCompleteTextForTest'
            if($complete -isnot [bool]){throw 'Complete-text geometry probe did not return a Boolean.'}
            Check ('EventDetail.CompleteText.'+$action+$suffix) $complete
        }
    }
    [string](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailGeometryForTest') | Set-Content (Join-Path $reportRoot 'restored-geometry.tsv')
    if(-not ('DetailNativeLayout' -as [type])) {
        Add-Type @'
using System; using System.Runtime.InteropServices;
public static class DetailNativeLayout {
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int command);
    [DllImport("user32.dll")] public static extern bool IsZoomed(IntPtr h);
    [StructLayout(LayoutKind.Sequential)] public struct Rect { public int Left,Top,Right,Bottom; }
    [StructLayout(LayoutKind.Sequential)] public struct Point { public int X,Y; }
    [StructLayout(LayoutKind.Sequential)] public struct Mouse { public int X,Y; public uint Data,Flags,Time; public UIntPtr Extra; }
    [StructLayout(LayoutKind.Sequential)] public struct Input { public uint Type; public Mouse Mouse; }
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern IntPtr GetAncestor(IntPtr h,uint flags);
    [DllImport("user32.dll")] static extern bool GetClientRect(IntPtr h,out Rect r);
    [DllImport("user32.dll")] static extern bool ClientToScreen(IntPtr h,ref Point p);
    [DllImport("user32.dll")] static extern IntPtr WindowFromPoint(Point p);
    [DllImport("user32.dll")] static extern bool SetCursorPos(int x,int y);
    [DllImport("user32.dll")] static extern uint SendInput(uint n,Input[] inputs,int size);
    static uint Owner(IntPtr h) {uint p;GetWindowThreadProcessId(h,out p);return p;}
    public static void ScrollClick(IntPtr form,IntPtr excel,double x,double y,double width,double height) {
        Rect r;
        if(Owner(form)==0 || Owner(form)!=Owner(excel) || GetForegroundWindow()!=form || !GetClientRect(form,out r))
            throw new Exception("Detail scroll requires the foreground owned form.");
        var point=new Point {X=(int)(x*(r.Right-r.Left)/width),Y=(int)(y*(r.Bottom-r.Top)/height)};
        if(!ClientToScreen(form,ref point))throw new Exception("Detail coordinates unavailable.");
        var target=WindowFromPoint(point);
        if(Owner(target)!=Owner(form) || GetAncestor(target,2)!=form)
            throw new Exception("Detail scroll point is outside the owned form.");
        if(!SetCursorPos(point.X,point.Y))throw new Exception("Detail cursor positioning failed.");
        Input[] inputs={new Input {Mouse=new Mouse {Flags=2}},new Input {Mouse=new Mouse {Flags=4}}};
        if(SendInput(2,inputs,Marshal.SizeOf(typeof(Input)))!=2)
            throw new Exception("Detail scroll input delivery failed.");
    }
}
'@
    }
    $window=[long](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailWindowForTest')
    foreach($command in @(3,9)) {
        [void][DetailNativeLayout]::ShowWindow([IntPtr]$window,$command)
        Start-Sleep -Milliseconds 200
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailWindowForTest')
        $name=if($command -eq 3){'Maximize'}else{'Restore'}
        Check ('EventDetail.Native'+$name) ([DetailNativeLayout]::IsZoomed([IntPtr]$window) -eq ($command -eq 3) -and [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailActionForTest' @('FitCurrent')))
        $complete=Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailCompleteTextForTest'
        if($complete -isnot [bool]){throw 'Native complete-text geometry probe did not return a Boolean.'}
        Check ('EventDetail.CompleteText.Native'+$name) $complete
        if($CaptureEvidence) { CaptureFormEvidence '' ('event-detail-'+$name.ToLowerInvariant()+'.png') $window }
    }
    foreach($module in @('modEventDetailSettings','modInventoryViewerData')) {
        $readCount=Run 'invSys.Core.xlam' ($module+'.DetailReadCounterForTest')
        if($readCount -isnot [int]){throw 'Detail read counter did not return an integer.'}
        Check ('EventDetail.LayoutDoesNotRead.'+$module) ($readCount -eq 0)
    }
    Check 'EventDetail.FilterDoesNotDropContributingLines' ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailFixtureFactForTest' @('EveryKey')))
    if($CaptureEvidence) {
        $window=[long](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailWindowForTest')
        CaptureFormEvidence '' 'event-detail-default.png' $window
        $point=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailOverflowPointForTest')
        if($point -ne '') {
            $parts=$point -split "`t"
            if($parts.Count -ne 4){throw 'Unexpected detail scroll geometry.'}
            for($scroll=0;$scroll -lt 3;$scroll++){
                [DetailNativeLayout]::ScrollClick([IntPtr]$window,[IntPtr]$excel.Hwnd,[double]$parts[0],[double]$parts[1],[double]$parts[2],[double]$parts[3])
                [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailWindowForTest')
            }
            CaptureFormEvidence '' 'event-detail-horizontal-scroll.png' $window
        }
    }
    foreach($state in @(@('BOX_DESIGNED','Boxing'),@('SHIP_HELD','Shipping'))) {
        [void](Run 'invSys.Core.xlam' 'modInventoryViewerData.SetDetailStateTypeForTest' @($state[0]))
        Check ('EventDetail.StateSelected.'+$state[0]) ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.SelectStateFixtureForTest'))
        Check ('EventDetail.CurrentStateClassification.'+$state[0]) ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailStateFactForTest' @($state[1],$true)))
        Check ('EventDetail.NoFabricatedStateIdentityOrBlankGrouping.'+$state[0]) ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailStateFactForTest' @($state[1],$false)))
    }
    [void](Run 'invSys.Core.xlam' 'modInventoryViewerData.SetDetailStateTypeForTest' @(''))
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.CloseDetailFixtureForTest')
    Check 'EventDetail.ParentCloseReleasesDetail' ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailActionForTest' @('Closed')))
    SelectTarget $Fixture
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OpenInventoryViewer')
    Check 'EventDetail.AdminProjectionSelected' ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.SelectDetailFixtureForTest'))
    Check 'EventDetail.ProfileFixtureSavedByAuthorizedCoreCommand' ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.SaveDetailFixtureProfileForTest'))
    $pins[$Fixture.Config]=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    Check 'EventDetail.ProfileChangeWaitsForExplicitRefresh' ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailProfileStateForTest' @(0,$true)))
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailActionForTest' @('Refresh'))
    Check 'EventDetail.RefreshAppliesSavedVisibilityOrderAndVersion' ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailProfileStateForTest' @(1,$false,$true)))
    [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
    Check 'EventDetail.SignOutResizeInvalidatesContent' ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailActionForTest' @('ResizeAfterSignOut')))
    Check 'EventDetail.RepeatedInvalidatedResizeRemainsSafe' ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailActionForTest' @('ResizeAgainAfterSignOut')))
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.CloseDetailFixtureForTest')
    SelectTarget $Fixture
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.OpenInventoryViewer')
    Check 'EventDetail.ReopenedForSignOutLineCheck' ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.SelectDetailFixtureForTest'))
    [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
    Check 'EventDetail.SignedOutLineSelectionInvalidatesContent' ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailActionForTest' @('Invalidated')))
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.CloseDetailFixtureForTest')
    $unchanged = $true
    foreach($path in $pins.Keys) { if((Get-FileHash -LiteralPath $path).Hash -ne $pins[$path]) { $unchanged=$false } }
    Check 'EventDetail.GeneratedAuthorityAndUnknownColumnBytesUnchanged' $unchanged
}
