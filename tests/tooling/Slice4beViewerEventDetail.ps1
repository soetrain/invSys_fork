# D18 published-fixture selection through the packaged Viewer. No canonical
# inventory rows are fabricated: only the disposable published test projection
# receives synthetic edge cases after Admin Generate Warehouse created it.
function Capture-DetailOwnedEvidence([string]$FileName,[long]$WindowHandle) {
    Initialize-SettingsCapture
    $owned=[InvSysSettingsCapture]::OwnedVisibleForm('Event Detail',[IntPtr]$excel.Hwnd).ToInt64()
    if($owned -ne $WindowHandle){throw 'Detail input does not identify the unique owned form.'}
    $clicked=[InvSysSettingsCapture]::ActivateByCaptionClick([IntPtr]$WindowHandle,[IntPtr]$excel.Hwnd)
    [pscustomobject]@{Image=$FileName;OwnedCaptionClick=$clicked}|
        ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'detail-capture-activation.jsonl')
    if($clicked){Start-Sleep -Milliseconds 200}
    CaptureOwnedFormEvidence 'Event Detail' $FileName $WindowHandle
}

function Test-Slice4beViewerEventDetail($Fixture,$OtherFixture) {
    if($DetailScrollLockDiagnostic -and -not $CaptureEvidence){throw 'The lock diagnostic requires owned visible captures.'}
    $detailHost=$null
    try {
    if($CaptureEvidence){
        $detailHost=$excel.Workbooks.Add()
        if($null -eq $detailHost){throw 'Disposable detail capture workbook is unavailable.'}
        $excel.Visible=$true
        $visible=$excel.Visible
        [pscustomobject]@{VisibleType=if($null -eq $visible){'Null'}else{$visible.GetType().FullName};Visible=$visible;DisposableWorkbook=$true}|
            ConvertTo-Json|Set-Content (Join-Path $reportRoot 'detail-capture-host.json')
        if($visible -isnot [bool] -or -not $visible){throw 'Visible detail capture host is unavailable.'}
    }
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
Public Function FocusDetailFieldsForTest() As Boolean
    Dim instance As Object, fields As Object
    Set instance = DetailFixtureFormForTest()
    Set fields = instance.Controls("lstEventFields")
    fields.SetFocus
    FocusDetailFieldsForTest = (instance.ActiveControl Is fields)
End Function
Public Function DetailFieldLockForTest(ByVal change As Boolean, ByVal locked As Boolean) As Boolean
    Dim instance As Object, fields As Object
    Set instance = DetailFixtureFormForTest()
    If instance Is Nothing Then Err.Raise 5, , "Detail fixture is unavailable."
    Set fields = instance.Controls("lstEventFields")
    If change Then fields.Locked = locked
    DetailFieldLockForTest = fields.Locked
End Function
Public Function DetailFieldValuesForTest() As String
    Dim instance As Object, fields As Object, row As Long, column As Long, value As String
    Set instance = DetailFixtureFormForTest()
    If instance Is Nothing Then Err.Raise 5, , "Detail fixture is unavailable."
    Set fields = instance.Controls("lstEventFields")
    For row = 0 To fields.ListCount - 1
        For column = 0 To fields.ColumnCount - 1
            value = CStr(fields.List(row, column))
            DetailFieldValuesForTest = DetailFieldValuesForTest & CStr(Len(value)) & ":" & value
        Next column
    Next row
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
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h,out Rect r);
    [DllImport("user32.dll")] static extern bool ClientToScreen(IntPtr h,ref Point p);
    [DllImport("user32.dll")] static extern uint GetDpiForWindow(IntPtr h);
    [DllImport("user32.dll")] static extern int GetSystemMetricsForDpi(int index,uint dpi);
    [DllImport("user32.dll")] static extern IntPtr WindowFromPoint(Point p);
    [DllImport("user32.dll")] static extern bool SetCursorPos(int x,int y);
    [DllImport("user32.dll")] static extern uint SendInput(uint n,Input[] inputs,int size);
    static uint Owner(IntPtr h) {uint p;GetWindowThreadProcessId(h,out p);return p;}
    public static string ScrollFieldBounds(IntPtr form,IntPtr excel,double[] bounds) {
        Rect client,whole;var origin=new Point();
        if(Owner(form)==0 || Owner(form)!=Owner(excel) || GetForegroundWindow()!=form ||
            !GetClientRect(form,out client) || !GetWindowRect(form,out whole) || !ClientToScreen(form,ref origin) || bounds.Length!=6 || bounds[4]<=0 || bounds[5]<=0)
            throw new Exception("Owned detail client geometry is unavailable.");
        double sx=client.Right/bounds[4],sy=client.Bottom/bounds[5];
        var rect=new Rect {Left=origin.X+(int)Math.Round(bounds[0]*sx),Top=origin.Y+(int)Math.Round(bounds[1]*sy),
            Right=origin.X+(int)Math.Round((bounds[0]+bounds[2])*sx),Bottom=origin.Y+(int)Math.Round((bounds[1]+bounds[3])*sy)};
        int thickness=GetSystemMetricsForDpi(3,GetDpiForWindow(form));
        if(thickness<=0 || rect.Right-rect.Left<4*thickness || rect.Bottom-rect.Top<4*thickness ||
            rect.Left<origin.X || rect.Top<origin.Y || rect.Right>origin.X+client.Right || rect.Bottom>origin.Y+client.Bottom)
            throw new Exception("Detail list scrollbar geometry is unavailable.");
        var point=new Point {X=rect.Right-2*thickness,Y=rect.Bottom-thickness/2-1};
        if(GetAncestor(WindowFromPoint(point),2)!=form)throw new Exception("Detail scroll point is outside the owned form.");
        if(!SetCursorPos(point.X,point.Y))throw new Exception("Detail scroll cursor positioning failed.");
        if(GetForegroundWindow()!=form || GetAncestor(WindowFromPoint(point),2)!=form)throw new Exception("Detail scroll point became covered.");
        Input[] down={new Input {Mouse=new Mouse {Flags=2}}},up={new Input {Mouse=new Mouse {Flags=4}}};
        try {
            if(SendInput(1,down,Marshal.SizeOf(typeof(Input)))!=1)throw new Exception("Detail scroll press delivery failed.");
            System.Threading.Thread.Sleep(150);
        } finally {
            if(SendInput(1,up,Marshal.SizeOf(typeof(Input)))!=1)throw new Exception("Detail scroll release delivery failed.");
        }
        return (rect.Right-rect.Left)+"|"+(rect.Bottom-rect.Top)+"|"+thickness+"|"+(point.X-rect.Left)+"|"+(point.Y-rect.Top)+
            "|screen="+point.X+","+point.Y+"|list="+rect.Left+","+rect.Top+","+rect.Right+","+rect.Bottom+
            "|form="+whole.Left+","+whole.Top+","+whole.Right+","+whole.Bottom+"|dpi="+GetDpiForWindow(form);
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
        if($CaptureEvidence) { Capture-DetailOwnedEvidence ('event-detail-'+$name.ToLowerInvariant()+'.png') $window }
    }
    Check 'EventDetail.FilterDoesNotDropContributingLines' ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailFixtureFactForTest' @('EveryKey')))
    if($CaptureEvidence) {
        $window=[long](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailWindowForTest')
        Capture-DetailOwnedEvidence 'event-detail-default.png' $window
        $focused=Run 'invSys.Operations.xlam' 'modInventoryViewer.FocusDetailFieldsForTest'
        if($focused -isnot [bool] -or -not $focused){throw 'Actual detail list focus is unavailable.'}
        $geometry=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailGeometryForTest')
        $fieldRows=@($geometry -split "`n" | Where-Object {$_ -like "lstEventFields`t*"})
        if($fieldRows.Count -ne 1){throw 'Actual detail field bounds are ambiguous.'}
        [double[]]$fieldBounds=@(($fieldRows[0] -split "`t")[1..6] | ForEach-Object {[double]::Parse($_,[Globalization.CultureInfo]::CurrentCulture)})
        for($scroll=0;$scroll -lt 3;$scroll++){
            [DetailNativeLayout]::ScrollFieldBounds([IntPtr]$window,[IntPtr]$excel.Hwnd,$fieldBounds)|
                Add-Content (Join-Path $reportRoot 'detail-native-scroll-geometry.tsv')
            Start-Sleep -Milliseconds 150
            CaptureFormEvidence 'Event Detail' ('event-detail-scroll-input-'+$scroll+'.png') $window
            [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailWindowForTest')
        }
        Capture-DetailOwnedEvidence 'event-detail-horizontal-scroll.png' $window
        if($DetailScrollLockDiagnostic){
            $originalLock=Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailFieldLockForTest' @($false,$false)
            if($originalLock -isnot [bool] -or -not $originalLock){throw 'Original detail lock state is not verified.'}
            $originalValues=Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailFieldValuesForTest'
            if($originalValues -isnot [string] -or $originalValues.Length -eq 0){throw 'Original detail values are unavailable.'}
            try {
                $unlocked=Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailFieldLockForTest' @($true,$false)
                if($unlocked -isnot [bool] -or $unlocked){throw 'Disposable detail unlock was not verified.'}
                $focused=Run 'invSys.Operations.xlam' 'modInventoryViewer.FocusDetailFieldsForTest'
                if($focused -isnot [bool] -or -not $focused){throw 'Unlocked detail focus is unavailable.'}
                Capture-DetailOwnedEvidence 'event-detail-unlocked-before.png' $window
                for($scroll=0;$scroll -lt 3;$scroll++){
                    [DetailNativeLayout]::ScrollFieldBounds([IntPtr]$window,[IntPtr]$excel.Hwnd,$fieldBounds)|
                        Add-Content (Join-Path $reportRoot 'detail-unlocked-scroll-geometry.tsv')
                    Start-Sleep -Milliseconds 150
                    CaptureFormEvidence 'Event Detail' ('event-detail-unlocked-input-'+$scroll+'.png') $window
                    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailWindowForTest')
                }
                $valuesPreserved=(Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailFieldValuesForTest') -ceq $originalValues
                if(-not $valuesPreserved){throw 'Disposable lock diagnostic changed field values.'}
            } finally {
                $restoredLock=Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailFieldLockForTest' @($true,$originalLock)
                if($restoredLock -isnot [bool] -or $restoredLock -ne $originalLock){throw 'Original detail lock state was not restored.'}
            }
            [pscustomobject]@{OriginalLocked=$originalLock;TemporarilyUnlocked=$true;RestoredLocked=$restoredLock;ValuesPreserved=$valuesPreserved;RuntimeChanged=$false;VisibleMovementRequiresImageReview=$true}|
                ConvertTo-Json|Set-Content (Join-Path $reportRoot 'detail-lock-diagnostic.json')
            Capture-DetailOwnedEvidence 'event-detail-lock-restored.png' $window
        }
    }
    foreach($module in @('modEventDetailSettings','modInventoryViewerData')) {
        $readCount=Run 'invSys.Core.xlam' ($module+'.DetailReadCounterForTest')
        if($readCount -isnot [int]){throw 'Detail read counter did not return an integer.'}
        Check ('EventDetail.LayoutDoesNotRead.'+$module) ($readCount -eq 0)
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
    } finally {
        if($null -ne $detailHost){
            try {$detailHost.Close($false)}
            finally {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($detailHost)}
        }
    }
}
