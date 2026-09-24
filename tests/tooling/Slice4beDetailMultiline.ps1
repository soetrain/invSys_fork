# Observes the real Viewer-owned form after its existing contributing-line handler.
function Install-Slice4beDetailMultilineProbe {
    $code=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modInventoryViewer').CodeModule
    $code.AddFromString(@'
Public Function DetailMultilineFactForTest(ByVal fact As String, Optional ByVal fixtureIndex As Long = 2) As Boolean
    Dim instance As Object, area As Object, control As Object, value As Object, heading As Object, fields As Object, status As Object
    Dim expected As String, rendered As String, measure As Object, lineCount As Long, second As Object, secondHeading As Object, extent As Single
    Set instance = DetailFixtureFormForTest()
    If instance Is Nothing Then Err.Raise 5, , "Published Detail form is unavailable."
    For Each control In instance.Controls
        If control.Name = "fraDetailMultiline" Then Set area = control
    Next control
    If area Is Nothing Then Exit Function
    For Each control In area.Controls
        If control.Name = "lblMultilineValue1" Then Set value = control
        If control.Name = "lblMultilineCaption1" Then Set heading = control
        If control.Name = "lblMultilineValue2" Then Set second = control
        If control.Name = "lblMultilineCaption2" Then Set secondHeading = control
    Next control
    Select Case fact
        Case "Present": DetailMultilineFactForTest = (TypeName(area) = "Frame" And area.Visible)
        Case "Content"
            If value Is Nothing Or heading Is Nothing Then Exit Function
            ' Office normalizes Label separators; the original list string is tested separately by binary equality.
            expected = Replace$(Replace$(DetailOriginalTextForTest(fixtureIndex), vbCrLf, vbLf), vbCr, vbLf)
            rendered = Replace$(value.Caption, vbCrLf, vbLf)
            lineCount = UBound(Split(expected, vbLf)) + 1
            Set measure = instance.Controls.Add("Forms.Label.1", "MultilineLineMeasureForTest", False)
            measure.Font.Name = value.Font.Name: measure.Font.Size = value.Font.Size
            measure.WordWrap = False: measure.AutoSize = True: measure.Caption = "A"
            DetailMultilineFactForTest = (heading.Caption = "Item name" And StrComp(rendered, expected, vbBinaryCompare) = 0 And value.Height >= measure.Height * lineCount - 1)
            instance.Controls.Remove "MultilineLineMeasureForTest"
        Case "ReadOnly"
            For Each control In area.Controls
                If TypeName(control) <> "Label" Then Exit Function
                If Len(control.Accelerator) <> 0 Then Exit Function
                If InStr(1, control.Caption, "DO-NOT-DISPLAY", vbBinaryCompare) > 0 Then Exit Function
            Next control
            DetailMultilineFactForTest = Not value Is Nothing
        Case "MultipleFields"
            If value Is Nothing Or heading Is Nothing Or second Is Nothing Or secondHeading Is Nothing Then Exit Function
            If heading.Caption <> "Item name" Or secondHeading.Caption <> "Location" Or area.Controls.Count <> 4 Then Exit Function
            If secondHeading.Top < value.Top + value.Height Or second.Top < secondHeading.Top + secondHeading.Height Then Exit Function
            If second.Top + second.Height > area.ScrollHeight Or second.Left + second.Width > area.ScrollWidth Then Exit Function
            If StrComp(Replace$(second.Caption, vbCrLf, vbLf), DetailMultilineLocationForTest(1), vbBinaryCompare) <> 0 Then Exit Function
            Set fields = instance.Controls("lstEventFields")
            For lineCount = 0 To fields.ListCount - 1
                If fields.List(lineCount, 0) = "Location" Then
                    DetailMultilineFactForTest = (StrComp(fields.List(lineCount, 1), DetailMultilineLocationForTest(1), vbBinaryCompare) = 0)
                    Exit Function
                End If
            Next lineCount
        Case "Empty", "Invalidated"
            If Not value Is Nothing Then Exit Function
            DetailMultilineFactForTest = (area.ScrollTop = 0 And area.ScrollLeft = 0)
        Case "Geometry"
            Set fields = instance.Controls("lstEventFields"): Set status = instance.Controls("lblDetailStatus")
            If area.Left < 0 Or area.Top < fields.Top + fields.Height Or area.Top + area.Height > status.Top Then Exit Function
            If area.Left + area.Width > instance.InsideWidth + 1 Then Exit Function
            For Each control In area.Controls
                If control.Left + control.Width > area.ScrollWidth Or control.Top + control.Height > area.ScrollHeight Then Exit Function
            Next control
            DetailMultilineFactForTest = True
        Case "Overflow"
            If value Is Nothing Then Exit Function
            DetailMultilineFactForTest = (value.Height > area.InsideHeight And value.Width > area.InsideWidth And area.ScrollBars = 3)
        Case "Bottom"
            If value Is Nothing Then Exit Function
            For Each control In area.Controls
                If control.Top + control.Height > extent Then extent = control.Top + control.Height
            Next control
            DetailMultilineFactForTest = (area.ScrollTop > 0 And extent <= area.ScrollTop + area.InsideHeight - 12)
        Case "Right"
            If value Is Nothing Then Exit Function
            DetailMultilineFactForTest = (area.ScrollLeft > 0 And value.Left + value.Width <= area.ScrollLeft + area.InsideWidth - 12)
    End Select
End Function
Public Function SelectDetailMultilineLineForTest() As Boolean
    Dim instance As Object, lines As Object, i As Long
    Set instance = DetailFixtureFormForTest()
    If instance Is Nothing Then Err.Raise 5, , "Published Detail form is unavailable."
    Set lines = instance.Controls("lstEventLines")
    For i = 0 To lines.ListCount - 1
        If CStr(lines.List(i, 0)) = "SYS-DETAIL-B" Then
            lines.ListIndex = i: SelectDetailMultilineLineForTest = True: Exit Function
        End If
    Next i
End Function
Public Function SetDetailMultilineProfileForTest(ByVal hide As Boolean, ByVal expectedVersion As Long) As Boolean
    Dim request As String, report As String
    request = modEventDetailSettings.StageField(modEventDetailSettings.DefaultRequest(), "Receiving", "ItemName", Not hide)
    SetDetailMultilineProfileForTest = modEventDetailSettings.SaveProfile(modActivity.CaptureContext(), expectedVersion, request, report)
End Function
'@)
}

function Test-Slice4beDetailMultiline([string]$Role) {
    foreach($index in @(1,2)){
        if(-not [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailOriginalTextMatchesForTest' @($index))){throw 'Original published multiline fixture is not preserved.'}
        Check ('EventDetail.Multiline.'+$Role+'.OriginalRendered.'+$index) ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailMultilineFactForTest' @('Content',$index)))
        if($Role -ceq 'Operations' -and $CaptureEvidence){
            $window=[long](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailWindowForTest')
            Capture-DetailOwnedEvidence ('event-detail-multiline-case-'+$index+'.png') $window
        }
        if($index -eq 1){
            Check ('EventDetail.Multiline.'+$Role+'.MultipleFieldsInProfileOrder') ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailMultilineFactForTest' @('MultipleFields')))
            if($Role -ceq 'Operations' -and $CaptureEvidence){
                $present=[bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailMultilineFactForTest' @('Present'))
                if($present){Invoke-DetailMultilineNativeScroll $window 'event-detail-multiline-multiple' @('Bottom')}
                Check 'EventDetail.Multiline.NativeReachesLastOfMultipleFields' ($present -and [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailMultilineFactForTest' @('Bottom')))
            }
        }
    }
    foreach($fact in @('Present','ReadOnly','Geometry','Overflow')){
        Check ('EventDetail.Multiline.'+$Role+'.'+$fact) ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailMultilineFactForTest' @($fact)))
    }
    $present=[bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailMultilineFactForTest' @('Present'))
    if($Role -ceq 'Operations'){
        $window=[long](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailWindowForTest')
        if($CaptureEvidence){Capture-DetailOwnedEvidence 'event-detail-multiline-before.png' $window}
        foreach($size in @('FitDefault','FitLarger','FitDefault')){
            if(-not [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailActionForTest' @($size))){throw 'Existing Detail layout failed.'}
            $suffix=if($size -ceq 'FitDefault' -and $results.Check -contains 'EventDetail.Multiline.Layout.FitDefault'){'.Restored'}else{''}
            Check ('EventDetail.Multiline.Layout.'+$size+$suffix) ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailMultilineFactForTest' @('Geometry')))
        }
        foreach($command in @(3,9)){
            [void][DetailNativeLayout]::ShowWindow([IntPtr]$window,$command)
            Start-Sleep -Milliseconds 200
            [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailWindowForTest')
            $name=if($command -eq 3){'Maximize'}else{'Restore'}
            Check ('EventDetail.Multiline.Layout.Native'+$name) ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailMultilineFactForTest' @('Geometry')))
            if($CaptureEvidence){Capture-DetailOwnedEvidence ('event-detail-multiline-'+$name.ToLowerInvariant()+'.png') $window}
        }
        if($present -and $CaptureEvidence){
            Invoke-DetailMultilineNativeScroll $window
        }
        if($CaptureEvidence){
            Check 'EventDetail.Multiline.NativeVerticalReachability' ($present -and [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailMultilineFactForTest' @('Bottom')))
            Check 'EventDetail.Multiline.NativeHorizontalReachability' ($present -and [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailMultilineFactForTest' @('Right')))
            Check 'EventDetail.Multiline.NativeScrollPreservesOriginal' ($present -and [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailMultilineFactForTest' @('Content',2)))
        }
    }
    if(-not [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailOriginalTextMatchesForTest' @(0))){throw 'Single-line fixture unavailable.'}
    Check ('EventDetail.Multiline.'+$Role+'.SingleLineClearsPreview') ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailMultilineFactForTest' @('Empty')))
}

function Invoke-DetailMultilineNativeScroll([long]$Window,[string]$Prefix='event-detail-multiline',[string[]]$Axes=@('Bottom','Right')) {
    if(-not ('DetailMultilineInput' -as [type])){
        Add-Type @'
using System; using System.Runtime.InteropServices;
public static class DetailMultilineInput {
 [StructLayout(LayoutKind.Sequential)] public struct Point { public int X,Y; }
 [StructLayout(LayoutKind.Sequential)] public struct Rect { public int Left,Top,Right,Bottom; }
 [StructLayout(LayoutKind.Sequential)] public struct Mouse { public int X,Y; public uint Data,Flags,Time; public UIntPtr Extra; }
 [StructLayout(LayoutKind.Sequential)] public struct Input { public uint Type; public Mouse Mouse; }
 [DllImport("user32.dll")] static extern bool GetClientRect(IntPtr h,out Rect r);
 [DllImport("user32.dll")] static extern bool ClientToScreen(IntPtr h,ref Point p);
 [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")] static extern IntPtr WindowFromPoint(Point p);
 [DllImport("user32.dll")] static extern IntPtr GetAncestor(IntPtr h,uint flag);
 [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
 [DllImport("user32.dll")] static extern IntPtr MonitorFromPoint(Point p,uint flags);
 [DllImport("user32.dll")] static extern bool SetCursorPos(int x,int y);
 [DllImport("user32.dll")] static extern bool GetCursorPos(out Point p);
 [DllImport("user32.dll")] static extern uint SendInput(uint n,Input[] input,int size);
 [DllImport("user32.dll")] static extern uint GetDpiForWindow(IntPtr h);
 [DllImport("user32.dll")] static extern int GetSystemMetricsForDpi(int index,uint dpi);
 public static string Click(IntPtr form,IntPtr excel,double[] b,bool vertical) {
  uint owner,expected; GetWindowThreadProcessId(form,out owner); GetWindowThreadProcessId(excel,out expected);
  if(owner==0 || owner!=expected || GetForegroundWindow()!=form)throw new Exception("Multiline input target is not the owned foreground form.");
  Rect client; Point origin=new Point();
  if(!GetClientRect(form,out client) || !ClientToScreen(form,ref origin))throw new Exception("Form client geometry unavailable.");
  double sx=(client.Right-client.Left)/b[4],sy=(client.Bottom-client.Top)/b[5];
  int right=origin.X+(int)Math.Round((b[0]+b[2])*sx),bottom=origin.Y+(int)Math.Round((b[1]+b[3])*sy);
  uint dpi=GetDpiForWindow(form); int sw=GetSystemMetricsForDpi(2,dpi),sh=GetSystemMetricsForDpi(3,dpi);
  Point point=new Point {X=right-(vertical?sw/2:sw+sw/2)-2,Y=bottom-(vertical?sh+sh/2:sh/2)-2};
  if(MonitorFromPoint(point,0)==IntPtr.Zero || GetAncestor(WindowFromPoint(point),2)!=form)throw new Exception("Multiline scroll point is not exposed on the owned form.");
  Point actual;
  if(!SetCursorPos(point.X,point.Y) || !GetCursorPos(out actual) || actual.X!=point.X || actual.Y!=point.Y)throw new Exception("Multiline cursor placement failed.");
  if(GetForegroundWindow()!=form || GetAncestor(WindowFromPoint(point),2)!=form)throw new Exception("Multiline scroll point became covered.");
  Input[] down={new Input {Mouse=new Mouse {Flags=2}}},up={new Input {Mouse=new Mouse {Flags=4}}};
  try {if(SendInput(1,down,Marshal.SizeOf(typeof(Input)))!=1)throw new Exception("Multiline press failed.");System.Threading.Thread.Sleep(20);}
  finally {if(SendInput(1,up,Marshal.SizeOf(typeof(Input)))!=1)throw new Exception("Multiline release failed.");}
  return point.X+"|"+point.Y+"|"+sw+"|"+sh;
 }
}
'@
    }
    Capture-DetailOwnedEvidence ($Prefix+'-scroll-start.png') $Window
    $geometry=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailGeometryForTest')
    $rows=@($geometry -split "`n"|Where-Object {$_ -like "fraDetailMultiline`t*"})
    if($rows.Count -ne 1){throw 'Multiline bounds are ambiguous.'}
    [double[]]$bounds=@(($rows[0] -split "`t")[1..6]|ForEach-Object {[double]::Parse($_,[Globalization.CultureInfo]::CurrentCulture)})
    foreach($axis in $Axes){
        for($click=0;$click -lt 150;$click++){
            if([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailMultilineFactForTest' @($axis))){break}
            [DetailMultilineInput]::Click([IntPtr]$Window,[IntPtr]$excel.Hwnd,$bounds,($axis -ceq 'Bottom'))|
                Add-Content (Join-Path $reportRoot 'detail-multiline-input.tsv')
            [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.DetailWindowForTest')
        }
        Capture-DetailOwnedEvidence ($Prefix+'-'+$axis.ToLowerInvariant()+'.png') $Window
    }
}
