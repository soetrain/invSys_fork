# D13 visual geometry measured from packaged controls and their actual font.
# The modal public launcher is driven at activation; business handlers are intact.
function Install-DetailColumnsProbe($TestModule,$FormCode) {
    $TestModule.CodeModule.AddFromString(@'
Private mDetailColumnsArmed As Boolean
Private mDetailColumnsResults As String
Private mDetailColumnsEntries As Long
Public Sub ArmDetailColumns()
    mDetailColumnsArmed = True
    mDetailColumnsResults = ""
End Sub
Public Sub DriveDetailColumns(ByVal instance As frmAdminSettings)
    Dim size As Long
    If Not mDetailColumnsArmed Then Exit Sub
    mDetailColumnsArmed = False
    mDetailColumnsEntries = mDetailColumnsEntries + 1
    For size = 0 To 2
        If size > 0 Then mDetailColumnsResults = mDetailColumnsResults & vbLf
        mDetailColumnsResults = mDetailColumnsResults & DetailColumnsForInstance(instance, size)
    Next size
    instance.PolicyTestCloseAction
End Sub
Public Function PublicDetailColumnsResults() As String
    PublicDetailColumnsResults = CStr(mDetailColumnsEntries) & vbLf & mDetailColumnsResults
End Function
Public Function DetailColumnsForSize(ByVal size As Long) As String
    DetailColumnsForSize = DetailColumnsForInstance(mForm, size)
End Function
Private Function DetailColumnsForInstance(ByVal instance As frmAdminSettings, ByVal size As Long) As String
    Dim page As Object, sections As Object, child As Object, fields As Object
    Dim before As String, widths As Variant, names As Variant, index As Long, offset As Double
    Dim result As String
    before = instance.DetailTestRequest()
    instance.Width = 744: instance.Height = 696
    If size = 1 Then instance.Width = 864: instance.Height = 776
    Set page = TrackingPage(instance, "Event Tracking")
    For Each child In instance.Controls
        If TypeName(child) = "MultiPage" Then child.Value = page.Index
    Next child
    Set sections = TrackingControl(page, "MultiPage", "", "mpEventTracking")
    For Each child In sections.Pages
        If child.Caption = "Event Detail" Then
            sections.Value = child.Index
            Set page = child
            Exit For
        End If
    Next child
    instance.Repaint
    Set fields = TrackingControl(page, "ListBox", "", "lstDetailFields")
    widths = Split(fields.ColumnWidths, ";")
    names = Array("Field", "Show", "Order", "Required")
    For index = 0 To 3
        If index > 0 Then result = result & "|"
        result = result & CStr(DetailHeadingFits(page, fields, CStr(names(index)), offset, Val(widths(index + 1))))
        offset = offset + Val(widths(index + 1))
    Next index
    result = result & "|" & CStr(fields.ColumnCount = 5 And Val(widths(0)) = 0 And fields.ListCount > 0)
    result = result & "|" & CStr(before = instance.DetailTestRequest())
    DetailColumnsForInstance = result
End Function
Private Function DetailHeadingFits(ByVal page As Object, ByVal fields As Object, ByVal heading As String, _
                                   ByVal offset As Double, ByVal width As Double) As Boolean
    Dim control As Object, position As Long, headingLeft As Double, textWidth As Double
    Dim prefix As String
    For Each control In page.Controls
        If TypeName(control) = "Label" Then
            If control.Visible And control.Top >= fields.Top - 22 And control.Top < fields.Top Then
                position = InStr(1, control.Caption, heading, vbBinaryCompare)
                If position > 0 Then
                    prefix = Left$(control.Caption, position - 1)
                    headingLeft = control.Left + DetailTextWidth(page, control, prefix & "X") - DetailTextWidth(page, control, "X")
                    textWidth = DetailTextWidth(page, control, heading)
                    DetailHeadingFits = Abs(headingLeft - (fields.Left + offset)) <= 3 And _
                        textWidth <= width And control.Top + control.Height <= fields.Top And _
                        headingLeft >= fields.Left - 3 And headingLeft + textWidth <= fields.Left + offset + width + 3
                    Exit Function
                End If
            End If
        End If
    Next control
End Function
Private Function DetailTextWidth(ByVal page As Object, ByVal source As Object, ByVal value As String) As Double
    Dim measure As Object
    Set measure = page.Controls.Add("Forms.Label.1", "detailColumnMeasureForTest", False)
    With measure
        .Font.Name = source.Font.Name: .Font.Size = source.Font.Size
        .Font.Bold = source.Font.Bold: .Font.Italic = source.Font.Italic
        .WordWrap = False: .Caption = value: .AutoSize = True
        DetailTextWidth = .Width
    End With
    page.Controls.Remove "detailColumnMeasureForTest"
End Function
'@)
    # Run after the original activation setup, without changing its return paths.
    $line=$FormCode.ProcStartLine('UserForm_Activate',0)
    $count=$FormCode.ProcCountLines('UserForm_Activate',0)
    $body=$FormCode.Lines($line,$count)
    $body=$body.Replace('    mResizeInitialized = True',"    mResizeInitialized = True`r`n    TestD5Commands.DriveDetailColumns Me")
    if($body -notmatch 'DriveDetailColumns Me'){throw 'Settings activation seam missing.'}
    $FormCode.DeleteLines($line,$count)
    $FormCode.InsertLines($line,$body)
}

function Assert-DetailColumns([string]$Prefix,[string]$Result) {
    $flags=$Result.Split('|')
    if($flags.Count -ne 6 -or @($flags|Where-Object{$_ -cnotin @('True','False')}).Count){throw 'Invalid fixed column geometry result.'}
    $names=@('FieldAlignedAndFits','ShowAlignedAndFits','OrderAlignedAndFits','RequiredAlignedAndFits','HiddenIdentityAndRowsPreserved','StagedRequestPreserved')
    for($i=0;$i -lt $names.Count;$i++){Check ($Prefix+'.'+$names[$i]) ($flags[$i] -ceq 'True')}
}

function Test-DetailColumns($Fixture) {
    $before=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $sizes=@('Default','Enlarged','Restored')
    for($size=0;$size -lt 3;$size++){
        Assert-DetailColumns ('DetailColumns.Surface.'+$sizes[$size]) ([string](Run 'invSys.Admin.xlam' 'TestD5Commands.DetailColumnsForSize' @($size)))
        if($CaptureEvidence){CaptureOwnedFormByCaptionEvidence 'invSys Settings' ('detail-columns-'+$sizes[$size].ToLowerInvariant()+'.png')}
    }
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.ArmDetailColumns')
    [void](Run 'invSys.Admin.xlam' 'modAdmin.Open_Settings')
    $public=([string](Run 'invSys.Admin.xlam' 'TestD5Commands.PublicDetailColumnsResults')) -split '\r?\n'
    if($public.Count -ne 4){throw 'Public Settings activation geometry unavailable.'}
    Check 'DetailColumns.ActualPublicLauncherObserved' ($public[0] -ceq '1')
    for($size=0;$size -lt 3;$size++){Assert-DetailColumns ('DetailColumns.PublicLauncher.'+$sizes[$size]) $public[$size+1]}
    Check 'DetailColumns.SavedConfigBytesPreserved' ($before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
}
