Attribute VB_Name = "modAdminSettingsControls"
Option Explicit
Option Private Module

Public Function AddControl(ByVal page As Object, ByVal kind As String, ByVal name As String, _
                           ByVal caption As String, ByVal x As Single, ByVal y As Single, _
                           ByVal width As Single, ByVal height As Single) As Object
    Dim control As Object
    Set control = page.Controls.Add("Forms." & kind & ".1", name, True)
    control.Left = x: control.Top = y: control.Width = width: control.Height = height
    If caption <> "" Then control.Caption = caption
    Set AddControl = control
End Function
