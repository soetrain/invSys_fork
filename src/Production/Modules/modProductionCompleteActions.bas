Attribute VB_Name = "modProductionCompleteActions"
Option Explicit
Option Private Module

' Keep one operator completion active across the owner's yielding work.
Public Sub Execute(ByVal owner As frmProduction, ByRef loading As Boolean, ByRef busy As Boolean)
    Dim priorLoading As Boolean, priorBusy As Boolean
    Dim number As Long, source As String, description As String, helpFile As String, helpContext As Long
    If loading Or busy Then Exit Sub
    priorLoading = loading: priorBusy = busy
    On Error GoTo Failed
    busy = True
    owner.CompleteProductionRun
Done:
    loading = priorLoading: busy = priorBusy
    If number <> 0 Then
        On Error GoTo 0
        Err.Raise number, source, description, helpFile, helpContext
    End If
    Exit Sub
Failed:
    number = Err.Number: source = Err.Source: description = Err.Description
    helpFile = Err.HelpFile: helpContext = Err.HelpContext
    Resume Done
End Sub
