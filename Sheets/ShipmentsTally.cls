VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShipmentsTally"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'===================
' ShipmentsTally.cls
' ==================
Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Application.EnableEvents = False Then Application.EnableEvents = True
    If Target.Cells.CountLarge > 1 Then Exit Sub

    Dim lo As ListObject
    On Error Resume Next
    Set lo = Target.ListObject
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    If LCase$(lo.Name) <> "shipmentstally" Then Exit Sub

    Dim itemsCol As ListColumn
    On Error Resume Next
    Set itemsCol = lo.ListColumns("ITEMS")
    On Error GoTo 0
    If itemsCol Is Nothing Then Exit Sub
    If Target.Column <> itemsCol.Range.Column Then Exit Sub
    If Target.Row <= lo.HeaderRowRange.Row Then Exit Sub

    Set gSelectedCell = Target
    modTS_Shipments.ShowDynamicItemSearch Target
End Sub



