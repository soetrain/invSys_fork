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

Private Sub Worksheet_Activate()
    On Error Resume Next
    modTS_Shipments.InitializeShipmentsUI
    On Error GoTo 0
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Application.EnableEvents = False Then Application.EnableEvents = True
    If Target.Cells.CountLarge > 1 Then Exit Sub

    Dim lo As ListObject
    On Error Resume Next
    Set lo = Target.ListObject
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub

    Dim loName As String
    loName = LCase$(lo.Name)

    Select Case loName
        Case "shipmentstally"
            Dim itemsCol As ListColumn
            On Error Resume Next
            Set itemsCol = lo.ListColumns("ITEMS")
            On Error GoTo 0
            If itemsCol Is Nothing Then Exit Sub
            If Target.Column <> itemsCol.Range.Column Then Exit Sub
            If Target.Row <= lo.HeaderRowRange.Row Then Exit Sub
            Set gSelectedCell = Target
            modTS_Shipments.ShowDynamicItemSearch Target
        Case "boxbom"
            Dim bomItemsCol As ListColumn
            On Error Resume Next
            Set bomItemsCol = lo.ListColumns("ITEM")
            On Error GoTo 0
            If bomItemsCol Is Nothing Then Exit Sub
            If Target.Column <> bomItemsCol.Range.Column Then Exit Sub
            If Target.Row <= lo.HeaderRowRange.Row Then Exit Sub
            Set gSelectedCell = Target
            modTS_Shipments.ShowDynamicItemSearch Target
    End Select
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ExitHandler
    If Target Is Nothing Then Exit Sub
    If Target.Cells.CountLarge > 50 Then Exit Sub

    Dim lo As ListObject
    On Error Resume Next
    Set lo = Me.ListObjects("ShipmentsTally")
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim qtyCol As ListColumn
    On Error Resume Next
    Set qtyCol = lo.ListColumns("QUANTITY")
    On Error GoTo 0
    If qtyCol Is Nothing Then Exit Sub

    Dim hit As Range
    Set hit = Application.Intersect(Target, qtyCol.DataBodyRange)
    If hit Is Nothing Then Exit Sub

    Application.EnableEvents = False
    modTS_Shipments.InvalidateAggregates True
ExitHandler:
    Application.EnableEvents = True
End Sub


