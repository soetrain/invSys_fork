VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReceivedTally"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'=================
'ReceivedTally.cls
'=================
' Simple event handler for both worksheet class modules
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Only for single cell selection
    If Target.Cells.count > 1 Then Exit Sub
    ' Check if cell is in ITEMS column
    If modTS_Data.IsInItemsColumn(Target) Then
        ' Set the global selected cell
        Set gSelectedCell = Target
        ' Use dynamic search UI for ReceivedTally
        On Error Resume Next
        modTS_Received.ShowDynamicItemSearch Target
        If Err.Number <> 0 Then
            MsgBox "Item Search could not open: " & Err.Description, vbExclamation
            Err.Clear
        End If
        On Error GoTo 0
    End If
End Sub

' Keep AggregateReceived in sync when REF_NUMBER/ITEMS/QUANTITY change
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ExitHandler
    If Target Is Nothing Then Exit Sub
    ' Limit to small ranges
    If Target.Cells.CountLarge > 200 Then Exit Sub
    Dim lo As ListObject
    On Error Resume Next
    Set lo = Me.ListObjects("ReceivedTally")
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Dim intersectRange As Range
    Set intersectRange = Application.Intersect(Target, lo.DataBodyRange)
    If intersectRange Is Nothing Then Exit Sub
    ' Rebuild aggregation whenever staging edits occur
    Application.EnableEvents = False
    modTS_Received.RebuildAggregation
ExitHandler:
    Application.EnableEvents = True
End Sub
