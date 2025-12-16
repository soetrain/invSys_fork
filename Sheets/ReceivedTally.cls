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
    ' Make sure events are on (they can be left off after runtime errors)
    If Application.EnableEvents = False Then Application.EnableEvents = True
    ' Only for single cell selection
    If Target.Cells.count > 1 Then Exit Sub
    Dim lo As ListObject
    On Error Resume Next
    Set lo = Target.ListObject
    On Error GoTo 0
    If lo Is Nothing Then
        Exit Sub
    End If
    If LCase$(lo.Name) <> "receivedtally" Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    ' must be in data body and in ITEMS column
    If Application.Intersect(Target, lo.DataBodyRange) Is Nothing Then Exit Sub
    Dim itemsCol As ListColumn
    On Error Resume Next
    Set itemsCol = lo.ListColumns("ITEMS")
    On Error GoTo 0
    If itemsCol Is Nothing Then Exit Sub
    If Target.Column <> itemsCol.Range.Column Then Exit Sub

    ' Set the global selected cell
    Set gSelectedCell = Target
    ' Use dynamic search UI for ReceivedTally (bubble up any error now)
    modTS_Received.ShowDynamicItemSearch Target
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
