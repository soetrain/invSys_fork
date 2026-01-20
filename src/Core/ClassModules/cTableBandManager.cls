VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTableBandManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Manages "process bands" on the Production sheet.
' A band = all ListObjects that share the same process key in their name.

Private mWs As Worksheet

Public Sub Init(ByVal ws As Worksheet)
    Set mWs = ws
End Sub

Public Sub ExpandBandForTable(ByVal lo As ListObject, ByVal rowsAdded As Long)
    If mWs Is Nothing Then Exit Sub
    If lo Is Nothing Then Exit Sub
    If rowsAdded <= 0 Then Exit Sub

    Dim processKey As String
    processKey = ExtractProcessKey(lo.Name)

    Dim bandTables As Collection
    If processKey = "" Then
        Set bandTables = New Collection
        bandTables.Add lo
    Else
        Set bandTables = GetBandTables(processKey)
    End If
    If bandTables Is Nothing Then Exit Sub
    If bandTables.Count = 0 Then Exit Sub

    Dim bandTop As Long, bandBottom As Long, bandLeft As Long, bandRight As Long
    ComputeBandBounds bandTables, bandTop, bandBottom, bandLeft, bandRight
    If bandTop = 0 Or bandBottom = 0 Or bandLeft = 0 Or bandRight = 0 Then Exit Sub

    Dim insertTop As Long: insertTop = bandBottom + 1
    If insertTop > mWs.Rows.Count Then Exit Sub

    Dim insertRange As Range
    Set insertRange = mWs.Range( _
        mWs.Cells(insertTop, bandLeft), _
        mWs.Cells(insertTop + rowsAdded - 1, bandRight))

    On Error Resume Next
    insertRange.Insert Shift:=xlShiftDown
    On Error GoTo 0
End Sub

Public Function GetBandTables(ByVal processKey As String) As Collection
    Dim result As New Collection
    If mWs Is Nothing Then
        Set GetBandTables = result
        Exit Function
    End If

    Dim lo As ListObject
    For Each lo In mWs.ListObjects
        If processKey = "" Then
            ' If no process key, only return the matching table name.
            If lo.Name = processKey Then result.Add lo
        ElseIf StrComp(ExtractProcessKey(lo.Name), processKey, vbTextCompare) = 0 Then
            result.Add lo
        End If
    Next lo

    Set GetBandTables = result
End Function

Public Sub ComputeBandBounds(ByVal bandTables As Collection, ByRef bandTop As Long, _
    ByRef bandBottom As Long, ByRef bandLeft As Long, ByRef bandRight As Long)

    Dim lo As ListObject
    Dim topSet As Boolean
    For Each lo In bandTables
        If lo Is Nothing Then GoTo NextLo
        Dim rTop As Long: rTop = lo.Range.Row
        Dim rBottom As Long: rBottom = lo.Range.Row + lo.Range.Rows.Count - 1
        Dim cLeft As Long: cLeft = lo.Range.Column
        Dim cRight As Long: cRight = lo.Range.Column + lo.Range.Columns.Count - 1
        If Not topSet Then
            bandTop = rTop: bandBottom = rBottom
            bandLeft = cLeft: bandRight = cRight
            topSet = True
        Else
            If rTop < bandTop Then bandTop = rTop
            If rBottom > bandBottom Then bandBottom = rBottom
            If cLeft < bandLeft Then bandLeft = cLeft
            If cRight > bandRight Then bandRight = cRight
        End If
NextLo:
    Next lo
End Sub

Public Function ExtractProcessKey(ByVal tableName As String) As String
    Dim nm As String: nm = tableName
    If LCase$(Left$(nm, 5)) <> "proc_" Then Exit Function
    Dim parts As Variant: parts = Split(nm, "_")
    If UBound(parts) < 2 Then Exit Function

    Dim i As Long
    For i = 1 To UBound(parts) - 1
        If ExtractProcessKey <> "" Then ExtractProcessKey = ExtractProcessKey & "_"
        ExtractProcessKey = ExtractProcessKey & parts(i)
    Next i
End Function
