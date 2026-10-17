VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "cInventoryTableInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Core-only transient input state. Values stay in memory; canonical row position
' is never identity. A staging row index addresses only its captured local table.
Private mOriginal As ListObject
Private mKeyHeader As String
Private mState As Object
Private mPlanned As Object
Private mNewHeaders As Object

Public Function Capture(ByVal table As ListObject, ByVal keyHeader As String, ByVal trackObject As Boolean) As Boolean
    mKeyHeader = UCase$(Trim$(keyHeader))
    Set mState = ReadState(table)
    If mState Is Nothing Then Exit Function
    If trackObject Then Set mOriginal = table Else Set mOriginal = Nothing
    Set mPlanned = CreateObject("Scripting.Dictionary"): mPlanned.CompareMode = vbBinaryCompare
    Set mNewHeaders = CreateObject("Scripting.Dictionary")
    Capture = True
End Function

Public Sub PlanValue(ByVal rowKey As String, ByVal header As String, ByVal value As Variant)
    Dim fields As Object
    If IsObject(value) Or IsArray(value) Or IsError(value) Or IsNull(value) Then Err.Raise 5, , "A planned scalar is required."
    header = UCase$(Trim$(header))
    If Not mPlanned.Exists(rowKey) Then
        Set fields = CreateObject("Scripting.Dictionary")
        mPlanned.Add rowKey, fields
    End If
    mPlanned(rowKey)(header) = value
    mNewHeaders(header) = True
End Sub

Public Function KeysForSku(ByVal sku As String) As Collection
    Dim result As New Collection, key As Variant
    For Each key In mState("Skus").Keys
        If StrComp(CStr(mState("Skus")(key)), sku, vbTextCompare) = 0 Then result.Add CStr(key)
    Next key
    Set KeysForSku = result
End Function

Public Function PlanNewRow(ByVal rowKey As String) As Boolean
    Dim header As Variant
    If rowKey = "" Or mState("Rows").Exists(rowKey) Or mPlanned.Exists(rowKey) Then Exit Function
    For Each header In mState("Headers").Keys
        PlanValue rowKey, CStr(header), ""
    Next header
    PlanNewRow = True
End Function

Public Function Scalar(ByVal rowKey As String, ByVal header As String) As Variant
    Dim fields As Object
    On Error GoTo Invalid
    If Not mState("Rows").Exists(rowKey) Then Exit Function
    Set fields = mState("Rows")(rowKey)
    header = UCase$(Trim$(header))
    If Not fields.Exists(header) Then Exit Function
    Scalar = fields(header)(3)
Invalid:
End Function

Public Function Matches(ByVal table As ListObject, Optional ByVal selectedSku As String = "", _
        Optional ByVal headersOnly As Boolean = False, Optional ByVal requireResult As Boolean = False) As Boolean
    On Error GoTo Invalid
    If Not mOriginal Is Nothing Then
        If Not (table Is mOriginal) Then Exit Function
    End If
    Matches = MatchesState(table, selectedSku, headersOnly, requireResult)
Invalid:
End Function

' Only the coordinator's independent saved-byte copy uses content comparison.
' Ordinary live-input checks above retain the captured object identity guard.
Public Function MatchesPersistedCopy(ByVal table As ListObject) As Boolean
    MatchesPersistedCopy = MatchesState(table, "", False, True)
End Function

Private Function MatchesState(ByVal table As ListObject, ByVal selectedSku As String, _
        ByVal headersOnly As Boolean, ByVal requireResult As Boolean) As Boolean
    Dim current As Object, header As Variant, key As Variant, before As Object, actual As Object
    Dim keys As Object, relevant As Boolean, plannedValue As Variant
    On Error GoTo Invalid
    If mState Is Nothing Then Exit Function
    Set current = ReadState(table)
    If current Is Nothing Then Exit Function
    For Each header In mState("Headers").Keys
        If Not current("Headers").Exists(header) Then Exit Function
    Next header
    For Each header In current("Headers").Keys
        If Not mState("Headers").Exists(header) And Not mNewHeaders.Exists(header) Then Exit Function
    Next header
    If headersOnly Then MatchesState = True: Exit Function
    Set keys = CreateObject("Scripting.Dictionary"): keys.CompareMode = vbBinaryCompare
    For Each key In mState("Rows").Keys: keys(key) = True: Next key
    For Each key In current("Rows").Keys: keys(key) = True: Next key
    For Each key In mPlanned.Keys: keys(key) = True: Next key
    For Each key In keys.Keys
        relevant = (selectedSku = "")
        If mState("Skus").Exists(key) Then relevant = relevant Or _
            (StrComp(CStr(mState("Skus")(key)), selectedSku, vbTextCompare) = 0)
        If current("Skus").Exists(key) Then relevant = relevant Or _
            (StrComp(CStr(current("Skus")(key)), selectedSku, vbTextCompare) = 0)
        If mPlanned.Exists(key) Then relevant = True
        If relevant Then
            If Not current("Rows").Exists(key) Then
                If mState("Rows").Exists(key) Or requireResult Then Exit Function
            Else
                Set actual = current("Rows")(key): Set before = Nothing
                If mState("Rows").Exists(key) Then Set before = mState("Rows")(key)
                If before Is Nothing And Not mPlanned.Exists(key) Then Exit Function
                For Each header In current("Headers").Keys
                    If mPlanned.Exists(key) Then
                        If mPlanned(key).Exists(header) Then
                            plannedValue = mPlanned(key)(header)
                            If PlannedValueMatches(actual(header), plannedValue) Then GoTo NextField
                            If requireResult Then Exit Function
                        End If
                    End If
                    If before Is Nothing Then
                        If actual(header)(0) <> "" Or actual(header)(1) <> "" Then Exit Function
                    Else
                        If before.Exists(header) Then
                            If actual(header)(2) <> before(header)(2) Then Exit Function
                        ElseIf actual(header)(0) <> "" Or actual(header)(1) <> "" Then
                            Exit Function
                        End If
                    End If
NextField:
                Next header
                If requireResult And mPlanned.Exists(key) Then
                    For Each header In mPlanned(key).Keys
                        If Not actual.Exists(header) Then Exit Function
                    Next header
                End If
            End If
        End If
    Next key
    MatchesState = True
Invalid:
End Function

Private Function PlannedValueMatches(ByVal actual As Variant, ByVal expected As Variant) As Boolean
    On Error GoTo Invalid
    ' Keep the existing Excel scalar coercions used by the catalog writer.
    ' Keep the typed Excel scalar: CStr(Double) can lose date/number precision.
    If IsNumeric(actual(3)) Then
        If IsNumeric(expected) Then
            PlannedValueMatches = (CDbl(actual(3)) = CDbl(expected)): Exit Function
        ElseIf IsDate(expected) Then
            PlannedValueMatches = (CDbl(actual(3)) = CDbl(CDate(expected))): Exit Function
        End If
    End If
    If actual(0) = CStr(expected) Or actual(1) = CStr(expected) Then PlannedValueMatches = True: Exit Function
    If LCase$(expected) = "true" Or LCase$(expected) = "false" Then _
        PlannedValueMatches = (LCase$(CStr(actual(0))) = LCase$(expected))
Invalid:
End Function

Public Function Confirm(ByVal table As ListObject, Optional ByVal selectedSku As String = "") As Boolean
    If Not Matches(table, selectedSku, False, True) Then Exit Function
    Confirm = Capture(table, mKeyHeader, Not mOriginal Is Nothing)
End Function

Public Sub ReleaseObjectReference()
    Set mOriginal = Nothing
End Sub

Public Function BindObject(ByVal table As ListObject) As Boolean
    If Not Matches(table) Then Exit Function
    Set mOriginal = table
    BindObject = True
End Function

Private Function ReadState(ByVal table As ListObject) As Object
    Dim state As Object, headers As Object, rows As Object, skus As Object, fields As Object
    Dim column As ListColumn, key As String, row As Long, name As Variant, cell As Range
    Dim text As String, formula As String, populated As Boolean, values As Variant
    On Error GoTo Invalid
    If table Is Nothing Then Exit Function
    Set headers = CreateObject("Scripting.Dictionary")
    For Each column In table.ListColumns
        key = UCase$(Trim$(column.Name))
        If key = "" Or key = "ROW" Or headers.Exists(key) Then Exit Function
        headers.Add key, column.Index
    Next column
    If mKeyHeader <> "" Then
        If Not headers.Exists(mKeyHeader) Then Exit Function
    End If
    Set rows = CreateObject("Scripting.Dictionary"): rows.CompareMode = vbBinaryCompare
    Set skus = CreateObject("Scripting.Dictionary")
    For row = 1 To table.ListRows.Count
        Set fields = CreateObject("Scripting.Dictionary"): populated = False
        For Each name In headers.Keys
            Set cell = table.DataBodyRange.Cells(row, headers(name))
            If IsError(cell.Value2) Then text = cell.Text Else text = CStr(cell.Value2)
            If IsError(cell.Formula) Then formula = cell.Text Else formula = CStr(cell.Formula)
            values = Array(text, formula, CStr(VarType(cell.Value2)) & ":" & CStr(Len(text)) & ":" & text & ":" & formula, cell.Value2)
            fields.Add CStr(name), values
            If text <> "" Or formula <> "" Then populated = True
        Next name
        If populated Then
            If mKeyHeader = "" Then key = CStr(row) Else key = CStr(fields(mKeyHeader)(0))
            If key = "" Or rows.Exists(key) Then Exit Function
            rows.Add key, fields
            If headers.Exists("SKU") Then skus.Add key, CStr(fields("SKU")(0))
        End If
    Next row
    Set state = CreateObject("Scripting.Dictionary")
    state.Add "Headers", headers: state.Add "Rows", rows: state.Add "Skus", skus
    Set ReadState = state
Invalid:
End Function
