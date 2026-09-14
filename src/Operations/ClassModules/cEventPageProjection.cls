VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "cEventPageProjection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Pure loaded-projection grouping. Indexes are local UI references, never identity.
Private mRows As Variant
Private mGroups As Object
Private mSearch As Object
Private mTimes As Object
Private mKinds As Object
Private mOrdered() As String
Private mMatches As Collection
Private mRowGroups() As String
Private mFields As Object

Public Sub LoadProjection(ByVal rows As Variant, ByVal header As Variant)
    Dim index As Long, column As Long, source As String, id As String, key As String
    Dim members As Collection, text As String, stamp As String, keys As Variant
    Dim fieldIds As Variant, kindColumn As Long, kind As String, bomColumn As Long, versionColumn As Long
    Dim bom As String, version As String
    mRows = rows
    Set mGroups = CreateObject("Scripting.Dictionary"): mGroups.CompareMode = vbBinaryCompare
    Set mSearch = CreateObject("Scripting.Dictionary"): mSearch.CompareMode = vbBinaryCompare
    Set mTimes = CreateObject("Scripting.Dictionary"): mTimes.CompareMode = vbBinaryCompare
    Set mKinds = CreateObject("Scripting.Dictionary"): mKinds.CompareMode = vbBinaryCompare
    Set mFields = CreateObject("Scripting.Dictionary"): mFields.CompareMode = vbBinaryCompare
    Set mMatches = New Collection
    Erase mOrdered
    Erase mRowGroups
    If IsEmpty(rows) Then Exit Sub
    ReDim mRowGroups(LBound(rows, 1) To UBound(rows, 1))
    If UBound(header) = 11 Then
        If CStr(header(4)) = "EVENTS1" Then
            fieldIds = Split(CStr(header(9)), ",")
            For column = 0 To UBound(fieldIds)
                mFields.Add CStr(fieldIds(column)), 19 + column
                Select Case CStr(fieldIds(column))
                    Case "SourceKind": kindColumn = 19 + column
                    Case "BomId": bomColumn = 19 + column
                    Case "BomVersion": versionColumn = 19 + column
                End Select
            Next column
        End If
    End If
    For index = LBound(rows, 1) To UBound(rows, 1)
        source = CStr(rows(index, 18)): id = CStr(rows(index, 11))
        kind = "": bom = "": version = ""
        If kindColumn > 0 Then
            kind = CStr(rows(index, kindColumn))
        ElseIf source = "Inventory" Then
            kind = "Business event"
        End If
        If source = "ShippingBOM" And kind = "Current state" Then
            If bomColumn > 0 And versionColumn > 0 Then
                bom = CStr(rows(index, bomColumn)): version = CStr(rows(index, versionColumn))
            End If
        End If
        If source = "" Or id = "" Then
            key = Chr$(0) & CStr(index)
            If bom <> "" And version <> "" Then key = Chr$(1) & CStr(Len(bom)) & ":" & bom & CStr(Len(version)) & ":" & version
        Else
            key = CStr(Len(source)) & ":" & source & CStr(Len(id)) & ":" & id
        End If
        If Not mGroups.Exists(key) Then
            Set members = New Collection
            mGroups.Add key, members
            mSearch.Add key, "": mTimes.Add key, ""
            mKinds.Add key, kind
        End If
        Set members = mGroups(key): members.Add index
        mRowGroups(index) = key
        stamp = RecordedStamp(CStr(rows(index, 1)))
        If stamp <> "" And (mTimes(key) = "" Or StrComp(stamp, mTimes(key), vbBinaryCompare) < 0) Then mTimes(key) = stamp
        text = ""
        For column = LBound(rows, 2) To UBound(rows, 2)
            text = text & Chr$(30) & CStr(rows(index, column))
        Next column
        mSearch(key) = CStr(mSearch(key)) & text
    Next index
    keys = mGroups.Keys
    ReDim mOrdered(0 To UBound(keys))
    For index = 0 To UBound(keys): mOrdered(index) = CStr(keys(index)): Next index
    SortGroups 0, UBound(mOrdered)
End Sub

Public Function SourceMembers(ByVal sourceIndex As Long) As Collection
    Dim result As New Collection, member As Variant
    If Not IsEmpty(mRows) Then
        If sourceIndex >= LBound(mRowGroups) And sourceIndex <= UBound(mRowGroups) Then
            For Each member In mGroups(mRowGroups(sourceIndex)): result.Add CLng(member): Next member
        End If
    End If
    Set SourceMembers = result
End Function

Public Sub Match(ByVal searchText As String, ByVal restrictDates As Boolean, ByVal earliest As Date, ByVal latest As Date, ByVal verifiedUtcNow As Date, _
                 ByVal family As String, ByVal source As String, ByVal outcome As String, ByVal allPublished As Boolean)
    Dim key As Variant, include As Boolean, stamp As String, index As Long, firstDate As Date, lastDate As Date
    Set mMatches = New Collection
    If mGroups Is Nothing Then Exit Sub
    If mGroups.Count = 0 Then Exit Sub
    For index = 0 To UBound(mOrdered)
        key = mOrdered(index)
        include = (searchText = "" Or InStr(1, CStr(mSearch(key)), searchText, vbTextCompare) > 0)
        include = include And MatchesFacet(CStr(key), "EventFamily", family) And MatchesFacet(CStr(key), "Source", source) And MatchesFacet(CStr(key), "Outcome", outcome)
        ' The default operator view excludes internal reservations, retaining
        ' those published lines in the loaded model for the all-events view.
        If Not allPublished And CStr(mRows(CLng(mGroups(key)(1)), 13)) = "SHIP_RESERVE" Then include = False
        If include And restrictDates Then
            stamp = CStr(mTimes(key))
            firstDate = earliest: lastDate = latest
            If InStr(1, CStr(mRows(CLng(mGroups(key)(1)), 1)), " UTC", vbBinaryCompare) > 0 Then
                lastDate = verifiedUtcNow
                firstDate = verifiedUtcNow - (latest - earliest)
            End If
            If IsDate(Left$(stamp, 19)) Then
                include = (CDate(Left$(stamp, 19)) >= firstDate And CDate(Left$(stamp, 19)) <= lastDate)
            Else
                include = False
            End If
        End If
        If include Then mMatches.Add CStr(key)
    Next index
End Sub

Public Function FacetChoices(ByVal field As String) As Collection
    Dim choices As Object, result As New Collection, key As Variant, member As Variant, value As String
    Set choices = CreateObject("Scripting.Dictionary"): choices.CompareMode = vbBinaryCompare
    If Not mGroups Is Nothing Then
        For Each key In mGroups.Keys
            For Each member In mGroups(key)
                value = FacetValue(CLng(member), field)
                If value <> "" And Not choices.Exists(value) Then choices.Add value, True
            Next member
            If MatchesFacet(CStr(key), field, "Unavailable") And Not choices.Exists("Unavailable") Then choices.Add "Unavailable", True
        Next key
    End If
    For Each key In choices.Keys: result.Add CStr(key): Next key
    Set FacetChoices = result
End Function

Private Function MatchesFacet(ByVal key As String, ByVal field As String, ByVal selected As String) As Boolean
    Dim member As Variant, value As String, supplied As Boolean
    If selected = "" Then MatchesFacet = True: Exit Function
    For Each member In mGroups(key)
        value = FacetValue(CLng(member), field)
        If value <> "" Then supplied = True
        If StrComp(value, selected, vbBinaryCompare) = 0 Then MatchesFacet = True: Exit Function
    Next member
    MatchesFacet = (selected = "Unavailable" And Not supplied)
End Function

Private Function FacetValue(ByVal index As Long, ByVal field As String) As String
    If field = "Source" Then
        FacetValue = CStr(mRows(index, 18))
    ElseIf mFields.Exists(field) Then
        FacetValue = CStr(mRows(index, CLng(mFields(field))))
    End If
End Function

Public Property Get AvailableCount() As Long
    If Not mGroups Is Nothing Then AvailableCount = mGroups.Count
End Property

Public Property Get MatchingCount() As Long
    If Not mMatches Is Nothing Then MatchingCount = mMatches.Count
End Property

Public Function SourceIndex(ByVal matchIndex As Long) As Long
    SourceIndex = CLng(mGroups(CStr(mMatches(matchIndex)))(1))
End Function

Public Function SummaryValue(ByVal matchIndex As Long, ByVal column As Long) As String
    Dim key As String, members As Collection, index As Variant, value As String
    key = CStr(mMatches(matchIndex)): Set members = mGroups(key)
    SummaryValue = CStr(mRows(CLng(members(1)), column))
    If column = 1 Then
        If CStr(mTimes(key)) <> "" Then
            SummaryValue = Left$(CStr(mTimes(key)), 19)
            If InStr(1, CStr(mRows(CLng(members(1)), 1)), " UTC", vbBinaryCompare) > 0 Then SummaryValue = SummaryValue & " UTC"
        End If
        Exit Function
    End If
    For Each index In members
        value = CStr(mRows(CLng(index), column))
        If StrComp(value, SummaryValue, vbBinaryCompare) <> 0 Then SummaryValue = "Multiple": Exit Function
    Next index
    If column = 5 And members.Count > 1 And SummaryValue <> "" Then SummaryValue = "Multiple"
End Function

Private Function RecordedStamp(ByVal value As String) As String
    value = Replace$(Replace$(value, " UTC", ""), "T", " ")
    If IsDate(Left$(value, 19)) Then
        RecordedStamp = Format$(CDate(Left$(value, 19)), "yyyy-mm-dd hh:nn:ss")
        If Len(value) >= 23 Then
            If Mid$(value, 20, 4) Like ".###" Then RecordedStamp = RecordedStamp & Mid$(value, 20, 4)
        End If
    End If
End Function

Private Function Before(ByVal first As String, ByVal second As String) As Boolean
    Dim comparison As Long
    comparison = StrComp(CStr(mTimes(first)), CStr(mTimes(second)), vbBinaryCompare)
    If comparison <> 0 Then Before = (comparison > 0): Exit Function
    comparison = StrComp(CStr(mKinds(first)), CStr(mKinds(second)), vbBinaryCompare)
    If comparison <> 0 Then Before = (comparison < 0): Exit Function
    comparison = StrComp(CStr(mRows(CLng(mGroups(first)(1)), 11)), CStr(mRows(CLng(mGroups(second)(1)), 11)), vbBinaryCompare)
    If comparison <> 0 Then Before = (comparison < 0): Exit Function
    comparison = StrComp(CStr(mRows(CLng(mGroups(first)(1)), 18)), CStr(mRows(CLng(mGroups(second)(1)), 18)), vbBinaryCompare)
    If comparison <> 0 Then Before = (comparison < 0): Exit Function
    Before = (CLng(mGroups(first)(1)) < CLng(mGroups(second)(1)))
End Function

Private Sub SortGroups(ByVal first As Long, ByVal last As Long)
    Dim left As Long, right As Long, pivot As String, temporary As String
    left = first: right = last: pivot = mOrdered((first + last) \ 2)
    Do While left <= right
        Do While Before(mOrdered(left), pivot): left = left + 1: Loop
        Do While Before(pivot, mOrdered(right)): right = right - 1: Loop
        If left <= right Then
            temporary = mOrdered(left): mOrdered(left) = mOrdered(right): mOrdered(right) = temporary
            left = left + 1: right = right - 1
        End If
    Loop
    If first < right Then SortGroups first, right
    If left < last Then SortGroups left, last
End Sub
