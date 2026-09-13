VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "cEventDetailController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mForm As frmEventDetail
Private mData As Variant
Private mSelected As Collection
Private mContext As String
Private mWarehouse As String
Private mLoaded As String
Private mProfile As String
Private mProfileVersion As Long
Private mProfileStatus As String
Private mStale As Boolean

Public Sub LoadProjection(ByVal rows As Variant, ByVal header As Variant, ByVal context As String)
    Dim version As Long, profile As String, report As String, oldId As String, oldSource As String, i As Long
    If Not mSelected Is Nothing Then
        If mSelected.Count > 0 And Not IsEmpty(mData) Then
            oldId = CStr(mData(mSelected(1), 11)): oldSource = CStr(mData(mSelected(1), 18))
        End If
    End If
    modEventDetailSettings.ReadViewer context, version, profile, report
    If context = "" Or context <> modActivity.CaptureContext() Or profile = "" Then Invalidate: Exit Sub
    mData = rows: mContext = context: mWarehouse = CStr(header(1))
    mProfile = profile: mProfileVersion = version: mProfileStatus = report: mStale = False
    mLoaded = "Unavailable"
    If UBound(header) >= 5 Then mLoaded = Replace$(Replace$(CStr(header(5)), "T", " "), "Z", " UTC")
    Set mSelected = Nothing
    If mForm Is Nothing Then Exit Sub
    If oldId <> "" And Not IsEmpty(mData) Then
        For i = LBound(mData, 1) To UBound(mData, 1)
            If CStr(mData(i, 11)) = oldId And CStr(mData(i, 18)) = oldSource Then ShowLine i: Exit Sub
        Next i
    End If
    mForm.ClearContent "Unavailable. The selected event is not in this loaded projection."
End Sub

Public Sub ShowLine(ByVal sourceIndex As Long)
    Dim i As Long, id As String, source As String
    If Not ContextValid() Or IsEmpty(mData) Then Exit Sub
    If sourceIndex < LBound(mData, 1) Or sourceIndex > UBound(mData, 1) Then Exit Sub
    id = CStr(mData(sourceIndex, 11)): source = CStr(mData(sourceIndex, 18))
    Set mSelected = New Collection
    If id = "" Or source = "" Then
        mSelected.Add sourceIndex
    Else
        For i = LBound(mData, 1) To UBound(mData, 1)
            If StrComp(CStr(mData(i, 11)), id, vbBinaryCompare) = 0 And StrComp(CStr(mData(i, 18)), source, vbBinaryCompare) = 0 Then mSelected.Add i
        Next i
    End If
    If mForm Is Nothing Then Set mForm = New frmEventDetail
    mForm.Bind Me, "Detail profile " & CStr(mProfileVersion) & ". " & mProfileStatus
    If Not mForm.Visible Then mForm.Show vbModeless
End Sub

Public Function ContextValid() As Boolean
    ContextValid = (mContext <> "" And mContext = modActivity.CaptureContext())
    If Not ContextValid Then Invalidate
End Function

Public Sub Invalidate()
    If mContext = "" And IsEmpty(mData) And mSelected Is Nothing Then Exit Sub
    mData = Empty: Set mSelected = Nothing
    mContext = "": mProfile = "": mLoaded = "": mWarehouse = ""
    If Not mForm Is Nothing Then mForm.ClearContent "Unavailable. The invSys session or warehouse changed. Reopen Viewer."
End Sub

Public Sub MarkStale()
    mStale = True
    If Not mForm Is Nothing Then mForm.RefreshFields
End Sub

Public Function Keys() As Collection
    Dim result As New Collection, i As Variant
    If ContextValid() Then
        If Not mSelected Is Nothing Then
            For Each i In mSelected
                result.Add Available(CStr(mData(CLng(i), 12)))
            Next i
        End If
    End If
    Set Keys = result
End Function

Public Function Fields(ByVal selection As Long) As Variant
    Dim rows As String, definition As Variant, definitionValues As Variant, result() As Variant, count As Long, row As Long
    If Not ContextValid() Then Exit Function
    If mSelected Is Nothing Then Exit Function
    If selection < 1 Or selection > mSelected.Count Then Exit Function
    row = CLng(mSelected(selection))
    rows = modEventDetailSettings.FieldRows(mProfile, Family(CStr(mData(row, 13)), CStr(mData(row, 2))))
    If rows = "" Then Exit Function
    ReDim result(0 To UBound(Split(rows, vbLf)), 0 To 1)
    For Each definition In Split(rows, vbLf)
        definitionValues = Split(CStr(definition), vbTab)
        If definitionValues(2) = "True" Then
            result(count, 0) = definitionValues(1)
            result(count, 1) = Available(FieldValue(row, CStr(definitionValues(0))))
            count = count + 1
        End If
    Next definition
    Fields = result
End Function

Private Function FieldValue(ByVal row As Long, ByVal field As String) As String
    Select Case field
        Case "SourceId": FieldValue = CStr(mData(row, 11))
        Case "EventCode": FieldValue = CStr(mData(row, 13))
        Case "DataEffect": FieldValue = "Unknown"
        Case "OwnerId": If mData(row, 18) = "Inventory" Then FieldValue = "Inventory Domain"
        Case "EventFamily": FieldValue = Family(CStr(mData(row, 13)), CStr(mData(row, 2)))
        Case "EventType": FieldValue = CStr(mData(row, 2))
        Case "SourceKind"
            If mData(row, 18) = "Inventory" Then
                FieldValue = "Business event / Inventory"
            ElseIf mData(row, 2) = "BOX_DESIGNED" Or mData(row, 2) = "SHIP_HELD" Then
                FieldValue = "Current state"
            End If
        Case "WarehouseId": FieldValue = mWarehouse
        Case "OccurredAt": FieldValue = CStr(mData(row, 16))
        Case "AppliedAt": FieldValue = CStr(mData(row, 17))
        Case "TimeProvenance": FieldValue = "Recorded time (zone unavailable). Ordering and date filters across sources are approximate."
        Case "Coverage": FieldValue = CStr(mSelected.Count) & " contributing line(s) in this loaded projection. Complete publication coverage unavailable."
        Case "LoadedAt": FieldValue = mLoaded
        Case "Freshness": FieldValue = IIf(mStale, "Stale. Refresh failed; showing retained data.", "Publication freshness unavailable.")
        Case "Explanation"
            If mData(row, 2) = "BOX_DESIGNED" Or mData(row, 2) = "SHIP_HELD" Then
                FieldValue = "Current-state supplement. Historical control usage and completed workflow outcomes are unavailable."
            Else
                FieldValue = "Published source lines do not establish a completed workflow conclusion without owning outcome evidence."
            End If
        Case "NextStep": FieldValue = "Inspect the owning workflow for completion evidence."
        Case "System_Key": FieldValue = CStr(mData(row, 12))
        Case "Reference": FieldValue = CStr(mData(row, 3))
        Case "ItemCode": FieldValue = CStr(mData(row, 14))
        Case "ItemName": FieldValue = CStr(mData(row, 4))
        Case "Quantity": FieldValue = CStr(mData(row, 5))
        Case "Uom": FieldValue = CStr(mData(row, 6))
        Case "Location": FieldValue = CStr(mData(row, 7))
        Case "Condition": FieldValue = CStr(mData(row, 8))
        Case "ActorId": FieldValue = CStr(mData(row, 9))
        Case "StationId": FieldValue = CStr(mData(row, 15))
    End Select
End Function

Private Function Family(ByVal eventType As String, Optional ByVal stateType As String = "") As String
    If eventType = "" Then
        If stateType = "BOX_DESIGNED" Then Family = "Boxing": Exit Function
        If stateType = "SHIP_HELD" Then Family = "Shipping": Exit Function
    End If
    Select Case UCase$(eventType)
        Case "RECEIVE", "RETURN", "DUMP": Family = "Receiving"
        Case "SHIP", "SHIP_RESERVE", "SHIP_RELEASE": Family = "Shipping"
        Case "BOX_BUILD", "BOX_UNBOX": Family = "Boxing"
        Case "PROD_CONSUME", "PROD_COMPLETE": Family = "Production"
        Case Else: Family = "Inventory"
    End Select
End Function

Private Function Available(ByVal value As String) As String
    Available = value
    If value = "" Then Available = "Unavailable"
End Function

Public Sub ReleaseForm(ByVal instance As frmEventDetail)
    If Not mForm Is Nothing Then
        If mForm Is instance Then Set mForm = Nothing
    End If
End Sub

Public Sub CloseDetail()
    Dim instance As frmEventDetail
    Set instance = mForm: Set mForm = Nothing
    If Not instance Is Nothing Then Unload instance
End Sub
