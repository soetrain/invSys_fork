Attribute VB_Name = "modInventoryItemFields"
Option Explicit
Option Private Module

Public Function IsNonCountedCustomFieldsAdmin(ByVal customFields As Object) As Boolean
    Dim trackQty As String
    Dim itemKind As String

    On Error Resume Next
    If Not customFields Is Nothing Then
        trackQty = UCase$(Trim$(CStr(customFields("TRACK_QTY"))))
        itemKind = UCase$(Trim$(CStr(customFields("ITEM_KIND"))))
    End If
    On Error GoTo 0

    IsNonCountedCustomFieldsAdmin = (trackQty = "FALSE" Or trackQty = "NO" Or trackQty = "0" _
                                     Or itemKind = "UTILITY" Or itemKind = "SERVICE" Or itemKind = "NON_COUNTED")
End Function

Public Function NonCountedItemKindAdmin(ByVal customFields As Object) As String
    Dim itemKind As String

    On Error Resume Next
    If Not customFields Is Nothing Then itemKind = UCase$(Trim$(CStr(customFields("ITEM_KIND"))))
    On Error GoTo 0

    Select Case itemKind
        Case "UTILITY", "SERVICE", "NON_COUNTED"
            NonCountedItemKindAdmin = itemKind
        Case Else
            NonCountedItemKindAdmin = "NON_COUNTED"
    End Select
End Function

Public Function InventoryWorksheetColumnIndex(ByVal lo As ListObject, _
                                               ByVal headerName As String) As Long
    Dim lc As ListColumn
    Dim normalized As String

    normalized = NormalizeInventoryWorksheetHeader(headerName)
    For Each lc In lo.ListColumns
        If NormalizeInventoryWorksheetHeader(CStr(lc.Name)) = normalized Then
            InventoryWorksheetColumnIndex = lc.Index
            Exit Function
        End If
    Next lc
End Function

Public Function NormalizeInventoryWorksheetHeader(ByVal headerName As String) As String
    headerName = UCase$(Trim$(headerName))
    headerName = Replace$(headerName, " ", "")
    headerName = Replace$(headerName, "_", "")
    headerName = Replace$(headerName, "-", "")
    headerName = Replace$(headerName, "/", "")
    headerName = Replace$(headerName, "(", "")
    headerName = Replace$(headerName, ")", "")
    NormalizeInventoryWorksheetHeader = headerName
End Function

Public Function InventoryWorksheetText(ByVal lo As ListObject, _
                                        ByVal rowIndex As Long, _
                                        ByVal headerName As String) As String
    Dim colIndex As Long

    colIndex = InventoryWorksheetColumnIndex(lo, headerName)
    If colIndex = 0 Or lo.DataBodyRange Is Nothing Then Exit Function
    InventoryWorksheetText = Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, colIndex).Value2))
End Function
