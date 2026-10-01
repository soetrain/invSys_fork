Attribute VB_Name = "modProductionCheckInActions"
Option Explicit
Option Private Module

' Check In correctness precedes its separate activity observation contract.
Public Sub Execute(ByVal owner As frmProduction, ByVal context As String, _
                   ByVal operatorBook As Workbook, ByRef loading As Boolean, ByRef busy As Boolean)
    Dim priorLoading As Boolean, number As Long, source As String, description As String
    If loading Or busy Then Exit Sub
    If Not modProductionRunBinding.RequireCurrentContext(owner, context, operatorBook) Then Exit Sub
    On Error GoTo Failed
    priorLoading = loading: busy = True
    owner.CheckInProductionRun
Done:
    loading = priorLoading: busy = False
    If number <> 0 Then
        On Error GoTo 0
        Err.Raise number, source, description
    End If
    Exit Sub
Failed:
    number = Err.Number: source = Err.Source: description = Err.Description
    Resume Done
End Sub

Public Function ClearManagedCheckRows(ByVal table As ListObject) As Boolean
    Dim headers As Variant, header As Variant, column As Long
    If table Is Nothing Then Exit Function
    headers = Array("System_Key", "ITEM_CODE", "ITEM", "UOM", "USED", "TOTAL INV")
    For Each header In headers
        If mProduction.ColumnIndex(table, CStr(header)) = 0 Then Exit Function
    Next header
    For Each header In headers
        column = mProduction.ColumnIndex(table, CStr(header))
        If Not table.ListColumns(column).DataBodyRange Is Nothing Then _
            table.ListColumns(column).DataBodyRange.ClearContents
    Next header
    ClearManagedCheckRows = True
End Function

' Retain the existing projection/Domain read paths, but never substitute an entity
' merely because its SKU, name or location matches the selected row.
Public Function ResolveSelectedKey(ByVal owner As frmProduction, ByVal inventoryTable As ListObject, _
                                   ByVal selectedKey As String, ByVal itemCode As String, _
                                   ByVal itemName As String, ByVal locationValue As String) As String
    Dim entities As Variant, cSystemKey As Long, cItemCode As Long, cItem As Long, cLocation As Long
    Dim r As Long, codeMatches As Boolean, nameMatches As Boolean, locationMatches As Boolean
    If selectedKey = "" Then Exit Function
    If Not inventoryTable Is Nothing Then
        If Not inventoryTable.DataBodyRange Is Nothing Then
            cSystemKey = mProduction.ColumnIndex(inventoryTable, "System_Key")
            cItemCode = mProduction.ColumnIndex(inventoryTable, "ITEM_CODE")
            If cItemCode = 0 Then cItemCode = mProduction.ColumnIndex(inventoryTable, "SKU")
            cItem = mProduction.ColumnIndex(inventoryTable, "ITEM")
            If cItem = 0 Then cItem = mProduction.ColumnIndex(inventoryTable, "ItemName")
            cLocation = mProduction.ColumnIndex(inventoryTable, "LOCATION")
            If cSystemKey > 0 Then
                For r = 1 To inventoryTable.ListRows.Count
                    codeMatches = False
                    If Trim$(itemCode) <> "" And cItemCode > 0 Then _
                        codeMatches = (StrComp(Trim$(owner.NzStr(inventoryTable.DataBodyRange.Cells(r, cItemCode).Value)), Trim$(itemCode), vbTextCompare) = 0)
                    nameMatches = False
                    If Trim$(itemName) <> "" And cItem > 0 Then _
                        nameMatches = (StrComp(Trim$(owner.NzStr(inventoryTable.DataBodyRange.Cells(r, cItem).Value)), Trim$(itemName), vbTextCompare) = 0)
                    locationMatches = True
                    If Trim$(locationValue) <> "" And cLocation > 0 Then _
                        locationMatches = (StrComp(Trim$(owner.NzStr(inventoryTable.DataBodyRange.Cells(r, cLocation).Value)), Trim$(locationValue), vbTextCompare) = 0)
                    If (codeMatches Or nameMatches) And locationMatches And _
                       StrComp(owner.NzStr(inventoryTable.DataBodyRange.Cells(r, cSystemKey).Value), selectedKey, vbBinaryCompare) = 0 Then
                        ResolveSelectedKey = selectedKey
                        Exit Function
                    End If
                Next r
            End If
        End If
    End If

    On Error GoTo CleanFail
    entities = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge(itemCode)
    If Not IsArray(entities) Then Exit Function
    For r = LBound(entities, 1) To UBound(entities, 1)
        codeMatches = (Trim$(itemCode) <> "" And _
            (StrComp(Trim$(owner.NzStr(entities(r, 3))), Trim$(itemCode), vbTextCompare) = 0 Or _
             StrComp(Trim$(owner.NzStr(entities(r, 2))), Trim$(itemCode), vbTextCompare) = 0))
        nameMatches = (Trim$(itemName) <> "" And _
            StrComp(Trim$(owner.NzStr(entities(r, 4))), Trim$(itemName), vbTextCompare) = 0)
        locationMatches = (Trim$(locationValue) = "" Or _
            StrComp(Trim$(owner.NzStr(entities(r, 7))), Trim$(locationValue), vbTextCompare) = 0)
        If (codeMatches Or nameMatches) And locationMatches And _
           StrComp(owner.NzStr(entities(r, 1)), selectedKey, vbBinaryCompare) = 0 Then
            ResolveSelectedKey = selectedKey
            Exit Function
        End If
    Next r
CleanFail:
End Function
