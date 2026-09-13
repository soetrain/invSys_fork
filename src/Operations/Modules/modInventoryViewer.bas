Attribute VB_Name = "modInventoryViewer"
Option Explicit

Private mInventoryViewer As frmInventoryViewer
Private mViewerGeneration As Long

Public Sub OpenInventoryViewer()
    Dim warehouseId As String

    If Not modAuth.IsSignedIn() Then
        MsgBox "Sign in to invSys before opening Inventory Viewer.", _
            vbInformation, "invSys Inventory Viewer"
        Exit Sub
    End If
    warehouseId = Trim$(modNasConnection.GetCurrentTargetWarehouseId())
    If warehouseId = "" Then
        MsgBox "Select a warehouse before opening Inventory Viewer.", _
            vbInformation, "invSys Inventory Viewer"
        Exit Sub
    End If

    If mInventoryViewer Is Nothing Then
        Set mInventoryViewer = New frmInventoryViewer
        mViewerGeneration = mViewerGeneration + 1
        mInventoryViewer.SetGeneration mViewerGeneration
    End If
    mInventoryViewer.SetWarehouse warehouseId
    mInventoryViewer.RefreshInventory
    If Not mInventoryViewer.Visible Then mInventoryViewer.Show vbModeless
End Sub

Public Function LoadInventoryViewerEvents(Optional ByVal publishedPayload As String = "") As String
    Dim corePayload As String
    Dim shippingPayload As String
    Dim coreLines As Variant
    Dim shippingLines As Variant
    Dim coreHeader As Variant
    Dim shippingHeader As Variant
    Dim resultText As String
    Dim lineIndex As Long
    Dim rowCount As Long

    corePayload = publishedPayload
    If Trim$(corePayload) = "" Then corePayload = modInventoryViewerData.LoadCurrentInventoryEventViewerData()
    LoadInventoryViewerEvents = "FAIL" & vbTab & "Published Events are unavailable. Try Refresh after publication is restored."
    If Trim$(corePayload) = "" Then Exit Function
    coreLines = Split(corePayload, vbCrLf)
    coreHeader = Split(CStr(coreLines(0)), vbTab)
    If UBound(coreHeader) < 3 Or StrComp(CStr(coreHeader(0)), "OK", vbTextCompare) <> 0 Then
        Exit Function
    End If
    shippingPayload = modTS_Shipments.LoadShippingViewerSupplementEvents()
    shippingLines = Split(shippingPayload, vbCrLf)
    shippingHeader = Split(CStr(shippingLines(0)), vbTab)

    resultText = CStr(coreLines(0))
    rowCount = CLng(Val(CStr(coreHeader(3))))
    For lineIndex = 1 To UBound(coreLines)
        If Trim$(CStr(coreLines(lineIndex))) <> "" Then resultText = resultText & vbCrLf & CStr(coreLines(lineIndex))
    Next lineIndex

    If UBound(shippingHeader) >= 1 And StrComp(CStr(shippingHeader(0)), "OK", vbTextCompare) = 0 Then
        If UBound(shippingHeader) >= 3 Then rowCount = rowCount + CLng(Val(CStr(shippingHeader(3))))
        For lineIndex = 1 To UBound(shippingLines)
            If Trim$(CStr(shippingLines(lineIndex))) <> "" Then resultText = resultText & vbCrLf & CStr(shippingLines(lineIndex))
        Next lineIndex
    End If
    coreLines = Split(resultText, vbCrLf)
    coreHeader = Split(CStr(coreLines(0)), vbTab)
    If UBound(coreHeader) >= 3 Then
        coreHeader(3) = CStr(rowCount)
        coreLines(0) = Join(coreHeader, vbTab)
        resultText = Join(coreLines, vbCrLf)
    End If
    LoadInventoryViewerEvents = resultText
End Function

Public Sub UnregisterInventoryViewer(ByVal formInstance As Object)
    On Error Resume Next
    If Not mInventoryViewer Is Nothing Then
        If formInstance Is Nothing Or mInventoryViewer Is formInstance Then Set mInventoryViewer = Nothing
    End If
    On Error GoTo 0
End Sub

Public Function RunInventoryViewerActionForTest() As String
    OpenInventoryViewer
    If mInventoryViewer Is Nothing Then
        RunInventoryViewerActionForTest = "FAIL|FormNotOpen"
    Else
        RunInventoryViewerActionForTest = mInventoryViewer.TestReport()
    End If
End Function

Public Function RunInventoryViewerFilterForTest(ByVal filterText As String) As String
    If mInventoryViewer Is Nothing Then OpenInventoryViewer
    If mInventoryViewer Is Nothing Then
        RunInventoryViewerFilterForTest = "FAIL|FormNotOpen"
    Else
        RunInventoryViewerFilterForTest = mInventoryViewer.TestApplySearch(filterText)
    End If
End Function

Public Function RunInventoryViewerEventsForTest(Optional ByVal rangeText As String = "") As String
    If mInventoryViewer Is Nothing Then OpenInventoryViewer
    If mInventoryViewer Is Nothing Then
        RunInventoryViewerEventsForTest = "FAIL|FormNotOpen"
    Else
        RunInventoryViewerEventsForTest = mInventoryViewer.TestEventsReport(rangeText)
    End If
End Function

Public Function ExportDeclaredListBoxToTable(ByVal listBoxName As String, _
                                             ByRef report As String) As Boolean
    Dim openForm As Object
    Dim sourceList As Object
    Dim matchedList As Object
    Dim sourceFormName As String
    Dim headers As Variant
    Dim matches As Long

    listBoxName = Trim$(listBoxName)
    If listBoxName = "" Then
        report = "Enter a ListBox name, for example lstInventory or lstRunPalette."
        Exit Function
    End If
    If StrComp(listBoxName, "lstInventory", vbTextCompare) = 0 And Not mInventoryViewer Is Nothing Then
        ExportDeclaredListBoxToTable = mInventoryViewer.ExportViewerListBoxToTable(listBoxName, report)
        Exit Function
    End If
    For Each openForm In VBA.UserForms
        On Error Resume Next
        Set sourceList = openForm.Controls(listBoxName)
        On Error GoTo 0
        If Not sourceList Is Nothing Then
            If TypeName(sourceList) = "ListBox" Then
                matches = matches + 1
                sourceFormName = openForm.Name
                Set matchedList = sourceList
            End If
        End If
        If matches > 1 Then Exit For
        Set sourceList = Nothing
    Next openForm
    If matches = 0 Then
        report = "No open declared ListBox named " & listBoxName & " was found."
        Exit Function
    End If
    Set sourceList = matchedList
    If matches > 1 Then
        report = "More than one open form contains " & listBoxName & ". Use Viewer lstInventory or close the duplicate source."
        Exit Function
    End If
    If InStr(1, sourceFormName, "Admin", vbTextCompare) > 0 Then
        If Not modRoleUiAccess.CanCurrentUserPerformCapabilityCached("ADMIN_MAINT", report) Then Exit Function
    End If
    If Not DeclaredListBoxHeaders(sourceFormName, listBoxName, headers) Then
        report = listBoxName & " is not an export-declared ListBox."
        Exit Function
    End If
    ExportDeclaredListBoxToTable = modListBoxTableExport.ExportVisibleListBoxToNewTable( _
        sourceList, headers, report)
End Function

Private Function DeclaredListBoxHeaders(ByVal formName As String, _
                                        ByVal listBoxName As String, _
                                        ByRef headers As Variant) As Boolean
    Select Case UCase$(Trim$(listBoxName))
        Case "LSTRUNPALETTE"
            headers = Array("", "", "Process / Ingredient", "", "Inventory Stock", "% Req", "Qty", "Stock / Requirement UOM", "Native / Requirement Available", "Location")
        Case "LSTMANAGERCHECK"
            headers = Array("Type", "Process / Requirement", "Source Process / Output", "System_Key", "Code", "Item", "UOM", "Committed / Used", "Remaining Balance")
        Case "LSTMANAGEROUTPUT"
            headers = Array("Process", "Output", "UOM", "Last Actual", "Batch", "Used Goods", "Process Total", "Recall", "System_Key")
        Case "LSTRUNINSTRUCTIONS"
            headers = Array("Step", "Instruction")
        Case "LSTRECEIVEITEMS"
            headers = Array("Item Code", "Item", "UOM", "Qty", "Location", "Lot", "Condition", "Vendor", "Description", "System_Key")
        Case "LSTSHIPPABLES"
            headers = Array("Box", "Alternative", "NAS Inv", "Projected Inv", "Locked", "UOM", "Location", "System_Key")
        Case "LSTSHIPMENTS", "LSTHOLD"
            headers = Array("Reference", "Box", "Qty", "UOM", "Area", "Locked", "System_Key", "Alternative", "Carrier", "", "", "")
        Case Else
            Exit Function
    End Select
    DeclaredListBoxHeaders = True
End Function

Public Function RunInventoryViewerListBoxTableActionForTest() As String
    Dim report As String
    Dim exported As Boolean

    If mInventoryViewer Is Nothing Then OpenInventoryViewer
    If mInventoryViewer Is Nothing Then
        RunInventoryViewerListBoxTableActionForTest = "FAIL|FormNotOpen"
        Exit Function
    End If
    exported = mInventoryViewer.TestListBoxTableAction(report)
    RunInventoryViewerListBoxTableActionForTest = "OK|ListBox->Table=True|Exported=" & _
        CStr(exported) & "|" & report
End Function

Public Sub CloseInventoryViewerForTest()
    If Not mInventoryViewer Is Nothing Then Unload mInventoryViewer
End Sub
