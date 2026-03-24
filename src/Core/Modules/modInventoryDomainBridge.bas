Attribute VB_Name = "modInventoryDomainBridge"
Option Explicit

Public Const CORE_APPLY_STATUS_APPLIED As String = "APPLIED"
Public Const CORE_APPLY_STATUS_SKIP_DUP As String = "SKIP_DUP"

Public Const CORE_EVENT_TYPE_RECEIVE As String = "RECEIVE"
Public Const CORE_EVENT_TYPE_SHIP As String = "SHIP"
Public Const CORE_EVENT_TYPE_PROD_CONSUME As String = "PROD_CONSUME"
Public Const CORE_EVENT_TYPE_PROD_COMPLETE As String = "PROD_COMPLETE"

Private Const INVENTORY_DOMAIN_ADDIN_NAME As String = "invSys.Inventory.Domain.xlam"

Public Function ResolveInventoryWorkbookBridge(Optional ByVal warehouseId As String = "", _
                                              Optional ByVal inventoryWb As Workbook = Nothing) As Workbook
    Dim result As Variant
    Dim report As String

    If Not inventoryWb Is Nothing Then
        Set ResolveInventoryWorkbookBridge = inventoryWb
        Exit Function
    End If

    Set ResolveInventoryWorkbookBridge = OpenOrCreateCanonicalInventoryWorkbookLocal(warehouseId, report)
    If Not ResolveInventoryWorkbookBridge Is Nothing Then Exit Function

    On Error GoTo FailResolve
    result = RunInventoryDomainMacro1("modInventoryBridgeApi.ResolveInventoryWorkbookBridgeResult", warehouseId)
    If IsObject(result) Then Set ResolveInventoryWorkbookBridge = result
    Exit Function

FailResolve:
    Set ResolveInventoryWorkbookBridge = Nothing
End Function

Public Function EnsureInventorySchemaBridge(Optional ByVal targetWb As Workbook = Nothing, _
                                           Optional ByRef report As String = "") As Boolean
    On Error GoTo FailEnsure
    If Not targetWb Is Nothing Then
        EnsureInventorySchemaBridge = EnsureInventorySchemaLocal(targetWb, report)
        Exit Function
    End If

    EnsureInventorySchemaBridge = CBool(RunInventoryDomainMacro1("modInventoryBridgeApi.EnsureInventorySchemaBridgeSuccess", targetWb))
    report = CStr(RunInventoryDomainMacro1("modInventoryBridgeApi.EnsureInventorySchemaBridgeReport", targetWb))
    Exit Function

FailEnsure:
    report = Err.Description
    EnsureInventorySchemaBridge = False
End Function

Public Function ApplyInventoryEventBridge(ByVal evt As Object, _
                                         Optional ByVal inventoryWb As Workbook = Nothing, _
                                         Optional ByVal runId As String = "", _
                                         Optional ByRef statusOut As String = "", _
                                         Optional ByRef errorCode As String = "", _
                                         Optional ByRef errorMessage As String = "") As Boolean
    Dim result As Variant

    On Error GoTo FailApply
    If inventoryWb Is Nothing Then Set inventoryWb = ResolveInventoryWorkbookBridge(GetBridgeString(evt, "WarehouseId"))

    result = RunInventoryDomainMacro3("modInventoryBridgeApi.ApplyEventBridgeResult", evt, inventoryWb, runId)
    If IsObject(result) Then
        ApplyInventoryEventBridge = GetBridgeBool(result, "Success")
        statusOut = GetBridgeString(result, "StatusOut")
        errorCode = GetBridgeString(result, "ErrorCode")
        errorMessage = GetBridgeString(result, "ErrorMessage")
    End If
    Exit Function

FailApply:
    errorCode = "INVENTORY_DOMAIN_CALL_FAILED"
    errorMessage = Err.Description
    ApplyInventoryEventBridge = False
End Function

Public Function RemoveLastBulkLogEntriesBridge(ByVal countToRemove As Long) As Collection
    Dim result As Variant

    On Error GoTo FailRemove
    result = RunInventoryDomainMacro1("modInventoryBridgeApi.RemoveLastBulkLogEntriesBridgeResult", countToRemove)
    If IsObject(result) Then Set RemoveLastBulkLogEntriesBridge = result
    Exit Function

FailRemove:
    Set RemoveLastBulkLogEntriesBridge = New Collection
End Function

Public Sub ReAddBulkLogEntriesBridge(ByVal logDataCollection As Collection)
    On Error Resume Next
    Call RunInventoryDomainMacro1("modInventoryBridgeApi.ReAddBulkLogEntriesBridgeResult", logDataCollection)
    On Error GoTo 0
End Sub

Private Function RunInventoryDomainMacro0(ByVal macroName As String) As Variant
    RunInventoryDomainMacro0 = Application.Run(ResolveInventoryDomainMacroName(macroName))
End Function

Private Function RunInventoryDomainMacro1(ByVal macroName As String, ByVal arg0 As Variant) As Variant
    RunInventoryDomainMacro1 = Application.Run(ResolveInventoryDomainMacroName(macroName), arg0)
End Function

Private Function RunInventoryDomainMacro2(ByVal macroName As String, ByVal arg0 As Variant, ByVal arg1 As Variant) As Variant
    RunInventoryDomainMacro2 = Application.Run(ResolveInventoryDomainMacroName(macroName), arg0, arg1)
End Function

Private Function RunInventoryDomainMacro3(ByVal macroName As String, ByVal arg0 As Variant, ByVal arg1 As Variant, ByVal arg2 As Variant) As Variant
    RunInventoryDomainMacro3 = Application.Run(ResolveInventoryDomainMacroName(macroName), arg0, arg1, arg2)
End Function

Private Function ResolveInventoryDomainMacroName(ByVal macroName As String) As String
    Dim hostName As String

    hostName = FindInventoryDomainMacroHostName()
    If hostName = "" Then
        Err.Raise vbObjectError + 2601, "modInventoryDomainBridge.ResolveInventoryDomainMacroName", _
                  "Inventory Domain add-in is not open."
    End If

    ResolveInventoryDomainMacroName = "'" & hostName & "'!" & macroName
End Function

Private Function FindInventoryDomainAddin() As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.Name, INVENTORY_DOMAIN_ADDIN_NAME, vbTextCompare) = 0 Then
            Set FindInventoryDomainAddin = wb
            Exit Function
        End If
    Next wb

    For Each wb In Application.Workbooks
        If InStr(1, wb.Name, "Inventory.Domain", vbTextCompare) > 0 Then
            Set FindInventoryDomainAddin = wb
            Exit Function
        End If
    Next wb
End Function

Private Function FindInventoryDomainMacroHostName() As String
    Dim wb As Workbook
    Dim addin As AddIn

    For Each wb In Application.Workbooks
        If StrComp(wb.Name, INVENTORY_DOMAIN_ADDIN_NAME, vbTextCompare) = 0 Then
            FindInventoryDomainMacroHostName = wb.Name
            Exit Function
        End If
    Next wb

    For Each wb In Application.Workbooks
        If InStr(1, wb.Name, "Inventory.Domain", vbTextCompare) > 0 Then
            FindInventoryDomainMacroHostName = wb.Name
            Exit Function
        End If
    Next wb

    On Error Resume Next
    For Each addin In Application.AddIns
        If addin Is Nothing Then GoTo NextAddIn
        If Not addin.Installed Then GoTo NextAddIn
        If StrComp(addin.Name, INVENTORY_DOMAIN_ADDIN_NAME, vbTextCompare) = 0 Then
            FindInventoryDomainMacroHostName = addin.Name
            Exit Function
        End If
        If InStr(1, addin.Name, "Inventory.Domain", vbTextCompare) > 0 Then
            FindInventoryDomainMacroHostName = addin.Name
            Exit Function
        End If
NextAddIn:
    Next addin
    On Error GoTo 0
End Function

Private Function FindInventoryWorkbookLocal(ByVal warehouseId As String) As Workbook
    Dim wb As Workbook
    Dim targetPath As String

    targetPath = BuildCanonicalInventoryPathLocal(warehouseId)
    For Each wb In Application.Workbooks
        If targetPath <> "" Then
            If StrComp(wb.FullName, targetPath, vbTextCompare) = 0 Then
                Set FindInventoryWorkbookLocal = wb
                Exit Function
            End If
        End If

        If IsInventoryWorkbookNameLocal(wb.Name, warehouseId) Then
            Set FindInventoryWorkbookLocal = wb
            Exit Function
        End If
    Next wb
End Function

Private Function IsInventoryWorkbookNameLocal(ByVal wbName As String, ByVal warehouseId As String) As Boolean
    Dim n As String

    n = LCase$(Trim$(wbName))
    If Not ((n Like "wh*.invsys.data.inventory.xlsb") Or _
            (n Like "wh*.invsys.data.inventory.xlsx") Or _
            (n Like "wh*.invsys.data.inventory.xlsm")) Then Exit Function

    If Trim$(warehouseId) = "" Then
        IsInventoryWorkbookNameLocal = True
    Else
        IsInventoryWorkbookNameLocal = (InStr(1, wbName, warehouseId, vbTextCompare) > 0)
    End If
End Function

Private Function WorkbookHasListObjectLocal(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    WorkbookHasListObjectLocal = Not (FindListObjectByNameLocal(wb, tableName) Is Nothing)
End Function

Private Function FindListObjectByNameLocal(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    On Error Resume Next
    For Each ws In wb.Worksheets
        Set FindListObjectByNameLocal = ws.ListObjects(tableName)
        If Not FindListObjectByNameLocal Is Nothing Then Exit Function
    Next ws
    On Error GoTo 0
End Function

Private Function EnsureInventorySchemaLocal(ByVal wb As Workbook, ByRef report As String) As Boolean
    On Error GoTo FailEnsure
    Dim issues As Collection

    Set issues = New Collection
    EnsureTableWithHeadersLocal wb, "InventoryLog", "tblInventoryLog", _
        Array("EventID", "UndoOfEventId", "AppliedSeq", "EventType", "OccurredAtUTC", "AppliedAtUTC", _
              "WarehouseId", "StationId", "UserId", "SKU", "QtyDelta", "Location", "Note"), issues
    EnsureTableWithHeadersLocal wb, "AppliedEvents", "tblAppliedEvents", _
        Array("EventID", "UndoOfEventId", "AppliedSeq", "AppliedAtUTC", "RunId", "SourceInbox", "Status"), issues
    EnsureTableWithHeadersLocal wb, "Locks", "tblLocks", _
        Array("LockName", "OwnerStationId", "OwnerUserId", "RunId", "AcquiredAtUTC", "ExpiresAtUTC", "HeartbeatAtUTC", "Status"), issues

    report = JoinIssuesLocal(issues)
    EnsureInventorySchemaLocal = True
    Exit Function

FailEnsure:
    report = "EnsureInventorySchemaLocal failed: " & Err.Description
End Function

Private Sub EnsureTableWithHeadersLocal(ByVal wb As Workbook, ByVal sheetName As String, ByVal tableName As String, ByVal headers As Variant, ByVal issues As Collection)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim i As Long
    Dim startCell As Range
    Dim tableRange As Range

    Set ws = EnsureWorksheetLocal(wb, sheetName)
    EnsureWorksheetEditableLocal ws
    Set lo = FindListObjectByNameLocal(wb, tableName)

    If lo Is Nothing Then
        Set startCell = GetNextTableStartCellLocal(ws)
        For i = LBound(headers) To UBound(headers)
            startCell.Offset(0, i - LBound(headers)).Value = headers(i)
        Next i

        Set tableRange = ws.Range(startCell, startCell.Offset(1, UBound(headers) - LBound(headers)))
        Set lo = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
        lo.Name = tableName
        issues.Add tableName & " created"
    End If

    For i = LBound(headers) To UBound(headers)
        EnsureListColumnLocal lo, CStr(headers(i)), issues
    Next i

    RemoveBlankSeedRowLocal lo
End Sub

Private Function EnsureWorksheetLocal(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureWorksheetLocal = wb.Worksheets(sheetName)
    On Error GoTo 0

    If EnsureWorksheetLocal Is Nothing Then
        Set EnsureWorksheetLocal = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureWorksheetLocal.Name = sheetName
    End If
End Function

Private Sub EnsureWorksheetEditableLocal(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    If Not ws.ProtectContents Then Exit Sub

    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    If ws.ProtectContents Then
        Err.Raise vbObjectError + 2603, "modInventoryDomainBridge.EnsureWorksheetEditableLocal", _
                  "Worksheet '" & ws.Name & "' is protected and could not be unprotected."
    End If
End Sub

Private Function GetNextTableStartCellLocal(ByVal ws As Worksheet) As Range
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        Set GetNextTableStartCellLocal = ws.Range("A1")
    Else
        Set GetNextTableStartCellLocal = ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(2, 0)
    End If
End Function

Private Sub EnsureListColumnLocal(ByVal lo As ListObject, ByVal columnName As String, ByVal issues As Collection)
    If GetColumnIndexLocal(lo, columnName) > 0 Then Exit Sub

    lo.ListColumns.Add lo.ListColumns.Count + 1
    lo.ListColumns(lo.ListColumns.Count).Name = columnName
    issues.Add lo.Name & "." & columnName & " created"
End Sub

Private Function GetColumnIndexLocal(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long

    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            GetColumnIndexLocal = i
            Exit Function
        End If
    Next i
End Function

Private Sub RemoveBlankSeedRowLocal(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If lo.ListRows.Count <> 1 Then Exit Sub
    If Not TableRowIsBlankLocal(lo, 1) Then Exit Sub
    lo.ListRows(1).Delete
End Sub

Private Function TableRowIsBlankLocal(ByVal lo As ListObject, ByVal rowIndex As Long) As Boolean
    Dim c As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If rowIndex <= 0 Or rowIndex > lo.ListRows.Count Then Exit Function

    TableRowIsBlankLocal = True
    For c = 1 To lo.ListColumns.Count
        If SafeTrim(lo.DataBodyRange.Cells(rowIndex, c).Value) <> "" Then
            TableRowIsBlankLocal = False
            Exit Function
        End If
    Next c
End Function

Private Function JoinIssuesLocal(ByVal issues As Collection) As String
    Dim issue As Variant

    For Each issue In issues
        If Len(JoinIssuesLocal) > 0 Then JoinIssuesLocal = JoinIssuesLocal & "; "
        JoinIssuesLocal = JoinIssuesLocal & CStr(issue)
    Next issue
End Function

Private Function GetBridgeString(ByVal payload As Object, ByVal key As String) As String
    On Error Resume Next
    If Not payload Is Nothing Then
        If payload.Exists(key) Then GetBridgeString = CStr(payload(key))
    End If
    On Error GoTo 0
End Function

Private Function GetBridgeBool(ByVal payload As Object, ByVal key As String) As Boolean
    On Error Resume Next
    If Not payload Is Nothing Then
        If payload.Exists(key) Then GetBridgeBool = CBool(payload(key))
    End If
    On Error GoTo 0
End Function

Private Function OpenOrCreateCanonicalInventoryWorkbookLocal(ByVal warehouseId As String, ByRef report As String) As Workbook
    On Error GoTo FailOpen

    Dim targetPath As String
    Dim wb As Workbook
    Dim prevEvents As Boolean
    Dim eventsSuppressed As Boolean
    Dim wasCreated As Boolean

    targetPath = BuildCanonicalInventoryPathLocal(warehouseId)
    If targetPath = "" Then Exit Function

    Set wb = FindOpenWorkbookByFullNameLocal(targetPath)
    If wb Is Nothing Then
        EnsureFolderRecursiveLocal GetParentFolderLocal(targetPath)
        If Len(Dir$(targetPath)) > 0 Then
            Set wb = Application.Workbooks.Open(targetPath)
        Else
            prevEvents = Application.EnableEvents
            Application.EnableEvents = False
            eventsSuppressed = True
            Set wb = Application.Workbooks.Add(xlWBATWorksheet)
            wb.SaveAs Filename:=targetPath, FileFormat:=50
            wasCreated = True
            Application.EnableEvents = prevEvents
            eventsSuppressed = False
        End If
    End If

    If Not EnsureInventorySchemaLocal(wb, report) Then Exit Function
    If wasCreated Then wb.Save
    Set OpenOrCreateCanonicalInventoryWorkbookLocal = wb
    Exit Function

FailOpen:
    On Error Resume Next
    If eventsSuppressed Then Application.EnableEvents = prevEvents
    On Error GoTo 0
    report = "Inventory workbook open/create failed: " & Err.Description
End Function

Private Function BuildCanonicalInventoryPathLocal(ByVal warehouseId As String) As String
    Dim resolvedWh As String
    Dim rootPath As String

    resolvedWh = Trim$(warehouseId)
    If resolvedWh = "" Then resolvedWh = SafeTrim(modConfig.GetString("WarehouseId", "WH1"))
    If resolvedWh = "" Then resolvedWh = "WH1"

    rootPath = SafeTrim(GetCoreDataRootOverride())
    If rootPath = "" Then rootPath = SafeTrim(modConfig.GetString("PathDataRoot", ""))
    If rootPath = "" Then rootPath = DefaultInventoryRootLocal(resolvedWh)

    BuildCanonicalInventoryPathLocal = NormalizeFolderPathLocal(rootPath) & resolvedWh & ".invSys.Data.Inventory.xlsb"
End Function

Private Function DefaultInventoryRootLocal(ByVal warehouseId As String) As String
    DefaultInventoryRootLocal = "C:\invSys\" & warehouseId & "\"
End Function

Private Function FindOpenWorkbookByFullNameLocal(ByVal fullNameIn As String) As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullNameIn, vbTextCompare) = 0 Then
            Set FindOpenWorkbookByFullNameLocal = wb
            Exit Function
        End If
    Next wb
End Function

Private Function NormalizeFolderPathLocal(ByVal folderPath As String) As String
    folderPath = Trim$(folderPath)
    If folderPath = "" Then Exit Function
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    NormalizeFolderPathLocal = folderPath
End Function

Private Function GetParentFolderLocal(ByVal fullPath As String) As String
    Dim lastSlash As Long

    lastSlash = InStrRev(fullPath, "\")
    If lastSlash > 0 Then GetParentFolderLocal = Left$(fullPath, lastSlash - 1)
End Function

Private Sub EnsureFolderRecursiveLocal(ByVal folderPath As String)
    Dim parentPath As String

    folderPath = Trim$(folderPath)
    If folderPath = "" Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) > 0 Then Exit Sub

    parentPath = GetParentFolderLocal(folderPath)
    If parentPath <> "" And Len(Dir$(parentPath, vbDirectory)) = 0 Then EnsureFolderRecursiveLocal parentPath

    On Error Resume Next
    MkDir folderPath
    On Error GoTo 0
End Sub

Private Function SafeTrim(ByVal valueIn As Variant) As String
    On Error Resume Next
    SafeTrim = Trim$(CStr(valueIn))
    On Error GoTo 0
End Function
