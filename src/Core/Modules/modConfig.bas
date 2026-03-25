Attribute VB_Name = "modConfig"
Option Explicit

Public Const ERR_CONFIG_NOT_LOADED As Long = vbObjectError + 7100
Public Const ERR_CONFIG_KEY_MISSING As Long = vbObjectError + 7101

Private mConfigCache As Object
Private mValidationIssues As Collection
Private mWarehouseId As String
Private mStationId As String
Private mResolvedWorkbook As String
Private mIsLoaded As Boolean

Public Function LoadConfig(Optional ByVal whId As String = "", Optional ByVal stId As String = "") As Boolean
    On Error GoTo FailLoad

    Dim wb As Workbook
    Dim loWh As ListObject
    Dim loSt As ListObject
    Dim whRow As Long
    Dim stRow As Long
    Dim whValues As Object
    Dim stValues As Object
    Dim defs() As ConfigKeyDef
    Dim defCount As Long
    Dim i As Long
    Dim rawVal As Variant
    Dim valOut As Variant
    Dim hasVal As Boolean

    InitializeState

    Set wb = ResolveConfigWorkbook(whId, stId)
    If wb Is Nothing Then
        AddValidationIssue "ERROR", "CONFIG_MISSING", "No open config workbook found."
        GoTo FailSoft
    End If
    mResolvedWorkbook = wb.Name

    If Not EnsureConfigSchema(wb, whId, stId) Then
        AddValidationIssue "ERROR", "CONFIG_SELF_HEAL_FAILED", "Failed to create/repair config tables."
        GoTo FailSoft
    End If

    Set loWh = FindListObjectByName(wb, "tblWarehouseConfig")
    Set loSt = FindListObjectByName(wb, "tblStationConfig")
    If loWh Is Nothing Then
        AddValidationIssue "ERROR", "CONFIG_TABLE_MISSING", "tblWarehouseConfig not found."
        GoTo FailSoft
    End If
    If loSt Is Nothing Then
        AddValidationIssue "ERROR", "CONFIG_TABLE_MISSING", "tblStationConfig not found."
        GoTo FailSoft
    End If

    EnsureTableHasRow loWh
    EnsureTableHasRow loSt

    whRow = ResolveWarehouseRow(loWh, whId, wb.Name)
    If whRow = 0 Then
        AddValidationIssue "ERROR", "CONFIG_WAREHOUSE_NOT_FOUND", "Warehouse row not found."
        GoTo FailSoft
    End If

    Set whValues = BuildRowDictionary(loWh, whRow)
    mWarehouseId = SafeTrim(GetDictionaryValue(whValues, "WarehouseId"))

    stRow = ResolveStationRow(loSt, stId, mWarehouseId)
    If stRow = 0 Then
        AddValidationIssue "ERROR", "CONFIG_STATION_NOT_FOUND", "Station row not found."
        GoTo FailSoft
    End If

    Set stValues = BuildRowDictionary(loSt, stRow)
    mStationId = SafeTrim(GetDictionaryValue(stValues, "StationId"))

    defCount = GetConfigSchema(defs)
    For i = 1 To defCount
        hasVal = False
        If UCase$(defs(i).Scope) = CONFIG_SCOPE_STATION Then
            If TryGetDictionaryValue(stValues, defs(i).Key, rawVal) And Not IsBlankValue(rawVal) Then
                hasVal = True
            ElseIf TryGetDictionaryValue(whValues, defs(i).Key, rawVal) And Not IsBlankValue(rawVal) Then
                hasVal = True
            End If
        Else
            If TryGetDictionaryValue(whValues, defs(i).Key, rawVal) And Not IsBlankValue(rawVal) Then
                hasVal = True
            End If
        End If

        If hasVal Then
            If TryCoerceValue(defs(i).DataType, rawVal, valOut) Then
                mConfigCache(defs(i).Key) = valOut
            Else
                HandleMalformedKey defs(i), rawVal
            End If
        Else
            HandleMissingKey defs(i)
        End If
    Next i

    If SafeTrim(GetString("WarehouseId", "")) = "" Then
        AddValidationIssue "ERROR", "CONFIG_KEY_MISSING", "WarehouseId is required."
    End If
    If SafeTrim(GetString("StationId", "")) = "" Then
        AddValidationIssue "ERROR", "CONFIG_KEY_MISSING", "StationId is required."
    End If

    mWarehouseId = GetString("WarehouseId", "")
    mStationId = GetString("StationId", "")

    If CountFatalIssues() > 0 Then
        GoTo FailSoft
    End If

    mIsLoaded = True
    LoadConfig = True
    Exit Function

FailSoft:
    mIsLoaded = False
    LoadConfig = False
    Exit Function

FailLoad:
    AddValidationIssue "ERROR", "CONFIG_LOAD_EXCEPTION", Err.Description
    Resume FailSoft
End Function

Public Function EnsureConfigSchema(Optional ByVal targetWb As Workbook = Nothing, _
                                   Optional ByVal warehouseId As String = "", _
                                   Optional ByVal stationId As String = "", _
                                   Optional ByRef report As String = "") As Boolean
    On Error GoTo FailEnsure

    Dim wb As Workbook

    If targetWb Is Nothing Then
        Set wb = ThisWorkbook
    Else
        Set wb = targetWb
    End If

    If Not EnsureConfigTables(wb) Then GoTo FailSoft
    SeedConfigDefaults wb, warehouseId, stationId
    FormatConfigSurface wb

    EnsureConfigSchema = True
    Exit Function

FailSoft:
    report = "EnsureConfigSchema failed."
    Exit Function

FailEnsure:
    report = "EnsureConfigSchema failed: " & Err.Description
End Function

Public Function Reload() As Boolean
    Reload = LoadConfig(mWarehouseId, mStationId)
End Function

Public Function IsLoaded() As Boolean
    IsLoaded = mIsLoaded
End Function

Public Function GetRequired(ByVal key As String) As Variant
    Dim v As Variant
    If Not mIsLoaded Then
        Err.Raise ERR_CONFIG_NOT_LOADED, "modConfig.GetRequired", "Config is not loaded."
    End If
    If Not TryGet(key, v) Or IsBlankValue(v) Then
        Err.Raise ERR_CONFIG_KEY_MISSING, "modConfig.GetRequired", "Missing required key: " & key
    End If
    GetRequired = v
End Function

Public Function TryGet(ByVal key As String, ByRef outVal As Variant) As Boolean
    If mConfigCache Is Nothing Then Exit Function
    If mConfigCache.Exists(key) Then
        outVal = mConfigCache(key)
        TryGet = True
    End If
End Function

Public Function Validate() As String
    Dim itm As Variant
    Dim parts() As String
    Dim lineOut As String

    If mValidationIssues Is Nothing Or mValidationIssues.Count = 0 Then Exit Function

    For Each itm In mValidationIssues
        parts = Split(CStr(itm), "|")
        If UBound(parts) >= 2 Then
            lineOut = parts(0) & " " & parts(1) & ": " & parts(2)
        Else
            lineOut = CStr(itm)
        End If

        If Len(Validate) > 0 Then Validate = Validate & "; "
        Validate = Validate & lineOut
    Next itm
End Function

Public Function GetWarehouseId() As String
    GetWarehouseId = mWarehouseId
End Function

Public Function GetStationId() As String
    GetStationId = mStationId
End Function

Public Function GetLong(ByVal key As String, ByVal defaultVal As Long) As Long
    Dim v As Variant
    If TryGet(key, v) Then
        If IsNumeric(v) Then
            GetLong = CLng(v)
            Exit Function
        End If
    End If
    GetLong = defaultVal
End Function

Public Function GetBool(ByVal key As String, ByVal defaultVal As Boolean) As Boolean
    Dim v As Variant
    Dim parsed As Variant
    If TryGet(key, v) Then
        If TryCoerceValue(CONFIG_TYPE_BOOLEAN, v, parsed) Then
            GetBool = CBool(parsed)
            Exit Function
        End If
    End If
    GetBool = defaultVal
End Function

Public Function GetString(ByVal key As String, ByVal defaultVal As String) As String
    Dim v As Variant
    If TryGet(key, v) Then
        GetString = SafeTrim(v)
    Else
        GetString = defaultVal
    End If
End Function

Private Sub InitializeState()
    Set mConfigCache = CreateObject("Scripting.Dictionary")
    mConfigCache.CompareMode = vbTextCompare
    Set mValidationIssues = New Collection
    mWarehouseId = vbNullString
    mStationId = vbNullString
    mResolvedWorkbook = vbNullString
    mIsLoaded = False
End Sub

Private Sub HandleMalformedKey(ByRef def As ConfigKeyDef, ByVal rawVal As Variant)
    Dim v As Variant
    If def.DefaultVal <> "" And TryCoerceValue(def.DataType, def.DefaultVal, v) Then
        mConfigCache(def.Key) = v
        AddValidationIssue "WARN", "CONFIG_KEY_DEFAULT", def.Key & " malformed (" & CStr(rawVal) & "), default applied."
    ElseIf def.Required Then
        AddValidationIssue "ERROR", "CONFIG_KEY_INVALID", def.Key & " has invalid type and no default."
    Else
        AddValidationIssue "WARN", "CONFIG_KEY_INVALID", def.Key & " has invalid type and no default."
    End If
End Sub

Private Sub HandleMissingKey(ByRef def As ConfigKeyDef)
    Dim v As Variant
    If def.DefaultVal <> "" And TryCoerceValue(def.DataType, def.DefaultVal, v) Then
        mConfigCache(def.Key) = v
        AddValidationIssue "WARN", "CONFIG_KEY_DEFAULT", def.Key & " missing, default applied."
    ElseIf def.Required Then
        AddValidationIssue "ERROR", "CONFIG_KEY_MISSING", def.Key & " is required."
    Else
        mConfigCache(def.Key) = Empty
    End If
End Sub

Private Function CountFatalIssues() As Long
    Dim itm As Variant
    Dim parts() As String
    For Each itm In mValidationIssues
        parts = Split(CStr(itm), "|")
        If UBound(parts) >= 0 Then
            If UCase$(parts(0)) = "ERROR" Then CountFatalIssues = CountFatalIssues + 1
        End If
    Next itm
End Function

Private Sub AddValidationIssue(ByVal severity As String, ByVal code As String, ByVal message As String)
    If mValidationIssues Is Nothing Then Set mValidationIssues = New Collection
    mValidationIssues.Add severity & "|" & code & "|" & message
End Sub

Private Function ResolveConfigWorkbook(ByVal whId As String, ByVal stId As String) As Workbook
    Dim wb As Workbook
    Dim bootstrapReport As String

    Set ResolveConfigWorkbook = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime(whId, stId, "", bootstrapReport)
    If Not ResolveConfigWorkbook Is Nothing Then Exit Function

    For Each wb In Application.Workbooks
        If IsConfigWorkbookName(wb.Name) Then
            If whId = "" Or InStr(1, wb.Name, whId, vbTextCompare) > 0 Then
                Set ResolveConfigWorkbook = wb
                Exit Function
            End If
        End If
    Next wb

    If whId <> "" Then
        For Each wb In Application.Workbooks
            If WorkbookHasListObject(wb, "tblWarehouseConfig") And WorkbookHasListObject(wb, "tblStationConfig") Then
                If WorkbookHasWarehouseConfigRow(wb, whId) Then
                    Set ResolveConfigWorkbook = wb
                    Exit Function
                End If
            End If
        Next wb
    End If

    For Each wb In Application.Workbooks
        If WorkbookHasListObject(wb, "tblWarehouseConfig") And WorkbookHasListObject(wb, "tblStationConfig") Then
            Set ResolveConfigWorkbook = wb
            Exit Function
        End If
    Next wb

    Set ResolveConfigWorkbook = modRuntimeWorkbooks.OpenFirstRuntimeConfigWorkbook(bootstrapReport)
    If Not ResolveConfigWorkbook Is Nothing Then Exit Function
End Function

Private Function IsConfigWorkbookName(ByVal wbName As String) As Boolean
    Dim n As String
    n = LCase$(wbName)
    IsConfigWorkbookName = (n Like "wh*.invsys.config.xlsb") Or _
                           (n Like "wh*.invsys.config.xlsx") Or _
                           (n Like "wh*.invsys.config.xlsm")
End Function

Private Function WorkbookHasListObject(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    WorkbookHasListObject = Not (FindListObjectByName(wb, tableName) Is Nothing)
End Function

Private Function WorkbookHasWarehouseConfigRow(ByVal wb As Workbook, ByVal whId As String) As Boolean
    Dim lo As ListObject
    Set lo = FindListObjectByName(wb, "tblWarehouseConfig")
    If lo Is Nothing Then Exit Function
    WorkbookHasWarehouseConfigRow = (FindRowByValue(lo, "WarehouseId", whId) > 0)
End Function

Private Function ResolveWarehouseRow(ByVal loWh As ListObject, ByVal whId As String, ByVal wbName As String) As Long
    Dim inferredWh As String

    If whId <> "" Then
        ResolveWarehouseRow = FindRowByValue(loWh, "WarehouseId", whId)
        If ResolveWarehouseRow > 0 Then Exit Function
    End If

    inferredWh = InferWarehouseIdFromWorkbookName(wbName)
    If inferredWh <> "" Then
        ResolveWarehouseRow = FindRowByValue(loWh, "WarehouseId", inferredWh)
        If ResolveWarehouseRow > 0 Then Exit Function
    End If

    If loWh.ListRows.Count > 0 Then ResolveWarehouseRow = 1
End Function

Private Function ResolveStationRow(ByVal loSt As ListObject, ByVal stId As String, ByVal whId As String) As Long
    Dim computerName As String

    If stId <> "" Then
        ResolveStationRow = FindRowByValue(loSt, "StationId", stId)
        If ResolveStationRow > 0 Then Exit Function
    End If

    computerName = Environ$("COMPUTERNAME")
    If computerName <> "" Then
        ResolveStationRow = FindRowByValue(loSt, "StationName", computerName)
        If ResolveStationRow > 0 Then Exit Function
        ResolveStationRow = FindRowByValue(loSt, "StationId", computerName)
        If ResolveStationRow > 0 Then Exit Function
    End If

    If whId <> "" Then
        ResolveStationRow = FindFirstRowByValue(loSt, "WarehouseId", whId)
        If ResolveStationRow > 0 Then Exit Function
    End If

    If loSt.ListRows.Count > 0 Then ResolveStationRow = 1
End Function

Private Function InferWarehouseIdFromWorkbookName(ByVal wbName As String) As String
    Dim p As Long
    p = InStr(1, wbName, ".", vbTextCompare)
    If p > 1 Then InferWarehouseIdFromWorkbookName = Left$(wbName, p - 1)
End Function

Private Function FindRowByValue(ByVal lo As ListObject, ByVal columnName As String, ByVal matchValue As String) As Long
    Dim colIndex As Long
    Dim i As Long
    Dim v As String

    colIndex = GetColumnIndex(lo, columnName)
    If colIndex = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To lo.ListRows.Count
        v = SafeTrim(lo.DataBodyRange.Cells(i, colIndex).Value)
        If StrComp(v, matchValue, vbTextCompare) = 0 Then
            FindRowByValue = i
            Exit Function
        End If
    Next i
End Function

Private Function FindFirstRowByValue(ByVal lo As ListObject, ByVal columnName As String, ByVal matchValue As String) As Long
    FindFirstRowByValue = FindRowByValue(lo, columnName, matchValue)
End Function

Private Function BuildRowDictionary(ByVal lo As ListObject, ByVal rowIndex As Long) As Object
    Dim d As Object
    Dim col As ListColumn
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    If lo Is Nothing Then
        Set BuildRowDictionary = d
        Exit Function
    End If
    If lo.DataBodyRange Is Nothing Then
        Set BuildRowDictionary = d
        Exit Function
    End If
    If rowIndex < 1 Or rowIndex > lo.ListRows.Count Then
        Set BuildRowDictionary = d
        Exit Function
    End If

    For Each col In lo.ListColumns
        d(col.Name) = lo.DataBodyRange.Cells(rowIndex, col.Index).Value
    Next col
    Set BuildRowDictionary = d
End Function

Private Function TryGetDictionaryValue(ByVal d As Object, ByVal key As String, ByRef outVal As Variant) As Boolean
    If d Is Nothing Then Exit Function
    If d.Exists(key) Then
        outVal = d(key)
        TryGetDictionaryValue = True
    End If
End Function

Private Function GetDictionaryValue(ByVal d As Object, ByVal key As String) As Variant
    Dim tmp As Variant
    If TryGetDictionaryValue(d, key, tmp) Then
        GetDictionaryValue = tmp
    End If
End Function

Private Function TryCoerceValue(ByVal dataType As String, ByVal rawValue As Variant, ByRef outVal As Variant) As Boolean
    Dim t As String

    If IsError(rawValue) Then Exit Function
    t = UCase$(dataType)

    Select Case t
        Case CONFIG_TYPE_STRING
            outVal = SafeTrim(rawValue)
            TryCoerceValue = True
        Case CONFIG_TYPE_LONG
            If IsNumeric(rawValue) Then
                outVal = CLng(rawValue)
                TryCoerceValue = True
            End If
        Case CONFIG_TYPE_BOOLEAN
            If VarType(rawValue) = vbBoolean Then
                outVal = CBool(rawValue)
                TryCoerceValue = True
            Else
                Select Case UCase$(SafeTrim(rawValue))
                    Case "TRUE", "1", "YES", "Y", "ON"
                        outVal = True
                        TryCoerceValue = True
                    Case "FALSE", "0", "NO", "N", "OFF"
                        outVal = False
                        TryCoerceValue = True
                End Select
            End If
        Case CONFIG_TYPE_DATETIME
            If IsDate(rawValue) Then
                outVal = CDate(rawValue)
                TryCoerceValue = True
            End If
        Case Else
            outVal = rawValue
            TryCoerceValue = True
    End Select
End Function

Private Function IsBlankValue(ByVal v As Variant) As Boolean
    If IsError(v) Then
        IsBlankValue = True
    ElseIf IsEmpty(v) Or IsNull(v) Then
        IsBlankValue = True
    Else
        IsBlankValue = (SafeTrim(v) = "")
    End If
End Function

Private Function SafeTrim(ByVal v As Variant) As String
    On Error Resume Next
    SafeTrim = Trim$(CStr(v))
End Function

Private Function EnsureConfigTables(ByVal wb As Workbook) As Boolean
    On Error GoTo FailEnsure

    Dim whHeaders As Variant
    Dim stHeaders As Variant

    whHeaders = Array( _
        "WarehouseId", "WarehouseName", "Timezone", "DefaultLocation", _
        "BatchSize", "LockTimeoutMinutes", "HeartbeatIntervalSeconds", "MaxLockHoldMinutes", _
        "SnapshotCadence", "BackupCadence", "PathDataRoot", "PathBackupRoot", "PathSharePointRoot", _
        "DesignsEnabled", "PoisonRetryMax", "AuthCacheTTLSeconds", "ProcessorServiceUserId", _
        "FF_DesignsEnabled", "FF_OutlookAlerts", "FF_AutoSnapshot")
    stHeaders = Array("StationId", "WarehouseId", "StationName", "RoleDefault")

    EnsureListObjectWithHeaders wb, "WarehouseConfig", "tblWarehouseConfig", whHeaders
    EnsureListObjectWithHeaders wb, "StationConfig", "tblStationConfig", stHeaders

    EnsureConfigTables = True
    Exit Function

FailEnsure:
    EnsureConfigTables = False
End Function

Private Sub EnsureListObjectWithHeaders(ByVal wb As Workbook, _
                                        ByVal sheetName As String, _
                                        ByVal tableName As String, _
                                        ByVal headers As Variant)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim i As Long
    Dim dataRange As Range
    Dim startCell As Range

    Set ws = EnsureConfigWorksheet(wb, sheetName, tableName)
    EnsureWorksheetEditableConfig ws
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0

    If lo Is Nothing Then
        Set startCell = GetNextTableStartCell(ws)
        For i = LBound(headers) To UBound(headers)
            startCell.Offset(0, i - LBound(headers)).Value = headers(i)
        Next i

        Set dataRange = ws.Range(startCell, startCell.Offset(1, UBound(headers) - LBound(headers)))
        Set lo = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
        lo.Name = tableName
        AddValidationIssue "WARN", "CONFIG_TABLE_CREATED", tableName & " created."
    End If

    For i = LBound(headers) To UBound(headers)
        EnsureListColumn lo, CStr(headers(i))
    Next i

    EnsureTableHasRow lo
End Sub

Private Function EnsureConfigWorksheet(ByVal wb As Workbook, ByVal sheetName As String, ByVal tableName As String) As Worksheet
    Dim ws As Worksheet

    Set ws = FindWorksheetByNameConfig(wb, sheetName)
    If ws Is Nothing Then
        Set EnsureConfigWorksheet = EnsureWorksheet(wb, sheetName)
        Exit Function
    End If

    If WorksheetHasListObjectConfig(ws, tableName) Then
        Set EnsureConfigWorksheet = ws
        Exit Function
    End If

    If WorksheetHasUnexpectedContentConfig(ws) Then
        QuarantineConfigWorksheet ws
        AddValidationIssue "WARN", "CONFIG_SHEET_QUARANTINED", sheetName & " contained unexpected content and was quarantined."
        Set EnsureConfigWorksheet = EnsureWorksheet(wb, sheetName)
    Else
        Set EnsureConfigWorksheet = ws
    End If
End Function

Private Function EnsureWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureWorksheet = wb.Worksheets(sheetName)
    On Error GoTo 0

    If EnsureWorksheet Is Nothing Then
        Set EnsureWorksheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureWorksheet.Name = sheetName
    End If
End Function

Private Function FindWorksheetByNameConfig(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set FindWorksheetByNameConfig = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function WorksheetHasListObjectConfig(ByVal ws As Worksheet, ByVal tableName As String) As Boolean
    On Error Resume Next
    WorksheetHasListObjectConfig = Not (ws.ListObjects(tableName) Is Nothing)
    On Error GoTo 0
End Function

Private Function WorksheetHasUnexpectedContentConfig(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    If ws.ListObjects.Count > 0 Then
        WorksheetHasUnexpectedContentConfig = True
        Exit Function
    End If
    WorksheetHasUnexpectedContentConfig = (Application.WorksheetFunction.CountA(ws.Cells) > 0)
End Function

Private Sub QuarantineConfigWorksheet(ByVal ws As Worksheet)
    Dim baseName As String
    Dim suffix As Long
    Dim candidate As String

    If ws Is Nothing Then Exit Sub

    baseName = Left$(ws.Name & "_Stale", 31)
    candidate = baseName
    suffix = 1

    Do While Not FindWorksheetByNameConfig(ws.Parent, candidate) Is Nothing
        candidate = Left$(baseName, 31 - Len(CStr(suffix)) - 1) & "_" & CStr(suffix)
        suffix = suffix + 1
    Loop

    ws.Name = candidate
End Sub

Private Sub SeedConfigDefaults(ByVal wb As Workbook, ByVal warehouseId As String, ByVal stationId As String)
    Dim loWh As ListObject
    Dim loSt As ListObject
    Dim resolvedWh As String
    Dim resolvedSt As String

    Set loWh = FindListObjectByName(wb, "tblWarehouseConfig")
    Set loSt = FindListObjectByName(wb, "tblStationConfig")
    If loWh Is Nothing Or loSt Is Nothing Then Exit Sub

    EnsureTableHasRow loWh
    EnsureTableHasRow loSt

    resolvedWh = ResolveSeedWarehouseIdConfig(warehouseId, wb.Name)
    resolvedSt = ResolveSeedStationIdConfig(stationId)

    EnsureConfigCellDefault loWh, 1, "WarehouseId", resolvedWh
    EnsureConfigCellDefault loWh, 1, "WarehouseName", resolvedWh
    EnsureConfigCellDefault loWh, 1, "Timezone", "UTC"
    EnsureConfigCellDefault loWh, 1, "DefaultLocation", "A1"
    EnsureConfigCellDefault loWh, 1, "BatchSize", 500
    EnsureConfigCellDefault loWh, 1, "LockTimeoutMinutes", 3
    EnsureConfigCellDefault loWh, 1, "HeartbeatIntervalSeconds", 30
    EnsureConfigCellDefault loWh, 1, "MaxLockHoldMinutes", 2
    EnsureConfigCellDefault loWh, 1, "SnapshotCadence", "PER_BATCH"
    EnsureConfigCellDefault loWh, 1, "BackupCadence", "DAILY"
    EnsureConfigCellDefault loWh, 1, "PathDataRoot", Replace$("C:\invSys\{WarehouseId}\", "{WarehouseId}", resolvedWh)
    EnsureConfigCellDefault loWh, 1, "PathBackupRoot", Replace$("C:\invSys\Backups\{WarehouseId}\", "{WarehouseId}", resolvedWh)
    EnsureConfigCellDefault loWh, 1, "PathSharePointRoot", ""
    EnsureConfigCellDefault loWh, 1, "DesignsEnabled", False
    EnsureConfigCellDefault loWh, 1, "PoisonRetryMax", 3
    EnsureConfigCellDefault loWh, 1, "AuthCacheTTLSeconds", 300
    EnsureConfigCellDefault loWh, 1, "ProcessorServiceUserId", "svc_processor"
    EnsureConfigCellDefault loWh, 1, "FF_DesignsEnabled", False
    EnsureConfigCellDefault loWh, 1, "FF_OutlookAlerts", False
    EnsureConfigCellDefault loWh, 1, "FF_AutoSnapshot", True

    EnsureConfigCellDefault loSt, 1, "StationId", resolvedSt
    EnsureConfigCellDefault loSt, 1, "WarehouseId", resolvedWh
    EnsureConfigCellDefault loSt, 1, "StationName", Environ$("COMPUTERNAME")
    EnsureConfigCellDefault loSt, 1, "RoleDefault", "RECEIVE"
End Sub

Private Sub EnsureConfigCellDefault(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal defaultValue As Variant)
    Dim idx As Long
    Dim existingValue As Variant

    idx = GetColumnIndex(lo, columnName)
    If idx = 0 Then Exit Sub
    existingValue = lo.DataBodyRange.Cells(rowIndex, idx).Value
    If IsBlankValue(existingValue) Then
        lo.DataBodyRange.Cells(rowIndex, idx).Value = defaultValue
    End If
End Sub

Private Function ResolveSeedWarehouseIdConfig(ByVal warehouseId As String, ByVal workbookName As String) As String
    ResolveSeedWarehouseIdConfig = SafeTrim(warehouseId)
    If ResolveSeedWarehouseIdConfig = "" Then ResolveSeedWarehouseIdConfig = InferWarehouseIdFromWorkbookName(workbookName)
    If ResolveSeedWarehouseIdConfig = "" Then ResolveSeedWarehouseIdConfig = "WH1"
End Function

Private Function ResolveSeedStationIdConfig(ByVal stationId As String) As String
    ResolveSeedStationIdConfig = SafeTrim(stationId)
    If ResolveSeedStationIdConfig = "" Then ResolveSeedStationIdConfig = "S1"
End Function

Private Sub FormatConfigSurface(ByVal wb As Workbook)
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        ws.Cells.EntireColumn.AutoFit
        ws.Rows(1).Font.Bold = True
    Next ws
End Sub

Private Function GetNextTableStartCell(ByVal ws As Worksheet) As Range
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        Set GetNextTableStartCell = ws.Range("A1")
    Else
        Set GetNextTableStartCell = ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(2, 0)
    End If
End Function

Private Sub EnsureListColumn(ByVal lo As ListObject, ByVal columnName As String)
    Dim idx As Long
    idx = GetColumnIndex(lo, columnName)
    If idx > 0 Then Exit Sub

    lo.ListColumns.Add lo.ListColumns.Count + 1
    lo.ListColumns(lo.ListColumns.Count).Name = columnName
    AddValidationIssue "WARN", "CONFIG_COLUMN_CREATED", lo.Name & "." & columnName & " created."
End Sub

Private Sub EnsureTableHasRow(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then lo.ListRows.Add
End Sub

Private Sub EnsureWorksheetEditableConfig(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    If Not ws.ProtectContents Then Exit Sub

    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    If ws.ProtectContents Then
        Err.Raise vbObjectError + 2701, "modConfig.EnsureWorksheetEditableConfig", _
                  "Worksheet '" & ws.Name & "' is protected and could not be unprotected before updating config tables."
    End If
End Sub

Private Function GetColumnIndex(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long
    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            GetColumnIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function FindListObjectByName(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In wb.Worksheets
        Set FindListObjectByName = ws.ListObjects(tableName)
        If Not FindListObjectByName Is Nothing Then Exit Function
    Next ws
    On Error GoTo 0
End Function
