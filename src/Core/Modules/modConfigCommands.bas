Attribute VB_Name = "modConfigCommands"
Option Explicit

' D5: configuration reads stay in modConfig; ordinary mutation enters here.
Public Function UpdateConfigValue(ByVal key As String, ByVal rawValue As Variant, _
                                 Optional ByRef report As String = "", _
                                 Optional ByVal warehouseId As String = "", _
                                 Optional ByVal stationId As String = "") As Boolean
    Dim target As WarehouseTarget
    Dim changes As Object
    On Error GoTo Failed
    If Not AuthorizeCommand(warehouseId, stationId, False, target, report) Then Exit Function
    Set changes = CreateObject("Scripting.Dictionary")
    changes.CompareMode = vbTextCompare
    changes.Add Trim$(key), rawValue
    UpdateConfigValue = PersistChanges(target, changes, report)
    Exit Function
Failed:
    report = "Configuration save failed. No success was recorded."
End Function

Public Function PublishUomCatalogRows(ByVal rows As Variant, _
                                     Optional ByRef report As String = "") As Boolean
    Dim target As WarehouseTarget
    Dim packedUoms As String, packedConversions As String
    Dim changes As Object
    Dim version As Long
    On Error GoTo Failed
    If Not AuthorizeCommand("", "", True, target, report) Then Exit Function
    If Not modUomSettings.PrepareUomCatalogRows(rows, packedUoms, packedConversions, report) Then Exit Function
    If Not modConfig.LoadConfig(target.WarehouseId, target.StationId) Then
        report = "Configuration could not be read for UOM publication."
        Exit Function
    End If
    version = modConfig.GetLong("UomConversionCatalogVersion", 1)
    If modUomSettings.UomCatalogMatches(packedUoms, packedConversions) Then
        report = "UOM Catalog version " & CStr(version) & " is unchanged."
        PublishUomCatalogRows = True
        Exit Function
    End If
    Set changes = CreateObject("Scripting.Dictionary")
    changes.Add "UomCatalog", packedUoms
    changes.Add "UomConversionCatalog", packedConversions
    changes.Add "UomConversionCatalogVersion", version + 1
    PublishUomCatalogRows = PersistChanges(target, changes, report)
    If PublishUomCatalogRows Then report = "UOM Catalog version " & CStr(version + 1) & " published."
    Exit Function
Failed:
    report = "UOM Catalog publication failed. No success was recorded."
End Function

Private Function AuthorizeCommand(ByVal warehouseId As String, ByVal stationId As String, _
                                  ByVal allowProduction As Boolean, ByRef target As WarehouseTarget, _
                                  ByRef report As String) As Boolean
    Dim allowed As Boolean
    Dim ignored As String
    If Not modAuth.IsSignedIn() Then
        report = "Sign in to invSys before changing configuration."
        Exit Function
    End If
    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then
        report = "Select a connected warehouse before changing configuration."
        Exit Function
    End If
    If (Trim$(warehouseId) <> "" And StrComp(Trim$(warehouseId), target.WarehouseId, vbTextCompare) <> 0) Or _
       (Trim$(stationId) <> "" And StrComp(Trim$(stationId), target.StationId, vbTextCompare) <> 0) Then
        report = "Warehouse or station changed. Reopen Settings before saving."
        Exit Function
    End If
    allowed = modRoleUiAccess.CanCurrentUserPerformCapabilityCached("ADMIN_MAINT", ignored)
    If Not allowed And allowProduction Then
        allowed = modRoleUiAccess.CanCurrentUserPerformCapabilityCached("PROD_POST", ignored)
    End If
    If Not allowed Then
        report = "The signed-in user lacks the required configuration capability."
        Exit Function
    End If
    If Trim$(target.ConfigPath) = "" Then
        report = "The selected warehouse has no configuration path."
        Exit Function
    End If
    AuthorizeCommand = True
End Function

Private Function PersistChanges(ByVal target As WarehouseTarget, ByVal changes As Object, _
                                ByRef report As String) As Boolean
    Dim wb As Workbook, candidate As Workbook
    Dim lo As ListObject, ws As Worksheet, table As ListObject
    Dim defs() As ConfigKeyDef, defCount As Long, definition As Long, i As Long
    Dim key As Variant, value As Variant, oldValues As Collection, cells As Collection
    Dim cell As Range, row As Long, column As Long, tableName As String
    Dim opened As Boolean, wrote As Boolean, saved As Boolean, unchanged As Boolean
    On Error GoTo Failed
    Set cells = New Collection
    Set oldValues = New Collection
    defCount = modConfigDefaults.GetConfigSchema(defs)
    For Each candidate In Application.Workbooks
        If StrComp(candidate.FullName, target.ConfigPath, vbTextCompare) = 0 Then Set wb = candidate: Exit For
    Next candidate
    If wb Is Nothing Then
        If Len(Dir$(target.ConfigPath)) = 0 Then report = "Configuration workbook is missing.": Exit Function
        Set wb = Application.Workbooks.Open(target.ConfigPath, UpdateLinks:=0, ReadOnly:=False, AddToMru:=False)
        opened = True
    End If
    If wb.ReadOnly Or Not wb.Saved Then
        report = "Configuration is read-only, locked, or has unsaved changes. Close it before saving Settings."
        GoTo CleanExit
    End If
    If Not modConfig.LoadConfig(target.WarehouseId, target.StationId) Then
        report = "Configuration is invalid. Use explicit setup before saving Settings."
        GoTo CleanExit
    End If
    ' Validate the complete change set before changing a cell or adding a header.
    For Each key In changes.Keys
        definition = 0
        For i = 1 To defCount
            If StrComp(Trim$(CStr(key)), defs(i).Key, vbTextCompare) = 0 Then definition = i: Exit For
        Next i
        If definition = 0 Then report = "Unknown configuration key.": GoTo CleanExit
        If defs(definition).Key = "WarehouseId" Or defs(definition).Key = "StationId" Then
            report = "WarehouseId and StationId cannot be renamed from Settings."
            GoTo CleanExit
        End If
        If IsEmpty(changes(key)) Or (VarType(changes(key)) = vbString And Len(Trim$(CStr(changes(key)))) = 0) Then
            If defs(definition).Required Then report = "A required configuration value is blank.": GoTo CleanExit
            value = ""
        ElseIf Not modConfig.TryCoerceConfigValue(defs(definition).DataType, changes(key), value) Then
            report = defs(definition).Key & " requires a " & LCase$(defs(definition).DataType) & " value."
            GoTo CleanExit
        End If
        tableName = "tblWarehouseConfig"
        If defs(definition).Scope = CONFIG_SCOPE_STATION Then tableName = "tblStationConfig"
        Set lo = Nothing
        For Each ws In wb.Worksheets
            For Each table In ws.ListObjects
                If StrComp(table.Name, tableName, vbTextCompare) = 0 Then Set lo = table
            Next table
        Next ws
        If lo Is Nothing Then report = "Configuration table is missing; use explicit setup to repair it.": GoTo CleanExit
        row = TargetRow(lo, target)
        column = ColumnIndex(lo, defs(definition).Key)
        If row = 0 Or column = 0 Then report = "Configuration row or header is missing or ambiguous.": GoTo CleanExit
        If lo.Parent.ProtectContents Then report = "Configuration worksheet is protected.": GoTo CleanExit
        Set cell = lo.DataBodyRange.Cells(row, column)
        cells.Add cell
        oldValues.Add cell.Formula
        changes(key) = value
    Next key
    unchanged = True
    i = 0
    For Each key In changes.Keys
        i = i + 1
        Set cell = cells(i)
        If CStr(cell.Value2) <> CStr(changes(key)) Then unchanged = False
    Next key
    If Not unchanged Then
        wrote = True
        i = 0
        For Each key In changes.Keys
            i = i + 1
            Set cell = cells(i)
            cell.Value2 = changes(key)
        Next key
        wb.Save
        saved = True
    End If
    PersistChanges = True
    report = "Configuration saved."
    If Not modConfig.LoadConfig(target.WarehouseId, target.StationId) Then
        report = "Configuration saved; reload failed. Reopen Settings to verify the current values."
    End If
CleanExit:
    If opened And Not wb Is Nothing Then wb.Close SaveChanges:=False
    Exit Function
Failed:
    On Error Resume Next
    If wrote And Not saved Then
        For i = 1 To cells.Count
            Set cell = cells(i)
            cell.Formula = oldValues(i)
        Next i
        If Not wb Is Nothing Then wb.Saved = True
    End If
    report = "Configuration save failed. Reopen Settings to verify the current values."
    PersistChanges = False
    Resume CleanExit
End Function

Private Function ColumnIndex(ByVal table As ListObject, ByVal header As String) As Long
    Dim column As ListColumn
    For Each column In table.ListColumns
        If StrComp(Trim$(column.Name), header, vbTextCompare) = 0 Then
            If ColumnIndex > 0 Then ColumnIndex = 0: Exit Function
            ColumnIndex = column.Index
        End If
    Next column
End Function

Private Function TargetRow(ByVal table As ListObject, ByVal target As WarehouseTarget) As Long
    Dim whColumn As Long, stColumn As Long, row As Long
    whColumn = ColumnIndex(table, "WarehouseId")
    If whColumn = 0 Or table.DataBodyRange Is Nothing Then Exit Function
    If table.Name = "tblStationConfig" Then
        stColumn = ColumnIndex(table, "StationId")
        If stColumn = 0 Then Exit Function
    End If
    For row = 1 To table.ListRows.Count
        If StrComp(Trim$(CStr(table.DataBodyRange.Cells(row, whColumn).Value2)), target.WarehouseId, vbTextCompare) = 0 Then
            If stColumn = 0 Then
                If TargetRow > 0 Then TargetRow = 0: Exit Function
                TargetRow = row
            ElseIf StrComp(Trim$(CStr(table.DataBodyRange.Cells(row, stColumn).Value2)), target.StationId, vbTextCompare) = 0 Then
                If TargetRow > 0 Then TargetRow = 0: Exit Function
                TargetRow = row
            End If
        End If
    Next row
End Function
