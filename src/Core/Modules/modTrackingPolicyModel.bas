Attribute VB_Name = "modTrackingPolicyModel"
Option Explicit
Option Private Module

Public Function AuthorizeEditor(ByVal context As String, ByRef target As WarehouseTarget, ByRef report As String) As Boolean
    report = "Session or warehouse changed. Reopen Settings."
    If context = "" Or context <> modActivity.CaptureContext() Then Exit Function
    If Not modRoleUiAccess.CanCurrentUserPerformCapabilityCached("ADMIN_MAINT", report) Then Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then Exit Function
    AuthorizeEditor = True
End Function

Public Function Defaults(Optional ByVal enableDefaults As Boolean = True) As Object
    Dim model As Object, rows As Collection, row As Object, definition As Object, id As Variant
    Set model = CreateObject("Scripting.Dictionary")
    model.Add "SchemaVersion", 1
    model.Add "CatalogVersion", modActivityCatalog.CATALOG_VERSION
    model.Add "DefaultView", "How-To"
    model.Add "ViewerActionPathCaptureEnabled", False
    model.Add "AdminViewerEventLoggingEnabled", True
    Set rows = New Collection
    For Each id In modActivityCatalog.ControlIds()
        Set definition = modActivityCatalog.Control(CStr(id))
        Set row = CreateObject("Scripting.Dictionary")
        row.Add "ControlId", CStr(id)
        row.Add "Collect", enableDefaults And definition("Class") = "Command"
        row.Add "Visible", enableDefaults
        row.Add "SequenceEligible", enableDefaults
        rows.Add row
    Next id
    model.Add "Controls", rows
    Set Defaults = model
End Function

Public Function Decode(ByVal request As String) As Object
    Dim model As Object, rows As Collection, row As Variant, seen As Object, id As String
    On Error GoTo Invalid
    If LenB(request) > 1048576 Then Exit Function
    Set model = modTrainingJson.DecodeObject(request)
    If Not HasFields(model, Array("SchemaVersion", "CatalogVersion", "DefaultView", _
        "ViewerActionPathCaptureEnabled", "AdminViewerEventLoggingEnabled", "Controls")) Then Exit Function
    If Not ExactInteger(model("SchemaVersion"), 1) Then Exit Function
    If Not ExactInteger(model("CatalogVersion"), modActivityCatalog.CATALOG_VERSION) Then Exit Function
    If VarType(model("DefaultView")) <> vbString Then Exit Function
    If model("DefaultView") <> "How-To" And model("DefaultView") <> "Diagnostic" And _
       model("DefaultView") <> "Compare both" Then Exit Function
    If VarType(model("ViewerActionPathCaptureEnabled")) <> vbBoolean Then Exit Function
    If VarType(model("AdminViewerEventLoggingEnabled")) <> vbBoolean Then Exit Function
    If TypeName(model("Controls")) <> "Collection" Then Exit Function
    Set rows = model("Controls")
    Set seen = CreateObject("Scripting.Dictionary")
    For Each row In rows
        If Not IsObject(row) Then Exit Function
        If Not HasFields(row, Array("ControlId", "Collect", "Visible", "SequenceEligible")) Then Exit Function
        If VarType(row("ControlId")) <> vbString Then Exit Function
        id = row("ControlId")
        If seen.Exists(id) Then Exit Function
        If modActivityCatalog.Control(id) Is Nothing Then Exit Function
        If VarType(row("Collect")) <> vbBoolean Or VarType(row("Visible")) <> vbBoolean Or _
           VarType(row("SequenceEligible")) <> vbBoolean Then Exit Function
        seen.Add id, True
    Next row
    If rows.Count <> UBound(modActivityCatalog.ControlIds()) + 1 Then Exit Function
    Set Decode = model
Invalid:
End Function

Private Function HasFields(ByVal model As Object, ByVal fields As Variant) As Boolean
    Dim field As Variant
    If model Is Nothing Then Exit Function
    If TypeName(model) <> "Dictionary" Then Exit Function
    If model.Count <> UBound(fields) + 1 Then Exit Function
    For Each field In fields
        If Not model.Exists(CStr(field)) Then Exit Function
    Next field
    HasFields = True
End Function

Private Function ExactInteger(ByVal value As Variant, ByVal expected As Long) As Boolean
    Select Case VarType(value)
        Case vbInteger, vbLong, vbDouble: ExactInteger = (value = expected)
    End Select
End Function

Public Sub Project(ByVal headers As ListObject, ByVal controls As ListObject, _
                   ByVal selected As Long, ByVal version As Long, ByVal output As Object)
    Dim model As Object, row As Variant, index As Long, key As Variant, field As Variant
    Set model = Defaults(headers Is Nothing)
    If Not headers Is Nothing Then
        For Each field In Array("DefaultView", "ViewerActionPathCaptureEnabled", "AdminViewerEventLoggingEnabled")
            model(field) = TableValue(headers, selected, CStr(field))
        Next field
        For Each row In model("Controls")
            For index = 1 To controls.ListRows.Count
                If TableValue(controls, index, "PolicyVersion") = version And _
                   TableValue(controls, index, "ControlId") = row("ControlId") Then
                    For Each field In Array("Collect", "Visible", "SequenceEligible")
                        row(field) = TableValue(controls, index, CStr(field))
                    Next field
                End If
            Next index
        Next row
    End If
    For Each key In model.Keys
        output.Add key, model(key)
    Next key
    output.Add "SavedCatalogVersion", modActivityCatalog.CATALOG_VERSION
    If Not headers Is Nothing Then output("SavedCatalogVersion") = TableValue(headers, selected, "CatalogVersion")
End Sub

Public Function TableColumn(ByVal table As ListObject, ByVal header As String) As Long
    Dim column As ListColumn
    For Each column In table.ListColumns
        If StrComp(Trim$(column.Name), header, vbTextCompare) = 0 Then
            If TableColumn <> 0 Then Err.Raise 5, , "Ambiguous tracking policy header."
            TableColumn = column.Index
        End If
    Next column
    If TableColumn = 0 Then Err.Raise 5, , "Missing tracking policy header."
End Function

Public Function TableValue(ByVal table As ListObject, ByVal row As Long, ByVal header As String) As Variant
    TableValue = table.DataBodyRange.Cells(row, TableColumn(table, header)).Value2
End Function
