Attribute VB_Name = "modAggregationAuthority"
Option Explicit
Option Private Module

Private mAggregationForms As cCapturedFormStore

Public Function Capture(ByVal bookName As String) As String
    Dim context As String, target As WarehouseTarget, book As Workbook, entry As Object, key As String
    Dim feedPath As String, raw As String
    On Error GoTo Failed
    context = modAuthUserCommands.CaptureContext()
    If context = "" Then Exit Function
    If bookName = "" Then Set book = Application.ActiveWorkbook Else Set book = Application.Workbooks(bookName)
    If book Is Nothing Then Exit Function
    If book.IsAddin Then Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    If modAggregationFiles.SameFile(book.FullName, target.ConfigPath) Or modAggregationFiles.SameFile(book.FullName, target.AuthPath) Then Exit Function
    If modAggregationFiles.SameFile(book.FullName, target.RuntimeRoot & "\" & target.WarehouseId & ".invSys.Data.Inventory.xlsb") Then Exit Function
    If modAggregationFiles.Dirty(target.AuthPath) Or modAggregationFiles.Dirty(target.ConfigPath) Then Exit Function
    If Not modAuth.LoadAuth(target.WarehouseId) Then Exit Function
    If Not modAuth.CanPerform("ADMIN_MAINT", modAuth.GetCurrentUserId(), target.WarehouseId, target.StationId, "HQ_AGGREGATE") Then Exit Function
    If Not modConfig.LoadConfig(target.WarehouseId, target.StationId) Then Exit Function
    feedPath = modAggregationSources.Feed(target.ConfigPath, target.WarehouseId, raw)
    If feedPath = "" Or Not CreateObject("Scripting.FileSystemObject").FolderExists(feedPath) Then Exit Function
    Set entry = CreateObject("Scripting.Dictionary")
    entry.Add "Context", context: entry.Add "Workbook", book
    entry.Add "Warehouse", target.WarehouseId: entry.Add "Station", target.StationId
    entry.Add "User", modAuth.GetCurrentUserId(): entry.Add "Root", target.RuntimeRoot
    entry.Add "Config", target.ConfigPath: entry.Add "Auth", target.AuthPath
    entry.Add "Feed", raw: entry.Add "Destination", modAggregationFiles.Canonical(feedPath & "\Global\invSys.Global.InventorySnapshot.xlsb")
    entry.Add "Used", False: entry.Add "Sources", CreateObject("Scripting.Dictionary")
    entry("Sources").CompareMode = vbTextCompare
    If mAggregationForms Is Nothing Then Set mAggregationForms = New cCapturedFormStore
    key = modSystemIdentity.NewId(): mAggregationForms.Add key, entry
    If Not Check(key) Then mAggregationForms.Remove key: Exit Function
    Capture = key
Failed:
End Function

Public Function Entry(ByVal key As String) As Object
    If mAggregationForms Is Nothing Then Exit Function
    Set Entry = mAggregationForms.Item(key)
End Function

Public Sub Release(ByVal key As String)
    If mAggregationForms Is Nothing Then Exit Sub
    mAggregationForms.Remove key
End Sub

Public Function Check(ByVal key As String) As Boolean
    Dim saved As Object, book As Workbook, original As Workbook, found As Boolean, raw As String, feedPath As String, root As String
    On Error GoTo Failed
    Set saved = Entry(key)
    If saved Is Nothing Then Exit Function
    If saved("Used") Or saved("Context") <> modAuthUserCommands.CaptureContext() Then Exit Function
    Set original = saved("Workbook")
    For Each book In Application.Workbooks
        If book Is original Then found = True
    Next book
    If Not found Then Exit Function
    root = modRuntimeWorkbooks.GetCoreDataRootOverride()
    If root <> "" Then
        If StrComp(modAggregationFiles.Canonical(root), modAggregationFiles.Canonical(CStr(saved("Root"))), vbTextCompare) <> 0 Then Exit Function
    End If
    If modAggregationFiles.Dirty(CStr(saved("Auth"))) Or modAggregationFiles.Dirty(CStr(saved("Config"))) Then Exit Function
    If Not modAuth.LoadAuth(CStr(saved("Warehouse"))) Then Exit Function
    If Not modAuth.CanPerform("ADMIN_MAINT", modAuth.GetCurrentUserId(), CStr(saved("Warehouse")), CStr(saved("Station")), "HQ_AGGREGATE") Then Exit Function
    feedPath = modAggregationSources.Feed(CStr(saved("Config")), CStr(saved("Warehouse")), raw)
    If feedPath = "" Or raw <> saved("Feed") Then
        ' Restoring the setting cannot revive an observed-untrusted capture.
        saved("Used") = True
        Exit Function
    End If
    Check = (saved("Context") = modAuthUserCommands.CaptureContext())
Failed:
End Function

Public Function SelectSources(ByVal key As String, ByVal paths As String) As Collection
    Dim saved As Object, result As New Collection, path As Variant, normalized As String, record As Object, seen As Object
    On Error GoTo Invalid
    Set saved = Entry(key)
    If saved Is Nothing Then Exit Function
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
    For Each path In Split(Replace$(paths, vbCrLf, vbLf), vbLf)
        normalized = modAggregationFiles.Canonical(CStr(path))
        If normalized = "" Or Not saved("Sources").Exists(normalized) Then Exit Function
        Set record = saved("Sources")(normalized)
        If seen.Exists(record("Warehouse")) Then Exit Function
        seen.Add record("Warehouse"), True
        result.Add record
    Next path
    If result.Count > 0 Then Set SelectSources = result
Invalid:
End Function

Public Function Joint(ByVal key As String, ByVal sources As Collection) As Boolean
    Dim source As Variant, saved As Object, destination As String, fso As Object
    On Error GoTo Failed
    If Not Check(key) Then Exit Function
    Set saved = Entry(key): destination = CStr(saved("Destination"))
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not modAggregationFiles.Available(destination) Then Exit Function
    If modAggregationFiles.Under(destination, CStr(saved("Root"))) Then Exit Function
    For Each source In sources
        If modAggregationFiles.Under(destination, fso.GetParentFolderName(CStr(source("Config")))) Then Exit Function
        If modAggregationFiles.Under(destination, fso.GetParentFolderName(CStr(source("Path")))) Then Exit Function
        If modAggregationFiles.SameFile(destination, CStr(source("Path"))) Then Exit Function
        If Not modAggregationSources.Unchanged(source) Then Exit Function
    Next source
    Joint = Check(key)
Failed:
End Function
