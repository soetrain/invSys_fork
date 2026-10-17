Attribute VB_Name = "modInventoryCommands"
Option Explicit

' D8-INV-01 interactive facade. Handles are transient and grant no headless authority.
Private mForms As cCapturedFormStore
Private Const RESULT_CATALOG As Long = 1
Private Const RESULT_STAGE As Long = 2
Private Const RESULT_QUANTITY As Long = 3
Private Const RESULT_CREATE As Long = 4
Private Const RESULT_RETIREMENT As Long = 5
Private Const RESULT_STAGE_SAVED As Long = 6
Private Const PHASE_CATALOG As String = "Catalog"
Private Const PHASE_QUANTITY As String = "Quantity"
Private Const PHASE_CREATE As String = "Create"
Private Const PHASE_RETIREMENT As String = "Retirement"

Public Function Capture(ByVal workbookName As String) As String
    Dim book As Workbook, target As WarehouseTarget, entry As Object
    Dim context As String, key As String, report As String
    Dim inputs As cInventoryCommandInputs
    On Error GoTo Failed
    Set book = Application.Workbooks(workbookName)
    If book.IsAddin Or book.Path = "" Or book.ReadOnly Then Exit Function
    Set inputs = New cInventoryCommandInputs
    If Not inputs.CaptureSelection(book) Then Exit Function
    context = modAuthUserCommands.CaptureContext()
    If context = "" Then Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    If modAggregationFiles.SameFile(book.FullName, target.ConfigPath) Or _
       modAggregationFiles.SameFile(book.FullName, target.AuthPath) Then Exit Function
    If LCase$(book.Name) Like "*.invsys.data.*" Or _
       LCase$(book.Name) Like "*.invsys.snapshot.*" Or _
       LCase$(book.Name) = "invsys.global.inventorysnapshot.xlsb" Then Exit Function
    If Not modConfig.LoadConfig(target.WarehouseId, target.StationId) Then Exit Function
    Set entry = CreateObject("Scripting.Dictionary")
    entry.Add "Context", context: entry.Add "Workbook", book
    entry.Add "WorkbookPath", book.FullName
    entry.Add "Warehouse", target.WarehouseId: entry.Add "Station", target.StationId
    entry.Add "User", modAuth.GetCurrentUserId(): entry.Add "Root", target.RuntimeRoot
    entry.Add "Config", target.ConfigPath: entry.Add "Auth", target.AuthPath
    entry.Add "Feed", modConfig.GetString("PathDataRoot", "")
    entry.Add "Expired", False
    If mForms Is Nothing Then Set mForms = New cCapturedFormStore
    key = modSystemIdentity.NewId(): mForms.Add key, entry
    If Not CheckAuthority(key, report) Then Release key: Exit Function
    If Not inputs.CaptureSource(target.WarehouseId) Then Release key: Exit Function
    If modAggregationFiles.SameFile(book.FullName, inputs.SourcePath) Then Release key: Exit Function
    entry.Add "Inputs", inputs
    If Not CheckForm(key, report) Then Release key: Exit Function
    Capture = key
Failed:
End Function

' Declared cross-package values: a primitive handle and a three-string array.
Public Function SourceValues(ByVal key As String) As Variant
    Dim report As String, entry As Object
    If Not CheckForm(key, report) Then Exit Function
    Set entry = mForms.Item(key)
    SourceValues = Array(CStr(entry("Warehouse")), CStr(entry("Station")), CStr(entry("User")))
End Function

Public Function CheckForm(ByVal key As String, ByRef report As String) As Boolean
    Dim entry As Object, inputs As cInventoryCommandInputs
    On Error GoTo Refused
    If Not CheckAuthority(key, report) Then Exit Function
    Set entry = mForms.Item(key): Set inputs = entry("Inputs")
    If Not inputs.Unchanged() Then GoTo Refused
    CheckForm = CheckAuthority(key, report)
    Exit Function
Refused:
    Expire key, report
End Function

Private Function CheckAuthority(ByVal key As String, ByRef report As String) As Boolean
    Dim entry As Object, original As Workbook, book As Workbook, found As Boolean
    Dim root As String
    report = "The invSys session or inventory source changed. Reopen Add/Edit Inventory."
    On Error GoTo Refused
    If mForms Is Nothing Then Exit Function
    Set entry = mForms.Item(key)
    If entry Is Nothing Then Exit Function
    If entry("Expired") Then Exit Function
    If Not modAuth.IsSignedIn() Then
        report = "Current invSys user is not signed in."
        GoTo Refused
    End If
    If entry("Context") <> modAuthUserCommands.CaptureContext() Then GoTo Refused
    Set original = entry("Workbook")
    If original.ReadOnly Then GoTo Refused
    For Each book In Application.Workbooks
        If book Is original Then found = True
    Next book
    If Not found Then GoTo Refused
    If Not modAggregationFiles.SameFile(original.FullName, CStr(entry("WorkbookPath"))) Then GoTo Refused
    root = modRuntimeWorkbooks.GetCoreDataRootOverride()
    If root <> "" Then
        If StrComp(modAggregationFiles.Canonical(root), _
                   modAggregationFiles.Canonical(CStr(entry("Root"))), vbTextCompare) <> 0 Then GoTo Refused
    End If
    If modAggregationFiles.Dirty(CStr(entry("Config"))) Or _
       modAggregationFiles.Dirty(CStr(entry("Auth"))) Then GoTo Refused
    If Not modAuth.LoadAuth(CStr(entry("Warehouse"))) Then GoTo Refused
    If Not modAuth.CanPerform("ADMIN_MAINT", modAuth.GetCurrentUserId(), _
            CStr(entry("Warehouse")), CStr(entry("Station")), "INVENTORY_COMMAND") Then
        report = "Current user lacks ADMIN_MAINT for the captured warehouse."
        GoTo Refused
    End If
    If Not modConfig.LoadConfig(CStr(entry("Warehouse")), CStr(entry("Station"))) Then GoTo Refused
    If modConfig.GetString("PathDataRoot", "") <> entry("Feed") Then GoTo Refused
    If entry("Context") <> modAuthUserCommands.CaptureContext() Then GoTo Refused
    report = "": CheckAuthority = True
    Exit Function
Refused:
    If Not entry Is Nothing Then entry("Expired") = True
End Function

Private Sub Expire(ByVal key As String, ByRef report As String)
    Dim entry As Object
    report = "The selected inventory source, catalog, entity or staging table changed. Reopen Add/Edit Inventory."
    On Error Resume Next
    Set entry = mForms.Item(key)
    If Not entry Is Nothing Then entry("Expired") = True
End Sub

Public Function SelectSku(ByVal key As String, ByVal sku As String, ByRef report As String) As Boolean
    Dim entry As Object, inputs As cInventoryCommandInputs
    On Error GoTo Refused
    If Not CheckForm(key, report) Then Exit Function
    Set entry = mForms.Item(key): Set inputs = entry("Inputs")
    If Not inputs.BindSku(sku) Then GoTo Refused
    SelectSku = CheckAuthority(key, report)
    Exit Function
Refused:
    Expire key, report
End Function

Public Function SourceInventoryPath(ByVal key As String) As String
    Dim entry As Object, inputs As cInventoryCommandInputs, report As String
    If Not CheckForm(key, report) Then Exit Function
    Set entry = mForms.Item(key): Set inputs = entry("Inputs")
    SourceInventoryPath = inputs.SourcePath
End Function

Public Function OperatorWorkbookName(ByVal key As String) As String
    Dim entry As Object, book As Workbook, report As String
    If Not CheckForm(key, report) Then Exit Function
    Set entry = mForms.Item(key): Set book = entry("Workbook")
    OperatorWorkbookName = book.Name
End Function

Public Function ConfirmStageSaved(ByVal key As String, ByRef report As String) As Boolean
    ConfirmStageSaved = ConfirmResult(key, RESULT_STAGE_SAVED, "", report)
End Function

Public Function CheckOperatorWorkbook(ByVal key As String, ByVal workbookName As String, ByRef report As String) As Boolean
    Dim entry As Object, original As Workbook, candidate As Workbook
    On Error GoTo Refused
    If Not CheckForm(key, report) Then Exit Function
    Set entry = mForms.Item(key): Set original = entry("Workbook")
    Set candidate = Application.Workbooks(workbookName)
    If Not (candidate Is original) Then GoTo Refused
    CheckOperatorWorkbook = True
    Exit Function
Refused:
    Expire key, report
End Function

Public Function StagingTableName(ByVal key As String) As String
    Dim report As String, entry As Object, inputs As cInventoryCommandInputs
    If Not CheckForm(key, report) Then Exit Function
    Set entry = mForms.Item(key): Set inputs = entry("Inputs")
    StagingTableName = inputs.StagingName
End Function

Public Function PlanStageCell(ByVal key As String, ByVal row As Long, ByVal header As String, _
        ByVal value As String, ByRef report As String) As Boolean
    Dim entry As Object, inputs As cInventoryCommandInputs
    On Error GoTo Refused
    If Not CheckForm(key, report) Then Exit Function
    Set entry = mForms.Item(key): Set inputs = entry("Inputs")
    If Not inputs.PlanStageCell(row, header, value) Then GoTo Refused
    PlanStageCell = True
    Exit Function
Refused:
    Expire key, report
End Function

Public Function ConfirmStageCell(ByVal key As String, ByRef report As String) As Boolean
    ConfirmStageCell = ConfirmResult(key, RESULT_STAGE, "", report)
End Function

Public Function BeginCatalog(ByVal key As String, ByVal sku As String, ByRef report As String) As Boolean
    Dim entry As Object, inputs As cInventoryCommandInputs
    On Error GoTo Refused
    If Not CheckForm(key, report) Then Exit Function
    Set entry = mForms.Item(key): Set inputs = entry("Inputs")
    If Not inputs.BeginCatalog(sku) Then GoTo Refused
    BeginCatalog = True
    Exit Function
Refused:
    Expire key, report
End Function

Public Function BindCatalogWriter(ByVal key As String, ByVal workbookName As String, ByRef report As String) As Boolean
    Dim entry As Object, inputs As cInventoryCommandInputs
    On Error GoTo Refused
    If Not CheckForm(key, report) Then Exit Function
    Set entry = mForms.Item(key): Set inputs = entry("Inputs")
    If Not inputs.BindCatalogWriter(Application.Workbooks(workbookName)) Then GoTo Refused
    BindCatalogWriter = True
    Exit Function
Refused:
    Expire key, report
End Function

Public Function PlanCatalogField(ByVal key As String, ByVal sku As String, ByVal header As String, _
        ByVal value As String, ByRef report As String) As Boolean
    Dim entry As Object, inputs As cInventoryCommandInputs
    On Error GoTo Refused
    If mForms Is Nothing Then GoTo Refused
    Set entry = mForms.Item(key)
    If entry Is Nothing Then GoTo Refused
    If entry("Expired") Then GoTo Refused
    Set inputs = entry("Inputs")
    inputs.PlanCatalogField sku, header, value
    PlanCatalogField = True
    Exit Function
Refused:
    Expire key, report
End Function

Public Function BeginCatalogSave(ByVal key As String, ByRef report As String) As Boolean
    BeginCatalogSave = BeginResultSave(key, PHASE_CATALOG, report)
End Function

Public Function BeginQuantity(ByVal key As String, ByVal systemKey As String, _
        ByVal target As Double, ByRef report As String) As Boolean
    Dim entry As Object, inputs As cInventoryCommandInputs
    On Error GoTo Refused
    If Not CheckForm(key, report) Then Exit Function
    Set entry = mForms.Item(key): Set inputs = entry("Inputs")
    If Not inputs.BeginQuantity(systemKey, target) Then GoTo Refused
    BeginQuantity = CheckForm(key, report)
    Exit Function
Refused:
    Expire key, report
End Function

Public Function BeginRetirement(ByVal key As String, ByVal sku As String, ByVal reason As String, _
        ByRef payload As String, ByRef count As Long, ByRef report As String) As Boolean
    Dim entry As Object, inputs As cInventoryCommandInputs
    On Error GoTo Refused
    If Not CheckForm(key, report) Then Exit Function
    Set entry = mForms.Item(key): Set inputs = entry("Inputs")
    If Not inputs.BeginRetirement(sku, reason, payload, count) Then GoTo Refused
    BeginRetirement = CheckForm(key, report)
    Exit Function
Refused:
    Expire key, report
End Function

Public Function BeginRetirementSave(ByVal key As String, ByRef report As String) As Boolean
    BeginRetirementSave = BeginResultSave(key, PHASE_RETIREMENT, report)
End Function

Public Function ConfirmRetirement(ByVal key As String, ByVal eventId As String, ByRef report As String) As Boolean
    ConfirmRetirement = ConfirmResult(key, RESULT_RETIREMENT, eventId, report)
End Function

Public Function BeginCreate(ByVal key As String, ByVal payload As String, ByRef report As String) As Boolean
    Dim entry As Object, inputs As cInventoryCommandInputs
    On Error GoTo Refused
    If Not CheckForm(key, report) Then Exit Function
    Set entry = mForms.Item(key): Set inputs = entry("Inputs")
    If Not inputs.BeginCreate(payload) Then GoTo Refused
    BeginCreate = CheckForm(key, report)
    Exit Function
Refused:
    Expire key, report
End Function

Public Function BeginCreateSave(ByVal key As String, ByRef report As String) As Boolean
    BeginCreateSave = BeginResultSave(key, PHASE_CREATE, report)
End Function

Public Function ConfirmCreate(ByVal key As String, ByVal eventId As String, ByRef report As String) As Boolean
    ConfirmCreate = ConfirmResult(key, RESULT_CREATE, eventId, report)
End Function

Public Function BeginQuantitySave(ByVal key As String, ByRef report As String) As Boolean
    BeginQuantitySave = BeginResultSave(key, PHASE_QUANTITY, report)
End Function

Private Function BeginResultSave(ByVal key As String, ByVal phase As String, ByRef report As String) As Boolean
    Dim entry As Object, inputs As cInventoryCommandInputs
    On Error GoTo Refused
    If Not CheckForm(key, report) Then Exit Function
    Set entry = mForms.Item(key): Set inputs = entry("Inputs")
    If Not inputs.BeginSave(phase) Then GoTo Refused
    BeginResultSave = True
    Exit Function
Refused:
    Expire key, report
End Function

Public Function ConfirmCatalog(ByVal key As String, ByRef report As String) As Boolean
    ConfirmCatalog = ConfirmResult(key, RESULT_CATALOG, "", report)
End Function

Public Function ConfirmQuantity(ByVal key As String, ByVal eventId As String, ByRef report As String) As Boolean
    ConfirmQuantity = ConfirmResult(key, RESULT_QUANTITY, eventId, report)
End Function

Private Function ConfirmResult(ByVal key As String, ByVal kind As Long, _
        ByVal eventId As String, ByRef report As String) As Boolean
    Dim entry As Object, inputs As cInventoryCommandInputs
    Dim confirmed As Boolean
    On Error GoTo Refused
    If Not CheckAuthority(key, report) Then Exit Function
    Set entry = mForms.Item(key): Set inputs = entry("Inputs")
    Select Case kind
        Case RESULT_CATALOG: confirmed = inputs.ConfirmCatalog()
        Case RESULT_STAGE: confirmed = inputs.ConfirmStageCell()
        Case RESULT_STAGE_SAVED: confirmed = inputs.ConfirmStageSaved()
        Case RESULT_QUANTITY: confirmed = inputs.ConfirmEventResult(eventId, PHASE_QUANTITY)
        Case RESULT_CREATE: confirmed = inputs.ConfirmEventResult(eventId, PHASE_CREATE)
        Case RESULT_RETIREMENT: confirmed = inputs.ConfirmEventResult(eventId, PHASE_RETIREMENT)
    End Select
    If Not confirmed Then GoTo Refused
    ConfirmResult = CheckForm(key, report)
    Exit Function
Refused:
    Expire key, report
End Function

Public Sub Release(ByVal key As String)
    If Not mForms Is Nothing Then mForms.Remove key
End Sub
