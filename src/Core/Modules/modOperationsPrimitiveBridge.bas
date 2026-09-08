Attribute VB_Name = "modOperationsPrimitiveBridge"
Option Explicit

Public Function EnsureProductionWorkbookSurface(ByVal workbookName As String, _
                                                ByRef report As String) As Boolean
    Dim wb As Workbook

    Set wb = ResolveOpenWorkbook(workbookName)
    If wb Is Nothing Then
        report = "Production operator workbook is not open: " & Trim$(workbookName)
        Exit Function
    End If
    EnsureProductionWorkbookSurface = _
        modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface(wb, report)
End Function

Public Function EnsureReceivingWorkbookSurface(ByVal workbookName As String, _
                                               ByRef report As String) As Boolean
    Dim wb As Workbook

    Set wb = ResolveOpenWorkbook(workbookName)
    If wb Is Nothing Then
        report = "Receiving operator workbook is not open: " & Trim$(workbookName)
        Exit Function
    End If
    EnsureReceivingWorkbookSurface = _
        modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wb, report)
End Function

Public Function EnsureShippingWorkbookSurface(ByVal workbookName As String, _
                                              ByRef report As String) As Boolean
    Dim wb As Workbook

    Set wb = ResolveOpenWorkbook(workbookName)
    If wb Is Nothing Then
        report = "Shipping operator workbook is not open: " & Trim$(workbookName)
        Exit Function
    End If
    EnsureShippingWorkbookSurface = _
        modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface(wb, report)
End Function

Public Function ResolveEligibleRoleOperatorWorkbookName( _
        ByVal preferredWorkbookName As String, _
        ByVal roleCode As String, _
        ByRef workbookNameOut As String, _
        ByRef report As String) As Boolean
    Dim wb As Workbook
    Dim candidate As Workbook
    Dim candidateCount As Long

    workbookNameOut = vbNullString
    report = vbNullString
    roleCode = UCase$(Trim$(roleCode))
    If Not IsSupportedRoleCodePrimitive(roleCode) Then
        report = "Unsupported Operations role: " & roleCode
        Exit Function
    End If

    Set wb = ResolveOpenWorkbook(preferredWorkbookName)
    If IsEligibleRoleOperatorWorkbookPrimitive(wb, roleCode) Then
        workbookNameOut = wb.Name
        report = "OK"
        ResolveEligibleRoleOperatorWorkbookName = True
        Exit Function
    End If

    For Each wb In Application.Workbooks
        If IsEligibleRoleOperatorWorkbookPrimitive(wb, roleCode) Then
            candidateCount = candidateCount + 1
            Set candidate = wb
        End If
    Next wb

    If candidateCount = 1 Then
        workbookNameOut = candidate.Name
        report = "OK"
        ResolveEligibleRoleOperatorWorkbookName = True
    ElseIf candidateCount > 1 Then
        report = "Multiple eligible " & RoleDisplayNamePrimitive(roleCode) & _
                 " operator workbooks are open. Activate the intended workbook and try again."
    Else
        report = "Open a saved " & RoleDisplayNamePrimitive(roleCode) & _
                 " operator workbook before using this control."
    End If
End Function

Public Function OpenOrCreateCurrentReceivingOperatorWorkbook( _
        ByVal preferredWorkbookName As String, _
        ByRef workbookNameOut As String, _
        ByRef report As String) As Boolean
    OpenOrCreateCurrentReceivingOperatorWorkbook = _
        OpenOrCreateCurrentRoleOperatorWorkbook( _
            preferredWorkbookName, "RECEIVING", workbookNameOut, report)
End Function

Public Function OpenOrCreateCurrentRoleOperatorWorkbook( _
        ByVal preferredWorkbookName As String, _
        ByVal roleCode As String, _
        ByRef workbookNameOut As String, _
        ByRef report As String) As Boolean
    If ResolveEligibleRoleOperatorWorkbookName( _
            preferredWorkbookName, roleCode, workbookNameOut, report) Then
        OpenOrCreateCurrentRoleOperatorWorkbook = True
        Exit Function
    End If
    If InStr(1, report, "Multiple eligible ", vbTextCompare) = 1 Then Exit Function

    OpenOrCreateCurrentRoleOperatorWorkbook = _
        modWarehouseBootstrap.OpenOrCreateRoleOperatorWorkbookForCurrentTarget( _
            roleCode, workbookNameOut, report)
End Function

Public Function ShouldBootstrapRoleWorkbookSurface(ByVal workbookName As String) As Boolean
    Dim wb As Workbook

    Set wb = ResolveOpenWorkbook(workbookName)
    If wb Is Nothing Then Exit Function
    ShouldBootstrapRoleWorkbookSurface = _
        modRoleWorkbookSurfaces.ShouldBootstrapRoleWorkbookSurface(wb)
End Function

Public Function BeginQuietUiForWorkbook(ByVal workbookName As String) As Boolean
    Dim wb As Workbook

    Set wb = ResolveOpenWorkbook(workbookName)
    If wb Is Nothing Then Exit Function
    modUiQuiet.BeginQuietUi wb
    BeginQuietUiForWorkbook = True
End Function

Public Function InitializeProductionAutoSnapshot(ByVal workbookName As String) As Boolean
    Dim wb As Workbook

    Set wb = ResolveOpenWorkbook(workbookName)
    If wb Is Nothing Then Exit Function
    modOperatorReadModel.InitializeAutoSnapshotForWorkbook wb
    InitializeProductionAutoSnapshot = True
End Function

Public Function InitializeReceivingAutoSnapshot(ByVal workbookName As String) As Boolean
    Dim wb As Workbook

    Set wb = ResolveOpenWorkbook(workbookName)
    If wb Is Nothing Then Exit Function
    modOperatorReadModel.InitializeAutoSnapshotForWorkbook wb
    InitializeReceivingAutoSnapshot = True
End Function

Public Sub UnregisterProductionAutoSnapshot(ByVal workbookName As String)
    Dim wb As Workbook

    Set wb = ResolveOpenWorkbook(workbookName)
    If wb Is Nothing Then Exit Sub
    modOperatorReadModel.UnregisterAutoSnapshotWorkbook wb
End Sub

Public Sub UnregisterReceivingAutoSnapshot(ByVal workbookName As String)
    Dim wb As Workbook

    Set wb = ResolveOpenWorkbook(workbookName)
    If wb Is Nothing Then Exit Sub
    modOperatorReadModel.UnregisterAutoSnapshotWorkbook wb
End Sub

Public Function RefreshInventoryReadModel(ByVal workbookName As String, _
                                          Optional ByVal warehouseId As String = "", _
                                          Optional ByVal sourceType As String = "LOCAL", _
                                          Optional ByRef report As String = "", _
                                          Optional ByRef refreshState As String = "") As Boolean
    Dim wb As Workbook
    refreshState = "FAILED"
    Set wb = ResolveOpenWorkbook(workbookName)
    If wb Is Nothing Then
        report = "Operator workbook is not open: " & Trim$(workbookName)
        Exit Function
    End If
    RefreshInventoryReadModel = _
        modOperatorReadModel.RefreshInventoryReadModelForWorkbook( _
            wb, warehouseId, sourceType, report, refreshState)
End Function

Public Function DiagnoseInventoryReadModel(ByVal workbookName As String, _
                                           Optional ByVal warehouseId As String = "", _
                                           Optional ByVal sourceType As String = "LOCAL") As String
    Dim wb As Workbook

    Set wb = ResolveOpenWorkbook(workbookName)
    If wb Is Nothing Then
        DiagnoseInventoryReadModel = _
            "Operator workbook is not open: " & Trim$(workbookName)
        Exit Function
    End If
    DiagnoseInventoryReadModel = _
        modOperatorReadModel.DiagnoseInventoryReadModelRefresh( _
            wb, warehouseId, sourceType)
End Function

Public Function RunBatchAndRefreshOperatorWorkbook(ByVal workbookName As String, _
                                                   Optional ByVal warehouseId As String = "", _
                                                   Optional ByVal sourceType As String = "LOCAL", _
                                                   Optional ByRef report As String = "", _
                                                   Optional ByVal requireQueuedWork As Boolean = True, _
                                                   Optional ByRef queuedWorkHandled As Boolean = False) As Boolean
    Dim wb As Workbook

    Set wb = ResolveOpenWorkbook(workbookName)
    If wb Is Nothing Then
        report = "Operator workbook is not open: " & Trim$(workbookName)
        Exit Function
    End If
    RunBatchAndRefreshOperatorWorkbook = _
        modOperatorReadModel.RunBatchAndRefreshOperatorWorkbook( _
            wb, warehouseId, sourceType, report, requireQueuedWork, queuedWorkHandled)
End Function

Public Sub ApplyShapeCapability(ByVal workbookName As String, _
                                ByVal worksheetName As String, _
                                ByVal shapeName As String, _
                                ByVal capability As String)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim shp As Shape
    Dim errorMessage As String

    Set wb = ResolveOpenWorkbook(workbookName)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    Set ws = wb.Worksheets(worksheetName)
    If Not ws Is Nothing Then Set shp = ws.Shapes(shapeName)
    On Error GoTo 0
    If shp Is Nothing Then Exit Sub
    shp.Visible = IIf( _
        modRoleUiAccess.CanCurrentUserPerformCapabilityCached(capability, errorMessage), _
        -1, _
        0)
End Sub

Public Function ListDesigns(Optional ByVal statusFilter As String = "") As Variant
    ListDesigns = modDesignsDomainBridge.ListDesignsBridge(Nothing, statusFilter)
End Function

Public Function GetDesignBom(ByVal designId As String, _
                             ByVal designVersion As String) As Variant
    GetDesignBom = _
        modDesignsDomainBridge.GetDesignBOMBridge(designId, designVersion, Nothing)
End Function

Public Function GetDesignBomForStatus(ByVal designId As String, _
                                      ByVal designVersion As String, _
                                      ByVal requiredStatus As String) As Variant
    GetDesignBomForStatus = _
        modDesignsDomainBridge.GetDesignBOMForStatusBridge( _
            designId, designVersion, requiredStatus, Nothing)
End Function

Public Function ListProcesses(Optional ByVal statusFilter As String = "", _
                              Optional ByVal designsWb As Workbook = Nothing) As Variant
    ListProcesses = modDesignsDomainBridge.ListProcessesBridge(designsWb, statusFilter)
End Function

Public Function GetProcessVersion(ByVal processId As String, _
                                  ByVal processVersion As String, _
                                  Optional ByVal designsWb As Workbook = Nothing) As String
    GetProcessVersion = modDesignsDomainBridge.GetProcessVersionBridge( _
        processId, processVersion, designsWb)
End Function

Public Function ListRecipes(Optional ByVal statusFilter As String = "", _
                            Optional ByVal designsWb As Workbook = Nothing) As Variant
    ListRecipes = modDesignsDomainBridge.ListRecipesBridge(designsWb, statusFilter)
End Function

Public Function GetRecipeGraph(ByVal recipeId As String, _
                               ByVal recipeVersion As String, _
                               Optional ByVal designsWb As Workbook = Nothing) As String
    GetRecipeGraph = modDesignsDomainBridge.GetRecipeGraphBridge( _
        recipeId, recipeVersion, designsWb)
End Function

Public Function ValidateReleasedRecipe(ByVal recipeId As String, _
                                       ByVal recipeVersion As String, _
                                       Optional ByVal designsWb As Workbook = Nothing) As String
    ValidateReleasedRecipe = modDesignsDomainBridge.ValidateReleasedRecipeBridgeEncoded( _
        recipeId, recipeVersion, designsWb)
End Function

Private Function ResolveOpenWorkbook(ByVal workbookName As String) As Workbook
    Dim wb As Workbook

    workbookName = Trim$(workbookName)
    If workbookName = "" Then Exit Function

    On Error Resume Next
    Set ResolveOpenWorkbook = Application.Workbooks(workbookName)
    On Error GoTo 0
    If Not ResolveOpenWorkbook Is Nothing Then Exit Function

    For Each wb In Application.Workbooks
        On Error Resume Next
        If StrComp(wb.FullName, workbookName, vbTextCompare) = 0 Then
            Set ResolveOpenWorkbook = wb
            Exit Function
        End If
        On Error GoTo 0
    Next wb
End Function

Private Function IsSupportedRoleCodePrimitive(ByVal roleCode As String) As Boolean
    IsSupportedRoleCodePrimitive = _
        (roleCode = "RECEIVING" Or roleCode = "PRODUCTION" Or roleCode = "SHIPPING")
End Function

Private Function RoleDisplayNamePrimitive(ByVal roleCode As String) As String
    Select Case roleCode
        Case "RECEIVING"
            RoleDisplayNamePrimitive = "Receiving"
        Case "PRODUCTION"
            RoleDisplayNamePrimitive = "Production"
        Case "SHIPPING"
            RoleDisplayNamePrimitive = "Shipping"
        Case Else
            RoleDisplayNamePrimitive = "role"
    End Select
End Function

Private Function IsEligibleRoleOperatorWorkbookPrimitive(ByVal wb As Workbook, _
                                                         ByVal roleCode As String) As Boolean
    If IsRejectedOperatorAuthorityWorkbookPrimitive(wb) Then Exit Function

    Select Case roleCode
        Case "RECEIVING"
            IsEligibleRoleOperatorWorkbookPrimitive = _
                (Not FindListObjectPrimitive(wb, "ReceivedTally") Is Nothing) And _
                (Not FindListObjectPrimitive(wb, "invSys") Is Nothing)
        Case "PRODUCTION"
            IsEligibleRoleOperatorWorkbookPrimitive = _
                (Not FindListObjectPrimitive(wb, "RB_AddRecipeName") Is Nothing) And _
                (Not FindListObjectPrimitive(wb, "ProductionOutput") Is Nothing)
        Case "SHIPPING"
            IsEligibleRoleOperatorWorkbookPrimitive = _
                (Not FindListObjectPrimitive(wb, "ShipmentsTally") Is Nothing) And _
                (Not FindListObjectPrimitive(wb, "NotShipped") Is Nothing)
    End Select
End Function

Private Function IsRejectedOperatorAuthorityWorkbookPrimitive(ByVal wb As Workbook) As Boolean
    Dim wbName As String

    If wb Is Nothing Then
        IsRejectedOperatorAuthorityWorkbookPrimitive = True
        Exit Function
    End If
    If wb.IsAddin Or wb.ReadOnly Then
        IsRejectedOperatorAuthorityWorkbookPrimitive = True
        Exit Function
    End If
    If Trim$(wb.Path) = "" Then
        IsRejectedOperatorAuthorityWorkbookPrimitive = True
        Exit Function
    End If

    wbName = LCase$(Trim$(wb.Name))
    If wbName = "" Or Left$(wbName, 2) = "~$" Or wbName = "personal.xlsb" Then
        IsRejectedOperatorAuthorityWorkbookPrimitive = True
        Exit Function
    End If
    If wbName Like "*.xla" Or wbName Like "*.xlam" Then
        IsRejectedOperatorAuthorityWorkbookPrimitive = True
        Exit Function
    End If
    If wbName Like "invsys.inbox.*" _
       Or InStr(1, wbName, ".invsys.config.", vbTextCompare) > 0 _
       Or InStr(1, wbName, ".invsys.auth.", vbTextCompare) > 0 _
       Or InStr(1, wbName, ".invsys.data.", vbTextCompare) > 0 _
       Or InStr(1, wbName, ".invsys.snapshot.", vbTextCompare) > 0 _
       Or InStr(1, wbName, ".outbox.", vbTextCompare) > 0 Then
        IsRejectedOperatorAuthorityWorkbookPrimitive = True
    End If
End Function

Private Function FindListObjectPrimitive(ByVal wb As Workbook, _
                                         ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindListObjectPrimitive = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindListObjectPrimitive Is Nothing Then Exit Function
    Next ws
End Function
