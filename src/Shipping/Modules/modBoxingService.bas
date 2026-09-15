Attribute VB_Name = "modBoxingService"
Option Explicit

Public Function LoadBoxDesigns(ByVal operatorWb As Workbook, _
                               Optional ByVal includeActive As Boolean = True, _
                               Optional ByVal includeArchived As Boolean = False) As Variant
    If operatorWb Is Nothing Then Exit Function
    LoadBoxDesigns = modTS_Shipments.BoxBuilderFormLoadSavedBoxes( _
        includeActive, includeArchived, False, operatorWb)
End Function

Public Function LoadBoxMakerChoices(ByVal operatorWb As Workbook, _
                                    Optional ByRef report As String = "") As Variant
    If operatorWb Is Nothing Then
        report = "The captured Shipping operator workbook was not provided."
        Exit Function
    End If
    LoadBoxMakerChoices = modTS_Shipments.BoxMakerFormLoadSavedBoxes(operatorWb, False)
    report = "OK"
End Function

Public Function LoadBoxDesignVersions(ByVal operatorWb As Workbook, _
                                      ByVal packageSystemKey As String) As Variant
    If operatorWb Is Nothing Or Trim$(packageSystemKey) = "" Then Exit Function
    LoadBoxDesignVersions = modTS_Shipments.BoxBuilderFormLoadVersions( _
        packageSystemKey, operatorWb)
End Function

Public Function LoadBoxDesignComponents(ByVal operatorWb As Workbook, _
                                        ByVal packageSystemKey As String, _
                                        ByVal versionLabel As String) As Variant
    If operatorWb Is Nothing Or Trim$(packageSystemKey) = "" Then Exit Function
    LoadBoxDesignComponents = modTS_Shipments.BoxBuilderFormLoadVersionComponents( _
        packageSystemKey, versionLabel, operatorWb)
End Function

Public Function LoadBoxMakerVersions(ByVal operatorWb As Workbook, _
                                     ByVal packageSystemKey As String) As Variant
    If operatorWb Is Nothing Or Trim$(packageSystemKey) = "" Then Exit Function
    LoadBoxMakerVersions = modTS_Shipments.BoxMakerFormLoadVersions( _
        packageSystemKey, operatorWb)
End Function

Public Function LoadBoxMakerComponents(ByVal operatorWb As Workbook, _
                                       ByVal packageSystemKey As String, _
                                       ByVal versionLabel As String) As Variant
    If operatorWb Is Nothing Or Trim$(packageSystemKey) = "" Then Exit Function
    LoadBoxMakerComponents = modTS_Shipments.BoxMakerFormLoadVersionComponents( _
        packageSystemKey, versionLabel, operatorWb)
End Function

Public Function LoadComponentChoices(ByVal operatorWb As Workbook) As Variant
    If operatorWb Is Nothing Then Exit Function
    LoadComponentChoices = modTS_Shipments.LoadShippingComponentPickerItems(operatorWb)
End Function

Public Function SaveBoxDesign(ByVal operatorWb As Workbook, _
                              ByVal boxName As String, _
                              ByVal boxUom As String, _
                              ByVal boxLocation As String, _
                              ByVal boxDescription As String, _
                              ByVal componentRows As Variant, _
                              ByVal saveAction As String, _
                              ByVal versionLabel As String, _
                              ByVal statusText As String, _
                              ByRef report As String) As Boolean
    On Error GoTo Fail

    Dim quietStarted As Boolean
    Dim failureReason As String

    If operatorWb Is Nothing Then
        report = "The captured Shipping operator workbook was not provided."
        Exit Function
    End If
    If Not modRoleUiAccess.RequireCurrentUserCapability("SHIP_POST") Then
        report = "SHIP_POST capability is required."
        Exit Function
    End If
    If Trim$(boxName) = "" Or Trim$(boxUom) = "" Or IsEmpty(componentRows) Then
        report = "Box name, UOM, and at least one component are required."
        Exit Function
    End If

    modUiQuiet.BeginQuietUi operatorWb
    quietStarted = True
    modTS_Shipments.CommitBoxBuilderFormState _
        boxName, boxUom, boxLocation, boxDescription, componentRows, _
        saveAction, versionLabel, statusText, operatorWb
    modUiQuiet.EndQuietUi
    quietStarted = False
    report = "Box design action completed."
    SaveBoxDesign = True
    Exit Function

Fail:
    failureReason = Err.Description
    On Error Resume Next
    If quietStarted Then modUiQuiet.EndQuietUi
    On Error GoTo 0
    report = "Box design action failed: " & failureReason
End Function

Public Function DeleteBoxDesignVersion(ByVal operatorWb As Workbook, _
                                       ByVal packageSystemKey As String, _
                                       ByVal versionLabel As String, _
                                       ByRef report As String) As Boolean
    If Not ValidateBoxMaintenanceRequest(operatorWb, packageSystemKey, report) Then Exit Function
    If Trim$(versionLabel) = "" Then
        report = "Select a box alternative before deleting."
        Exit Function
    End If
    DeleteBoxDesignVersion = modTS_Shipments.DeleteBoxDesignVersionForWorkbook( _
        operatorWb, packageSystemKey, versionLabel, report)
End Function

Public Function ArchiveBoxDesign(ByVal operatorWb As Workbook, _
                                 ByVal packageSystemKey As String, _
                                 ByRef report As String) As Boolean
    If Not ValidateBoxMaintenanceRequest(operatorWb, packageSystemKey, report) Then Exit Function
    ArchiveBoxDesign = modTS_Shipments.ArchiveBoxDesignForWorkbook( _
        operatorWb, packageSystemKey, report)
End Function

Public Function DeleteBoxDesign(ByVal operatorWb As Workbook, _
                                ByVal packageSystemKey As String, _
                                ByRef report As String) As Boolean
    If Not ValidateBoxMaintenanceRequest(operatorWb, packageSystemKey, report) Then Exit Function
    DeleteBoxDesign = modTS_Shipments.DeleteBoxDesignForWorkbook( _
        operatorWb, packageSystemKey, report)
End Function

Private Function ValidateBoxMaintenanceRequest(ByVal operatorWb As Workbook, _
                                               ByVal packageSystemKey As String, _
                                               ByRef report As String) As Boolean
    If operatorWb Is Nothing Then
        report = "The captured Shipping operator workbook was not provided."
        Exit Function
    End If
    If Trim$(packageSystemKey) = "" Then
        report = "Select a saved box design."
        Exit Function
    End If
    If Not modRoleUiAccess.RequireCurrentUserCapability("ADMIN_MAINT") Then
        report = "ADMIN_MAINT capability is required."
        Exit Function
    End If
    ValidateBoxMaintenanceRequest = True
End Function

Public Function PostBoxMakerAction(ByVal operatorWb As Workbook, _
                                   ByVal packageSystemKey As String, _
                                   ByVal boxName As String, _
                                   ByVal boxUom As String, _
                                   ByVal boxLocation As String, _
                                   ByVal boxDescription As String, _
                                   ByVal versionLabel As String, _
                                   ByVal boxQty As Double, _
                                   ByVal componentRows As Variant, _
                                   ByVal actionText As String, _
                                   ByRef report As String, _
                                   Optional ByVal facts As cShippingOwnerFacts = Nothing) As Boolean
    On Error GoTo Fail

    Dim syncCompleted As Boolean
    Dim quietStarted As Boolean
    Dim failureReason As String

    If Not facts Is Nothing Then facts.BeginOwner
    If operatorWb Is Nothing Then
        report = "The captured Shipping operator workbook was not provided."
        Exit Function
    End If
    If Not modRoleUiAccess.RequireCurrentUserCapability("SHIP_POST") Then
        report = "SHIP_POST capability is required."
        If Not facts Is Nothing Then facts.RejectBeforeMutation True
        Exit Function
    End If
    modUiQuiet.BeginQuietUi operatorWb
    quietStarted = True
    PostBoxMakerAction = modTS_Shipments.CommitBoxMakerFormAction( _
        packageSystemKey, boxName, boxUom, boxLocation, boxDescription, _
        versionLabel, boxQty, componentRows, report, actionText, _
        syncCompleted, Empty, operatorWb, facts)
    modUiQuiet.EndQuietUi
    quietStarted = False
    Exit Function

Fail:
    failureReason = Err.Description
    If Not facts Is Nothing Then facts.CompleteOwner False
    On Error Resume Next
    If quietStarted Then modUiQuiet.EndQuietUi
    On Error GoTo 0
    report = "Box Maker action failed: " & failureReason
End Function

Public Function RunRelease1BoxingActionForTest(ByVal operatorWb As Workbook, _
                                               ByVal componentSku As String, _
                                               ByVal packageSku As String, _
                                               ByVal versionLabel As String, _
                                               ByVal boxQty As Double, _
                                               ByVal componentQtyPerBox As Double) As String
    Dim componentRows(1 To 1, 1 To 8) As Variant
    Dim componentSystemKey As String
    Dim packageSystemKey As String
    Dim report As String
    Dim succeeded As Boolean

    componentSystemKey = _
        modTS_Shipments.ResolveBoxingInventorySystemKeyForSku( _
            operatorWb, componentSku)
    If componentSystemKey = "" Then
        RunRelease1BoxingActionForTest = _
            "FAIL|Component System_Key was not resolved for " & componentSku
        Exit Function
    End If
    packageSystemKey = modRoleEventWriter.CreateSystemKey()
    If packageSystemKey = "" Then
        RunRelease1BoxingActionForTest = _
            "FAIL|Package System_Key could not be created."
        Exit Function
    End If

    componentRows(1, 2) = componentSku
    componentRows(1, 3) = componentSku
    componentRows(1, 4) = componentSystemKey
    componentRows(1, 5) = componentQtyPerBox
    componentRows(1, 6) = "EA"
    componentRows(1, 7) = "BIN-A"
    componentRows(1, 8) = "Release 1 packaged component"

    succeeded = PostBoxMakerAction(operatorWb, packageSystemKey, packageSku, "EA", _
        "BIN-B", "Release 1 versioned box", versionLabel, boxQty, _
        componentRows, "MAKE", report)
    If succeeded Then
        RunRelease1BoxingActionForTest = "OK|BomVersion=" & _
            versionLabel & "|" & report
    Else
        RunRelease1BoxingActionForTest = "FAIL|" & report
    End If
End Function

Public Function NasInventoryIsReadOnly() As Boolean
    ' Boxing services consume NAS Inv as a read-only canonical input.
    NasInventoryIsReadOnly = True
End Function

Public Function ProjectedComponentInventory(ByVal nasInventory As Double, _
                                            ByVal pendingBuildQty As Double, _
                                            ByVal pendingUnboxQty As Double) As Double
    ProjectedComponentInventory = nasInventory - pendingBuildQty + pendingUnboxQty
End Function

Public Function ProjectedComponentInventoryTextForTest(ByVal rowValue As Long, _
                                                       ByVal backendText As String, _
                                                       ByVal requiredQty As Double) As String
    Dim projectedQty As Double

    ProjectedComponentInventoryTextForTest = Trim$(backendText)
    If rowValue <= 0 Or requiredQty <= 0 Then Exit Function
    If Trim$(backendText) = "" Then Exit Function
    If Not IsNumeric(Replace$(backendText, ",", "")) Then Exit Function

    projectedQty = CDbl(Replace$(backendText, ",", "")) - requiredQty
    If projectedQty < 0 Then projectedQty = 0
    ProjectedComponentInventoryTextForTest = FormatBoxingQuantityText(projectedQty)
End Function

Public Function RenderedComponentInventoryAfterPendingActionForTest( _
                                        ByVal backendText As String, _
                                        ByVal perBoxQty As Double, _
                                        ByVal qtyMade As Double, _
                                        ByVal actionText As String) As String
    Dim projectedQty As Double
    Dim requiredQty As Double

    If Trim$(backendText) = "" Or _
       Not IsNumeric(Replace$(backendText, ",", "")) Then
        RenderedComponentInventoryAfterPendingActionForTest = _
            "NAS=unknown;PROJECTED=unknown"
        Exit Function
    End If

    projectedQty = CDbl(Replace$(backendText, ",", ""))
    requiredQty = perBoxQty * qtyMade
    If UCase$(Trim$(actionText)) = "UNMAKE" Or _
       UCase$(Trim$(actionText)) = "UNBOX" Then
        projectedQty = projectedQty + requiredQty
    Else
        projectedQty = projectedQty - requiredQty
        If projectedQty < 0 Then projectedQty = 0
    End If
    RenderedComponentInventoryAfterPendingActionForTest = _
        "NAS=" & FormatBoxingQuantityText(projectedQty) & _
        ";PROJECTED=" & FormatBoxingQuantityText(projectedQty)
End Function

Private Function FormatBoxingQuantityText(ByVal value As Double) As String
    If Abs(value - Fix(value)) < 0.0000001 Then
        FormatBoxingQuantityText = Format$(value, "0")
    Else
        FormatBoxingQuantityText = Format$(value, "0.###")
    End If
End Function
