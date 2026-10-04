Attribute VB_Name = "modWarehouseProvisioning"
Option Explicit
Option Private Module

' Creation-only Config command. The bootstrap owner rejects existing runtimes
' before calling this module; ordinary Settings cannot change warehouse purpose.
Public Function NormalizePurpose(ByRef purpose As String, ByRef report As String) As Boolean
    Select Case UCase$(Trim$(purpose))
        Case "", "OPERATIONAL": purpose = "Operational"
        Case "TRAINING": purpose = "Training"
        Case Else
            report = "Warehouse purpose must be Operational or Training."
            Exit Function
    End Select
    NormalizePurpose = True
End Function

Public Function StampConfig(ByVal wbCfg As Workbook, ByRef spec As WarehouseSpec, _
                            ByRef report As String) As Boolean
    Dim loWh As ListObject, loSt As ListObject, purposeColumn As ListColumn
    On Error GoTo Failed
    If Not NormalizePurpose(spec.WarehousePurpose, report) Then Exit Function
    If wbCfg Is Nothing Then GoTo Failed
    If wbCfg.ReadOnly Or Len(wbCfg.Path) = 0 Then GoTo Failed
    Set loWh = wbCfg.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    Set loSt = wbCfg.Worksheets("StationConfig").ListObjects("tblStationConfig")
    SetProvisioningCell loWh, "WarehouseId", spec.WarehouseId
    SetProvisioningCell loWh, "WarehouseName", IIf(Trim$(spec.WarehouseName) = "", spec.WarehouseId, spec.WarehouseName)
    SetProvisioningCell loWh, "PathDataRoot", spec.PathLocal
    SetProvisioningCell loWh, "PathSharePointRoot", spec.PathSharePoint
    If ProvisioningColumnIndex(loWh, "WarehousePurpose") = 0 Then
        Set purposeColumn = loWh.ListColumns.Add
        purposeColumn.Name = "WarehousePurpose"
    End If
    SetProvisioningCell loWh, "WarehousePurpose", spec.WarehousePurpose
    SetProvisioningCell loSt, "StationId", spec.StationId
    SetProvisioningCell loSt, "WarehouseId", spec.WarehouseId
    SetProvisioningCell loSt, "StationName", spec.AdminUser
    SetProvisioningCell loSt, "PathInboxRoot", spec.PathLocal & "\inbox\"
    SetProvisioningCell loSt, "RoleDefault", "RECEIVE"
    wbCfg.Save
    StampConfig = True
    Exit Function
Failed:
    report = "Creation configuration could not be saved."
End Function

Private Sub SetProvisioningCell(ByVal table As ListObject, ByVal header As String, ByVal value As String)
    Dim column As Long
    column = ProvisioningColumnIndex(table, header)
    If column = 0 Then Err.Raise vbObjectError + 7386, "modWarehouseProvisioning", "Missing configuration header."
    table.DataBodyRange.Cells(1, column).Value2 = value
End Sub

Private Function ProvisioningColumnIndex(ByVal table As ListObject, ByVal header As String) As Long
    Dim column As ListColumn
    For Each column In table.ListColumns
        If StrComp(Trim$(column.Name), header, vbTextCompare) = 0 Then
            If ProvisioningColumnIndex > 0 Then Err.Raise vbObjectError + 7387, "modWarehouseProvisioning", "Ambiguous configuration header."
            ProvisioningColumnIndex = column.Index
        End If
    Next column
End Function
