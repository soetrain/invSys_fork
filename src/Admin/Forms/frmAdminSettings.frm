VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAdminSettings
   Caption         =   "invSys Settings"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAdminSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private WithEvents mTxtCarrier As MSForms.TextBox
Private WithEvents mLstCarriers As MSForms.ListBox
Private WithEvents mLstConfig As MSForms.ListBox
Private WithEvents mTxtConfigKey As MSForms.TextBox
Private WithEvents mTxtConfigValue As MSForms.TextBox
Private WithEvents mBtnSaveConfig As MSForms.CommandButton
Private WithEvents mBtnReloadConfig As MSForms.CommandButton
Private WithEvents mBtnAdd As MSForms.CommandButton
Private WithEvents mBtnRemove As MSForms.CommandButton
Private WithEvents mBtnReset As MSForms.CommandButton
Private WithEvents mTxtUom As MSForms.TextBox
Private WithEvents mLstUoms As MSForms.ListBox
Private WithEvents mBtnUomAdd As MSForms.CommandButton
Private WithEvents mBtnUomRemove As MSForms.CommandButton
Private WithEvents mBtnUomReset As MSForms.CommandButton
Private WithEvents mBtnClose As MSForms.CommandButton
Private WithEvents mChkManualServerCredentials As MSForms.CheckBox
Private WithEvents mBtnSaveConnectionPolicy As MSForms.CommandButton

Private mLblStatus As MSForms.Label
Private mLblConfigWorkbook As MSForms.Label
Private mLoading As Boolean
Private mWarehouseId As String
Private mStationId As String
Private mActivityContext As String
Private mResizeInitialized As Boolean

Private Sub UserForm_Initialize()
    CaptureTargetContext
    BuildLayout
    LoadConfigRows
    LoadConnectionPolicy
    LoadCarriers
    LoadUoms
End Sub

Private Sub UserForm_Activate()
    If mResizeInitialized Then Exit Sub
    On Error Resume Next
    modUserFormResizeWin.EnableResizableUserForm Me, True, True
    On Error GoTo 0
    mResizeInitialized = True
End Sub

Private Sub CaptureTargetContext()
    Dim target As WarehouseTarget

    Set target = modNasConnection.GetCurrentTarget()
    If Not target Is Nothing Then
        mWarehouseId = Trim$(target.WarehouseId)
        mStationId = Trim$(target.StationId)
    End If
    If mWarehouseId = "" Then mWarehouseId = Trim$(modConfig.GetWarehouseId())
    If mStationId = "" Then mStationId = Trim$(modConfig.GetStationId())
    mActivityContext = modActivity.CaptureContext()
End Sub

Private Sub BuildLayout()
    Me.Caption = "invSys Settings"
    Me.Width = 720
    Me.Height = 630

    AddLabel "lblTitle", "Warehouse Settings", 12, 10, 180, 18, True
    Set mLblConfigWorkbook = AddLabel("lblConfigWorkbook", "", 200, 10, 490, 18, False)
    AddLabel "lblConfigKeyHeader", "Key", 14, 40, 185, 16, True
    AddLabel "lblConfigValueHeader", "Value", 205, 40, 260, 16, True
    AddLabel "lblConfigTypeHeader", "Type", 470, 40, 70, 16, True
    AddLabel "lblConfigScopeHeader", "Scope", 545, 40, 85, 16, True
    AddLabel "lblConfigRequiredHeader", "Required", 635, 40, 55, 16, True

    Set mLstConfig = AddListBox("lstConfig", 12, 58, 680, 225)
    With mLstConfig
        .ColumnCount = 5
        .ColumnWidths = "190 pt;265 pt;75 pt;90 pt;55 pt"
    End With

    AddLabel "lblSelectedKey", "Selected key", 12, 296, 85, 16, False
    Set mTxtConfigKey = AddTextBox("txtConfigKey", 100, 292, 190, 22)
    mTxtConfigKey.Locked = True
    AddLabel "lblSelectedValue", "Value", 302, 296, 40, 16, False
    Set mTxtConfigValue = AddTextBox("txtConfigValue", 344, 292, 210, 22)
    Set mBtnSaveConfig = AddButton("btnSaveConfig", "Save Value", 564, 290, 76, 26)
    Set mBtnReloadConfig = AddButton("btnReloadConfig", "Reload", 646, 290, 46, 26)

    AddLabel "lblServerConnection", "Server Connection", 12, 338, 150, 18, True
    Set mChkManualServerCredentials = AddCheckBox( _
        "chkManualServerCredentials", _
        "Always require manual server credential entry when Connect Server is clicked", _
        12, 364, 470, 20)
    Set mBtnSaveConnectionPolicy = AddButton("btnSaveConnectionPolicy", "Save Connection Option", 500, 358, 142, 28)
    AddLabel "lblServerConnectionScope", "Applies to this Windows user only; it does not inconvenience other stations.", 30, 386, 500, 18, False

    AddLabel "lblSection", "Shipping Carriers", 12, 420, 150, 18, True
    AddLabel "lblCarrier", "Carrier", 12, 448, 60, 18, False
    Set mTxtCarrier = AddTextBox("txtCarrier", 76, 444, 170, 22)
    Set mBtnAdd = AddButton("btnAdd", "Add", 252, 442, 42, 26)
    Set mBtnRemove = AddButton("btnRemove", "Remove", 300, 442, 54, 26)

    Set mLstCarriers = AddListBox("lstCarriers", 12, 476, 270, 72)
    With mLstCarriers
        .ColumnCount = 1
        .ColumnWidths = "245 pt"
    End With
    Set mBtnReset = AddButton("btnReset", "Reset", 288, 476, 66, 28)

    AddLabel "lblUomSection", "Recipe UOM Catalog", 365, 420, 170, 18, True
    AddLabel "lblUom", "UOM", 365, 448, 40, 18, False
    Set mTxtUom = AddTextBox("txtUom", 410, 444, 140, 22)
    Set mBtnUomAdd = AddButton("btnUomAdd", "Add", 556, 442, 42, 26)
    Set mBtnUomRemove = AddButton("btnUomRemove", "Remove", 604, 442, 58, 26)
    Set mLstUoms = AddListBox("lstUoms", 365, 476, 235, 72)
    With mLstUoms
        .ColumnCount = 1
        .ColumnWidths = "210 pt"
    End With
    Set mBtnUomReset = AddButton("btnUomReset", "Reset", 608, 476, 54, 28)

    Set mBtnClose = AddButton("btnClose", "Close", 626, 558, 66, 28)
    Set mLblStatus = AddLabel("lblStatus", "", 12, 558, 600, 36, False)
End Sub

Private Sub LoadConfigRows()
    Dim rows As Variant

    mLoading = True
    mLstConfig.Clear
    If Not modConfig.IsLoaded() _
       Or (mWarehouseId <> "" And StrComp(mWarehouseId, modConfig.GetWarehouseId(), vbTextCompare) <> 0) _
       Or (mStationId <> "" And StrComp(mStationId, modConfig.GetStationId(), vbTextCompare) <> 0) Then
        If Not modConfig.LoadConfig(mWarehouseId, mStationId) Then
            mLoading = False
            ShowStatus "Config load failed: " & modConfig.Validate()
            Exit Sub
        End If
    End If

    rows = modConfig.GetConfigEditorRows()
    If Not IsEmpty(rows) Then
        If IsArray(rows) Then mLstConfig.List = rows
    End If
    mLblConfigWorkbook.Caption = "Canonical workbook: " & modConfig.GetResolvedWorkbookName()
    mTxtConfigKey.Value = ""
    mTxtConfigValue.Value = ""
    mBtnSaveConfig.Enabled = False
    mLoading = False
End Sub

Private Sub mLstConfig_Click()
    Dim keyName As String

    If mLoading Then Exit Sub
    If mLstConfig.ListIndex < 0 Then Exit Sub
    keyName = CStr(mLstConfig.List(mLstConfig.ListIndex, 0))
    mTxtConfigKey.Value = keyName
    mTxtConfigValue.Value = CStr(mLstConfig.List(mLstConfig.ListIndex, 1))
    mBtnSaveConfig.Enabled = Not IsIdentityConfigKey(keyName)
    mTxtConfigValue.Enabled = mBtnSaveConfig.Enabled
    If mBtnSaveConfig.Enabled Then
        ShowStatus "Edit the selected value, then click Save Value."
    Else
        ShowStatus keyName & " is runtime identity and is read-only here."
    End If
End Sub

Private Sub mBtnSaveConfig_Click()
    Dim report As String
    Dim keyName As String
    keyName = Trim$(CStr(mTxtConfigKey.Value))
    If modAdminSettingsAction.SaveValue(keyName, mTxtConfigValue.Value, mWarehouseId, mStationId, mActivityContext, report) Then
        LoadConfigRows
        If StrComp(keyName, "UomCatalog", vbTextCompare) = 0 Then LoadUoms
    End If
    ShowStatus report
End Sub

Private Sub mBtnReloadConfig_Click()
    If modConfig.Reload() Then
        LoadConfigRows
        ShowStatus "Canonical config reloaded."
    Else
        ShowStatus "Config reload failed: " & modConfig.Validate()
    End If
End Sub

Private Function IsIdentityConfigKey(ByVal keyName As String) As Boolean
    IsIdentityConfigKey = (StrComp(keyName, "WarehouseId", vbTextCompare) = 0) _
                          Or (StrComp(keyName, "StationId", vbTextCompare) = 0)
End Function

Public Function TestInitializeConfigEditor() As String
    If mLstConfig Is Nothing Then BuildLayout
    LoadConfigRows
    LoadConnectionPolicy
    LoadUoms
    TestInitializeConfigEditor = "Rows=" & CStr(mLstConfig.ListCount) & _
                                 "|Workbook=" & modConfig.GetResolvedWorkbookName() & _
                                 "|ManualServerCredentials=" & _
                                 IIf(CBool(mChkManualServerCredentials.Value), "TRUE", "FALSE") & _
                                 "|Uoms=" & CStr(mLstUoms.ListCount)
End Function

Private Sub LoadConnectionPolicy()
    If mChkManualServerCredentials Is Nothing Then Exit Sub
    mChkManualServerCredentials.Value = modNasConnection.RequireManualServerCredentials()
End Sub

Private Sub mBtnSaveConnectionPolicy_Click()
    Dim report As String

    If Not modRoleUiAccess.CanCurrentUserPerformCapabilityCached("ADMIN_MAINT", report) Then
        ShowStatus report
        Exit Sub
    End If

    modNasConnection.SetRequireManualServerCredentials CBool(mChkManualServerCredentials.Value)
    ShowStatus "Server connection option saved for this Windows user."
End Sub

Private Sub LoadCarriers()
    Dim carriers As Variant
    Dim displayRows As Variant
    Dim idx As Long

    mLoading = True
    mLstCarriers.Clear
    carriers = modCarrierSettings.GetConfiguredCarriers()
    If Not IsEmpty(carriers) Then
        ReDim displayRows(0 To UBound(carriers) - 1, 0 To 0)
        For idx = LBound(carriers) To UBound(carriers)
            displayRows(idx - 1, 0) = CStr(carriers(idx))
        Next idx
        mLstCarriers.List = displayRows
    End If
    mLoading = False
End Sub

Private Sub mBtnAdd_Click()
    Dim carrierName As String

    carrierName = Trim$(CStr(mTxtCarrier.Value))
    If carrierName = "" Then
        ShowStatus "Enter a carrier."
        Exit Sub
    End If

    If modCarrierSettings.AddConfiguredCarrier(carrierName) Then
        mTxtCarrier.Value = ""
        LoadCarriers
        ShowStatus "Carrier added."
    Else
        ShowStatus "Carrier was not added."
    End If
End Sub

Private Sub mBtnRemove_Click()
    Dim carrierName As String

    If mLstCarriers.ListIndex < 0 Then
        ShowStatus "Select a carrier."
        Exit Sub
    End If

    carrierName = CStr(mLstCarriers.List(mLstCarriers.ListIndex, 0))
    If modCarrierSettings.RemoveConfiguredCarrier(carrierName) Then
        LoadCarriers
        ShowStatus "Carrier removed."
    Else
        ShowStatus "Carrier was not removed."
    End If
End Sub

Private Sub mBtnReset_Click()
    If MsgBox("Reset shipping carriers to defaults?", vbQuestion + vbYesNo, "invSys Settings") <> vbYes Then Exit Sub
    modCarrierSettings.ResetConfiguredCarriers
    LoadCarriers
    ShowStatus "Defaults restored."
End Sub

Private Sub mBtnClose_Click()
    Me.Hide
End Sub

Private Sub mLstCarriers_Click()
    If mLoading Then Exit Sub
    If mLstCarriers.ListIndex >= 0 Then mTxtCarrier.Value = CStr(mLstCarriers.List(mLstCarriers.ListIndex, 0))
End Sub

Private Sub LoadUoms()
    Dim uoms As Variant
    Dim displayRows As Variant
    Dim idx As Long

    mLoading = True
    mLstUoms.Clear
    uoms = modUomSettings.GetConfiguredUoms()
    If IsArray(uoms) Then
        ReDim displayRows(0 To UBound(uoms) - LBound(uoms), 0 To 0)
        For idx = LBound(uoms) To UBound(uoms)
            displayRows(idx - LBound(uoms), 0) = CStr(uoms(idx))
        Next idx
        mLstUoms.List = displayRows
    End If
    mLoading = False
End Sub

Private Sub mBtnUomAdd_Click()
    Dim report As String
    Dim uomName As String

    If Not modRoleUiAccess.CanCurrentUserPerformCapabilityCached("ADMIN_MAINT", report) Then
        ShowStatus report
        Exit Sub
    End If
    uomName = UCase$(Trim$(CStr(mTxtUom.Value)))
    If modUomSettings.AddConfiguredUom(uomName, report) Then
        mTxtUom.Value = ""
        LoadConfigRows
        LoadUoms
    End If
    ShowStatus report
End Sub

Private Sub mBtnUomRemove_Click()
    Dim report As String
    Dim uomName As String

    If mLstUoms.ListIndex < 0 Then
        ShowStatus "Select a UOM."
        Exit Sub
    End If
    If Not modRoleUiAccess.CanCurrentUserPerformCapabilityCached("ADMIN_MAINT", report) Then
        ShowStatus report
        Exit Sub
    End If
    uomName = CStr(mLstUoms.List(mLstUoms.ListIndex, 0))
    If modUomSettings.RemoveConfiguredUom(uomName, report) Then
        mTxtUom.Value = ""
        LoadConfigRows
        LoadUoms
    End If
    ShowStatus report
End Sub

Private Sub mBtnUomReset_Click()
    Dim report As String

    If Not modRoleUiAccess.CanCurrentUserPerformCapabilityCached("ADMIN_MAINT", report) Then
        ShowStatus report
        Exit Sub
    End If
    If MsgBox("Reset the warehouse UOM catalog to defaults?", vbQuestion + vbYesNo, "invSys Settings") <> vbYes Then Exit Sub
    If modUomSettings.ResetConfiguredUoms(report) Then
        LoadConfigRows
        LoadUoms
    End If
    ShowStatus report
End Sub

Private Sub mLstUoms_Click()
    If mLoading Then Exit Sub
    If mLstUoms.ListIndex >= 0 Then mTxtUom.Value = CStr(mLstUoms.List(mLstUoms.ListIndex, 0))
End Sub

Private Sub ShowStatus(ByVal message As String)
    If mLblStatus Is Nothing Then Exit Sub
    mLblStatus.Caption = message
End Sub

Private Function AddLabel(ByVal name As String, _
                          ByVal caption As String, _
                          ByVal leftPos As Single, _
                          ByVal topPos As Single, _
                          ByVal widthVal As Single, _
                          ByVal heightVal As Single, _
                          ByVal boldText As Boolean) As MSForms.Label
    Set AddLabel = Me.Controls.Add("Forms.Label.1", name, True)
    With AddLabel
        .Caption = caption
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
        .Font.Bold = boldText
    End With
End Function

Private Function AddTextBox(ByVal name As String, _
                            ByVal leftPos As Single, _
                            ByVal topPos As Single, _
                            ByVal widthVal As Single, _
                            ByVal heightVal As Single) As MSForms.TextBox
    Set AddTextBox = Me.Controls.Add("Forms.TextBox.1", name, True)
    With AddTextBox
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Function AddListBox(ByVal name As String, _
                            ByVal leftPos As Single, _
                            ByVal topPos As Single, _
                            ByVal widthVal As Single, _
                            ByVal heightVal As Single) As MSForms.ListBox
    Set AddListBox = Me.Controls.Add("Forms.ListBox.1", name, True)
    With AddListBox
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Function AddCheckBox(ByVal name As String, _
                             ByVal caption As String, _
                             ByVal leftPos As Single, _
                             ByVal topPos As Single, _
                             ByVal widthVal As Single, _
                             ByVal heightVal As Single) As MSForms.CheckBox
    Set AddCheckBox = Me.Controls.Add("Forms.CheckBox.1", name, True)
    With AddCheckBox
        .Caption = caption
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Function AddButton(ByVal name As String, _
                           ByVal caption As String, _
                           ByVal leftPos As Single, _
                           ByVal topPos As Single, _
                           ByVal widthVal As Single, _
                           ByVal heightVal As Single) As MSForms.CommandButton
    Set AddButton = Me.Controls.Add("Forms.CommandButton.1", name, True)
    With AddButton
        .Caption = caption
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function
