VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreateWarehouse 
   Caption         =   "UserForm1"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "frmCreateWarehouse.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCreateWarehouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mPathLocalTouched As Boolean
Private mLastSuggestedLocalPath As String
Private mFormBusy As Boolean
Private mLocalBootstrapComplete As Boolean
Private mCreatedSpec As modWarehouseBootstrap.WarehouseSpec

Private Const COLOR_ERROR As Long = 255
Private Const COLOR_SUCCESS As Long = 32768
Private Const COLOR_INFO As Long = 0

Private Sub UserForm_Initialize()
    Me.Caption = "Create Warehouse"
    Me.Width = 510
    Me.Height = 430
    Me.StartUpPosition = 1
    mFormBusy = True
    Me.txtStationId.Value = "S1"
    Me.txtAdminUser.Value = ResolveDefaultAdminUserForm()
    Me.txtPathSharePoint.Value = ResolveDefaultSharePointRootForm()
    Me.chkPublishInitial.Value = True
    mPathLocalTouched = False
    mLastSuggestedLocalPath = vbNullString
    RefreshSuggestedLocalPath True
    ClearValidationErrors
    ShowSummary "Enter the warehouse details, then click Create.", COLOR_INFO
    mFormBusy = False
End Sub

Private Sub txtWarehouseId_Change()
    If mFormBusy Then Exit Sub
    ClearErrorLabel Me.lblWarehouseIdError
    RefreshSuggestedLocalPath False
End Sub

Private Sub txtWarehouseName_Change()
    If mFormBusy Then Exit Sub
    ClearErrorLabel Me.lblWarehouseNameError
End Sub

Private Sub txtStationId_Change()
    If mFormBusy Then Exit Sub
    ClearErrorLabel Me.lblStationIdError
End Sub

Private Sub txtAdminUser_Change()
    If mFormBusy Then Exit Sub
    ClearErrorLabel Me.lblAdminUserError
End Sub

Private Sub txtPathLocal_Change()
    Dim currentValue As String

    If mFormBusy Then Exit Sub
    ClearErrorLabel Me.lblPathLocalError
    currentValue = Trim$(CStr(Me.txtPathLocal.Value))
    If mLastSuggestedLocalPath <> "" Then
        mPathLocalTouched = (StrComp(currentValue, mLastSuggestedLocalPath, vbTextCompare) <> 0)
    End If
End Sub

Private Sub txtPathSharePoint_Change()
    If mFormBusy Then Exit Sub
    ClearErrorLabel Me.lblPathSharePointError
End Sub

Private Sub chkPublishInitial_Click()
    If mFormBusy Then Exit Sub
    ClearErrorLabel Me.lblPathSharePointError
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    Dim spec As modWarehouseBootstrap.WarehouseSpec
    Dim summaryText As String

    If mLocalBootstrapComplete And StrComp(Me.btnOK.Caption, "Close", vbTextCompare) = 0 Then
        Unload Me
        Exit Sub
    End If

    ClearValidationErrors
    If Not BuildSpecFromForm(spec) Then
        ShowSummary "Fix the highlighted fields and try again.", COLOR_ERROR
        Exit Sub
    End If

    If mLocalBootstrapComplete Then
        spec.PathLocal = mCreatedSpec.PathLocal
        mCreatedSpec.PathSharePoint = spec.PathSharePoint
        If Not CBool(Me.chkPublishInitial.Value) Then
            ShowSummary "Local warehouse already exists. Check the publish box to retry SharePoint publish, or click Close.", COLOR_INFO
            Exit Sub
        End If

        If modWarehouseBootstrap.PublishInitialArtifacts(mCreatedSpec) Then
            summaryText = "Local warehouse already existed from this session." & vbCrLf & _
                          "Initial SharePoint publish complete." & vbCrLf & _
                          modWarehouseBootstrap.GetLastWarehouseBootstrapReport()
            MarkFormComplete summaryText, True
        Else
            summaryText = "Local warehouse was created, but SharePoint publish still failed." & vbCrLf & _
                          modWarehouseBootstrap.GetLastWarehouseBootstrapReport()
            Me.btnOK.Caption = "Retry Publish"
            Me.btnCancel.Caption = "Close"
            ShowSummary summaryText, COLOR_ERROR
        End If
        Exit Sub
    End If

    If modWarehouseBootstrap.WarehouseIdExists(spec.warehouseId) Then
        SetErrorCaption Me.lblWarehouseIdError, "WarehouseId already exists locally or on SharePoint."
        ShowSummary "Choose a different WarehouseId and try again.", COLOR_ERROR
        Exit Sub
    End If

    If Not modWarehouseBootstrap.BootstrapWarehouseLocal(spec) Then
        ShowSummary "Local bootstrap failed:" & vbCrLf & modWarehouseBootstrap.GetLastWarehouseBootstrapReport(), COLOR_ERROR
        Exit Sub
    End If

    mCreatedSpec = spec
    mLocalBootstrapComplete = True
    summaryText = "Local bootstrap complete for " & spec.warehouseId & "."

    If CBool(Me.chkPublishInitial.Value) Then
        If modWarehouseBootstrap.PublishInitialArtifacts(spec) Then
            summaryText = summaryText & vbCrLf & "Initial SharePoint publish complete." & vbCrLf & _
                          modWarehouseBootstrap.GetLastWarehouseBootstrapReport()
            MarkFormComplete summaryText, True
        Else
            mCreatedSpec = spec
            summaryText = summaryText & vbCrLf & "Initial SharePoint publish failed." & vbCrLf & _
                          modWarehouseBootstrap.GetLastWarehouseBootstrapReport() & vbCrLf & _
                          "Correct the SharePoint path or connectivity, then click Retry Publish."
            Me.btnOK.Caption = "Retry Publish"
            Me.btnCancel.Caption = "Close"
            ShowSummary summaryText, COLOR_ERROR
        End If
    Else
        summaryText = summaryText & vbCrLf & "Initial SharePoint publish skipped."
        MarkFormComplete summaryText, True
    End If
End Sub

Private Sub MarkFormComplete(ByVal summaryText As String, ByVal includeCloseHint As Boolean)
    If includeCloseHint Then summaryText = summaryText & vbCrLf & "Click Close to finish."
    Me.Tag = "COMPLETE"
    Me.btnOK.Caption = "Close"
    Me.btnCancel.Caption = "Close"
    ShowSummary summaryText, COLOR_SUCCESS
End Sub

Private Function BuildSpecFromForm(ByRef spec As modWarehouseBootstrap.WarehouseSpec) As Boolean
    Dim report As String
    Dim isValid As Boolean

    spec.warehouseId = Trim$(CStr(Me.txtWarehouseId.Value))
    spec.WarehouseName = Trim$(CStr(Me.txtWarehouseName.Value))
    spec.StationId = Trim$(CStr(Me.txtStationId.Value))
    spec.AdminUser = Trim$(CStr(Me.txtAdminUser.Value))
    spec.PathLocal = Trim$(CStr(Me.txtPathLocal.Value))
    spec.PathSharePoint = Trim$(CStr(Me.txtPathSharePoint.Value))

    isValid = modWarehouseBootstrap.ValidateWarehouseSpec(spec, report)
    If Not isValid Then SetErrorCaption Me.lblWarehouseIdError, report

    If spec.StationId = "" Then
        SetErrorCaption Me.lblStationIdError, "StationId is required."
        isValid = False
    End If
    If spec.AdminUser = "" Then
        SetErrorCaption Me.lblAdminUserError, "AdminUser is required."
        isValid = False
    End If
    If spec.PathLocal = "" Then
        SetErrorCaption Me.lblPathLocalError, "Local path is required."
        isValid = False
    End If
    If CBool(Me.chkPublishInitial.Value) And spec.PathSharePoint = "" Then
        SetErrorCaption Me.lblPathSharePointError, "SharePoint path is required when initial publish is enabled."
        isValid = False
    End If

    BuildSpecFromForm = isValid
End Function

Private Sub RefreshSuggestedLocalPath(ByVal forceApply As Boolean)
    Dim warehouseId As String
    Dim suggestedPath As String
    Dim currentValue As String

    warehouseId = Trim$(CStr(Me.txtWarehouseId.Value))
    suggestedPath = "C:\invSys"
    If warehouseId <> "" Then suggestedPath = suggestedPath & "\" & warehouseId

    currentValue = Trim$(CStr(Me.txtPathLocal.Value))
    If forceApply Or (Not mPathLocalTouched) Or currentValue = "" Or StrComp(currentValue, mLastSuggestedLocalPath, vbTextCompare) = 0 Then
        mFormBusy = True
        Me.txtPathLocal.Value = suggestedPath
        mFormBusy = False
        mPathLocalTouched = False
    End If
    mLastSuggestedLocalPath = suggestedPath
End Sub

Private Sub ClearValidationErrors()
    ClearErrorLabel Me.lblWarehouseIdError
    ClearErrorLabel Me.lblWarehouseNameError
    ClearErrorLabel Me.lblStationIdError
    ClearErrorLabel Me.lblAdminUserError
    ClearErrorLabel Me.lblPathLocalError
    ClearErrorLabel Me.lblPathSharePointError
End Sub

Private Sub ClearErrorLabel(ByVal lbl As MSForms.Label)
    If lbl Is Nothing Then Exit Sub
    lbl.Caption = ""
    lbl.foreColor = COLOR_ERROR
End Sub

Private Sub SetErrorCaption(ByVal lbl As MSForms.Label, ByVal messageText As String)
    If lbl Is Nothing Then Exit Sub
    lbl.Caption = Trim$(messageText)
    lbl.foreColor = COLOR_ERROR
End Sub

Private Sub ShowSummary(ByVal messageText As String, ByVal foreColor As Long)
    Me.lblSummary.Caption = Trim$(messageText)
    Me.lblSummary.foreColor = foreColor
End Sub

Private Function ResolveDefaultAdminUserForm() As String
    ResolveDefaultAdminUserForm = Trim$(Environ$("USERNAME"))
End Function

Private Function ResolveDefaultSharePointRootForm() As String
    On Error Resume Next
    If modConfig.IsLoaded() Then
        ResolveDefaultSharePointRootForm = Trim$(CStr(modConfig.GetString("PathSharePointRoot", "")))
    End If
    On Error GoTo 0
End Function

