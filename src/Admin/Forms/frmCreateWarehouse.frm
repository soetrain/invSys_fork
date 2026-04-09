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

Private WithEvents mBtnSharePointHelper As MSForms.CommandButton
Attribute mBtnSharePointHelper.VB_VarHelpID = -1

Private mPathLocalTouched As Boolean
Private mLastSuggestedLocalPath As String
Private mFormBusy As Boolean
Private mLocalBootstrapComplete As Boolean
Private mCreatedWarehouseId As String
Private mCreatedWarehouseName As String
Private mCreatedStationId As String
Private mCreatedAdminUser As String
Private mCreatedPathLocal As String
Private mCreatedPathSharePoint As String
Private mDefaultOkLeft As Single
Private mDefaultOkWidth As Single
Private mDefaultCancelLeft As Single
Private mDefaultCancelWidth As Single

Private Const COLOR_ERROR As Long = 255
Private Const COLOR_SUCCESS As Long = 32768
Private Const COLOR_INFO As Long = 0

Private Sub UserForm_Initialize()
    Me.Caption = "Create Warehouse"
    Me.Width = 620
    Me.Height = 470
    Me.StartUpPosition = 1
    mDefaultOkLeft = Me.btnOK.Left
    mDefaultOkWidth = Me.btnOK.Width
    mDefaultCancelLeft = Me.btnCancel.Left
    mDefaultCancelWidth = Me.btnCancel.Width
    ConfigureSummaryArea
    RestoreDefaultButtonLayout
    mFormBusy = True
    Me.txtStationId.Value = "S1"
    Me.txtAdminUser.Value = ResolveDefaultAdminUserForm()
    Me.txtPathSharePoint.Value = ResolveDefaultSharePointRootForm()
    ConfigureSharePointHelperButton
    Me.btnOK.Caption = "Create"
    Me.btnCancel.Caption = "Cancel"
    Me.chkPublishInitial.Value = True
    mPathLocalTouched = False
    mLastSuggestedLocalPath = vbNullString
    RefreshSuggestedLocalPath True
    ClearValidationErrors
    ShowSummary "Pick the locally synced invSys root that contains Addins, Events, Snapshots, and TesterPackage, then click Create.", COLOR_INFO
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

Private Sub mBtnSharePointHelper_Click()
    Dim candidate As String
    Dim warehouseId As String

    warehouseId = Trim$(CStr(Me.txtWarehouseId.Value))
    candidate = modTesterSetup.DetectSharePointRoot(warehouseId)
    If candidate = "" Then
        candidate = modTesterSetup.BrowseForSharePointRoot(Trim$(CStr(Me.txtPathSharePoint.Value)))
    End If

    If candidate = "" Then
        ShowSummary "SharePoint root was not detected. Pick the locally synced invSys root folder manually.", COLOR_INFO
        Exit Sub
    End If

    mFormBusy = True
    Me.txtPathSharePoint.Value = candidate
    mFormBusy = False
    ClearErrorLabel Me.lblPathSharePointError
    ShowSummary "SharePoint root detected.", COLOR_SUCCESS
End Sub

Private Sub chkPublishInitial_Click()
    If mFormBusy Then Exit Sub
    ClearErrorLabel Me.lblPathSharePointError
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    Dim warehouseId As String
    Dim warehouseName As String
    Dim stationId As String
    Dim adminUser As String
    Dim pathLocal As String
    Dim pathSharePoint As String
    Dim summaryText As String

    If mLocalBootstrapComplete And StrComp(Me.btnOK.Caption, "Close", vbTextCompare) = 0 Then
        Unload Me
        Exit Sub
    End If

    ClearValidationErrors
    If Not BuildSpecFromForm(warehouseId, warehouseName, stationId, adminUser, pathLocal, pathSharePoint) Then
        ShowSummary "Fix the highlighted fields and try again.", COLOR_ERROR
        Exit Sub
    End If

    If Not modLocalAddinsRegistration.EnsureLocalInvSysAddinsRegistered(pathSharePoint & "\Addins", summaryText) Then
        ShowSummary "invSys add-ins are not registered cleanly for this Excel session." & vbCrLf & summaryText, COLOR_ERROR
        Exit Sub
    End If

    If mLocalBootstrapComplete Then
        pathLocal = mCreatedPathLocal
        mCreatedPathSharePoint = pathSharePoint
        If Not CBool(Me.chkPublishInitial.Value) Then
            ShowSummary "Local warehouse already exists. Check the publish box to retry SharePoint publish, or click Close.", COLOR_INFO
            Exit Sub
        End If

        If modAdminConsole.PublishInitialArtifactsAdmin(mCreatedWarehouseId, mCreatedWarehouseName, mCreatedStationId, mCreatedAdminUser, mCreatedPathLocal, mCreatedPathSharePoint) Then
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

    If modWarehouseBootstrap.WarehouseIdExists(warehouseId) Then
        SetErrorCaption Me.lblWarehouseIdError, "WarehouseId already exists locally or on SharePoint."
        ShowSummary "Choose a different WarehouseId and try again.", COLOR_ERROR
        Exit Sub
    End If

    If Not modAdminConsole.BootstrapWarehouseLocalAdmin(warehouseId, warehouseName, stationId, adminUser, pathLocal, pathSharePoint) Then
        ShowSummary "Local bootstrap failed:" & vbCrLf & modWarehouseBootstrap.GetLastWarehouseBootstrapReport(), COLOR_ERROR
        Exit Sub
    End If

    mCreatedWarehouseId = warehouseId
    mCreatedWarehouseName = warehouseName
    mCreatedStationId = stationId
    mCreatedAdminUser = adminUser
    mCreatedPathLocal = pathLocal
    mCreatedPathSharePoint = pathSharePoint
    mLocalBootstrapComplete = True
    summaryText = "Local bootstrap complete for " & warehouseId & "."

    If CBool(Me.chkPublishInitial.Value) Then
        If modAdminConsole.PublishInitialArtifactsAdmin(warehouseId, warehouseName, stationId, adminUser, pathLocal, pathSharePoint) Then
            summaryText = summaryText & vbCrLf & "Initial SharePoint publish complete." & vbCrLf & _
                          modWarehouseBootstrap.GetLastWarehouseBootstrapReport()
            MarkFormComplete summaryText, True
        Else
            mCreatedPathSharePoint = pathSharePoint
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
    ConfigureCompleteButtonLayout
    ShowSummary summaryText, COLOR_SUCCESS
End Sub

Private Function BuildSpecFromForm(ByRef warehouseId As String, _
                                   ByRef warehouseName As String, _
                                   ByRef stationId As String, _
                                   ByRef adminUser As String, _
                                   ByRef pathLocal As String, _
                                   ByRef pathSharePoint As String) As Boolean
    Dim report As String
    Dim isValid As Boolean

    warehouseId = Trim$(CStr(Me.txtWarehouseId.Value))
    warehouseName = Trim$(CStr(Me.txtWarehouseName.Value))
    stationId = Trim$(CStr(Me.txtStationId.Value))
    adminUser = Trim$(CStr(Me.txtAdminUser.Value))
    pathLocal = Trim$(CStr(Me.txtPathLocal.Value))
    pathSharePoint = Trim$(CStr(Me.txtPathSharePoint.Value))

    isValid = modAdminConsole.ValidateWarehouseSpecAdmin(warehouseId, warehouseName, stationId, adminUser, pathLocal, pathSharePoint, report)
    If Not isValid Then SetErrorCaption Me.lblWarehouseIdError, report

    If stationId = "" Then
        SetErrorCaption Me.lblStationIdError, "StationId is required."
        isValid = False
    End If
    If adminUser = "" Then
        SetErrorCaption Me.lblAdminUserError, "AdminUser is required."
        isValid = False
    End If
    If pathLocal = "" Then
        SetErrorCaption Me.lblPathLocalError, "Local path is required."
        isValid = False
    End If
    If CBool(Me.chkPublishInitial.Value) And pathSharePoint = "" Then
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
    Me.lblSummary.Caption = FormatSummaryMessage(messageText)
    Me.lblSummary.foreColor = foreColor
    Me.Repaint
End Sub

Private Sub ConfigureSummaryArea()
    With Me.lblSummary
        .WordWrap = True
        .AutoSize = False
        .Left = 18
        .Top = 282
        .Width = Me.InsideWidth - 36
        .Height = 108
    End With
End Sub

Private Sub ConfigureCompleteButtonLayout()
    Me.btnCancel.Visible = False
    Me.btnCancel.Enabled = False
    Me.btnOK.Width = 96
    Me.btnOK.Left = Me.InsideWidth - Me.btnOK.Width - 24
End Sub

Private Sub RestoreDefaultButtonLayout()
    Me.btnCancel.Visible = True
    Me.btnCancel.Enabled = True
    Me.btnCancel.Left = mDefaultCancelLeft
    Me.btnCancel.Width = mDefaultCancelWidth
    Me.btnOK.Left = mDefaultOkLeft
    Me.btnOK.Width = mDefaultOkWidth
End Sub

Private Function FormatSummaryMessage(ByVal messageText As String) As String
    Dim lines As Collection
    Dim parsed As String
    Dim lineValue As Variant

    messageText = Trim$(messageText)
    If messageText = "" Then Exit Function

    parsed = FormatBootstrapReportForSummary(messageText)
    If parsed <> "" Then
        FormatSummaryMessage = parsed
        Exit Function
    End If

    Set lines = New Collection
    AppendWrappedSummaryLine lines, messageText, 92

    For Each lineValue In lines
        If FormatSummaryMessage <> "" Then FormatSummaryMessage = FormatSummaryMessage & vbCrLf
        FormatSummaryMessage = FormatSummaryMessage & CStr(lineValue)
    Next lineValue
End Function

Private Function FormatBootstrapReportForSummary(ByVal reportText As String) As String
    Dim parts() As String
    Dim i As Long
    Dim keyText As String
    Dim valueText As String
    Dim lines As Collection
    Dim lineValue As Variant
    Dim eqPos As Long

    reportText = Trim$(reportText)
    If reportText = "" Then Exit Function
    If InStr(1, reportText, "|", vbBinaryCompare) = 0 Then Exit Function

    parts = Split(reportText, "|")
    Set lines = New Collection

    If UBound(parts) >= 0 Then
        If StrComp(Trim$(parts(0)), "OK", vbTextCompare) = 0 Then
            AppendWrappedSummaryLine lines, "Status: OK", 92
        Else
            AppendWrappedSummaryLine lines, Trim$(parts(0)), 92
        End If
    End If

    For i = 1 To UBound(parts)
        eqPos = InStr(1, parts(i), "=", vbBinaryCompare)
        If eqPos > 1 Then
            keyText = Trim$(Left$(parts(i), eqPos - 1))
            valueText = Trim$(Mid$(parts(i), eqPos + 1))
            If valueText <> "" Then
                valueText = Replace$(valueText, "COPIED:", "copied to ", 1, 1, vbTextCompare)
                valueText = Replace$(valueText, "SKIPPED", "already current", 1, 1, vbTextCompare)
                AppendWrappedSummaryLine lines, FriendlySummaryLabel(keyText) & ": " & valueText, 92
            End If
        ElseIf Trim$(parts(i)) <> "" Then
            AppendWrappedSummaryLine lines, Trim$(parts(i)), 92
        End If
    Next i

    For Each lineValue In lines
        If FormatBootstrapReportForSummary <> "" Then FormatBootstrapReportForSummary = FormatBootstrapReportForSummary & vbCrLf
        FormatBootstrapReportForSummary = FormatBootstrapReportForSummary & CStr(lineValue)
    Next lineValue
End Function

Private Function FriendlySummaryLabel(ByVal keyText As String) As String
    Select Case UCase$(Trim$(keyText))
        Case "CONFIG"
            FriendlySummaryLabel = "Config artifact"
        Case "DISCOVERY"
            FriendlySummaryLabel = "Discovery file"
        Case Else
            FriendlySummaryLabel = Trim$(keyText)
    End Select
End Function

Private Sub AppendWrappedSummaryLine(ByVal lines As Collection, ByVal textLine As String, ByVal maxChars As Long)
    Dim working As String
    Dim breakPos As Long

    working = Trim$(textLine)
    If working = "" Then
        lines.Add vbNullString
        Exit Sub
    End If

    Do While Len(working) > maxChars
        breakPos = InStrRev(Left$(working, maxChars), " ")
        If breakPos <= 0 Then breakPos = maxChars
        lines.Add Trim$(Left$(working, breakPos))
        working = Trim$(Mid$(working, breakPos + 1))
    Loop

    If working <> "" Then lines.Add working
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

    If ResolveDefaultSharePointRootForm = "" Then
        ResolveDefaultSharePointRootForm = modTesterSetup.DetectSharePointRoot(Trim$(CStr(Me.txtWarehouseId.Value)))
    End If
End Function

Private Sub ConfigureSharePointHelperButton()
    If mBtnSharePointHelper Is Nothing Then
        Set mBtnSharePointHelper = Me.Controls.Add("Forms.CommandButton.1", "btnSharePointHelperRuntime", True)
    End If

    Me.txtPathSharePoint.Width = 258
    With mBtnSharePointHelper
        .Caption = "Find..."
        .Left = Me.txtPathSharePoint.Left + Me.txtPathSharePoint.Width + 8
        .Top = Me.txtPathSharePoint.Top - 1
        .Width = 72
        .Height = Me.txtPathSharePoint.Height + 2
        .ControlTipText = "Choose the locally synced invSys SharePoint root folder."
        .Visible = True
        .Enabled = True
    End With
End Sub

