VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRetireMigrateWarehouse 
   Caption         =   "Retire / Migrate Warehouse"
   ClientHeight    =   4200
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6400
   OleObjectBlob   =   "frmRetireMigrateWarehouse.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRetireMigrateWarehouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PANEL_SELECTION As String = "SELECTION"
Private Const PANEL_CONFIRM As String = "CONFIRM"
Private Const PANEL_RESULT As String = "RESULT"

Private Const COLOR_ERROR As Long = 255
Private Const COLOR_SUCCESS As Long = 32768
Private Const COLOR_INFO As Long = 0
Private Const COLOR_WARNING As Long = 192

Private mFormBusy As Boolean
Private mCurrentPanel As String
Private mReAuthPassed As Boolean
Private mPendingSourceWarehouseId As String
Private mPendingTargetWarehouseId As String
Private mPendingOperationMode As Long
Private mPendingAdminUser As String
Private mPendingArchiveDestPath As String
Private mPendingPublishTombstone As Boolean

Private Sub UserForm_Initialize()
    mFormBusy = True
    Me.Caption = "Retire / Migrate Warehouse"
    Me.StartUpPosition = 1
    ConfigureRetireMigrateLayout

    Me.optArchiveOnly.Value = True
    Me.chkPublishTombstone.Value = True
    Me.chkConfirmAction.Value = False
    Me.lblDeleteWarning.ForeColor = COLOR_ERROR

    ClearAllInlineErrors
    PopulateWarehouseDropdowns
    ApplyDefaultSelections
    ShowSelectionPanel
    ShowFormMessage "Select a source warehouse and operation mode, then click OK.", COLOR_INFO

    mFormBusy = False
End Sub

Private Sub ConfigureRetireMigrateLayout()
    Me.Width = 760
    Me.Height = 620

    Me.lblTitle.Left = 18
    Me.lblTitle.Top = 18
    Me.lblTitle.Width = 360

    Me.lblSelectionIntro.Left = 18
    Me.lblSelectionIntro.Top = 52
    Me.lblSelectionIntro.Width = 700
    Me.lblSelectionIntro.Height = 36

    Me.lblSourceWarehouse.Left = 18
    Me.lblSourceWarehouse.Top = 112
    Me.lblSourceWarehouse.Width = 180

    Me.cmbSourceWarehouse.Left = 220
    Me.cmbSourceWarehouse.Top = 108
    Me.cmbSourceWarehouse.Width = 190

    Me.lblSourceWarehouseError.Left = 220
    Me.lblSourceWarehouseError.Top = 134
    Me.lblSourceWarehouseError.Width = 460

    Me.lblTargetWarehouse.Left = 18
    Me.lblTargetWarehouse.Top = 164
    Me.lblTargetWarehouse.Width = 180

    Me.cmbTargetWarehouse.Left = 220
    Me.cmbTargetWarehouse.Top = 160
    Me.cmbTargetWarehouse.Width = 190

    Me.lblTargetWarehouseError.Left = 220
    Me.lblTargetWarehouseError.Top = 186
    Me.lblTargetWarehouseError.Width = 460

    Me.fraMode.Left = 18
    Me.fraMode.Top = 220
    Me.fraMode.Width = 700
    Me.fraMode.Height = 138

    Me.optArchiveOnly.Left = 18
    Me.optArchiveOnly.Top = 24
    Me.optArchiveOnly.Width = 250
    Me.optArchiveMigrate.Left = 18
    Me.optArchiveMigrate.Top = 48
    Me.optArchiveMigrate.Width = 250
    Me.optArchiveRetire.Left = 18
    Me.optArchiveRetire.Top = 72
    Me.optArchiveRetire.Width = 250
    Me.optArchiveRetireDelete.Left = 18
    Me.optArchiveRetireDelete.Top = 96
    Me.optArchiveRetireDelete.Width = 300

    Me.lblArchiveDestPath.Left = 18
    Me.lblArchiveDestPath.Top = 376
    Me.lblArchiveDestPath.Width = 180
    Me.txtArchiveDestPath.Left = 220
    Me.txtArchiveDestPath.Top = 372
    Me.txtArchiveDestPath.Width = 498
    Me.lblArchiveDestPathError.Left = 220
    Me.lblArchiveDestPathError.Top = 398
    Me.lblArchiveDestPathError.Width = 498

    Me.chkPublishTombstone.Left = 220
    Me.chkPublishTombstone.Top = 428
    Me.chkPublishTombstone.Width = 340
    Me.lblReAuthError.Left = 220
    Me.lblReAuthError.Top = 454
    Me.lblReAuthError.Width = 498
    Me.lblDeleteWarning.Left = 220
    Me.lblDeleteWarning.Top = 478
    Me.lblDeleteWarning.Width = 498

    Me.fraConfirm.Left = 18
    Me.fraConfirm.Top = 96
    Me.fraConfirm.Width = 700
    Me.fraConfirm.Height = 430
    Me.lblConfirmSummary.Left = 18
    Me.lblConfirmSummary.Top = 24
    Me.lblConfirmSummary.Width = 660
    Me.lblConfirmSummary.Height = 170
    Me.chkConfirmAction.Left = 18
    Me.chkConfirmAction.Top = 210
    Me.chkConfirmAction.Width = 320
    Me.lblConfirmError.Left = 18
    Me.lblConfirmError.Top = 238
    Me.lblConfirmError.Width = 660
    Me.lblDeleteWarning.Left = 18
    Me.lblDeleteWarning.Width = 660

    Me.fraResult.Left = 18
    Me.fraResult.Top = 96
    Me.fraResult.Width = 700
    Me.fraResult.Height = 430
    Me.lblResultSummary.Left = 18
    Me.lblResultSummary.Top = 24
    Me.lblResultSummary.Width = 660
    Me.lblResultSummary.Height = 360

    Me.btnBack.Left = 396
    Me.btnBack.Top = 540
    Me.btnBack.Width = 88
    Me.btnCancel.Left = 494
    Me.btnCancel.Top = 540
    Me.btnCancel.Width = 88
    Me.btnOK.Left = 592
    Me.btnOK.Top = 540
    Me.btnOK.Width = 88
End Sub

Private Sub cmbSourceWarehouse_Change()
    If mFormBusy Then Exit Sub
    ClearInlineError Me.lblSourceWarehouseError
    SuggestArchiveDestination False
End Sub

Private Sub cmbTargetWarehouse_Change()
    If mFormBusy Then Exit Sub
    ClearInlineError Me.lblTargetWarehouseError
End Sub

Private Sub txtArchiveDestPath_Change()
    If mFormBusy Then Exit Sub
    ClearInlineError Me.lblArchiveDestPathError
End Sub

Private Sub optArchiveOnly_Click()
    If mFormBusy Then Exit Sub
    UpdateModeUi
End Sub

Private Sub optArchiveMigrate_Click()
    If mFormBusy Then Exit Sub
    UpdateModeUi
End Sub

Private Sub optArchiveRetire_Click()
    If mFormBusy Then Exit Sub
    UpdateModeUi
End Sub

Private Sub optArchiveRetireDelete_Click()
    If mFormBusy Then Exit Sub
    UpdateModeUi
End Sub

Private Sub chkConfirmAction_Click()
    If mCurrentPanel <> PANEL_CONFIRM Then Exit Sub
    ClearInlineError Me.lblConfirmError
    UpdateConfirmOkState
End Sub

Private Sub btnBack_Click()
    If mCurrentPanel = PANEL_CONFIRM Then
        ShowSelectionPanel
        ShowFormMessage "Adjust the selection, then click OK to continue.", COLOR_INFO
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    Select Case mCurrentPanel
        Case PANEL_SELECTION
            HandleSelectionOk
        Case PANEL_CONFIRM
            HandleConfirmOk
        Case PANEL_RESULT
            Unload Me
    End Select
End Sub

Private Sub HandleSelectionOk()
    Dim sourceWarehouseId As String
    Dim targetWarehouseId As String
    Dim operationMode As Long
    Dim adminUser As String
    Dim archiveDestPath As String
    Dim publishTombstone As Boolean

    ClearAllInlineErrors
    If Not BuildSpecFromSelection(sourceWarehouseId, targetWarehouseId, operationMode, adminUser, archiveDestPath, publishTombstone) Then
        ShowFormMessage "Fix the highlighted fields and try again.", COLOR_ERROR
        Exit Sub
    End If

    If Not modWarehouseRetire.RequireReAuth("ADMIN_MAINT") Then
        SetInlineError Me.lblReAuthError, "Re-authentication required to continue"
        ShowFormMessage "Re-authentication required to continue.", COLOR_ERROR
        Exit Sub
    End If

    mPendingSourceWarehouseId = sourceWarehouseId
    mPendingTargetWarehouseId = targetWarehouseId
    mPendingOperationMode = operationMode
    mPendingAdminUser = adminUser
    mPendingArchiveDestPath = archiveDestPath
    mPendingPublishTombstone = publishTombstone
    mReAuthPassed = True
    ShowConfirmPanel
End Sub

Private Sub HandleConfirmOk()
    Dim summaryText As String
    Dim failureText As String

    ClearInlineError Me.lblConfirmError
    If Not mReAuthPassed Then
        SetInlineError Me.lblConfirmError, "Re-authentication required to continue"
        Exit Sub
    End If
    If Not CBool(Me.chkConfirmAction.Value) Then
        SetInlineError Me.lblConfirmError, "You must confirm this action before continuing."
        Exit Sub
    End If

    If Not modAdminConsole.ValidateRetireMigrateSpecAdmin(mPendingSourceWarehouseId, mPendingTargetWarehouseId, mPendingOperationMode, mPendingAdminUser, True, mPendingArchiveDestPath, mPendingPublishTombstone, failureText) Then
        SetInlineError Me.lblConfirmError, failureText
        Exit Sub
    End If

    If Not modAdminConsole.WriteArchivePackageAdmin(mPendingSourceWarehouseId, mPendingTargetWarehouseId, mPendingOperationMode, mPendingAdminUser, True, mPendingArchiveDestPath, mPendingPublishTombstone) Then
        ShowResultPanel False, "WriteArchivePackage failed: " & modWarehouseRetire.GetLastWarehouseRetireReport()
        Exit Sub
    End If
    summaryText = "WriteArchivePackage: " & modWarehouseRetire.GetLastWarehouseRetireReport()

    If mPendingOperationMode = modWarehouseRetire.MODE_ARCHIVE_MIGRATE Then
        If Not modAdminConsole.MigrateInventoryToTargetAdmin(mPendingSourceWarehouseId, mPendingTargetWarehouseId, mPendingOperationMode, mPendingAdminUser, True, mPendingArchiveDestPath, mPendingPublishTombstone) Then
            ShowResultPanel False, "MigrateInventoryToTarget failed: " & modWarehouseRetire.GetLastWarehouseRetireReport()
            Exit Sub
        End If
        summaryText = summaryText & vbCrLf & "MigrateInventoryToTarget: " & modWarehouseRetire.GetLastWarehouseRetireReport()
    End If

    If mPendingOperationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE Or _
       mPendingOperationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE Then
        If Not modAdminConsole.RetireSourceWarehouseAdmin(mPendingSourceWarehouseId, mPendingTargetWarehouseId, mPendingOperationMode, mPendingAdminUser, True, mPendingArchiveDestPath, mPendingPublishTombstone) Then
            ShowResultPanel False, "RetireSourceWarehouse failed: " & modWarehouseRetire.GetLastWarehouseRetireReport()
            Exit Sub
        End If
        summaryText = summaryText & vbCrLf & "RetireSourceWarehouse: " & modWarehouseRetire.GetLastWarehouseRetireReport()
    End If

    If mPendingOperationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE Then
        If Not modAdminConsole.DeleteLocalRuntimeAdmin(mPendingSourceWarehouseId, mPendingTargetWarehouseId, mPendingOperationMode, mPendingAdminUser, True, mPendingArchiveDestPath, mPendingPublishTombstone) Then
            ShowResultPanel False, "DeleteLocalRuntime failed: " & modWarehouseRetire.GetLastWarehouseRetireReport()
            Exit Sub
        End If
        summaryText = summaryText & vbCrLf & "DeleteLocalRuntime: " & modWarehouseRetire.GetLastWarehouseRetireReport()
    End If

    ShowResultPanel True, summaryText
End Sub

Private Function BuildSpecFromSelection(ByRef sourceWarehouseId As String, _
                                        ByRef targetWarehouseId As String, _
                                        ByRef operationMode As Long, _
                                        ByRef adminUser As String, _
                                        ByRef archiveDestPath As String, _
                                        ByRef publishTombstone As Boolean) As Boolean
    Dim isValid As Boolean

    sourceWarehouseId = Trim$(CStr(Me.cmbSourceWarehouse.Value))
    targetWarehouseId = Trim$(CStr(Me.cmbTargetWarehouse.Value))
    operationMode = ResolveSelectedMode()
    adminUser = ResolveCurrentUserForm()
    archiveDestPath = Trim$(CStr(Me.txtArchiveDestPath.Value))
    publishTombstone = CBool(Me.chkPublishTombstone.Value)

    If sourceWarehouseId = "" Then
        SetInlineError Me.lblSourceWarehouseError, "Source warehouse is required."
        isValid = False
    Else
        isValid = True
    End If

    If operationMode = modWarehouseRetire.MODE_ARCHIVE_MIGRATE Then
        If targetWarehouseId = "" Then
            SetInlineError Me.lblTargetWarehouseError, "Target warehouse is required for migrate mode."
            isValid = False
        ElseIf StrComp(sourceWarehouseId, targetWarehouseId, vbTextCompare) = 0 Then
            SetInlineError Me.lblTargetWarehouseError, "Target warehouse must be different from the source."
            isValid = False
        End If
    End If

    If archiveDestPath = "" Then
        SetInlineError Me.lblArchiveDestPathError, "Archive destination path is required."
        isValid = False
    End If

    If isValid Then
        If Not ValidateSelectionSpecForm(sourceWarehouseId, targetWarehouseId, operationMode, adminUser, archiveDestPath, publishTombstone) Then
            isValid = False
        End If
    End If

    BuildSpecFromSelection = isValid
End Function

Private Function ValidateSelectionSpecForm(ByVal sourceWarehouseId As String, _
                                           ByVal targetWarehouseId As String, _
                                           ByVal operationMode As Long, _
                                           ByVal adminUser As String, _
                                           ByVal archiveDestPath As String, _
                                           ByVal publishTombstone As Boolean) As Boolean
    Dim report As String

    If modAdminConsole.ValidateRetireMigrateSpecAdmin(sourceWarehouseId, targetWarehouseId, operationMode, adminUser, True, archiveDestPath, publishTombstone, report) Then
        ValidateSelectionSpecForm = True
        Exit Function
    End If

    If InStr(1, report, "SourceWarehouseId", vbTextCompare) > 0 Then
        SetInlineError Me.lblSourceWarehouseError, report
    ElseIf InStr(1, report, "TargetWarehouseId", vbTextCompare) > 0 Then
        SetInlineError Me.lblTargetWarehouseError, report
    ElseIf InStr(1, report, "ArchiveDestPath", vbTextCompare) > 0 Then
        SetInlineError Me.lblArchiveDestPathError, report
    Else
        ShowFormMessage report, COLOR_ERROR
    End If
End Function

Private Sub PopulateWarehouseDropdowns()
    Dim warehouseIds As Collection
    Dim item As Variant

    Set warehouseIds = DiscoverWarehouseIdsForm()
    Me.cmbSourceWarehouse.Clear
    Me.cmbTargetWarehouse.Clear

    For Each item In warehouseIds
        Me.cmbSourceWarehouse.AddItem CStr(item)
        Me.cmbTargetWarehouse.AddItem CStr(item)
    Next item
End Sub

Private Function DiscoverWarehouseIdsForm() As Collection
    Dim results As Collection
    Dim seen As Object
    Dim rootPath As String
    Dim scanRoots As Collection
    Dim candidate As Variant

    Set results = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Set scanRoots = New Collection

    rootPath = ResolveWarehouseScanRootForm()
    If rootPath <> "" Then scanRoots.Add rootPath
    If StrComp(rootPath, "C:\invSys", vbTextCompare) <> 0 Then scanRoots.Add "C:\invSys"

    For Each candidate In scanRoots
        AddWarehousesFromRootForm results, seen, CStr(candidate)
    Next candidate

    Set DiscoverWarehouseIdsForm = results
End Function

Private Function ResolveWarehouseScanRootForm() As String
    Dim rootPath As String
    Dim parentPath As String

    rootPath = Trim$(modRuntimeWorkbooks.GetCoreDataRootOverride())
    If rootPath = "" Then rootPath = Trim$(modRuntimeWorkbooks.ResolveCoreDataRoot("", ""))
    rootPath = NormalizePathForm(rootPath)
    If rootPath = "" Then
        ResolveWarehouseScanRootForm = "C:\invSys"
        Exit Function
    End If

    If LooksLikeWarehouseRuntimeRootForm(rootPath) Then
        parentPath = GetParentFolderForm(rootPath)
        If parentPath <> "" Then
            ResolveWarehouseScanRootForm = parentPath
            Exit Function
        End If
    End If

    parentPath = GetParentFolderForm(rootPath)
    If parentPath = "" Or StrComp(parentPath, "C:", vbTextCompare) = 0 Then
        ResolveWarehouseScanRootForm = rootPath
    Else
        ResolveWarehouseScanRootForm = parentPath
    End If
End Function

Private Sub AddWarehousesFromRootForm(ByVal results As Collection, ByVal seen As Object, ByVal rootPath As String)
    Dim fso As Object
    Dim rootFolder As Object
    Dim subFolder As Object
    Dim folderName As String
    Dim configPath As String

    rootPath = NormalizePathForm(rootPath)
    If rootPath = "" Then Exit Sub
    If Not FolderExistsForm(rootPath) Then Exit Sub

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Sub
    Set rootFolder = fso.GetFolder(rootPath)
    On Error GoTo 0
    If rootFolder Is Nothing Then Exit Sub

    For Each subFolder In rootFolder.SubFolders
        folderName = CStr(subFolder.Name)
        configPath = NormalizePathForm(CStr(subFolder.Path)) & "\" & folderName & ".invSys.Config.xlsb"
        If FileExistsForm(configPath) Then
            If Not seen.Exists(folderName) Then
                seen.Add folderName, True
                results.Add folderName
            End If
        End If
    Next subFolder
End Sub

Private Sub ApplyDefaultSelections()
    If Me.cmbSourceWarehouse.ListCount > 0 Then
        Me.cmbSourceWarehouse.ListIndex = 0
    End If
    If Me.cmbTargetWarehouse.ListCount > 0 Then
        Me.cmbTargetWarehouse.ListIndex = 0
    End If
    SuggestArchiveDestination True
    UpdateModeUi
End Sub

Private Sub SuggestArchiveDestination(ByVal forceApply As Boolean)
    Dim suggestedPath As String

    suggestedPath = ResolveArchiveDefaultForm(Trim$(CStr(Me.cmbSourceWarehouse.Value)))
    If forceApply Or Trim$(CStr(Me.txtArchiveDestPath.Value)) = "" Then
        Me.txtArchiveDestPath.Value = suggestedPath
    End If
End Sub

Private Function ResolveArchiveDefaultForm(ByVal warehouseId As String) As String
    Dim priorRoot As String
    Dim pathValue As String

    If warehouseId = "" Then
        ResolveArchiveDefaultForm = "C:\invSys\Archive"
        Exit Function
    End If

    priorRoot = modRuntimeWorkbooks.GetCoreDataRootOverride()
    On Error Resume Next
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    If modConfig.LoadConfig(warehouseId, "") Then
        pathValue = Trim$(modConfig.GetString("PathBackupRoot", ""))
    End If
    On Error GoTo 0
    RestoreRootOverrideForm priorRoot

    pathValue = NormalizePathForm(pathValue)
    If pathValue = "" Then
        ResolveArchiveDefaultForm = "C:\invSys\Archive"
    Else
        ResolveArchiveDefaultForm = pathValue
    End If
End Function

Private Sub UpdateModeUi()
    Dim migrateMode As Boolean
    Dim retireMode As Boolean
    Dim deleteMode As Boolean

    migrateMode = (ResolveSelectedMode() = modWarehouseRetire.MODE_ARCHIVE_MIGRATE)
    retireMode = (ResolveSelectedMode() = modWarehouseRetire.MODE_ARCHIVE_RETIRE Or ResolveSelectedMode() = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE)
    deleteMode = (ResolveSelectedMode() = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE)

    Me.cmbTargetWarehouse.Enabled = migrateMode
    Me.lblTargetWarehouse.Enabled = migrateMode
    Me.chkPublishTombstone.Visible = retireMode
    Me.chkPublishTombstone.Enabled = retireMode
    If Not retireMode Then Me.chkPublishTombstone.Value = False

    Me.lblDeleteWarning.Visible = deleteMode
    ClearInlineError Me.lblTargetWarehouseError
    ClearInlineError Me.lblReAuthError
End Sub

Private Function ResolveSelectedMode() As Long
    If CBool(Me.optArchiveMigrate.Value) Then
        ResolveSelectedMode = modWarehouseRetire.MODE_ARCHIVE_MIGRATE
    ElseIf CBool(Me.optArchiveRetire.Value) Then
        ResolveSelectedMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE
    ElseIf CBool(Me.optArchiveRetireDelete.Value) Then
        ResolveSelectedMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE
    Else
        ResolveSelectedMode = modWarehouseRetire.MODE_ARCHIVE_ONLY
    End If
End Function

Private Sub ShowSelectionPanel()
    mCurrentPanel = PANEL_SELECTION
    SetSelectionControlsVisible True
    Me.fraConfirm.Visible = False
    Me.fraResult.Visible = False
    Me.btnBack.Visible = False
    Me.btnCancel.Caption = "Cancel"
    Me.btnOK.Caption = "OK"
    UpdateModeUi
End Sub

Private Sub ShowConfirmPanel()
    mCurrentPanel = PANEL_CONFIRM
    SetSelectionControlsVisible False
    Me.fraConfirm.Visible = True
    Me.fraResult.Visible = False
    Me.btnBack.Visible = True
    Me.btnCancel.Caption = "Cancel"
    Me.btnOK.Caption = "Run"
    Me.chkConfirmAction.Value = False
    Me.lblConfirmSummary.Caption = BuildConfirmationSummaryForm(mPendingSourceWarehouseId, mPendingTargetWarehouseId, mPendingOperationMode, mPendingArchiveDestPath, mPendingPublishTombstone)
    Me.lblDeleteWarning.Visible = (mPendingOperationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE)
    ClearInlineError Me.lblConfirmError
    UpdateConfirmOkState
End Sub

Private Sub ShowResultPanel(ByVal wasSuccessful As Boolean, ByVal detailText As String)
    mCurrentPanel = PANEL_RESULT
    SetSelectionControlsVisible False
    Me.fraConfirm.Visible = False
    Me.fraResult.Visible = True
    Me.btnBack.Visible = False
    Me.btnCancel.Caption = "Close"
    Me.btnOK.Caption = "Close"
    Me.lblResultSummary.Caption = Trim$(detailText)
    Me.lblResultSummary.ForeColor = IIf(wasSuccessful, COLOR_SUCCESS, COLOR_ERROR)
End Sub

Private Function BuildConfirmationSummaryForm(ByVal sourceWarehouseId As String, _
                                              ByVal targetWarehouseId As String, _
                                              ByVal operationMode As Long, _
                                              ByVal archiveDestPath As String, _
                                              ByVal publishTombstone As Boolean) As String
    Select Case operationMode
        Case modWarehouseRetire.MODE_ARCHIVE_ONLY
            BuildConfirmationSummaryForm = _
                "Archive only will create an archive package for " & sourceWarehouseId & "." & vbCrLf & _
                "No migration, retirement, or deletion will occur." & vbCrLf & _
                "Archive destination: " & archiveDestPath
        Case modWarehouseRetire.MODE_ARCHIVE_MIGRATE
            BuildConfirmationSummaryForm = _
                "Archive + Migrate will archive " & sourceWarehouseId & " and seed current inventory into " & targetWarehouseId & "." & vbCrLf & _
                "The target remains locally authoritative. No auth, config identity, or inbox files are copied." & vbCrLf & _
                "Archive destination: " & archiveDestPath
        Case modWarehouseRetire.MODE_ARCHIVE_RETIRE
            BuildConfirmationSummaryForm = _
                "Archive + Retire will archive " & sourceWarehouseId & ", mark it RETIRED locally, and write a tombstone." & vbCrLf & _
                IIf(publishTombstone, "A best-effort SharePoint tombstone publish will also be attempted.", "SharePoint tombstone publish is disabled.") & vbCrLf & _
                "Archive destination: " & archiveDestPath
        Case modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE
            BuildConfirmationSummaryForm = _
                "Archive + Retire + Delete will archive " & sourceWarehouseId & ", mark it RETIRED, write a tombstone, then delete the local runtime folder." & vbCrLf & _
                IIf(publishTombstone, "A best-effort SharePoint tombstone publish will also be attempted before deletion.", "SharePoint tombstone publish is disabled.") & vbCrLf & _
                "Archive destination: " & archiveDestPath
    End Select
End Function

Private Sub UpdateConfirmOkState()
    Me.btnOK.Enabled = CBool(Me.chkConfirmAction.Value)
End Sub

Private Sub SetSelectionControlsVisible(ByVal isVisible As Boolean)
    Me.lblTitle.Visible = isVisible
    Me.lblSelectionIntro.Visible = isVisible
    Me.lblSourceWarehouse.Visible = isVisible
    Me.cmbSourceWarehouse.Visible = isVisible
    Me.lblSourceWarehouseError.Visible = isVisible
    Me.lblTargetWarehouse.Visible = isVisible
    Me.cmbTargetWarehouse.Visible = isVisible
    Me.lblTargetWarehouseError.Visible = isVisible
    Me.fraMode.Visible = isVisible
    Me.optArchiveOnly.Visible = isVisible
    Me.optArchiveMigrate.Visible = isVisible
    Me.optArchiveRetire.Visible = isVisible
    Me.optArchiveRetireDelete.Visible = isVisible
    Me.lblArchiveDestPath.Visible = isVisible
    Me.txtArchiveDestPath.Visible = isVisible
    Me.lblArchiveDestPathError.Visible = isVisible
    Me.chkPublishTombstone.Visible = isVisible And (ResolveSelectedMode() = modWarehouseRetire.MODE_ARCHIVE_RETIRE Or ResolveSelectedMode() = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE)
    Me.lblReAuthError.Visible = isVisible
End Sub

Private Sub ClearAllInlineErrors()
    ClearInlineError Me.lblSourceWarehouseError
    ClearInlineError Me.lblTargetWarehouseError
    ClearInlineError Me.lblArchiveDestPathError
    ClearInlineError Me.lblReAuthError
    ClearInlineError Me.lblConfirmError
End Sub

Private Sub ClearInlineError(ByVal lbl As MSForms.Label)
    If lbl Is Nothing Then Exit Sub
    lbl.Caption = ""
    lbl.ForeColor = COLOR_ERROR
End Sub

Private Sub SetInlineError(ByVal lbl As MSForms.Label, ByVal messageText As String)
    If lbl Is Nothing Then Exit Sub
    lbl.Caption = Trim$(messageText)
    lbl.ForeColor = COLOR_ERROR
End Sub

Private Sub ShowFormMessage(ByVal messageText As String, ByVal foreColor As Long)
    Me.lblSelectionIntro.Caption = Trim$(messageText)
    Me.lblSelectionIntro.ForeColor = foreColor
End Sub

Private Function ResolveCurrentUserForm() As String
    ResolveCurrentUserForm = Trim$(Environ$("USERNAME"))
    If ResolveCurrentUserForm = "" Then ResolveCurrentUserForm = Trim$(Application.UserName)
End Function

Private Function NormalizePathForm(ByVal pathText As String) As String
    pathText = Trim$(Replace$(pathText, "/", "\"))
    Do While Len(pathText) > 3 And Right$(pathText, 1) = "\"
        pathText = Left$(pathText, Len(pathText) - 1)
    Loop
    NormalizePathForm = pathText
End Function

Private Function GetParentFolderForm(ByVal pathText As String) As String
    Dim sepPos As Long

    pathText = NormalizePathForm(pathText)
    sepPos = InStrRev(pathText, "\")
    If sepPos = 3 And Mid$(pathText, 2, 2) = ":\" Then
        GetParentFolderForm = Left$(pathText, 3)
    ElseIf sepPos > 1 Then
        GetParentFolderForm = Left$(pathText, sepPos - 1)
    End If
End Function

Private Function GetLeafFolderNameForm(ByVal pathText As String) As String
    Dim sepPos As Long

    pathText = NormalizePathForm(pathText)
    If pathText = "" Then Exit Function
    sepPos = InStrRev(pathText, "\")
    If sepPos > 0 And sepPos < Len(pathText) Then
        GetLeafFolderNameForm = Mid$(pathText, sepPos + 1)
    Else
        GetLeafFolderNameForm = pathText
    End If
End Function

Private Function LooksLikeWarehouseRuntimeRootForm(ByVal rootPath As String) As Boolean
    Dim leafName As String

    rootPath = NormalizePathForm(rootPath)
    If rootPath = "" Then Exit Function
    leafName = Trim$(GetLeafFolderNameForm(rootPath))
    If leafName = "" Then Exit Function

    LooksLikeWarehouseRuntimeRootForm = FileExistsForm(rootPath & "\" & leafName & ".invSys.Config.xlsb")
End Function

Private Function FolderExistsForm(ByVal folderPath As String) As Boolean
    Dim fso As Object

    folderPath = NormalizePathForm(folderPath)
    If folderPath = "" Then Exit Function

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then FolderExistsForm = fso.FolderExists(folderPath)
    If Err.Number <> 0 Then
        Err.Clear
        FolderExistsForm = (Len(Dir$(folderPath, vbDirectory)) > 0)
    End If
    On Error GoTo 0
End Function

Private Function FileExistsForm(ByVal filePath As String) As Boolean
    Dim fso As Object

    filePath = NormalizePathForm(filePath)
    If filePath = "" Then Exit Function

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then FileExistsForm = fso.FileExists(filePath)
    If Err.Number <> 0 Then
        Err.Clear
        FileExistsForm = (Len(Dir$(filePath, vbNormal)) > 0)
    End If
    On Error GoTo 0
End Function

Private Sub RestoreRootOverrideForm(ByVal priorRoot As String)
    If Trim$(priorRoot) = "" Then
        modRuntimeWorkbooks.ClearCoreDataRootOverride
    Else
        modRuntimeWorkbooks.SetCoreDataRootOverride priorRoot
    End If
End Sub
