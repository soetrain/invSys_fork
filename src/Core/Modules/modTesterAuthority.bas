Attribute VB_Name = "modTesterAuthority"
Option Explicit
Option Private Module

Private mTesterForms As cCapturedFormStore

' A primitive handle retains the actual source workbook in its owning Core project.
Public Function Capture(Optional ByVal workbookName As String = "") As String
    Dim entry As Object, key As String, context As String, target As WarehouseTarget
    Dim workbook As Workbook
    On Error GoTo Failed
    context = modAuthUserCommands.CaptureContext()
    If context = "" Then Exit Function
    If workbookName = "" Then Set workbook = Application.ActiveWorkbook Else Set workbook = Application.Workbooks(workbookName)
    If workbook Is Nothing Then Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    Set entry = CreateObject("Scripting.Dictionary")
    entry.Add "Context", context
    entry.Add "Workbook", workbook
    entry.Add "SourceWarehouse", target.WarehouseId
    entry.Add "SourceRoot", target.RuntimeRoot
    entry.Add "CreatorUser", modAuth.GetCurrentUserId()
    entry.Add "CreatorGeneration", modAuthSession.Version()
    If mTesterForms Is Nothing Then Set mTesterForms = New cCapturedFormStore
    key = modSystemIdentity.NewId()
    mTesterForms.Add key, entry
    Capture = key
Failed:
End Function

Public Sub Release(ByVal key As String)
    If mTesterForms Is Nothing Then Exit Sub
    mTesterForms.Remove key
End Sub

Public Function Entry(ByVal key As String) As Object
    If mTesterForms Is Nothing Then Exit Function
    Set Entry = mTesterForms.Item(key)
End Function

Public Function Check(ByVal key As String, Optional ByVal provisioningRoot As String = "") As Boolean
    Dim saved As Object, target As WarehouseTarget, wb As Workbook, source As Workbook
    Dim found As Boolean, priorRoot As String, changedRoot As Boolean
    On Error GoTo Finished
    Set saved = Entry(key)
    If saved Is Nothing Then Exit Function
    priorRoot = modRuntimeWorkbooks.GetCoreDataRootOverride()
    If priorRoot <> "" And StrComp(NormalizeFolderPathTesterSetup(priorRoot), NormalizeFolderPathTesterSetup(CStr(saved("SourceRoot"))), vbTextCompare) <> 0 Then
        If provisioningRoot = "" Then Exit Function
        If StrComp(NormalizeFolderPathTesterSetup(priorRoot), NormalizeFolderPathTesterSetup(provisioningRoot), vbTextCompare) <> 0 Then Exit Function
    End If
    modRuntimeWorkbooks.SetCoreDataRootOverride CStr(saved("SourceRoot"))
    changedRoot = True
    If saved("Context") <> modAuthUserCommands.CaptureContext() Then GoTo Finished
    Set source = saved("Workbook")
    For Each wb In Application.Workbooks
        If wb Is source Then found = True
    Next wb
    If Not found Then GoTo Finished
    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then GoTo Finished
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, target.AuthPath, vbTextCompare) = 0 Then
            If Not wb.Saved Then GoTo Finished
        End If
    Next wb
    If Not modAuth.LoadAuth(target.WarehouseId) Then GoTo Finished
    If Not modAuth.CanPerform("ADMIN_MAINT", modAuth.GetCurrentUserId(), target.WarehouseId, target.StationId, "TESTER_SETUP") Then GoTo Finished
    Check = (saved("Context") = modAuthUserCommands.CaptureContext())
Finished:
    If changedRoot Then RestoreCoreRootOverrideTesterSetup priorRoot
End Function
