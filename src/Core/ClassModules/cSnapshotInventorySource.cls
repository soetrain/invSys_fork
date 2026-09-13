VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "cSnapshotInventorySource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mOwned As Workbook

Public Function Read(ByVal warehouseId As String, ByVal supplied As Workbook) As Workbook
    Dim target As WarehouseTarget, root As String, path As String, candidate As Workbook
    Dim events As Boolean, security As Long, guarded As Boolean, window As Window
    On Error GoTo Failed
    If Not supplied Is Nothing Then Set Read = supplied: Exit Function
    If Not modTrainingWire.ValidSegment(warehouseId) Then Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    If Not target Is Nothing Then
        If StrComp(target.WarehouseId, warehouseId, vbTextCompare) <> 0 Then Exit Function
        root = target.RuntimeRoot
    Else
        root = modRuntimeWorkbooks.GetCoreDataRootOverride()
        If root = "" Then root = modConfig.GetString("PathDataRoot", "")
    End If
    If root = "" Then Exit Function
    If Right$(root, 1) <> "\" Then root = root & "\"
    path = root & warehouseId & ".invSys.Data.Inventory.xlsb"
    For Each candidate In Application.Workbooks
        If StrComp(candidate.FullName, path, vbTextCompare) = 0 Then Set Read = candidate: Exit Function
    Next candidate
    If Not CreateObject("Scripting.FileSystemObject").FileExists(path) Then Exit Function
    events = Application.EnableEvents: security = Application.AutomationSecurity: guarded = True
    Application.EnableEvents = False: Application.AutomationSecurity = 3
    Set mOwned = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, _
        IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
    For Each window In mOwned.Windows
        window.Visible = False
    Next window
    Set Read = mOwned
CleanExit:
    If guarded Then
        Application.AutomationSecurity = security
        Application.EnableEvents = events
    End If
    Exit Function
Failed:
    Set Read = Nothing
    Resume CleanExit
End Function

Private Sub Class_Terminate()
    Dim events As Boolean
    If mOwned Is Nothing Then Exit Sub
    On Error Resume Next
    events = Application.EnableEvents
    Application.EnableEvents = False
    mOwned.Close SaveChanges:=False
    Application.EnableEvents = events
    Set mOwned = Nothing
End Sub
