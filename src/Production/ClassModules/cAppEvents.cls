VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As Application

Public Sub Init()
    Set App = Application
End Sub

Private Sub App_NewWorkbook(ByVal Wb As Workbook)
    Application.EnableEvents = True
    modProductionInit.EnsureProductionSurfaceForWorkbook Wb
End Sub

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    Application.EnableEvents = True
    modProductionInit.EnsureProductionSurfaceForWorkbook Wb
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    Application.EnableEvents = True
    modProductionInit.EnsureProductionSurfaceForWorkbook Wb
End Sub

Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    Application.EnableEvents = True
    modOperatorReadModel.UnregisterAutoSnapshotWorkbook Wb
End Sub

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Application.EnableEvents = True
    On Error Resume Next
    mProduction.HandleProductionSelectionChange Target
    On Error GoTo 0
End Sub

Private Sub App_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Application.EnableEvents = True
    On Error Resume Next
    mProduction.HandleProductionBeforeDoubleClick Target, Cancel
    On Error GoTo 0
End Sub

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Application.EnableEvents = True
    On Error Resume Next
    mProduction.HandleProductionChange Target
    On Error GoTo 0
End Sub
