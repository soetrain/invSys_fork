VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cInventoryAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As Application

Public Sub Init()
    Set App = Application
End Sub

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    modInventoryPublisher.HandlePotentialInventoryWorkbook Wb
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    modInventoryPublisher.HandlePotentialInventoryWorkbook Wb
End Sub
