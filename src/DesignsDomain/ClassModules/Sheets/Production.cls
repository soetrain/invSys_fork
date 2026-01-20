VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Production"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    On Error GoTo CleanExit
    mProduction.HandleProductionSelectionChange Target
CleanExit:
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    On Error GoTo CleanExit
    mProduction.HandleProductionBeforeDoubleClick Target, Cancel
CleanExit:
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo CleanExit
    mProduction.HandleProductionChange Target
CleanExit:
End Sub


