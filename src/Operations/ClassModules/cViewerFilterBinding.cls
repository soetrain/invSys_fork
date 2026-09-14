VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "cViewerFilterBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mControl As MSForms.ComboBox
Private mOwner As frmInventoryViewer

Public Sub Bind(ByVal control As MSForms.ComboBox, ByVal owner As frmInventoryViewer)
    Set mControl = control
    Set mOwner = owner
End Sub

Public Sub Disconnect()
    Set mControl = Nothing
    Set mOwner = Nothing
End Sub

Private Sub mControl_Change()
    If Not mOwner Is Nothing Then mOwner.ApplyEventFilters
End Sub
