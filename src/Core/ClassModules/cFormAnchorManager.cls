VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFormAnchorManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mForm As Object
Private mItems As Collection
Private mMinWidth As Double
Private mMinHeight As Double
Private mApplying As Boolean

Public Sub Initialize(ByVal targetForm As Object, _
                      Optional ByVal minWidth As Double = 0, _
                      Optional ByVal minHeight As Double = 0)
    Set mForm = targetForm
    Set mItems = New Collection
    mMinWidth = minWidth
    mMinHeight = minHeight
End Sub

Public Sub Add(ByVal targetControl As Object, ByVal anchorMask As Long)
    Dim item As cFormAnchorItem

    If mItems Is Nothing Then Set mItems = New Collection

    Set item = New cFormAnchorItem
    item.Capture targetControl, anchorMask
    mItems.Add item
End Sub

Public Sub ResizeControls()
    Dim item As cFormAnchorItem

    If mApplying Then Exit Sub
    If mForm Is Nothing Then Exit Sub
    If mItems Is Nothing Then Exit Sub

    mApplying = True
    ApplyMinimumFormSizeLocal

    For Each item In mItems
        item.ApplyLayout
    Next item

    mApplying = False
End Sub

Private Sub ApplyMinimumFormSizeLocal()
    On Error Resume Next
    If mMinWidth > 0 Then
        If CDbl(CallByName(mForm, "Width", VbGet)) < mMinWidth Then CallByName mForm, "Width", VbLet, mMinWidth
    End If
    If mMinHeight > 0 Then
        If CDbl(CallByName(mForm, "Height", VbGet)) < mMinHeight Then CallByName mForm, "Height", VbLet, mMinHeight
    End If
    On Error GoTo 0
End Sub
