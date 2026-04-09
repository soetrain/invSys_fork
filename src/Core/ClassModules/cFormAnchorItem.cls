VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFormAnchorItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mControl As Object
Private mParent As Object
Private mAnchorMask As Long
Private mStartingTop As Double
Private mStartingLeft As Double
Private mStartingHeight As Double
Private mStartingWidth As Double
Private mParentStartingHeight As Double
Private mParentStartingWidth As Double

Public Sub Capture(ByVal targetControl As Object, ByVal anchorMask As Long)
    Set mControl = targetControl
    Set mParent = ResolveParentLocal(targetControl)
    mAnchorMask = anchorMask

    mStartingTop = CDbl(targetControl.Top)
    mStartingLeft = CDbl(targetControl.Left)
    mStartingHeight = CDbl(targetControl.Height)
    mStartingWidth = CDbl(targetControl.Width)
    mParentStartingHeight = GetParentInsideHeightLocal(mParent)
    mParentStartingWidth = GetParentInsideWidthLocal(mParent)
End Sub

Public Sub ApplyLayout()
    Dim widthDelta As Double
    Dim heightDelta As Double
    Dim newLeft As Double
    Dim newTop As Double
    Dim newWidth As Double
    Dim newHeight As Double

    If mControl Is Nothing Or mParent Is Nothing Then Exit Sub

    widthDelta = GetParentInsideWidthLocal(mParent) - mParentStartingWidth
    heightDelta = GetParentInsideHeightLocal(mParent) - mParentStartingHeight

    newLeft = mStartingLeft
    newTop = mStartingTop
    newWidth = mStartingWidth
    newHeight = mStartingHeight

    If HasAnchorLocal(anchorLeft) And HasAnchorLocal(anchorRight) Then
        newWidth = MaxDoubleLocal(0, mStartingWidth + widthDelta)
    ElseIf (Not HasAnchorLocal(anchorLeft)) And HasAnchorLocal(anchorRight) Then
        newLeft = mStartingLeft + widthDelta
    End If

    If HasAnchorLocal(anchorTop) And HasAnchorLocal(anchorBottom) Then
        newHeight = MaxDoubleLocal(0, mStartingHeight + heightDelta)
    ElseIf (Not HasAnchorLocal(anchorTop)) And HasAnchorLocal(anchorBottom) Then
        newTop = mStartingTop + heightDelta
    End If

    On Error Resume Next
    mControl.Move newLeft, newTop, newWidth, newHeight
    If Err.Number <> 0 Then
        Err.Clear
        mControl.Left = newLeft
        mControl.Top = newTop
        mControl.Width = newWidth
        mControl.Height = newHeight
    End If
    On Error GoTo 0
End Sub

Private Function HasAnchorLocal(ByVal anchorFlag As UserFormAnchorEdge) As Boolean
    HasAnchorLocal = ((mAnchorMask And CLng(anchorFlag)) = CLng(anchorFlag))
End Function

Private Function ResolveParentLocal(ByVal targetControl As Object) As Object
    On Error Resume Next
    Set ResolveParentLocal = targetControl.Parent
    On Error GoTo 0
End Function

Private Function GetParentInsideWidthLocal(ByVal parentObj As Object) As Double
    On Error Resume Next
    GetParentInsideWidthLocal = CDbl(CallByName(parentObj, "InsideWidth", VbGet))
    If Err.Number <> 0 Then
        Err.Clear
        GetParentInsideWidthLocal = CDbl(CallByName(parentObj, "Width", VbGet))
    End If
    On Error GoTo 0
End Function

Private Function GetParentInsideHeightLocal(ByVal parentObj As Object) As Double
    On Error Resume Next
    GetParentInsideHeightLocal = CDbl(CallByName(parentObj, "InsideHeight", VbGet))
    If Err.Number <> 0 Then
        Err.Clear
        GetParentInsideHeightLocal = CDbl(CallByName(parentObj, "Height", VbGet))
    End If
    On Error GoTo 0
End Function

Private Function MaxDoubleLocal(ByVal leftValue As Double, ByVal rightValue As Double) As Double
    If leftValue >= rightValue Then
        MaxDoubleLocal = leftValue
    Else
        MaxDoubleLocal = rightValue
    End If
End Function
