Attribute VB_Name = "modReceivingFormWindow"
Option Explicit

#If Not Mac Then

#If VBA7 Then
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal flags As Long) As Long
    Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal unknownObject As IUnknown, ByRef hwnd As LongPtr) As Long
#Else
    Private Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal newStyle As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal flags As Long) As Long
    Private Declare Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal unknownObject As IUnknown, ByRef hwnd As Long) As Long
#End If

Private Const GWL_STYLE As Long = -16
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000

#End If

Public Function EnableReceivingResizable(ByVal receivingForm As Object, _
                                         Optional ByVal allowMinimize As Boolean = True, _
                                         Optional ByVal allowMaximize As Boolean = True) As Boolean
#If Mac Then
    EnableReceivingResizable = False
#Else
    Dim hwnd As LongPtr

    On Error GoTo Failed
    hwnd = ResolveReceivingFormWindowHandle(receivingForm)
    If hwnd = 0 Then Exit Function
    EnableReceivingResizable = ApplyReceivingWindowStyle(hwnd, allowMinimize, allowMaximize)
    Exit Function
Failed:
    EnableReceivingResizable = False
#End If
End Function

#If Not Mac Then

Private Function ApplyReceivingWindowStyle(ByVal hwnd As LongPtr, _
                                           ByVal allowMinimize As Boolean, _
                                           ByVal allowMaximize As Boolean) As Boolean
    Dim currentStyle As LongPtr
    Dim requestedStyle As LongPtr

    currentStyle = GetWindowLongPtr(hwnd, GWL_STYLE)
    requestedStyle = currentStyle Or WS_THICKFRAME
    If allowMinimize Then requestedStyle = requestedStyle Or WS_MINIMIZEBOX
    If allowMaximize Then requestedStyle = requestedStyle Or WS_MAXIMIZEBOX

    If requestedStyle <> currentStyle Then
        Call SetWindowLongPtr(hwnd, GWL_STYLE, requestedStyle)
        Call SetWindowPos(hwnd, 0, 0, 0, 0, 0, _
                          SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED)
    End If
    ApplyReceivingWindowStyle = True
End Function

Private Function ResolveReceivingFormWindowHandle(ByVal receivingForm As Object) As LongPtr
    On Error Resume Next
    IUnknown_GetWindow receivingForm, ResolveReceivingFormWindowHandle
    On Error GoTo 0
End Function

#End If

Public Function BuildReceivingHeaderCaption(ByVal targetList As MSForms.ListBox, _
                                              ByVal headings As Variant) As String
    Dim widths As Variant
    Dim i As Long
    Dim pointWidth As Double
    Dim charWidth As Long
    Dim headingText As String

    widths = Split(CStr(targetList.ColumnWidths), ";")
    For i = LBound(widths) To UBound(widths)
        pointWidth = Val(CStr(widths(i)))
        If pointWidth > 0 Then
            headingText = CStr(headings(i))
            charWidth = CLng(pointWidth / 5.25)
            If charWidth < 2 Then charWidth = 2
            If Len(headingText) >= charWidth Then
                BuildReceivingHeaderCaption = BuildReceivingHeaderCaption & _
                    Left$(headingText, charWidth - 1) & " "
            Else
                BuildReceivingHeaderCaption = BuildReceivingHeaderCaption & _
                    headingText & Space$(charWidth - Len(headingText))
            End If
        End If
    Next i
End Function
