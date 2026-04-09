Attribute VB_Name = "modUserFormResizeWin"
Option Explicit

Public Enum UserFormAnchorEdge
    anchorNone = 0
    anchorLeft = 1
    anchorTop = 2
    anchorRight = 4
    anchorBottom = 8
End Enum

#If Mac Then

Public Function EnableResizableUserForm(ByVal uForm As Object, _
                                        Optional ByVal allowMinimize As Boolean = False, _
                                        Optional ByVal allowMaximize As Boolean = False) As Boolean
End Function

Public Function GetUserFormWindowHandle(ByVal uForm As Object) As Long
End Function

#Else

#If VBA7 Then
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByRef hwnd As LongPtr) As Long
#Else
    Private Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As IUnknown, ByRef hwnd As Long) As Long
#End If

Private Const GWL_STYLE As Long = -16
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000

Public Function EnableResizableUserForm(ByVal uForm As Object, _
                                        Optional ByVal allowMinimize As Boolean = False, _
                                        Optional ByVal allowMaximize As Boolean = False) As Boolean
    Dim hwnd As LongPtr
    Dim currentStyle As LongPtr
    Dim updatedStyle As LongPtr

    On Error GoTo FailEnable

    hwnd = ResolveUserFormWindowHandleLocal(uForm)
    If hwnd = 0 Then Exit Function

    currentStyle = GetWindowLongPtr(hwnd, GWL_STYLE)
    updatedStyle = currentStyle Or WS_THICKFRAME
    If allowMinimize Then updatedStyle = updatedStyle Or WS_MINIMIZEBOX
    If allowMaximize Then updatedStyle = updatedStyle Or WS_MAXIMIZEBOX

    If updatedStyle <> currentStyle Then
        SetWindowLongPtr hwnd, GWL_STYLE, updatedStyle
        SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED
    End If

    EnableResizableUserForm = True
    Exit Function

FailEnable:
    EnableResizableUserForm = False
End Function

Public Function GetUserFormWindowHandle(ByVal uForm As Object) As LongPtr
    GetUserFormWindowHandle = ResolveUserFormWindowHandleLocal(uForm)
End Function

Private Function ResolveUserFormWindowHandleLocal(ByVal uForm As Object) As LongPtr
    On Error Resume Next
    IUnknown_GetWindow uForm, ResolveUserFormWindowHandleLocal
    On Error GoTo 0
End Function

#End If
