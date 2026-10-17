Attribute VB_Name = "modAggregationFiles"
Option Explicit
Option Private Module

Private Type FileIdentity
    Attributes As Long
    CreatedLow As Long
    CreatedHigh As Long
    AccessLow As Long
    AccessHigh As Long
    WriteLow As Long
    WriteHigh As Long
    Volume As Long
    SizeHigh As Long
    SizeLow As Long
    Links As Long
    IndexHigh As Long
    IndexLow As Long
End Type

Private Declare PtrSafe Function CreateFileW Lib "kernel32" (ByVal name As LongPtr, ByVal access As Long, ByVal share As Long, ByVal security As LongPtr, ByVal creation As Long, ByVal flags As Long, ByVal template As LongPtr) As LongPtr
Private Declare PtrSafe Function GetFileInformationByHandle Lib "kernel32" (ByVal handle As LongPtr, ByRef info As FileIdentity) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal handle As LongPtr) As Long
Private Declare PtrSafe Function MoveFileExW Lib "kernel32" (ByVal source As LongPtr, ByVal target As LongPtr, ByVal flags As Long) As Long
Private Declare PtrSafe Function ReplaceFileW Lib "kernel32" (ByVal target As LongPtr, ByVal source As LongPtr, ByVal backup As LongPtr, ByVal flags As Long, ByVal exclude As LongPtr, ByVal reserved As LongPtr) As Long

Public Function Canonical(ByVal path As String) As String
    Dim prefix As String, part As Variant, body As String, fso As Object
    On Error GoTo Invalid
    path = Replace$(Trim$(path), "/", "\")
    If Left$(path, 2) = "\\" Then prefix = "\\": path = Mid$(path, 3)
    Do While InStr(path, "\\") > 0: path = Replace$(path, "\\", "\"): Loop
    Do While Len(path) > 3 And Right$(path, 1) = "\": path = Left$(path, Len(path) - 1): Loop
    If prefix = "" And Mid$(path, 2, 2) <> ":\" Then Exit Function
    If InStr(path, "*") Or InStr(path, "?") Then Exit Function
    body = path
    If prefix = "" Then body = Mid$(path, 4)
    For Each part In Split(body, "\")
        If part = "." Or part = ".." Or InStr(CStr(part), ":") Then Exit Function
    Next part
    Set fso = CreateObject("Scripting.FileSystemObject")
    Canonical = fso.GetAbsolutePathName(prefix & path)
Invalid:
End Function

Public Function Identity(ByVal path As String) As String
    Dim handle As LongPtr, info As FileIdentity
    On Error GoTo Finished
    handle = CreateFileW(StrPtr(path), 0, 7, 0, 3, &H2000000, 0)
    If handle = -1 Or handle = 0 Then Exit Function
    If GetFileInformationByHandle(handle, info) = 0 Then GoTo Finished
    If info.IndexHigh = 0 And info.IndexLow = 0 Then GoTo Finished
    Identity = CStr(info.Volume) & ":" & CStr(info.IndexHigh) & ":" & CStr(info.IndexLow)
Finished:
    If handle <> 0 And handle <> -1 Then CloseHandle handle
End Function

Public Function SameFile(ByVal first As String, ByVal second As String) As Boolean
    Dim key As String
    If StrComp(Canonical(first), Canonical(second), vbTextCompare) = 0 Then SameFile = True: Exit Function
    key = Identity(first)
    If key <> "" Then SameFile = (key = Identity(second))
End Function

Public Function Under(ByVal path As String, ByVal folder As String) As Boolean
    Dim fso As Object, current As String, key As String
    path = Canonical(path): folder = Canonical(folder)
    If path = "" Or folder = "" Then Under = True: Exit Function
    If StrComp(Left$(path & "\", Len(folder) + 1), folder & "\", vbTextCompare) = 0 Then Under = True: Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    key = Identity(folder)
    If key = "" Then Under = True: Exit Function
    current = path
    Do While current <> ""
        If key <> "" And Identity(current) = key Then Under = True: Exit Function
        current = fso.GetParentFolderName(current)
    Loop
End Function

Public Function Dirty(ByVal path As String) As Boolean
    Dim book As Workbook
    For Each book In Application.Workbooks
        If SameFile(book.FullName, path) Then
            If Not book.Saved Then Dirty = True
        End If
    Next book
End Function

Public Function Available(ByVal path As String) As Boolean
    Dim book As Workbook, fso As Object, current As String
    Dim handle As LongPtr
    On Error GoTo Invalid
    If Canonical(path) = "" Then Exit Function
    For Each book In Application.Workbooks
        If SameFile(book.FullName, path) Then Exit Function
    Next book
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(path) Then Exit Function
    If fso.FileExists(path) Then
        If (fso.GetFile(path).Attributes And 1) <> 0 Then Exit Function
        ' OPEN_EXISTING reserves write access without writing or truncating.
        ' No sharing also detects another Excel instance or external holder.
        handle = CreateFileW(StrPtr(path), &H40000000, 0, 0, 3, 0, 0)
        If handle = -1 Or handle = 0 Then Exit Function
        CloseHandle handle: handle = 0
    End If
    current = fso.GetParentFolderName(path)
    Do While current <> ""
        If fso.FileExists(current) Then Exit Function
        If fso.FolderExists(current) Then
            If (fso.GetFolder(current).Attributes And 1024) <> 0 Then Exit Function
        End If
        current = fso.GetParentFolderName(current)
    Loop
    Available = True
Invalid:
    If handle <> 0 And handle <> -1 Then CloseHandle handle
End Function

Public Function OpenCopy(ByVal path As String, ByRef copyPath As String, _
        Optional ByVal uniqueName As Boolean = False) As Workbook
    Dim folder As String, fso As Object, security As Long, events As Boolean, book As Workbook, leaf As String
    On Error GoTo Failed
    security = Application.AutomationSecurity: events = Application.EnableEvents
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = fso.BuildPath(fso.GetSpecialFolder(2), "invSysAggRead_" & modSystemIdentity.NewId())
    fso.CreateFolder folder
    leaf = fso.GetFileName(path)
    If uniqueName Then leaf = "verified_" & modSystemIdentity.NewId() & "." & fso.GetExtensionName(path)
    copyPath = fso.BuildPath(folder, leaf)
    fso.CopyFile path, copyPath, False
    Application.AutomationSecurity = 3: Application.EnableEvents = False
    Set book = Application.Workbooks.Open(copyPath, UpdateLinks:=0, ReadOnly:=True, Notify:=False, AddToMru:=False, IgnoreReadOnlyRecommended:=True)
    book.Windows(1).Visible = False
    Set OpenCopy = book
Finished:
    Application.AutomationSecurity = security: Application.EnableEvents = events
    Exit Function
Failed:
    If Not book Is Nothing Then book.Close False
    Resume Finished
End Function

Public Sub CloseCopy(ByVal book As Workbook, ByVal path As String)
    Dim fso As Object, folder As String
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = fso.GetParentFolderName(path)
    If Left$(fso.GetFileName(folder), 14) = "invSysAggRead_" Then fso.DeleteFolder folder, True
End Sub

Public Function Promote(ByVal candidate As String, ByVal destination As String, ByVal backup As String) As Boolean
    If CreateObject("Scripting.FileSystemObject").FileExists(destination) Then
        Promote = (ReplaceFileW(StrPtr(destination), StrPtr(candidate), StrPtr(backup), 0, 0, 0) <> 0)
    Else
        Promote = (MoveFileExW(StrPtr(candidate), StrPtr(destination), 8) <> 0)
    End If
End Function

Public Function PublishRecord(ByVal source As String, ByVal destination As String) As Boolean
    PublishRecord = (MoveFileExW(StrPtr(source), StrPtr(destination), 9) <> 0)
End Function
