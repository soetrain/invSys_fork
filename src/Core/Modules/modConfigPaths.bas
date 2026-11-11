Attribute VB_Name = "modConfigPaths"
Option Explicit
Option Private Module

' Shared Config path helpers, moved without behavioral changes.
Public Function GetParentFolderConfig(ByVal pathIn As String) As String
    Dim sepPos As Long

    sepPos = InStrRev(pathIn, "\")
    If sepPos > 1 Then GetParentFolderConfig = Left$(pathIn, sepPos - 1)
End Function

Public Sub EnsureFolderRecursiveConfig(ByVal folderPath As String)
    Dim parentPath As String
    Dim sepPos As Long
    Dim fso As Object

    folderPath = NormalizeFolderPathConfig(folderPath, False)
    If folderPath = "" Then Exit Sub
    If FolderExistsConfig(folderPath) Then Exit Sub
    If IsUncShareRootConfig(folderPath) Then Exit Sub

    sepPos = InStrRev(folderPath, "\")
    If sepPos > 1 Then
        parentPath = Left$(folderPath, sepPos - 1)
        If Right$(parentPath, 1) = ":" Then parentPath = parentPath & "\"
        If parentPath <> "" And Not FolderExistsConfig(parentPath) Then EnsureFolderRecursiveConfig parentPath
    End If

    If FolderExistsConfig(folderPath) Then Exit Sub

    If IsUncPathConfig(folderPath) Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder folderPath
    Else
        MkDir folderPath
    End If
End Sub

Public Function NormalizeFolderPathConfig(ByVal folderPath As String, ByVal withTrailingSlash As Boolean) As String
    folderPath = CollapseBackslashesConfig(Trim$(Replace$(folderPath, "/", "\")))
    If folderPath = "" Then Exit Function

    Do While Len(folderPath) > 3 And Right$(folderPath, 1) = "\"
        folderPath = Left$(folderPath, Len(folderPath) - 1)
    Loop

    NormalizeFolderPathConfig = folderPath
    If withTrailingSlash And Right$(NormalizeFolderPathConfig, 1) <> "\" Then
        NormalizeFolderPathConfig = NormalizeFolderPathConfig & "\"
    End If
End Function

Public Function CombinePathConfig(ByVal basePath As String, ByVal childName As String) As String
    If basePath = "" Then
        CombinePathConfig = childName
    ElseIf Right$(basePath, 1) = "\" Then
        CombinePathConfig = basePath & childName
    Else
        CombinePathConfig = basePath & "\" & childName
    End If
End Function

Public Function FileExistsConfig(ByVal fullPath As String) As Boolean
    Dim fso As Object

    fullPath = Trim$(Replace$(fullPath, "/", "\"))
    If fullPath = "" Then Exit Function

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then FileExistsConfig = fso.FileExists(fullPath)
    If Err.Number <> 0 Then
        Err.Clear
        FileExistsConfig = (Len(Dir$(fullPath, vbNormal)) > 0)
    End If
    On Error GoTo 0
End Function

Public Function FolderExistsConfig(ByVal folderPath As String) As Boolean
    Dim fso As Object

    folderPath = NormalizeFolderPathConfig(folderPath, False)
    If folderPath = "" Then Exit Function

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then FolderExistsConfig = fso.FolderExists(folderPath)
    If Err.Number <> 0 Then
        Err.Clear
        FolderExistsConfig = (Len(Dir$(folderPath, vbDirectory)) > 0)
    End If
    On Error GoTo 0
End Function

Public Function IsUncPathConfig(ByVal folderPath As String) As Boolean
    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    IsUncPathConfig = (Left$(folderPath, 2) = "\\")
End Function

Public Function IsUncShareRootConfig(ByVal folderPath As String) As Boolean
    Dim trimmedPath As String
    Dim parts() As String

    trimmedPath = NormalizeFolderPathConfig(folderPath, False)
    If Left$(trimmedPath, 2) <> "\\" Then Exit Function

    trimmedPath = Mid$(trimmedPath, 3)
    If trimmedPath = "" Then Exit Function

    parts = Split(trimmedPath, "\")
    IsUncShareRootConfig = (UBound(parts) = 1)
End Function

Public Function CollapseBackslashesConfig(ByVal pathIn As String) As String
    Dim prefix As String
    Dim body As String

    pathIn = Trim$(pathIn)
    If pathIn = "" Then Exit Function

    If Left$(pathIn, 2) = "\\" Then
        prefix = "\\"
        body = Mid$(pathIn, 3)
        Do While InStr(body, "\\") > 0
            body = Replace$(body, "\\", "\")
        Loop
        CollapseBackslashesConfig = prefix & body
        Exit Function
    End If

    If Len(pathIn) >= 2 And Mid$(pathIn, 2, 1) = ":" Then
        prefix = Left$(pathIn, 2)
        body = Mid$(pathIn, 3)
        Do While Left$(body, 1) = "\"
            body = Mid$(body, 2)
        Loop
        Do While InStr(body, "\\") > 0
            body = Replace$(body, "\\", "\")
        Loop
        If body <> "" Then
            CollapseBackslashesConfig = prefix & "\" & body
        Else
            CollapseBackslashesConfig = prefix & "\"
        End If
        Exit Function
    End If

    body = pathIn
    Do While InStr(body, "\\") > 0
        body = Replace$(body, "\\", "\")
    Loop
    CollapseBackslashesConfig = body
End Function
