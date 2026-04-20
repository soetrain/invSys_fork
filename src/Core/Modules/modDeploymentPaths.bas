Attribute VB_Name = "modDeploymentPaths"
Option Explicit

Private Const DEFAULT_LOCAL_RUNTIME_HUB_ROOT As String = "C:\invSys"
Private Const DEFAULT_BACKUP_FOLDER_NAME As String = "Backups"
Private Const DEFAULT_ARCHIVE_FOLDER_NAME As String = "Archive"

Public Function DefaultRuntimeHubRootPath(Optional ByVal withTrailingSlash As Boolean = False) As String
    DefaultRuntimeHubRootPath = NormalizePolicyPath(DEFAULT_LOCAL_RUNTIME_HUB_ROOT, withTrailingSlash)
End Function

Public Function DefaultWarehouseRuntimeRootPath(ByVal warehouseId As String, _
                                                Optional ByVal withTrailingSlash As Boolean = True) As String
    Dim resolvedWarehouseId As String

    resolvedWarehouseId = Trim$(warehouseId)
    If resolvedWarehouseId = "" Then resolvedWarehouseId = "{WarehouseId}"
    DefaultWarehouseRuntimeRootPath = CombinePolicyPath(DefaultRuntimeHubRootPath(False), resolvedWarehouseId, withTrailingSlash)
End Function

Public Function DefaultWarehouseBackupRootPath(ByVal warehouseId As String, _
                                               Optional ByVal withTrailingSlash As Boolean = True) As String
    Dim resolvedWarehouseId As String

    resolvedWarehouseId = Trim$(warehouseId)
    If resolvedWarehouseId = "" Then resolvedWarehouseId = "{WarehouseId}"
    DefaultWarehouseBackupRootPath = CombinePolicyPath(DefaultRuntimeHubRootPath(False) & "\" & DEFAULT_BACKUP_FOLDER_NAME, resolvedWarehouseId, withTrailingSlash)
End Function

Public Function DefaultArchiveRootPath(Optional ByVal withTrailingSlash As Boolean = False) As String
    DefaultArchiveRootPath = CombinePolicyPath(DefaultRuntimeHubRootPath(False), DEFAULT_ARCHIVE_FOLDER_NAME, withTrailingSlash)
End Function

Public Function NormalizeManagedFolderPath(ByVal pathText As String, _
                                           Optional ByVal withTrailingSlash As Boolean = False) As String
    NormalizeManagedFolderPath = NormalizePolicyPath(pathText, withTrailingSlash)
End Function

Public Function FolderExistsManaged(ByVal folderPath As String) As Boolean
    Dim fso As Object

    folderPath = NormalizePolicyPath(folderPath, False)
    If folderPath = "" Then Exit Function

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then FolderExistsManaged = fso.FolderExists(folderPath)
    If Err.Number <> 0 Then Err.Clear
    If Not FolderExistsManaged Then FolderExistsManaged = (Len(Dir$(folderPath, vbDirectory)) > 0)
    On Error GoTo 0
End Function

Public Function FileExistsManaged(ByVal filePath As String) As Boolean
    Dim fso As Object

    filePath = Trim$(Replace$(filePath, "/", "\"))
    If filePath = "" Then Exit Function

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then FileExistsManaged = fso.FileExists(filePath)
    If Err.Number <> 0 Then Err.Clear
    If Not FileExistsManaged Then FileExistsManaged = (Len(Dir$(filePath, vbNormal)) > 0)
    On Error GoTo 0
End Function

Public Sub EnsureFolderRecursiveManaged(ByVal folderPath As String)
    Dim parentPath As String

    folderPath = NormalizePolicyPath(folderPath, False)
    If folderPath = "" Then Exit Sub
    If FolderExistsManaged(folderPath) Then Exit Sub
    If IsUncShareRootManaged(folderPath) Then Exit Sub

    parentPath = GetParentFolderManaged(folderPath)
    If parentPath <> "" And Not FolderExistsManaged(parentPath) Then EnsureFolderRecursiveManaged parentPath

    If FolderExistsManaged(folderPath) Then Exit Sub

    On Error Resume Next
    MkDir folderPath
    On Error GoTo 0
End Sub

Public Function GetParentFolderManaged(ByVal pathText As String) As String
    Dim normalized As String
    Dim lastSlash As Long

    normalized = NormalizePolicyPath(pathText, False)
    If normalized = "" Then Exit Function
    If IsUncShareRootManaged(normalized) Then Exit Function

    lastSlash = InStrRev(normalized, "\")
    If lastSlash <= 0 Then Exit Function

    If lastSlash = 3 And Mid$(normalized, 2, 1) = ":" Then
        GetParentFolderManaged = Left$(normalized, lastSlash)
    ElseIf lastSlash > 2 Then
        GetParentFolderManaged = Left$(normalized, lastSlash - 1)
    End If
End Function

Public Function IsUncPathManaged(ByVal pathText As String) As Boolean
    pathText = NormalizePolicyPath(pathText, False)
    If pathText = "" Then Exit Function
    IsUncPathManaged = (Left$(pathText, 2) = "\\")
End Function

Public Function IsUncShareRootManaged(ByVal pathText As String) As Boolean
    Dim body As String
    Dim parts() As String

    pathText = NormalizePolicyPath(pathText, False)
    If Left$(pathText, 2) <> "\\" Then Exit Function

    body = Mid$(pathText, 3)
    If body = "" Then Exit Function
    parts = Split(body, "\")
    IsUncShareRootManaged = (UBound(parts) = 1)
End Function

Public Function BrowseForFolderPath(Optional ByVal initialPath As String = "", _
                                    Optional ByVal dialogTitle As String = "Choose Folder") As String
    Dim picker As FileDialog

    On Error GoTo CleanFail

    Set picker = Application.FileDialog(msoFileDialogFolderPicker)
    If picker Is Nothing Then Exit Function

    With picker
        .Title = dialogTitle
        .AllowMultiSelect = False
        If Trim$(initialPath) <> "" Then .InitialFileName = NormalizePolicyPath(initialPath, True)
        If .Show <> -1 Then Exit Function
        If .SelectedItems.Count <= 0 Then Exit Function
        BrowseForFolderPath = NormalizePolicyPath(CStr(.SelectedItems(1)), False)
    End With
    Exit Function

CleanFail:
    BrowseForFolderPath = vbNullString
End Function

Private Function CombinePolicyPath(ByVal basePath As String, _
                                   ByVal childName As String, _
                                   ByVal withTrailingSlash As Boolean) As String
    basePath = NormalizePolicyPath(basePath, False)
    childName = Trim$(Replace$(childName, "/", "\"))

    If basePath = "" Then
        CombinePolicyPath = NormalizePolicyPath(childName, withTrailingSlash)
    ElseIf childName = "" Then
        CombinePolicyPath = NormalizePolicyPath(basePath, withTrailingSlash)
    Else
        CombinePolicyPath = NormalizePolicyPath(basePath & "\" & childName, withTrailingSlash)
    End If
End Function

Private Function NormalizePolicyPath(ByVal pathText As String, ByVal withTrailingSlash As Boolean) As String
    pathText = Trim$(Replace$(pathText, "/", "\"))
    If pathText = "" Then Exit Function

    Do While Len(pathText) > 3 And Right$(pathText, 1) = "\"
        pathText = Left$(pathText, Len(pathText) - 1)
    Loop

    If withTrailingSlash Then
        NormalizePolicyPath = pathText & "\"
    Else
        NormalizePolicyPath = pathText
    End If
End Function
