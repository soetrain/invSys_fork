Attribute VB_Name = "modLocalAddinsRegistration"
Option Explicit

Private Function RequiredInvSysAddinNamesLocal() As Variant
    RequiredInvSysAddinNamesLocal = Array( _
        "invSys.Core.xlam", _
        "invSys.Inventory.Domain.xlam", _
        "invSys.Designs.Domain.xlam", _
        "invSys.Receiving.xlam", _
        "invSys.Shipping.xlam", _
        "invSys.Production.xlam", _
        "invSys.Admin.xlam")
End Function

Public Function EnsureLocalInvSysAddinsRegistered(Optional ByVal preferredFolder As String = "", _
                                                  Optional ByRef report As String = "") As Boolean
    Dim addinsFolder As String
    Dim addinNames As Variant

    On Error GoTo FailEnsure

    addinsFolder = ResolvePreferredInvSysAddinsFolderLocal(preferredFolder)
    If addinsFolder = "" Then
        report = "Current invSys add-ins could not be resolved." & vbCrLf & _
                 "Open the full add-in set from deploy/current or from the synced SharePoint Addins folder first."
        Exit Function
    End If

    addinNames = RequiredInvSysAddinNamesLocal()
    If Not FolderHasRequiredInvSysAddinsLocal(addinsFolder, addinNames, report) Then Exit Function

    report = "OK|Folder=" & addinsFolder & "|SessionAddinsPreserved=True"
    EnsureLocalInvSysAddinsRegistered = True
    Exit Function

FailEnsure:
    report = "EnsureLocalInvSysAddinsRegistered failed: " & Err.Description
End Function

Private Function ResolvePreferredInvSysAddinsFolderLocal(ByVal preferredFolder As String) As String
    Dim candidate As String

    candidate = NormalizeFolderPathLocalAddins(preferredFolder, False)
    If candidate <> "" Then
        If FolderHasRequiredInvSysAddinsLocal(candidate, RequiredInvSysAddinNamesLocal()) Then
            ResolvePreferredInvSysAddinsFolderLocal = candidate
            Exit Function
        End If
    End If

    candidate = NormalizeFolderPathLocalAddins(ThisWorkbook.Path, False)
    If candidate <> "" Then
        If FolderHasRequiredInvSysAddinsLocal(candidate, RequiredInvSysAddinNamesLocal()) Then
            ResolvePreferredInvSysAddinsFolderLocal = candidate
            Exit Function
        End If
    End If

    candidate = ResolveOpenWorkbookFolderByNameLocal("invSys.Admin.xlam")
    If candidate <> "" Then
        If FolderHasRequiredInvSysAddinsLocal(candidate, RequiredInvSysAddinNamesLocal()) Then
            ResolvePreferredInvSysAddinsFolderLocal = candidate
            Exit Function
        End If
    End If

    candidate = NormalizeFolderPathLocalAddins(modConfig.GetString("PathSharePointRoot", ""), False)
    If candidate <> "" Then
        If FolderHasRequiredInvSysAddinsLocal(candidate & "\Addins", RequiredInvSysAddinNamesLocal()) Then
            ResolvePreferredInvSysAddinsFolderLocal = candidate & "\Addins"
            Exit Function
        End If
        If FolderHasRequiredInvSysAddinsLocal(candidate, RequiredInvSysAddinNamesLocal()) Then
            ResolvePreferredInvSysAddinsFolderLocal = candidate
            Exit Function
        End If
    End If
End Function

Private Function FolderHasRequiredInvSysAddinsLocal(ByVal folderPath As String, _
                                                    ByVal addinNames As Variant, _
                                                    Optional ByRef report As String = "") As Boolean
    Dim addinName As Variant
    Dim normalized As String
    Dim missing As String

    normalized = NormalizeFolderPathLocalAddins(folderPath, False)
    If normalized = "" Then Exit Function
    If Not FolderExistsLocalAddins(normalized) Then Exit Function

    For Each addinName In addinNames
        If Not FileExistsLocalAddins(normalized & "\" & CStr(addinName)) Then
            If missing <> "" Then missing = missing & ", "
            missing = missing & CStr(addinName)
        End If
    Next addinName

    If missing <> "" Then
        report = "Required add-ins were not found under " & normalized & "." & vbCrLf & _
                 "Missing: " & missing
        Exit Function
    End If

    FolderHasRequiredInvSysAddinsLocal = True
End Function

Private Function ShouldKeepInvSysAddinLocal(ByVal addinObj As AddIn, _
                                            ByVal targetPath As String, _
                                            ByVal addinNames As Variant) As Boolean
    Dim addinPath As String

    If addinObj Is Nothing Then Exit Function
    If Not IsRequiredInvSysAddinNameLocal(SafeTrimLocalAddins(addinObj.Name), addinNames) Then Exit Function

    addinPath = NormalizeFilePathLocalAddins(SafeAddinFullNameLocal(addinObj))
    targetPath = NormalizeFilePathLocalAddins(targetPath)

    If addinPath = "" Then Exit Function
    If targetPath <> "" Then
        If StrComp(addinPath, targetPath, vbTextCompare) = 0 Then
            ShouldKeepInvSysAddinLocal = True
            Exit Function
        End If
    End If

    If Not FileExistsLocalAddins(addinPath) Then Exit Function
    If IsOpenWorkbookAtPathLocal(addinPath) Then ShouldKeepInvSysAddinLocal = True
End Function

Private Function ResolveAddinByTargetPathLocal(ByVal targetPath As String, _
                                               ByVal addinName As String) As AddIn
    Dim addinObj As AddIn
    Dim targetNorm As String

    targetNorm = NormalizeFilePathLocalAddins(targetPath)

    On Error Resume Next
    For Each addinObj In Application.AddIns
        If StrComp(NormalizeFilePathLocalAddins(SafeAddinFullNameLocal(addinObj)), targetNorm, vbTextCompare) = 0 Then
            Set ResolveAddinByTargetPathLocal = addinObj
            Exit Function
        End If
    Next addinObj

    For Each addinObj In Application.AddIns
        If StrComp(SafeTrimLocalAddins(addinObj.Name), addinName, vbTextCompare) = 0 Then
            If Not addinObj.Installed Then
                Set ResolveAddinByTargetPathLocal = addinObj
                Exit Function
            End If
        End If
    Next addinObj
    On Error GoTo 0
End Function

Private Function ShouldSkipInstallToggleLocal(ByVal addinObj As AddIn, ByVal targetPath As String) As Boolean
    Dim currentPath As String

    If addinObj Is Nothing Then Exit Function
    currentPath = NormalizeFilePathLocalAddins(SafeAddinFullNameLocal(addinObj))
    targetPath = NormalizeFilePathLocalAddins(targetPath)

    If StrComp(SafeTrimLocalAddins(addinObj.Name), SafeTrimLocalAddins(ThisWorkbook.Name), vbTextCompare) = 0 Then
        If StrComp(currentPath, NormalizeFilePathLocalAddins(ThisWorkbook.FullName), vbTextCompare) = 0 Then
            ShouldSkipInstallToggleLocal = True
            Exit Function
        End If
    End If

    If IsOpenWorkbookAtPathLocal(targetPath) Then ShouldSkipInstallToggleLocal = True
End Function

Private Function ResolveOpenWorkbookFolderByNameLocal(ByVal workbookName As String) As String
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(SafeTrimLocalAddins(wb.Name), workbookName, vbTextCompare) = 0 Then
            ResolveOpenWorkbookFolderByNameLocal = NormalizeFolderPathLocalAddins(wb.Path, False)
            Exit Function
        End If
    Next wb
End Function

Private Function IsOpenWorkbookAtPathLocal(ByVal fullPath As String) As Boolean
    Dim wb As Workbook

    fullPath = NormalizeFilePathLocalAddins(fullPath)
    If fullPath = "" Then Exit Function

    For Each wb In Application.Workbooks
        If StrComp(NormalizeFilePathLocalAddins(wb.FullName), fullPath, vbTextCompare) = 0 Then
            IsOpenWorkbookAtPathLocal = True
            Exit Function
        End If
    Next wb
End Function

Private Function ShouldManageInvSysAddinLocal(ByVal addinObj As AddIn) As Boolean
    Dim addinName As String
    Dim addinPath As String

    If addinObj Is Nothing Then Exit Function
    addinName = LCase$(SafeTrimLocalAddins(addinObj.Name))
    addinPath = LCase$(NormalizeFilePathLocalAddins(SafeAddinFullNameLocal(addinObj)))

    ShouldManageInvSysAddinLocal = (InStr(1, addinName, "invsys", vbTextCompare) > 0) Or _
                                   (InStr(1, addinPath, "\invsys", vbTextCompare) > 0)
End Function

Private Function IsRequiredInvSysAddinNameLocal(ByVal addinName As String, ByVal addinNames As Variant) As Boolean
    Dim item As Variant

    For Each item In addinNames
        If StrComp(addinName, CStr(item), vbTextCompare) = 0 Then
            IsRequiredInvSysAddinNameLocal = True
            Exit Function
        End If
    Next item
End Function

Private Function SafeAddinFullNameLocal(ByVal addinObj As AddIn) As String
    On Error Resume Next
    SafeAddinFullNameLocal = SafeTrimLocalAddins(addinObj.FullName)
    On Error GoTo 0
End Function

Private Function NormalizeFolderPathLocalAddins(ByVal folderPath As String, ByVal keepTrailingSlash As Boolean) As String
    folderPath = SafeTrimLocalAddins(Replace$(folderPath, "/", "\"))
    If folderPath = "" Then Exit Function

    Do While Len(folderPath) > 3 And Right$(folderPath, 1) = "\"
        folderPath = Left$(folderPath, Len(folderPath) - 1)
    Loop

    If keepTrailingSlash Then
        NormalizeFolderPathLocalAddins = folderPath & "\"
    Else
        NormalizeFolderPathLocalAddins = folderPath
    End If
End Function

Private Function NormalizeFilePathLocalAddins(ByVal filePath As String) As String
    NormalizeFilePathLocalAddins = NormalizeFolderPathLocalAddins(filePath, False)
End Function

Private Function FolderExistsLocalAddins(ByVal folderPath As String) As Boolean
    folderPath = NormalizeFolderPathLocalAddins(folderPath, False)
    If folderPath = "" Then Exit Function

    On Error Resume Next
    FolderExistsLocalAddins = (Len(Dir$(folderPath, vbDirectory)) > 0)
    On Error GoTo 0
End Function

Private Function FileExistsLocalAddins(ByVal filePath As String) As Boolean
    filePath = NormalizeFilePathLocalAddins(filePath)
    If filePath = "" Then Exit Function

    On Error Resume Next
    FileExistsLocalAddins = (Len(Dir$(filePath, vbNormal)) > 0)
    On Error GoTo 0
End Function

Private Function SafeTrimLocalAddins(ByVal valueIn As Variant) As String
    On Error Resume Next
    SafeTrimLocalAddins = Trim$(CStr(valueIn))
    On Error GoTo 0
End Function
