Attribute VB_Name = "modDesignsDomainBridge"
Option Explicit

Private Const DESIGNS_DOMAIN_ADDIN As String = "invSys.Designs.Domain.xlam"

Public Function ResolveDesignsWorkbookBridge(Optional ByVal warehouseId As String = "") As Workbook
    On Error GoTo CleanFail
    Set ResolveDesignsWorkbookBridge = Application.Run( _
        ResolveDesignsDomainMacroName("modDesignsBridgeApi.ResolveDesignsWorkbookBridgeResult"), warehouseId)
CleanFail:
End Function

Public Function EnsureDesignsSchemaBridge(ByVal targetWb As Workbook, _
                                          Optional ByRef report As String = "") As Boolean
    On Error GoTo FailBridge

    Dim encoded As String
    Dim separatorAt As Long

    encoded = CStr(Application.Run( _
        ResolveDesignsDomainMacroName("modDesignsBridgeApi.EnsureDesignsSchemaBridgeEncoded"), targetWb))
    separatorAt = InStr(1, encoded, vbTab, vbBinaryCompare)
    If separatorAt > 0 Then
        report = Mid$(encoded, separatorAt + 1)
        EnsureDesignsSchemaBridge = (StrComp(Left$(encoded, separatorAt - 1), "OK", vbTextCompare) = 0)
    Else
        report = encoded
    End If
    Exit Function

FailBridge:
    report = "EnsureDesignsSchemaBridge failed: " & Err.Description
End Function

Public Function ValidateDesignsSchemaBridge(ByVal targetWb As Workbook) As String
    On Error GoTo CleanFail
    ValidateDesignsSchemaBridge = CStr(Application.Run( _
        ResolveDesignsDomainMacroName("modDesignsBridgeApi.ValidateDesignsSchemaBridgeResult"), targetWb))
CleanFail:
End Function

Public Function ApplyDesignEventBridge(ByVal evt As Object, ByVal designsWb As Workbook, _
                                       Optional ByVal runId As String = "", _
                                       Optional ByRef statusOut As String = "", _
                                       Optional ByRef errorCode As String = "", _
                                       Optional ByRef errorMessage As String = "") As Boolean
    On Error GoTo FailBridge

    Dim encoded As String
    Dim parts() As String

    encoded = CStr(Application.Run( _
        ResolveDesignsDomainMacroName("modDesignsBridgeApi.ApplyDesignEventBridgeEncoded"), _
        evt, designsWb, runId))
    parts = Split(encoded, vbTab)
    If UBound(parts) >= 0 Then ApplyDesignEventBridge = (Trim$(parts(0)) = "1")
    If UBound(parts) >= 1 Then statusOut = parts(1)
    If UBound(parts) >= 2 Then errorCode = parts(2)
    If UBound(parts) >= 3 Then errorMessage = parts(3)
    Exit Function

FailBridge:
    errorCode = "DESIGNS_DOMAIN_CALL_FAILED"
    errorMessage = Err.Description
End Function

Public Function RebuildDesignProjectionsBridge(ByVal designsWb As Workbook, _
                                               Optional ByRef report As String = "") As Boolean
    On Error GoTo FailBridge

    Dim encoded As String
    Dim separatorAt As Long

    encoded = CStr(Application.Run( _
        ResolveDesignsDomainMacroName("modDesignsBridgeApi.RebuildDesignProjectionsBridgeEncoded"), designsWb))
    separatorAt = InStr(1, encoded, vbTab, vbBinaryCompare)
    If separatorAt > 0 Then
        RebuildDesignProjectionsBridge = (Left$(encoded, separatorAt - 1) = "1")
        report = Mid$(encoded, separatorAt + 1)
    End If
    Exit Function

FailBridge:
    report = "RebuildDesignProjectionsBridge failed: " & Err.Description
End Function

Public Function ListDesignsBridge(Optional ByVal designsWb As Workbook = Nothing, _
                                  Optional ByVal statusFilter As String = "") As Variant
    On Error GoTo CleanFail
    ListDesignsBridge = RunDesignsQueryBridge( _
        "LIST_DESIGNS", statusFilter, "", "", designsWb)
CleanFail:
End Function

Public Function GetDesignBOMBridge(ByVal designId As String, ByVal designVersion As String, _
                                   Optional ByVal designsWb As Workbook = Nothing) As Variant
    On Error GoTo CleanFail
    GetDesignBOMBridge = RunDesignsQueryBridge( _
        "GET_BOM", designId, designVersion, "", designsWb)
CleanFail:
End Function

Public Function GetDesignBOMForStatusBridge(ByVal designId As String, _
                                            ByVal designVersion As String, _
                                            ByVal requiredStatus As String, _
                                            Optional ByVal designsWb As Workbook = Nothing) As Variant
    On Error GoTo CleanFail
    GetDesignBOMForStatusBridge = RunDesignsQueryBridge( _
        "GET_BOM_FOR_STATUS", designId, designVersion, requiredStatus, designsWb)
CleanFail:
End Function

Public Function ListProcessesBridge(Optional ByVal designsWb As Workbook = Nothing, _
                                    Optional ByVal statusFilter As String = "") As Variant
    On Error GoTo CleanFail
    Dim queryResult As Variant
    queryResult = RunDesignsQueryBridge( _
        "LIST_PROCESSES", statusFilter, "", "", designsWb)
    ListProcessesBridge = queryResult
CleanFail:
End Function

Public Function GetProcessVersionBridge(ByVal processId As String, _
                                        ByVal processVersion As String, _
                                        Optional ByVal designsWb As Workbook = Nothing) As String
    On Error GoTo CleanFail
    GetProcessVersionBridge = CStr(RunDesignsQueryBridge( _
        "GET_PROCESS_VERSION", processId, processVersion, "", designsWb))
CleanFail:
End Function

Public Function ListRecipesBridge(Optional ByVal designsWb As Workbook = Nothing, _
                                  Optional ByVal statusFilter As String = "") As Variant
    On Error GoTo CleanFail
    statusFilter = Trim$(statusFilter)
    ListRecipesBridge = RunDesignsQueryBridge( _
        "LIST_RECIPES", statusFilter, "", "", designsWb)
CleanFail:
End Function

Public Function GetRecipeGraphBridge(ByVal recipeId As String, _
                                     ByVal recipeVersion As String, _
                                     Optional ByVal designsWb As Workbook = Nothing) As String
    On Error GoTo CleanFail
    GetRecipeGraphBridge = CStr(RunDesignsQueryBridge( _
        "GET_RECIPE_GRAPH", recipeId, recipeVersion, "", designsWb))
CleanFail:
End Function

Public Function ValidateReleasedRecipeBridgeEncoded(ByVal recipeId As String, _
                                                    ByVal recipeVersion As String, _
                                                    Optional ByVal designsWb As Workbook = Nothing) As String
    On Error GoTo CleanFail
    recipeId = Trim$(recipeId)
    ValidateReleasedRecipeBridgeEncoded = CStr(RunDesignsQueryBridge( _
        "VALIDATE_RELEASED_RECIPE", recipeId, recipeVersion, "", designsWb))
CleanFail:
End Function

Public Function ReadDesignsEventsForPublication(ByVal warehouseId As String, ByVal runtimeRoot As String) As String
    On Error GoTo Failed
    ReadDesignsEventsForPublication = CStr(RunDesignsQueryBridge("PUBLICATION_EVENTS", warehouseId, runtimeRoot, "", Nothing))
Failed:
End Function

Private Function RunDesignsQueryBridge(ByVal queryName As String, _
                                       ByVal arg1 As String, _
                                       ByVal arg2 As String, _
                                       ByVal arg3 As String, _
                                       ByVal designsWb As Workbook) As Variant
    RunDesignsQueryBridge = Application.Run( _
        ResolveDesignsDomainMacroName("modDesignsBridgeApi.ReadDesignsQueryBridgeResult"), _
        queryName, arg1, arg2, arg3, designsWb)
End Function

Public Function DiagnoseDesignsDomainBridge() As String
    On Error GoTo FailBridge
    DiagnoseDesignsDomainBridge = CStr(Application.Run( _
        ResolveDesignsDomainMacroName("modDesignsBridgeApi.DiagnoseDesignsDomainBridgeResult")))
    Exit Function
FailBridge:
    DiagnoseDesignsDomainBridge = "Designs Domain unavailable: " & Err.Description
End Function

Private Function ResolveDesignsDomainMacroName(ByVal macroName As String) As String
    Dim hostName As String

    hostName = FindDesignsDomainMacroHostName()
    If hostName <> "" Then
        ResolveDesignsDomainMacroName = "'" & Replace$(hostName, "'", "''") & "'!" & macroName
        Exit Function
    End If

    If IsSourceImportedDesignsTestHarness() Then
        ResolveDesignsDomainMacroName = "'" & Replace$(ThisWorkbook.Name, "'", "''") & "'!" & macroName
        Exit Function
    End If

    Err.Raise vbObjectError + 2701, "modDesignsDomainBridge.ResolveDesignsDomainMacroName", _
              "Designs Domain add-in is not open and could not be loaded beside Core."
End Function

Private Function FindDesignsDomainMacroHostName() As String
    Dim wb As Workbook
    Dim addin As AddIn
    Dim peerPath As String
    Dim parentPath As String

    For Each wb In Application.Workbooks
        If StrComp(wb.Name, DESIGNS_DOMAIN_ADDIN, vbTextCompare) = 0 Then
            FindDesignsDomainMacroHostName = wb.Name
            Exit Function
        End If
    Next wb

    For Each wb In Application.Workbooks
        If InStr(1, wb.Name, "Designs.Domain", vbTextCompare) > 0 Then
            FindDesignsDomainMacroHostName = wb.Name
            Exit Function
        End If
    Next wb

    peerPath = ThisWorkbook.Path
    If Trim$(peerPath) <> "" Then
        If Right$(peerPath, 1) <> "\" Then peerPath = peerPath & "\"

        Set wb = OpenDesignsDomainMacroHostIfExists(peerPath & DESIGNS_DOMAIN_ADDIN)
        If Not wb Is Nothing Then
            FindDesignsDomainMacroHostName = wb.Name
            Exit Function
        End If

        If Right$(peerPath, 6) = ".refs\" Then
            parentPath = Left$(peerPath, Len(peerPath) - 6)
            Set wb = OpenDesignsDomainMacroHostIfExists(parentPath & DESIGNS_DOMAIN_ADDIN)
            If Not wb Is Nothing Then
                FindDesignsDomainMacroHostName = wb.Name
                Exit Function
            End If
        End If
    End If

    On Error Resume Next
    For Each addin In Application.AddIns
        If addin Is Nothing Then GoTo NextAddIn
        If Not addin.Installed Then GoTo NextAddIn
        If Len(Dir$(addin.FullName, vbNormal)) = 0 Then GoTo NextAddIn
        If StrComp(addin.Name, DESIGNS_DOMAIN_ADDIN, vbTextCompare) = 0 _
           Or InStr(1, addin.Name, "Designs.Domain", vbTextCompare) > 0 Then
            FindDesignsDomainMacroHostName = addin.Name
            Exit Function
        End If
NextAddIn:
    Next addin
    On Error GoTo 0
End Function

Private Function OpenDesignsDomainMacroHostIfExists(ByVal workbookPath As String) As Workbook
    Dim wb As Workbook

    If Len(Dir$(workbookPath, vbNormal)) = 0 Then Exit Function

    On Error Resume Next
    Set wb = Application.Workbooks.Open(Filename:=workbookPath, UpdateLinks:=0, ReadOnly:=False, _
                                        IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
    If Not wb Is Nothing Then wb.IsAddin = True
    On Error GoTo 0

    Set OpenDesignsDomainMacroHostIfExists = wb
End Function

Private Function IsSourceImportedDesignsTestHarness() As Boolean
    Dim workbookName As String

    workbookName = LCase$(Trim$(ThisWorkbook.Name))
    IsSourceImportedDesignsTestHarness = (InStr(1, workbookName, "harness", vbTextCompare) > 0)
End Function
