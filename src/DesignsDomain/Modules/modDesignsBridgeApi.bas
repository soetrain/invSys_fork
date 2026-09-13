Attribute VB_Name = "modDesignsBridgeApi"
Option Explicit

Public Function ResolveDesignsWorkbookBridgeResult(Optional ByVal warehouseId As String = "") As Workbook
    Dim report As String
    Set ResolveDesignsWorkbookBridgeResult = modDesignsRuntime.ResolveDesignsWorkbook(warehouseId, Nothing, report)
End Function

Public Function EnsureDesignsSchemaBridgeEncoded(ByVal targetWb As Workbook) As String
    Dim report As String
    Dim succeeded As Boolean

    succeeded = modDesignsSchema.EnsureDesignsSchema(targetWb, report)
    EnsureDesignsSchemaBridgeEncoded = IIf(succeeded, "OK", "FAIL") & vbTab & report
End Function

Public Function ValidateDesignsSchemaBridgeResult(ByVal targetWb As Workbook) As String
    ValidateDesignsSchemaBridgeResult = modDesignsSchema.ValidateDesignsSchema(targetWb)
End Function

Public Function ApplyDesignEventBridgeEncoded(ByVal evt As Object, ByVal designsWb As Workbook, _
                                              Optional ByVal runId As String = "") As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim succeeded As Boolean

    succeeded = modDesignsApply.ApplyDesignEvent(evt, designsWb, runId, statusOut, errorCode, errorMessage)
    ApplyDesignEventBridgeEncoded = CStr(Abs(CLng(succeeded))) & vbTab & statusOut & vbTab & errorCode & vbTab & errorMessage
End Function

Public Function RebuildDesignProjectionsBridgeEncoded(ByVal designsWb As Workbook) As String
    Dim report As String
    Dim succeeded As Boolean

    succeeded = modDesignsApply.RebuildDesignProjections(designsWb, report)
    RebuildDesignProjectionsBridgeEncoded = CStr(Abs(CLng(succeeded))) & vbTab & report
End Function

Public Function ListDesignsBridgeResult(Optional ByVal designsWb As Workbook = Nothing, _
                                        Optional ByVal statusFilter As String = "") As Variant
    ListDesignsBridgeResult = modDesignsQueries.ListDesigns(designsWb, statusFilter)
End Function

Public Function GetBOMBridgeResult(ByVal designId As String, ByVal designVersion As String, _
                                   Optional ByVal designsWb As Workbook = Nothing) As Variant
    GetBOMBridgeResult = modDesignsQueries.GetBOM(designId, designVersion, designsWb)
End Function

Public Function GetBOMForStatusBridgeResult(ByVal designId As String, _
                                            ByVal designVersion As String, _
                                            ByVal requiredStatus As String, _
                                            Optional ByVal designsWb As Workbook = Nothing) As Variant
    GetBOMForStatusBridgeResult = modDesignsQueries.GetBOMForStatus( _
        designId, designVersion, requiredStatus, designsWb)
End Function

Public Function ListProcessesBridgeResult(Optional ByVal designsWb As Workbook = Nothing, _
                                          Optional ByVal statusFilter As String = "") As Variant
    ListProcessesBridgeResult = modDesignsQueries.ListProcesses(designsWb, statusFilter)
End Function

Public Function GetProcessVersionBridgeEncoded(ByVal processId As String, _
                                               ByVal processVersion As String, _
                                               Optional ByVal designsWb As Workbook = Nothing) As String
    GetProcessVersionBridgeEncoded = modDesignsQueries.GetProcessVersion( _
        processId, processVersion, designsWb)
End Function

Public Function ListRecipesBridgeResult(Optional ByVal designsWb As Workbook = Nothing, _
                                        Optional ByVal statusFilter As String = "") As Variant
    ListRecipesBridgeResult = modDesignsQueries.ListRecipes(designsWb, statusFilter)
End Function

Public Function GetRecipeGraphBridgeEncoded(ByVal recipeId As String, _
                                            ByVal recipeVersion As String, _
                                            Optional ByVal designsWb As Workbook = Nothing) As String
    GetRecipeGraphBridgeEncoded = modDesignsQueries.GetRecipeGraph( _
        recipeId, recipeVersion, designsWb)
End Function

Public Function ValidateReleasedRecipeBridgeEncoded(ByVal recipeId As String, _
                                                    ByVal recipeVersion As String, _
                                                    Optional ByVal designsWb As Workbook = Nothing) As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim succeeded As Boolean

    succeeded = modDesignsQueries.ValidateReleasedRecipe( _
        recipeId, recipeVersion, designsWb, errorCode, errorMessage)
    If succeeded Then errorCode = "OK"
    ValidateReleasedRecipeBridgeEncoded = CStr(Abs(CLng(succeeded))) & vbTab & _
        EncodeDesignsBridgeField(errorCode) & vbTab & EncodeDesignsBridgeField(errorMessage)
End Function

Public Function ReadDesignsQueryBridgeResult(ByVal queryName As String, _
                                             Optional ByVal arg1 As String = "", _
                                             Optional ByVal arg2 As String = "", _
                                             Optional ByVal arg3 As String = "", _
                                             Optional ByVal designsWb As Workbook = Nothing) As Variant
    Dim publication As cDesignsEventPublicationSource
    Select Case UCase$(Trim$(queryName))
        Case "PUBLICATION_EVENTS"
            Set publication = New cDesignsEventPublicationSource
            ReadDesignsQueryBridgeResult = publication.Read(arg1, arg2)
        Case "LIST_DESIGNS"
            ReadDesignsQueryBridgeResult = modDesignsQueries.ListDesigns(designsWb, arg1)
        Case "GET_BOM"
            ReadDesignsQueryBridgeResult = modDesignsQueries.GetBOM(arg1, arg2, designsWb)
        Case "GET_BOM_FOR_STATUS"
            ReadDesignsQueryBridgeResult = modDesignsQueries.GetBOMForStatus( _
                arg1, arg2, arg3, designsWb)
        Case "LIST_PROCESSES"
            ReadDesignsQueryBridgeResult = modDesignsQueries.ListProcesses(designsWb, arg1)
        Case "GET_PROCESS_VERSION"
            ReadDesignsQueryBridgeResult = modDesignsQueries.GetProcessVersion(arg1, arg2, designsWb)
        Case "LIST_RECIPES"
            ReadDesignsQueryBridgeResult = modDesignsQueries.ListRecipes(designsWb, arg1)
        Case "GET_RECIPE_GRAPH"
            ReadDesignsQueryBridgeResult = modDesignsQueries.GetRecipeGraph(arg1, arg2, designsWb)
        Case "VALIDATE_RELEASED_RECIPE"
            ReadDesignsQueryBridgeResult = ValidateReleasedRecipeBridgeEncoded( _
                arg1, arg2, designsWb)
    End Select
End Function

Public Function DiagnoseDesignsDomainBridgeResult() As String
    DiagnoseDesignsDomainBridgeResult = modDesignsInit.DiagnoseDesignsDomain()
End Function

Private Function EncodeDesignsBridgeField(ByVal valueIn As String) As String
    EncodeDesignsBridgeField = Replace$(valueIn, vbTab, " ")
    EncodeDesignsBridgeField = Replace$(EncodeDesignsBridgeField, vbCr, " ")
    EncodeDesignsBridgeField = Replace$(EncodeDesignsBridgeField, vbLf, " ")
End Function
