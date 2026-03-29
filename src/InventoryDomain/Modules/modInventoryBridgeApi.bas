Attribute VB_Name = "modInventoryBridgeApi"
Option Explicit

Public Function ResolveInventoryWorkbookBridgeResult(Optional ByVal warehouseId As String = "", _
                                                    Optional ByVal inventoryWb As Workbook = Nothing) As Workbook
    Set ResolveInventoryWorkbookBridgeResult = modInventoryApply.ResolveInventoryWorkbook(warehouseId, inventoryWb)
End Function

Public Function EnsureInventorySchemaBridgeResult(Optional ByVal targetWb As Workbook = Nothing) As Object
    Dim result As Object
    Dim report As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    result("Success") = modInventorySchema.EnsureInventorySchema(targetWb, report)
    result("Report") = report
    Set EnsureInventorySchemaBridgeResult = result
End Function

Public Function EnsureInventorySchemaBridgeSuccess(Optional ByVal targetWb As Workbook = Nothing) As Boolean
    Dim report As String

    EnsureInventorySchemaBridgeSuccess = modInventorySchema.EnsureInventorySchema(targetWb, report)
End Function

Public Function EnsureInventorySchemaBridgeReport(Optional ByVal targetWb As Workbook = Nothing) As String
    Dim report As String

    Call modInventorySchema.EnsureInventorySchema(targetWb, report)
    EnsureInventorySchemaBridgeReport = report
End Function

Public Function ApplyEventBridgeResult(ByVal evt As Object, _
                                      Optional ByVal inventoryWb As Workbook = Nothing, _
                                      Optional ByVal runId As String = "") As Object
    Dim result As Object
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    result("Success") = modInventoryApply.ApplyEvent(evt, inventoryWb, runId, statusOut, errorCode, errorMessage)
    result("StatusOut") = statusOut
    result("ErrorCode") = errorCode
    result("ErrorMessage") = errorMessage
    Set ApplyEventBridgeResult = result
End Function

Public Function ApplyEventBridgeEncoded(ByVal evt As Object, _
                                        Optional ByVal inventoryWb As Workbook = Nothing, _
                                        Optional ByVal runId As String = "") As String
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim success As Boolean

    success = modInventoryApply.ApplyEvent(evt, inventoryWb, runId, statusOut, errorCode, errorMessage)
    ApplyEventBridgeEncoded = CStr(Abs(CLng(success))) & vbTab & statusOut & vbTab & errorCode & vbTab & errorMessage
End Function

Public Function RemoveLastBulkLogEntriesBridgeResult(ByVal countToRemove As Long) As Collection
    Dim result As Variant

    On Error Resume Next
    result = Application.Run("'" & ThisWorkbook.Name & "'!modInvMan.RemoveLastBulkLogEntries", countToRemove)
    On Error GoTo 0

    If IsObject(result) Then
        Set RemoveLastBulkLogEntriesBridgeResult = result
    Else
        Set RemoveLastBulkLogEntriesBridgeResult = New Collection
    End If
End Function

Public Sub ReAddBulkLogEntriesBridgeResult(ByVal logDataCollection As Collection)
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!modInvMan.ReAddBulkLogEntries", logDataCollection
    On Error GoTo 0
End Sub

Public Sub ScheduleSourceWorkbookSyncBridgeResult()
    modInventoryInit.ScheduleSourceWorkbookSync
End Sub
