Attribute VB_Name = "modRibbonRuntimeStatus"
Option Explicit

Public Function GetStatusLabel(ByVal controlId As String) As String
    EnsureRuntimeStatusConfigLoaded

    Select Case Trim$(controlId)
        Case "btnRuntimeWarehouse"
            GetStatusLabel = "Warehouse: " & ValueOrPlaceholderStatus(modConfig.GetWarehouseId()) & _
                             " | Station: " & ValueOrPlaceholderStatus(modConfig.GetStationId())
        Case "btnRuntimeDataRoot"
            GetStatusLabel = "Data root: " & ValueOrPlaceholderStatus(modConfig.GetString("PathDataRoot", ""))
        Case "btnRuntimeInboxRoot"
            GetStatusLabel = "Inbox root: " & ValueOrPlaceholderStatus(modConfig.GetString("PathInboxRoot", ""))
        Case "btnRuntimeProcessor"
            GetStatusLabel = "Processor: " & ValueOrPlaceholderStatus(modConfig.GetString("ProcessorServiceUserId", "svc_processor"))
        Case "btnRuntimeHqAggregator"
            GetStatusLabel = "HQ aggregator: " & ResolveHqAggregatorLabelStatus()
        Case Else
            GetStatusLabel = "Runtime context"
    End Select
End Function

Public Sub RefreshRuntimeContext()
    Dim report As String

    If modConfig.LoadConfig("", "") Then
        report = "Warehouse: " & ValueOrPlaceholderStatus(modConfig.GetWarehouseId()) & vbCrLf & _
                 "Station: " & ValueOrPlaceholderStatus(modConfig.GetStationId()) & vbCrLf & _
                 "Data root: " & ValueOrPlaceholderStatus(modConfig.GetString("PathDataRoot", "")) & vbCrLf & _
                 "Inbox root: " & ValueOrPlaceholderStatus(modConfig.GetString("PathInboxRoot", "")) & vbCrLf & _
                 "Processor: " & ValueOrPlaceholderStatus(modConfig.GetString("ProcessorServiceUserId", "svc_processor")) & vbCrLf & _
                 "HQ aggregator: " & ResolveHqAggregatorLabelStatus()
        MsgBox report, vbInformation, "invSys Runtime Context"
    Else
        MsgBox "Runtime config could not be loaded." & vbCrLf & vbCrLf & modConfig.Validate(), vbExclamation, "invSys Runtime Context"
    End If
End Sub

Private Sub EnsureRuntimeStatusConfigLoaded()
    If Trim$(modConfig.GetWarehouseId()) = "" Then
        On Error Resume Next
        Call modConfig.LoadConfig("", "")
        On Error GoTo 0
    End If
End Sub

Private Function ResolveHqAggregatorLabelStatus() As String
    Dim sharePointRoot As String

    sharePointRoot = Trim$(modConfig.GetString("PathSharePointRoot", ""))
    If sharePointRoot = "" Then
        ResolveHqAggregatorLabelStatus = "<not configured>"
    Else
        ResolveHqAggregatorLabelStatus = "Admin scheduled aggregation via " & NormalizeFolderForStatus(sharePointRoot) & "\Snapshots"
    End If
End Function

Private Function ValueOrPlaceholderStatus(ByVal valueIn As String) As String
    valueIn = Trim$(valueIn)
    If valueIn = "" Then
        ValueOrPlaceholderStatus = "<not configured>"
    Else
        ValueOrPlaceholderStatus = valueIn
    End If
End Function

Private Function NormalizeFolderForStatus(ByVal folderPath As String) As String
    NormalizeFolderForStatus = Trim$(Replace$(folderPath, "/", "\"))
    Do While Len(NormalizeFolderForStatus) > 3 And Right$(NormalizeFolderForStatus, 1) = "\"
        NormalizeFolderForStatus = Left$(NormalizeFolderForStatus, Len(NormalizeFolderForStatus) - 1)
    Loop
End Function
