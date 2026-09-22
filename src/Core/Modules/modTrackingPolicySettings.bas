Attribute VB_Name = "modTrackingPolicySettings"
Option Explicit

' D18 UI boundary: only primitive/serialized values cross the project reference.
Public Function ReadEditor(ByVal context As String, ByRef version As Long, _
                           ByRef request As String, ByRef report As String, _
                           Optional ByRef savedCatalogVersion As Long = 0) As Boolean
    Dim target As WarehouseTarget, model As Object, collect As Boolean, visible As Boolean
    On Error GoTo Failed
    request = "": version = 0
    If Not modTrackingPolicyModel.AuthorizeEditor(context, target, report) Then Exit Function
    Set model = CreateObject("Scripting.Dictionary")
    If Not modActivityPolicy.ReadPolicy(target, "ADMIN_SETTINGS_SAVE_VALUE", version, collect, visible, report, model) Then Exit Function
    savedCatalogVersion = CLng(model("SavedCatalogVersion"))
    model.Remove "SavedCatalogVersion"
    request = modTrainingJson.EncodeObject(model)
    report = IIf(version = 0, "Built-in defaults; no tracking policy saved.", "Tracking policy version " & CStr(version) & " loaded.")
    ReadEditor = True
    Exit Function
Failed:
    request = ""
    report = "Tracking policy unavailable. No configuration was changed."
End Function

Public Function DefaultRequest() As String
    DefaultRequest = modTrainingJson.EncodeObject(modTrackingPolicyModel.Defaults())
End Function

Public Function FieldValue(ByVal request As String, ByVal field As String) As String
    Dim model As Object
    Set model = modTrackingPolicyModel.Decode(request)
    If model Is Nothing Then Exit Function
    Select Case field
        Case "DefaultView", "ViewerActionPathCaptureEnabled", "AdminViewerEventLoggingEnabled"
            FieldValue = CStr(model(field))
    End Select
End Function

Public Function ControlRows(ByVal request As String, Optional ByVal savedCatalogVersion As Long = 0) As String
    Dim model As Object, row As Variant, definition As Object, result As String
    Set model = modTrackingPolicyModel.Decode(request)
    If model Is Nothing Then Exit Function
    For Each row In model("Controls")
        Set definition = modActivityCatalog.Control(CStr(row("ControlId")))
        If result <> "" Then result = result & vbLf
        result = result & row("ControlId") & vbTab & definition("OwnerId") & vbTab & definition("Caption") & vbTab & _
            CStr(row("Collect")) & vbTab & CStr(row("Visible")) & vbTab & CStr(row("SequenceEligible"))
        If savedCatalogVersion > 0 Then
            If modActivityCatalog.Control(CStr(row("ControlId")), savedCatalogVersion) Is Nothing Then
                result = result & vbTab & "Unavailable until policy update"
            Else
                result = result & vbTab & "Available"
            End If
        Else
            result = result & vbTab & "Staged defaults"
        End If
    Next row
    ControlRows = result
End Function

Public Function StageValue(ByVal request As String, ByVal field As String, ByVal value As String, _
                           Optional ByVal controlId As String = "") As String
    Dim model As Object, row As Variant, updated As Object, found As Boolean
    On Error GoTo Invalid
    Set model = modTrackingPolicyModel.Decode(request)
    If model Is Nothing Then Exit Function
    If controlId = "" Then
        Select Case field
            Case "DefaultView": model(field) = value
            Case "ViewerActionPathCaptureEnabled", "AdminViewerEventLoggingEnabled"
                If value <> "True" And value <> "False" Then Exit Function
                model(field) = CBool(value)
            Case Else: Exit Function
        End Select
    Else
        If field <> "Collect" And field <> "Visible" And field <> "SequenceEligible" Then Exit Function
        If value <> "True" And value <> "False" Then Exit Function
        For Each row In model("Controls")
            If row("ControlId") = controlId Then row(field) = CBool(value): found = True
        Next row
        If Not found Then Exit Function
    End If
    request = modTrainingJson.EncodeObject(model)
    Set updated = modTrackingPolicyModel.Decode(request)
    If Not updated Is Nothing Then StageValue = request
Invalid:
End Function

Public Function SavePolicy(ByVal context As String, ByVal expectedVersion As Long, _
                           ByVal request As String, ByRef report As String, Optional ByRef outcome As String = "") As Boolean
    SavePolicy = modTrackingPolicyCommand.Save(context, expectedVersion, request, report, outcome)
End Function
