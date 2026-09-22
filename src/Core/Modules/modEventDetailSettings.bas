Attribute VB_Name = "modEventDetailSettings"
Option Explicit

' D18 Admin bridge: only primitive/serialized values cross the package boundary.
Public Function ReadEditor(ByVal context As String, ByRef version As Long, ByRef request As String, ByRef report As String) As Boolean
    Dim target As WarehouseTarget
    On Error GoTo Failed
    request = "": version = 0
    If Not modTrackingPolicyModel.AuthorizeEditor(context, target, report) Then Exit Function
    ReadEditor = modEventDetailStore.ReadProfile(target, version, request, report)
    Exit Function
Failed:
    report = "Detail profile unavailable. No configuration was changed."
End Function

Public Function ReadViewer(ByVal context As String, ByRef version As Long, ByRef request As String, ByRef report As String) As Boolean
    Dim target As WarehouseTarget
    On Error GoTo Failed
    request = "": version = 0
    report = "Detail profile unavailable. Reopen Viewer in the current invSys session."
    If context = "" Or context <> modActivity.CaptureContext() Then Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    ReadViewer = modEventDetailStore.ReadProfile(target, version, request, report)
    If context = modActivity.CaptureContext() Then Exit Function
Failed:
    request = "": version = 0: ReadViewer = False
    report = "Detail profile unavailable. Reopen Viewer in the current invSys session."
End Function

Public Function DefaultRequest() As String
    DefaultRequest = modTrainingJson.EncodeObject(modEventDetailModel.Defaults())
End Function

Public Function FamilyChoices() As String
    FamilyChoices = Join(modEventDetailCatalog.Families(), vbLf)
End Function

Public Function FieldRows(ByVal request As String, ByVal family As String) As String
    Dim model As Object, catalog As Object, row As Variant, ordinal As Long, definition As Object, result As String
    Set model = modEventDetailModel.Decode(request)
    If model Is Nothing Then Exit Function
    Set catalog = modEventDetailCatalog.Fields()
    For ordinal = 1 To catalog.Count
        For Each row In model("Fields")
            If row("EventFamily") = family And row("DisplayOrder") = ordinal Then
                Set definition = catalog(row("FieldId"))
                If result <> "" Then result = result & vbLf
                result = result & row("FieldId") & vbTab & definition("Caption") & vbTab & _
                    CStr(row("Enabled")) & vbTab & CStr(ordinal) & vbTab & CStr(definition("Required"))
            End If
        Next row
    Next ordinal
    FieldRows = result
End Function

Public Function Preview(ByVal request As String, ByVal family As String) As String
    Dim rows As String, line As Variant, values As Variant, catalog As Object, definition As Object
    rows = FieldRows(request, family)
    If rows = "" Then Exit Function
    Set catalog = modEventDetailCatalog.Fields()
    Preview = "SYNTHETIC PREVIEW - no live event data" & vbCrLf & family & vbCrLf
    For Each line In Split(rows, vbLf)
        values = Split(CStr(line), vbTab)
        If values(2) = "True" Then
            Set definition = catalog(values(0))
            Preview = Preview & vbCrLf & definition("Caption") & ": " & definition("Sample")
        End If
    Next line
End Function

Public Function StageField(ByVal request As String, ByVal family As String, ByVal fieldId As String, _
                           ByVal enabled As Boolean, Optional ByVal move As Long = 0) As String
    Dim model As Object, row As Variant, selected As Object, ordinal As Long, catalog As Object, validated As Object
    Set model = modEventDetailModel.Decode(request)
    If model Is Nothing Or move < -1 Or move > 1 Then Exit Function
    Set catalog = modEventDetailCatalog.Fields()
    For Each row In model("Fields")
        If row("EventFamily") = family And row("FieldId") = fieldId Then Set selected = row
    Next row
    If selected Is Nothing Then Exit Function
    selected("Enabled") = enabled
    ordinal = selected("DisplayOrder")
    If ordinal + move < 1 Or ordinal + move > catalog.Count Then move = 0
    If move <> 0 Then
        For Each row In model("Fields")
            If row("EventFamily") = family And row("DisplayOrder") = ordinal + move Then
                row("DisplayOrder") = ordinal
                Exit For
            End If
        Next row
        selected("DisplayOrder") = ordinal + move
    End If
    request = modTrainingJson.EncodeObject(model)
    Set validated = modEventDetailModel.Decode(request)
    If Not validated Is Nothing Then StageField = request
End Function

Public Function SaveProfile(ByVal context As String, ByVal expectedVersion As Long, ByVal request As String, ByRef report As String, Optional ByRef outcome As String = "") As Boolean
    SaveProfile = modEventDetailCommand.Save(context, expectedVersion, request, report, outcome)
End Function
