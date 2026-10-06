Attribute VB_Name = "modGuideTransfer"
Option Explicit

' D18 primitive cross-XLAM boundary. Files, models and policy remain in Core.
Public Function CanImport(ByVal context As String, ByRef notice As String, Optional ByRef outcome As String = "") As Boolean
    Dim target As WarehouseTarget, policy As Object, version As Long
    On Error GoTo Invalid
    CanImport = ImportContext(context, target, policy, version, notice, outcome)
    Exit Function
Invalid:
    outcome = "FAILED"
    notice = "Unavailable: guide transfer requires a valid session and policy."
End Function

Private Function ImportContext(ByVal context As String, ByRef target As WarehouseTarget, _
                               ByRef policy As Object, ByRef version As Long, ByRef notice As String, ByRef outcome As String) As Boolean
    outcome = "REJECTED"
    If Not modTrainingReadContext.Read(context, target, policy, notice, version) Then Exit Function
    notice = "Unavailable: guide transfer requires ACTION_PATH_MAINT."
    If Not modAuth.CanPerform("ACTION_PATH_MAINT", modAuth.GetCurrentUserId(), target.WarehouseId, target.StationId) Then
        outcome = "DENIED": Exit Function
    End If
    ImportContext = (context = modActivity.CaptureContext())
    If ImportContext Then notice = ""
End Function

Private Function ExportSource(ByVal context As String, ByVal key As String, ByRef model As Object, _
                              ByRef serialized As String, ByRef notice As String) As Boolean
    Dim records As Collection, visible As Object, policyHash As String, version As Long
    ExportSource = modPublishedGuideDraftSource.Read(context, key, model, records, visible, policyHash, notice, version, Nothing, serialized)
    notice = Replace(Replace(Replace(notice, "editing", "exporting"), "Reopen Edit guide.", "Select the guide again."), "published guide", "guide")
End Function

Public Function CanExport(ByVal context As String, ByVal key As String, ByRef notice As String, Optional ByRef outcome As String = "") As Boolean
    Dim model As Object, serialized As String
    On Error GoTo Invalid
    If Not CanImport(context, notice, outcome) Then Exit Function
    CanExport = ExportSource(context, key, model, serialized, notice)
    Exit Function
Invalid:
    outcome = "FAILED"
    notice = "Unavailable: the exact guide could not be validated for export."
End Function

Public Function ExportGuide(ByVal context As String, ByVal key As String, ByVal path As String, ByRef notice As String, Optional ByRef outcome As String = "") As Boolean
    Dim model As Object, header As Object, raw As String, body As String, content As String
    On Error GoTo Invalid
    If Not CanImport(context, notice, outcome) Then Exit Function
    If Not ExportSource(context, key, model, raw, notice) Then Exit Function
    Set header = CreateObject("Scripting.Dictionary")
    header.Add "SchemaVersion", 1&: header.Add "RecordKind", "GuideTransfer"
    header.Add "TransferId", modTrainingWire.NewId(): header.Add "ExportedAtUTC", modTrainingWire.UtcTimestamp()
    header.Add "ExportedByUserId", modAuth.GetCurrentUserId(): header.Add "SourceWarehouseId", model("WarehouseId")
    body = modTrainingJson.EncodeObject(header)
    body = Left$(body, Len(body) - 1) & ",""Guide"":" & raw & "}"
    content = modGuideTransferWire.Seal(body)
    If content = "" Then notice = "Export exceeds 1 MiB. No guide content was truncated.": Exit Function
    If context <> modActivity.CaptureContext() Then GoTo Invalid
    ExportGuide = modGuideTransferWire.WriteNew(path, content, notice, outcome)
    Exit Function
Invalid:
    outcome = "FAILED"
    notice = "Unavailable: export failed or the session changed."
End Function

Public Function ImportGuide(ByVal context As String, ByVal path As String, ByRef key As String, ByRef notice As String, Optional ByRef outcome As String = "") As Boolean
    Dim target As WarehouseTarget, policy As Object, transfer As Object, model As Object, visible As Object, item As Object
    Dim origin As Object, source As Object, field As Variant, version As Long, hash As String
    On Error GoTo Invalid
    key = ""
    If Not ImportContext(context, target, policy, version, notice, outcome) Then Exit Function
    Set transfer = modGuideTransferWire.Read(path, notice)
    If transfer Is Nothing Then Exit Function
    If Not ImportContext(context, target, policy, version, notice, outcome) Then Exit Function
    Set visible = modEvaluationMatches.PolicyControls(policy, False)
    Set source = transfer("Guide")
    notice = "Hidden by policy: the complete guide cannot be imported. No content was removed."
    For Each field In Array("Steps", "Observations")
        For Each item In source(field)
            If Not modEvaluationMatches.Permitted(visible, CStr(item("ControlId"))) Then Exit Function
        Next item
    Next field
    For Each item In source("ExpectedConclusion")("Steps")
        If Not modEvaluationMatches.Permitted(visible, CStr(item("ControlId"))) Then Exit Function
    Next item
    Set origin = CreateObject("Scripting.Dictionary")
    origin.Add "TransferId", transfer("TransferId"): origin.Add "TransferSha256", transfer("ContentSha256")
    origin.Add "SourceWarehouseId", transfer("SourceWarehouseId")
    For Each field In Array("ActionPathId", "Version", "RecordId", "ContentSha256"): origin.Add CStr(field), source(field): Next field
    Set model = modTrainingJson.DecodeObject(modTrainingJson.EncodeObject(source))
    model.Remove "ContentSha256"
    model("SchemaVersion") = 2&: model("ActionPathId") = modTrainingWire.NewId(): model("RecordId") = modTrainingWire.NewId()
    model("Version") = 1&: model("PreviousRecordId") = "": model("PreviousSha256") = ""
    model("WarehouseId") = target.WarehouseId: model("CreatedByUserId") = modAuth.GetCurrentUserId()
    model("CreatedAtUTC") = modTrainingWire.UtcTimestamp(): model("PolicyVersion") = version
    Set model("TransferOrigin") = origin
    If context <> modActivity.CaptureContext() Then GoTo Invalid
    outcome = "FAILED"
    If Not modGuideStore.Append(target, model, hash, notice) Then Exit Function
    key = CStr(model("ActionPathId")) & "|1|" & hash
    notice = "Guide imported. Imported origin evidence; not locally observed."
    ImportGuide = True: outcome = "COMPLETED"
    Exit Function
Invalid:
    outcome = "FAILED"
    key = "": notice = "Unavailable: import failed or the session changed."
End Function
