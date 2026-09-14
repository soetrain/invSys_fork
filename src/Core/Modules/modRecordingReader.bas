Attribute VB_Name = "modRecordingReader"
Option Explicit
Option Private Module

Public Function Versions(ByVal target As WarehouseTarget) As Object
    Dim root As String, fso As Object, file As Object, parts As Variant, id As String, version As Long, result As Object
    On Error GoTo Invalid
    root = modRecordingJournal.JournalRoot(target, False)
    If root = "" Then Exit Function
    Set result = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each file In fso.GetFolder(root).Files
        If LCase$(fso.GetExtensionName(file.Name)) = "json" Then
            parts = Split(file.Name, ".")
            If UBound(parts) <> 2 Then GoTo Invalid
            id = CStr(parts(0))
            If Not modTrainingWire.ValidId(id) Then GoTo Invalid
            If Not IsNumeric(parts(1)) Then GoTo Invalid
            version = CLng(parts(1))
            If CStr(version) <> CStr(parts(1)) Or version < 1 Or version > 514 Then GoTo Invalid
            If (file.Attributes And &H400) <> 0 Then GoTo Invalid
            If Not result.Exists(id) Then result.Add id, version
            If version > CLng(result(id)) Then result(id) = version
        End If
    Next file
    Set Versions = result
Invalid:
End Function

Public Function ReadRun(ByVal target As WarehouseTarget, ByVal pathId As String, ByVal lastVersion As Long, _
                        ByRef header As Object, ByRef notice As String) As Collection
    Dim entry As Object, first As Object, previous As Object, observations As New Collection, record As Variant
    Dim ids As Object, actions As Object, outcomes As Object, version As Long, index As Long, field As Variant
    Dim id As String, body As String, attempt As Object
    On Error GoTo Invalid
    notice = "Incomplete evidence: the saved recording journal is missing or corrupt."
    Set header = Nothing
    If lastVersion < 1 Or lastVersion > 514 Or Not modTrainingWire.ValidId(pathId) Then Exit Function
    Set ids = CreateObject("Scripting.Dictionary"): Set actions = CreateObject("Scripting.Dictionary")
    Set outcomes = CreateObject("Scripting.Dictionary")
    For version = 1 To lastVersion
        Set entry = modRecordingJournal.ReadEntry(target, pathId, version)
        If entry Is Nothing Then Exit Function
        If ids.Exists(entry("RecordId")) Then Exit Function
        ids.Add entry("RecordId"), True
        If version = 1 Then
            Set first = entry
        Else
            If previous("Lifecycle") <> "Recording" Then Exit Function
            If entry("PreviousRecordId") <> previous("RecordId") Or entry("PreviousSha256") <> previous("ContentSha256") Then Exit Function
            If entry("CreatedAtUTC") < previous("CreatedAtUTC") Then Exit Function
            For Each field In Array("SchemaVersion", "ActionPathId", "SequenceId", "CreatedByUserId", "WarehouseId", "OriginWarehouseId", _
                "PolicyVersion", "CatalogVersion", "PackageSetVersion", "BuildIdentity")
                If entry(field) <> first(field) Then Exit Function
            Next field
        End If
        Select Case entry("RecordType")
            Case "Observation"
                Set record = entry("Observations")(1)
                If ids.Exists(record("RecordId")) Then Exit Function
                ids.Add record("RecordId"), True
                id = record("ActivityId")
                If record("OutcomeCode") = "REQUESTED" Then
                    If actions.Exists(id) Or record("Ordinal") <> actions.Count + 1 Then Exit Function
                    actions.Add id, record
                Else
                    If Not actions.Exists(id) Or outcomes.Exists(id) Then Exit Function
                    Set attempt = actions(id)
                    For Each field In Array("ActivityId", "SequenceId", "Ordinal", "ControlId", "UserId", "StationId", _
                        "WarehouseId", "SourceRole", "Caption", "Surface", "PolicyVersion", "CatalogVersion", "PackageSetVersion", "BuildIdentity")
                        If record(field) <> attempt(field) Then Exit Function
                    Next field
                    outcomes.Add id, True
                End If
                If entry("ActionCount") <> actions.Count Then Exit Function
                observations.Add record
            Case "Close"
                If version <> lastVersion Or entry("ActionCount") <> actions.Count Then Exit Function
                If entry("Observations").Count <> observations.Count Then Exit Function
                For index = 1 To observations.Count
                    If modTrainingJson.EncodeObject(entry("Observations")(index)) <> modTrainingJson.EncodeObject(observations(index)) Then Exit Function
                Next index
                If entry("Lifecycle") = "Stopped" And actions.Count <> outcomes.Count Then Exit Function
        End Select
        Set previous = entry
    Next version
    Set header = entry
    notice = "Interrupted: no closing record is available. Recording was not resumed."
    If entry("RecordType") = "Close" Then
        Select Case entry("Lifecycle")
            Case "Stopped": notice = "Stopped. Capture frozen; conclusion not evaluated."
            Case "Cancelled": notice = "Cancelled. Observed work was not undone."
            Case Else: notice = "Incomplete evidence: " & Replace$(LCase$(entry("ReasonCode")), "_", " ") & "."
        End Select
    End If
    Set ReadRun = observations
Invalid:
End Function
