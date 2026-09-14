Attribute VB_Name = "modActionPathRead"
Option Explicit

' D18 primitive cross-XLAM read boundary. No UI, writes or workflow dispatch.
Public Function ListPaths(ByVal context As String, ByVal search As String, ByRef rows As String, ByRef notice As String) As Boolean
    Dim target As WarehouseTarget, policy As Object, versions As Object, id As Variant, start As Object, line As String
    On Error GoTo Invalid
    rows = ""
    If Not ReadContext(context, target, policy, notice) Then Exit Function
    Set versions = modRecordingReader.Versions(target)
    notice = "Unavailable: the recording library could not be read."
    If versions Is Nothing Then Exit Function
    For Each id In versions.Keys
        Set start = modRecordingJournal.ReadEntry(target, CStr(id), 1)
        line = CStr(id) & vbTab & "Recorded sequence" & vbTab & "Unavailable"
        If Not start Is Nothing Then line = CStr(id) & vbTab & "Recorded sequence" & vbTab & CStr(start("CreatedAtUTC"))
        If search = "" Or InStr(1, line, search, vbTextCompare) > 0 Then rows = rows & line & vbCrLf
    Next id
    If context <> modActivity.CaptureContext() Then GoTo Invalid
    notice = "Select a recording to inspect its saved evidence."
    ListPaths = True
    Exit Function
Invalid:
    rows = "": notice = "Unavailable: the invSys session, policy or recording library changed. Reopen Viewer."
End Function

Public Function ReadPath(ByVal context As String, ByVal pathId As String, ByRef evidence As String, ByRef notice As String) As Boolean
    Dim target As WarehouseTarget, policy As Object, versions As Object, header As Object, records As Collection
    Dim record As Variant, row As Variant, visible As Object, definition As Object, hidden As Long, reference As Variant
    On Error GoTo Invalid
    evidence = ""
    If Not ReadContext(context, target, policy, notice) Then Exit Function
    Set versions = modRecordingReader.Versions(target)
    notice = "Incomplete evidence: the selected recording is unavailable."
    If versions Is Nothing Then Exit Function
    If Not versions.Exists(pathId) Then Exit Function
    Set records = modRecordingReader.ReadRun(target, pathId, CLng(versions(pathId)), header, notice)
    If records Is Nothing Then Exit Function
    Set visible = CreateObject("Scripting.Dictionary")
    For Each row In policy("Controls")
        Set definition = modActivityCatalog.Control(CStr(row("ControlId")), CLng(policy("SavedCatalogVersion")))
        If Not definition Is Nothing Then
            visible.Add row("ControlId"), CBool(row("Visible")) And _
                (definition("Role") <> "Admin" Or CBool(policy("AdminViewerEventLoggingEnabled")))
        End If
    Next row
    evidence = "Recorded sequence" & vbCrLf & "Action Path: " & pathId & vbCrLf & _
        "Sequence: " & CStr(header("SequenceId")) & vbCrLf & "Saved version: " & CStr(header("Version")) & vbCrLf
    For Each record In records
        If CanShow(record, visible) Then
            evidence = evidence & vbCrLf & CStr(record("Ordinal")) & ". " & CStr(record("Caption")) & _
                " - " & CStr(record("OutcomeCode")) & vbCrLf & CStr(record("ActivityId")) & vbCrLf & _
                CStr(record("OccurredAtUTC")) & vbCrLf
            For Each reference In record("SourceEventRefs")
                evidence = evidence & "Source event: " & CStr(reference("EventId")) & _
                    " (" & CStr(reference("SubmissionState")) & "; application not asserted)" & vbCrLf
            Next reference
        Else
            hidden = hidden + 1
        End If
    Next record
    If hidden > 0 Then notice = "Incomplete evidence: current tracking policy restricts " & CStr(hidden) & " observation(s)."
    If CLng(header("CatalogVersion")) < modActivityCatalog.CATALOG_VERSION Then
        notice = notice & " Older release."
    ElseIf CStr(header("PackageSetVersion")) <> CStr(ThisWorkbook.CustomDocumentProperties("invSysPackageSetVersion").Value) Or _
           CStr(header("BuildIdentity")) <> CStr(ThisWorkbook.CustomDocumentProperties("invSysBuildIdentity").Value) Then
        notice = notice & " Different release/build; relative age unavailable."
    End If
    If context <> modActivity.CaptureContext() Then GoTo Invalid
    ReadPath = True
    Exit Function
Invalid:
    evidence = "": notice = "Unavailable: the invSys session, policy or recording library changed. Reopen Viewer."
End Function

Private Function CanShow(ByVal record As Object, ByVal visible As Object) As Boolean
    If visible.Exists(record("ControlId")) Then CanShow = CBool(visible(record("ControlId")))
End Function

Private Function ReadContext(ByVal context As String, ByRef target As WarehouseTarget, ByRef policy As Object, ByRef notice As String) As Boolean
    Dim version As Long, collect As Boolean, visible As Boolean
    notice = "Unavailable: the invSys session or warehouse changed. Reopen Viewer."
    If context = "" Or context <> modActivity.CaptureContext() Then Exit Function
    Set target = modNasConnection.GetCurrentTarget(): Set policy = CreateObject("Scripting.Dictionary")
    If Not modActivityPolicy.ReadPolicy(target, "ADMIN_SETTINGS_SAVE_VALUE", version, collect, visible, notice, policy) Then Exit Function
    ReadContext = (context = modActivity.CaptureContext())
End Function
