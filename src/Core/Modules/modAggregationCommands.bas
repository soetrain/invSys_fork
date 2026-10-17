Attribute VB_Name = "modAggregationCommands"
Option Explicit

Private mWriting As Boolean
Private mContext As String
Private mSources As Collection
Private mCandidate As String

Public Function CaptureForm(Optional ByVal bookName As String = "") As String
    CaptureForm = modAggregationAuthority.Capture(bookName)
End Function

Public Sub ReleaseForm(ByVal context As String)
    modAggregationAuthority.Release context
End Sub

Public Function CheckForm(ByVal context As String, Optional ByVal paths As String = "") As Boolean
    Dim sources As Collection
    On Error GoTo Failed
    If paths = "" Then
        CheckForm = modAggregationAuthority.Check(context)
    Else
        Set sources = modAggregationAuthority.SelectSources(context, paths)
        If Not sources Is Nothing Then CheckForm = modAggregationAuthority.Joint(context, sources)
    End If
Failed:
End Function

Public Function RegisterSource(ByVal context As String, ByVal warehouse As String, ByVal configPath As String, ByVal snapshot As String) As String
    Dim saved As Object, source As Object, path As String
    On Error GoTo Failed
    If Not CheckForm(context) Then Exit Function
    Set saved = modAggregationAuthority.Entry(context)
    Set source = modAggregationSources.Inspect(warehouse, configPath, snapshot)
    If source Is Nothing Then Exit Function
    If Not CheckForm(context) Then Exit Function
    path = CStr(source("Path"))
    If saved("Sources").Exists(path) Then saved("Sources").Remove path
    saved("Sources").Add path, source
    RegisterSource = CStr(source("Hash"))
Failed:
End Function

Public Function AuditIdentity(ByVal context As String) As String
    Dim saved As Object
    Set saved = modAggregationAuthority.Entry(context)
    If Not saved Is Nothing Then AuditIdentity = saved("User") & vbTab & saved("Warehouse") & vbTab & saved("Station")
End Function

Public Function BoundWorkbookName(ByVal context As String) As String
    Dim saved As Object, book As Workbook
    On Error GoTo Unavailable
    Set saved = modAggregationAuthority.Entry(context)
    If saved Is Nothing Then Exit Function
    Set book = saved("Workbook")
    BoundWorkbookName = book.Name
Unavailable:
End Function

' These guards are used only by the active command's typed Core writer.
Public Function CheckPhase(ByVal context As String, ByVal candidate As String) As Boolean
    If Not mWriting Or context <> mContext Or candidate <> mCandidate Then Exit Function
    CheckPhase = modAggregationAuthority.Joint(context, mSources)
End Function

Public Function IsWriting() As Boolean
    ' Either surviving marker keeps the headless writer out of an active attempt.
    IsWriting = (mWriting Or mContext <> "")
End Function

Public Function Aggregate(ByVal context As String, ByVal paths As String, ByRef report As String, Optional ByRef outcome As String = "") As Boolean
    Dim saved As Object, record As Object, candidateHash As String, rawReport As String, promoted As Boolean
    Dim source As Variant, rows As Long, events As Boolean, alerts As Boolean, began As Boolean, state As String, sourcePaths As String
    On Error GoTo Failed
    report = "Aggregation is busy. Wait for the current attempt to finish.": outcome = "REJECTED"
    If mWriting Then Exit Function
    report = "The source session, feed or selected snapshots changed. Reopen Aggregation and refresh the selection."
    outcome = "DENIED"
    Set mSources = modAggregationAuthority.SelectSources(context, paths)
    If mSources Is Nothing Then Exit Function
    If Not modAggregationAuthority.Joint(context, mSources) Then Exit Function
    Set saved = modAggregationAuthority.Entry(context)
    report = "A previous attempt needs review. Its output and retry evidence are retained. Reopen Aggregation to resolve it."
    outcome = "INDETERMINATE"
    If modAggregationRetry.Pending(CStr(saved("Destination"))) Then Exit Function
    outcome = "DENIED"
    report = "The source session, feed or selected snapshots changed. Reopen Aggregation and refresh the selection."
    For Each source In mSources
        If sourcePaths <> "" Then sourcePaths = sourcePaths & vbLf
        sourcePaths = sourcePaths & CStr(source("Path"))
    Next source
    events = Application.EnableEvents: alerts = Application.DisplayAlerts: began = True
    Application.EnableEvents = False: Application.DisplayAlerts = False
    mWriting = True: mContext = context
    ' Inspection has completed. The next joint check precedes any output artifact.
    If Not modAggregationAuthority.Joint(context, mSources) Then GoTo Finished
    Set record = modAggregationRetry.Reserve(CStr(saved("Destination")), mSources)
    mCandidate = CStr(record("Candidate"))
    outcome = "FAILED": report = "Aggregation did not complete. Retry evidence is retained; no rollback is claimed."
    If Not CheckPhase(context, mCandidate) Then GoTo Finished
    If Not modHqAggregator.GenerateGlobalSnapshotFromFiles(sourcePaths, mCandidate, rawReport, context, True) Then GoTo Finished
    If Not CheckPhase(context, mCandidate) Then GoTo Finished
    If Not modAggregationSources.Complete(mCandidate, mSources) Then GoTo Finished
    candidateHash = modTrainingWire.FileSha256(mCandidate)
    If candidateHash = "" Then GoTo Finished
    record("CandidateHash") = candidateHash
    record("Phase") = "CANDIDATE_VERIFIED"
    If Not CheckPhase(context, mCandidate) Then GoTo Finished
    modAggregationRetry.Save record
    CreateObject("Scripting.FileSystemObject").CopyFile mCandidate, CStr(record("Promotion")), False
    If Not CheckPhase(context, mCandidate) Then GoTo Finished
    If modAggregationRetry.Persistence(record) <> "PRESERVED" Then GoTo Finished
    record("Phase") = "PROMOTION_PENDING": record("Persistence") = "UNKNOWN"
    modAggregationRetry.Save record
    If Not CheckPhase(context, mCandidate) Then GoTo Finished
    promoted = modAggregationFiles.Promote(CStr(record("Promotion")), CStr(record("Destination")), CStr(record("Backup")))
    ' Verify actual saved bytes even when the native call reports failure.
    state = modAggregationRetry.Persistence(record)
    If state <> "COMPLETE" Then outcome = "INDETERMINATE": GoTo Finished
    If Not modAggregationSources.Complete(CStr(record("Destination")), mSources) Then outcome = "INDETERMINATE": GoTo Finished
    If Not CheckPhase(context, mCandidate) Then
        outcome = "COMPLETED_STOPPED"
        report = "The complete output was saved. The source context changed; reopen Aggregation to review the retained attempt."
        GoTo Finished
    End If
    record("Phase") = "COMPLETE_VERIFIED": record("Persistence") = "COMPLETE"
    modAggregationRetry.Save record
    modAggregationRetry.ResolveComplete record
    For Each source In mSources: rows = rows + source("Rows").Count: Next source
    Aggregate = True: outcome = "COMPLETED"
    report = "Complete advisory output saved. Sources=" & CStr(mSources.Count) & "; Rows=" & CStr(rows) & "."
Finished:
    modAggregationOutput.CloseOwnedStaging
    If Not record Is Nothing And Not Aggregate And outcome <> "COMPLETED_STOPPED" Then
        state = modAggregationRetry.Persistence(record)
        If state = "PRESERVED" Then
            report = "Aggregation did not complete. Existing feed bytes are preserved; retry evidence is retained."
        ElseIf state = "COMPLETE" Then
            report = "The complete output was saved, but the attempt needs review. Retry evidence is retained."
        Else
            outcome = "INDETERMINATE"
            report = "Aggregation did not complete. Output persistence is unknown; its bytes and retry evidence are retained."
        End If
    End If
    If Not saved Is Nothing And Not record Is Nothing And Not Aggregate Then saved("Used") = True
    If began Then Application.EnableEvents = events: Application.DisplayAlerts = alerts
    mWriting = False: mContext = "": mCandidate = "": Set mSources = Nothing
    Exit Function
Failed:
    outcome = "FAILED"
    report = "Aggregation did not complete. Inspect the retained attempt before retrying; no rollback is claimed."
    Resume Finished
End Function

Public Function NeedsReview(ByVal context As String) As Boolean
    Dim saved As Object
    Set saved = modAggregationAuthority.Entry(context)
    If Not saved Is Nothing Then NeedsReview = modAggregationRetry.Pending(CStr(saved("Destination")))
End Function

Public Function ResolveRetry(ByVal context As String, ByVal paths As String, ByVal confirmed As Boolean, ByRef report As String) As Boolean
    Dim saved As Object, sources As Collection, folder As Object, file As Object, record As Object, records As New Collection
    Dim fso As Object, state As String, item As Variant, field As Variant, expected As String
    On Error GoTo Failed
    report = "Attempt review cancelled. Existing output and retry evidence are retained."
    If Not confirmed Then Exit Function
    report = "Aggregation is busy. Wait for the current attempt to finish."
    If mWriting Then Exit Function
    report = "The source session, feed or selected snapshots changed. Reopen Aggregation and refresh the selection."
    Set sources = modAggregationAuthority.SelectSources(context, paths)
    If sources Is Nothing Then Exit Function
    If Not modAggregationAuthority.Joint(context, sources) Then Exit Function
    Set saved = modAggregationAuthority.Entry(context)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(fso.GetParentFolderName(CStr(saved("Destination"))))
    For Each file In folder.Files
        If LCase$(file.Name) Like "invsys.aggregation.*.retry.json" Then
            Set record = modAggregationRetry.Read(CStr(file.Path))
            If record Is Nothing Then GoTo Failed
            If Not modTrainingWire.ValidId(CStr(record("Attempt"))) Then GoTo Failed
            If record("Version") <> 1 Or record("Destination") <> saved("Destination") Then GoTo Failed
            For Each field In Array("Candidate", "Promotion", "Backup", "Receipt")
                Select Case CStr(field)
                    Case "Candidate": expected = ".candidate.xlsb"
                    Case "Promotion": expected = ".promotion.xlsb"
                    Case "Backup": expected = ".previous.xlsb"
                    Case "Receipt": expected = ".retry.json"
                End Select
                expected = fso.BuildPath(folder.Path, "invSys.Aggregation." & record("Attempt") & expected)
                If record(CStr(field)) <> expected Then GoTo Failed
            Next field
            state = modAggregationRetry.Persistence(record)
            If state = "UNKNOWN" Then GoTo Failed
            record("Persistence") = state
            records.Add record
        End If
    Next file
    For Each item In records
        If Not modAggregationAuthority.Joint(context, sources) Then GoTo Failed
        item("Phase") = "RESOLVED_VERIFIED"
        modAggregationRetry.Save item
        For Each field In Array("Candidate", "Promotion", "Backup")
            If Not modAggregationAuthority.Joint(context, sources) Then GoTo Failed
            If fso.FileExists(CStr(item(CStr(field)))) Then fso.DeleteFile CStr(item(CStr(field))), False
        Next field
        If Not modAggregationAuthority.Joint(context, sources) Then GoTo Failed
        fso.DeleteFile CStr(item("Receipt")), False
    Next item
    ResolveRetry = True: report = "Known output state reviewed. A new authorized attempt may build fresh output."
    Exit Function
Failed:
    report = "Output state could not be verified or resolved. Its bytes and retry evidence are retained."
End Function
