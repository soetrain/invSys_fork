Attribute VB_Name = "modCuratedGuideSource"
Option Explicit
Option Private Module

' Private Core selection over one acknowledged, immutable loaded publication.
Private mContext As String, mToken As String, mLoadedKey As String, mPolicyHash As String
Private mPolicyVersion As Long, mRevision As Double, mRestricted As Long, mUnavailable As Long
Private mHeader As Object, mCandidates As Object, mSelected As Object

Private Function Current(ByVal context As String, ByRef target As WarehouseTarget, ByRef model As Object, _
                         ByRef header As Object, ByRef visible As Object, ByRef policyHash As String, _
                         ByRef version As Long, ByRef notice As String) As Boolean
    Dim policy As Object
    If Not modTrainingReadContext.Read(context, target, policy, notice, version) Then Exit Function
    notice = "Unavailable: Choose tracked actions requires ACTION_PATH_MAINT."
    If Not modAuth.CanPerform("ACTION_PATH_MAINT", modAuth.GetCurrentUserId(), target.WarehouseId, target.StationId) Then Exit Function
    notice = "Unavailable: load current published Events before choosing tracked actions."
    If Not modLoadedEvents.Read(context, model, header) Then Exit Function
    Set visible = modEvaluationMatches.PolicyControls(policy, False)
    policy("SavedCatalogVersion") = CLng(policy("SavedCatalogVersion"))
    policyHash = modTrainingWire.Sha256(modTrainingJson.EncodeObject(policy))
    Current = modEvaluationModel.IsHash(policyHash) And context = modActivity.CaptureContext()
End Function

Private Function LoadedKey(ByVal header As Object) As String
    LoadedKey = CStr(header("PublicationId")) & "|" & CStr(header("ContentSha256")) & "|" & CStr(header("LoadedAtUTC"))
End Function

Public Function CanOpen(ByVal context As String, ByRef notice As String) As Boolean
    Dim target As WarehouseTarget, model As Object, header As Object, visible As Object, hash As String, version As Long
    On Error GoTo Failed
    CanOpen = Current(context, target, model, header, visible, hash, version, notice)
    If CanOpen Then notice = "Choose original published actions without starting a recording."
Failed:
End Function

Public Function OpenSelection(ByVal context As String, ByRef token As String, ByRef notice As String) As Boolean
    Dim target As WarehouseTarget, model As Object, header As Object, visible As Object, hash As String, version As Long
    Dim group As Object, candidate As Object, id As String, restricted As Boolean
    On Error GoTo Failed
    token = ""
    If Not Current(context, target, model, header, visible, hash, version, notice) Then Exit Function
    If mToken <> "" And mContext = context And mLoadedKey = LoadedKey(header) And mPolicyHash = hash Then
        token = mToken: OpenSelection = True: Exit Function
    End If
    Discard
    Set mCandidates = CreateObject("Scripting.Dictionary"): Set mSelected = CreateObject("Scripting.Dictionary")
    For Each group In model("Groups")
        If group("Source") = "Activity" Then
            restricted = False
            Set candidate = ValidateAction(target, group, visible, restricted)
            If candidate Is Nothing Then
                If restricted Then mRestricted = mRestricted + 1 Else mUnavailable = mUnavailable + 1
            Else
                id = CStr(group("SourceId"))
                If mCandidates.Exists(id) Then GoTo Failed
                mCandidates.Add id, candidate
            End If
        End If
    Next group
    mContext = context: mToken = modTrainingWire.NewId(): mLoadedKey = LoadedKey(header)
    mPolicyHash = hash: mPolicyVersion = version: Set mHeader = header
    token = mToken: notice = "": OpenSelection = True
    Exit Function
Failed:
    Discard: token = "": notice = "Unavailable: published action sources could not be validated."
End Function

Private Function ValidateAction(ByVal target As WarehouseTarget, ByVal group As Object, ByVal visible As Object, ByRef restricted As Boolean) As Object
    Dim record As Object, envelope As Object, requested As Object, result As Object, ids As Object, field As Variant, outcomes As String
    Dim records As New Collection
    On Error GoTo Invalid
    If group("SourceKind") <> "User activity" Or Not modTrainingWire.ValidId(CStr(group("SourceId"))) Then Exit Function
    Set ids = CreateObject("Scripting.Dictionary")
    For Each envelope In group("Lines")
        Set record = VerifiedBody(target, envelope)
        If record Is Nothing Then Exit Function
        If record("ActivityId") <> group("SourceId") Or ids.Exists(CStr(record("RecordId"))) Then Exit Function
        ids.Add CStr(record("RecordId")), True
        If Not modEvaluationMatches.Permitted(visible, CStr(record("ControlId"))) Then restricted = True: Exit Function
        If record("OutcomeCode") = "REQUESTED" Then
            If Not requested Is Nothing Or record("RecordId") <> record("ActivityId") Then Exit Function
            Set requested = record
        End If
        If outcomes <> "" Then outcomes = outcomes & ", "
        outcomes = outcomes & CStr(record("OutcomeCode"))
        records.Add record
    Next envelope
    If requested Is Nothing Then Exit Function
    For Each record In records
        For Each field In Array("ControlId", "UserId", "StationId", "SequenceId", "Ordinal", "PolicyVersion")
            If record(field) <> requested(field) Then Exit Function
        Next field
    Next record
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "Requested", requested: result.Add "Records", records: result.Add "Outcomes", outcomes
    Set ValidateAction = result
Invalid:
End Function

Private Function VerifiedBody(ByVal target As WarehouseTarget, ByVal envelope As Object) As Object
    Dim record As Object, body As String, hash As String
    On Error GoTo Invalid
    If envelope.Count <> 27 Or Not envelope.Exists("ContentSha256") Then Exit Function
    If VarType(envelope("ContentSha256")) <> vbString Then Exit Function
    hash = CStr(envelope("ContentSha256"))
    If Not modEvaluationModel.IsHash(hash) Then Exit Function
    Set record = modTrainingJson.DecodeObject(modTrainingJson.EncodeObject(envelope))
    If record Is Nothing Then Exit Function
    record.Remove "ContentSha256"
    body = modTrainingJson.EncodeObject(record)
    If modTrainingWire.Sha256(body) <> hash Then Exit Function
    If Not modActivityStore.ValidBody(target, CStr(record("RecordId")), body) Then Exit Function
    Set VerifiedBody = record
Invalid:
End Function

Private Function Guard(ByVal context As String, ByVal token As String, ByRef visible As Object, ByRef notice As String) As Boolean
    Dim target As WarehouseTarget, model As Object, header As Object, hash As String, version As Long
    On Error GoTo Invalid
    notice = "Unavailable: the selected actions or captured context changed. Reopen Choose tracked actions."
    If token = "" Or token <> mToken Or context <> mContext Or mCandidates Is Nothing Then Exit Function
    If Not Current(context, target, model, header, visible, hash, version, notice) Then GoTo Invalid
    If LoadedKey(header) <> mLoadedKey Or hash <> mPolicyHash Or version <> mPolicyVersion Then GoTo Invalid
    Guard = True: Exit Function
Invalid:
    Discard
    notice = "Unavailable: the loaded publication, session, permission or policy changed. Reopen Choose tracked actions."
End Function

Public Function ReadRows(ByVal context As String, ByVal token As String, ByVal query As String, ByRef rows As String, _
                         ByRef source As String, ByRef selectedCount As Long, ByRef notice As String) As Boolean
    Dim visible As Object, id As Variant, candidate As Object, record As Object
    On Error GoTo Failed
    rows = "": source = "": selectedCount = 0
    If Not Guard(context, token, visible, notice) Then Exit Function
    For Each id In mCandidates.Keys
        Set candidate = mCandidates(id): Set record = candidate("Requested")
        If query = "" Or InStr(1, CStr(record("Caption")) & " " & CStr(record("ControlId")), query, vbTextCompare) > 0 Then
            rows = rows & CStr(id) & vbTab & CStr(record("Caption")) & vbTab & CStr(record("OccurredAtUTC")) & vbTab & _
                   CStr(candidate("Outcomes")) & vbTab & IIf(mSelected.Exists(id), "1", "0") & vbCrLf
        End If
    Next id
    source = SourceCaption(mHeader): selectedCount = mSelected.Count
    notice = CStr(selectedCount) & " selected. " & CStr(mRestricted) & " hidden by policy; " & CStr(mUnavailable) & " unavailable action group(s)."
    ReadRows = True: Exit Function
Failed:
    rows = "": source = "": selectedCount = 0: Discard
    notice = "Unavailable: the selected action list could not be read."
End Function

Public Function Choose(ByVal context As String, ByVal token As String, ByVal id As String, ByVal selected As Boolean, ByRef notice As String) As Boolean
    Dim visible As Object
    On Error GoTo Failed
    If Not Guard(context, token, visible, notice) Then Exit Function
    notice = "Unavailable: select an original permitted action from this publication."
    If Not mCandidates.Exists(id) Then Exit Function
    If selected <> mSelected.Exists(id) Then
        If selected Then mSelected.Add id, True Else mSelected.Remove id
        mRevision = mRevision + 1
    End If
    Choose = True: notice = ""
Failed:
End Function

Public Function ReadSelected(ByVal context As String, ByVal token As String, ByRef header As Object, ByRef records As Collection, _
                             ByRef visible As Object, ByRef policyHash As String, ByRef version As Long, ByRef binding As String, ByRef notice As String) As Boolean
    Dim id As Variant, record As Object, ids As Object
    On Error GoTo Failed
    Set header = Nothing: Set records = Nothing: policyHash = "": binding = "": version = 0
    If Not Guard(context, token, visible, notice) Then Exit Function
    notice = "Select at least one published action to create a guide."
    If mSelected.Count = 0 Then Exit Function
    Set records = New Collection: Set ids = CreateObject("Scripting.Dictionary")
    For Each id In mCandidates.Keys
        If mSelected.Exists(id) Then
            For Each record In mCandidates(id)("Records")
                If ids.Exists(CStr(record("RecordId"))) Then GoTo Failed
                ids.Add CStr(record("RecordId")), True: records.Add record
            Next record
        End If
    Next id
    Set header = mHeader: policyHash = mPolicyHash: version = mPolicyVersion
    binding = token & "|" & CStr(mRevision): notice = "": ReadSelected = True
    Exit Function
Failed:
    Set header = Nothing: Set records = Nothing: policyHash = "": binding = "": version = 0
    Discard: notice = "Unavailable: selected original observations could not be validated."
End Function

Public Function SourceCaption(ByVal header As Object) As String
    SourceCaption = "Selected published actions: " & CStr(header("PublicationId")) & vbCrLf & _
                    "SHA-256: " & CStr(header("ContentSha256")) & vbCrLf & _
                    "Loaded: " & CStr(header("LoadedAtUTC")) & ". No recorded sequence; authored order is not execution evidence."
End Function

Public Sub CloseSelection(ByVal context As String, ByVal token As String)
    If context = mContext And token = mToken Then Discard
End Sub

Private Sub Discard()
    mContext = "": mToken = "": mLoadedKey = "": mPolicyHash = "": mPolicyVersion = 0
    mRevision = 0: mRestricted = 0: mUnavailable = 0
    Set mHeader = Nothing: Set mCandidates = Nothing: Set mSelected = Nothing
End Sub
