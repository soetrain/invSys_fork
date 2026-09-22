Attribute VB_Name = "modGuideLibraryRead"
Option Explicit

' D18 read-only primitive cross-XLAM boundary; models and policy stay in Core.
Public Function ListGuides(ByVal context As String, ByVal search As String, ByRef rows As String, ByRef notice As String) As Boolean
    Dim target As WarehouseTarget, policy As Object, root As String, fso As Object, file As Object
    Dim entries As Object, parts As Variant, version As Long, sortKey As Variant, keys As Variant
    Dim model As Object, line As String, tags As String, invalid As Long, i As Long, j As Long, swap As Variant
    On Error GoTo Failed
    rows = ""
    If Not modTrainingReadContext.Read(context, target, policy, notice) Then Exit Function
    root = modRecordingJournal.ChildRoot(target, "Guides", False)
    notice = "Unavailable: no readable published-guide folder."
    If root = "" Then ListGuides = True: Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject"): Set entries = CreateObject("Scripting.Dictionary")
    For Each file In fso.GetFolder(root).Files
        If LCase$(fso.GetExtensionName(file.Name)) = "json" Then
            parts = Split(file.Name, ".")
            If UBound(parts) <> 2 Then GoTo BadEntry
            If Not modTrainingWire.ValidId(CStr(parts(0))) Then GoTo BadEntry
            If Not ParseVersion(CStr(parts(1)), version) Then GoTo BadEntry
            Set model = modGuideStore.ReadChain(target, CStr(parts(0)), version)
            If model Is Nothing Then GoTo BadEntry
            tags = modGuideModel.TagsText(model)
            If search = "" Or InStr(1, CStr(model("Name")) & " " & tags & " " & CStr(model("ActionPathId")), search, vbTextCompare) > 0 Then
                line = CStr(model("ActionPathId")) & "|" & CStr(version) & "|" & CStr(model("ContentSha256")) & vbTab & _
                    ListText(CStr(model("Name"))) & vbTab & "Version " & CStr(version) & " - " & ListText(tags)
                entries.Add CStr(model("ActionPathId")) & "|" & Format$(version, "0000000000"), line
            End If
        End If
        GoTo NextEntry
BadEntry:
        invalid = invalid + 1
NextEntry:
    Next file
    If entries.Count > 0 Then
        keys = entries.Keys
        For i = 1 To UBound(keys)
            swap = keys(i): j = i - 1
            Do While j >= 0
                If StrComp(CStr(keys(j)), CStr(swap), vbBinaryCompare) <= 0 Then Exit Do
                keys(j + 1) = keys(j): j = j - 1
            Loop
            keys(j + 1) = swap
        Next i
        For Each sortKey In keys: rows = rows & CStr(entries(sortKey)) & vbCrLf: Next sortKey
    End If
    notice = "Select a published guide version. Authored instructions do not prove that a task ran."
    If invalid > 0 Then notice = "Unavailable: " & CStr(invalid) & " guide version(s) have invalid or missing evidence."
    If context <> modActivity.CaptureContext() Then GoTo Failed
    ListGuides = True
    Exit Function
Failed:
    rows = "": notice = "Unavailable: the published-guide library, session or policy could not be validated."
End Function

Public Function ReadGuide(ByVal context As String, ByVal key As String, ByRef instructions As String, _
                          ByRef observations As String, ByRef provenance As String, ByRef notice As String) As Boolean
    Dim target As WarehouseTarget, policy As Object, model As Object, visible As Object, parts As Variant
    Dim version As Long, step As Object, record As Object, hidden As Long, position As Long, source As Object
    On Error GoTo Failed
    instructions = "": observations = "": provenance = ""
    If Not modTrainingReadContext.Read(context, target, policy, notice) Then Exit Function
    notice = "Incomplete evidence: the exact selected guide or its predecessor is missing, changed or corrupt."
    parts = Split(key, "|")
    If UBound(parts) <> 2 Then Exit Function
    If Not ParseVersion(CStr(parts(1)), version) Then Exit Function
    Set model = modGuideStore.ReadChain(target, CStr(parts(0)), version)
    If model Is Nothing Then Exit Function
    If CStr(model("ContentSha256")) <> CStr(parts(2)) Then Exit Function
    Set visible = modEvaluationMatches.PolicyControls(policy, False)
    instructions = "Authored instruction" & vbCrLf & CStr(model("Name")) & vbCrLf & _
        "Tags: " & modGuideModel.TagsText(model) & vbCrLf & CStr(model("Instructions")) & vbCrLf
    For Each step In model("Steps")
        position = position + 1
        instructions = instructions & vbCrLf & CStr(position) & ". "
        If modEvaluationMatches.Permitted(visible, CStr(step("ControlId"))) Then
            instructions = instructions & CStr(step("Caption")) & vbCrLf & CStr(step("Instruction")) & vbCrLf
        Else
            instructions = instructions & "Hidden by policy" & vbCrLf
        End If
    Next step
    For Each record In model("Observations")
        If Not modEvaluationMatches.Permitted(visible, CStr(record("ControlId"))) Then hidden = hidden + 1
    Next record
    observations = modGuideDraftSource.Evidence(model("Observations"), visible)
    provenance = "Guide: " & CStr(model("ActionPathId")) & "; version: " & CStr(version) & vbCrLf & _
        "SHA-256: " & CStr(model("ContentSha256")) & vbCrLf & _
        "Author: " & CStr(model("CreatedByUserId")) & "; saved: " & CStr(model("CreatedAtUTC"))
    Set source = model("SourceRun")
    If source.Count > 0 Then
        provenance = provenance & vbCrLf & "Source sequence: " & CStr(source("SequenceId")) & _
            "; journal version: " & CStr(source("JournalVersion"))
        hidden = hidden + CLng(source("RestrictedObservationCount"))
    End If
    notice = "Published guide. Authored instructions are not an observed run or a diagnostic conclusion."
    If hidden > 0 Then notice = "Hidden by policy. Incomplete evidence: " & CStr(hidden) & " restricted observation(s)."
    If CLng(model("CatalogVersion")) < modActivityCatalog.CATALOG_VERSION Then
        notice = notice & " Older release."
    ElseIf CStr(model("PackageSetVersion")) <> CStr(ThisWorkbook.CustomDocumentProperties("invSysPackageSetVersion").Value) Or _
           CStr(model("BuildIdentity")) <> CStr(ThisWorkbook.CustomDocumentProperties("invSysBuildIdentity").Value) Then
        notice = notice & " Different release/build; relative age unavailable."
    End If
    If context <> modActivity.CaptureContext() Then GoTo Failed
    ReadGuide = True
    Exit Function
Failed:
    instructions = "": observations = "": provenance = ""
    notice = "Unavailable: the selected guide, session or policy could not be validated."
End Function

Private Function ParseVersion(ByVal text As String, ByRef version As Long) As Boolean
    On Error GoTo Invalid
    If Not IsNumeric(text) Then Exit Function
    version = CLng(text)
    ParseVersion = (version > 0 And CStr(version) = text)
Invalid:
End Function

Private Function ListText(ByVal text As String) As String
    ListText = Replace(Replace(Replace(text, vbCr, " "), vbLf, " "), vbTab, " ")
End Function
