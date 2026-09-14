Attribute VB_Name = "modPathExpectation"
Option Explicit
Option Private Module

' Authored analysis intent for one validated library selection; never a guide save.
Private mContext As String
Private mPathId As String
Private mSequence As String
Private mBinding As String
Private mDefinition As Object
Private mSource As String
Private mHeader As Object

Public Sub BindRun(ByVal context As String, ByVal header As Object)
    Dim binding As String
    binding = CStr(header("RecordId")) & ":" & CStr(header("Version")) & ":" & CStr(header("ContentSha256"))
    If mContext = context And mPathId = CStr(header("ActionPathId")) And mBinding = binding Then Exit Sub
    ClearContext mContext
    mContext = context: mPathId = CStr(header("ActionPathId"))
    mSequence = CStr(header("SequenceId")): mBinding = binding
    Set mHeader = modTrainingJson.DecodeObject(modTrainingJson.EncodeObject(header))
    Set mDefinition = modExpectationModel.NoneDefinition()
    mSource = "No expected conclusion selected"
    If header.Exists("ExpectedConclusion") Then
        Set mDefinition = modTrainingJson.DecodeObject(modTrainingJson.EncodeObject(header("ExpectedConclusion")))
        If mDefinition("TerminalKind") <> "None" Then mSource = "Captured expectation"
    End If
End Sub

Public Function ReadDefinition(ByVal context As String, ByVal pathId As String, ByRef sequenceId As String, _
                               ByRef binding As String, ByRef definition As Object, ByRef notice As String) As Boolean
    Dim target As WarehouseTarget, policy As Object, version As Long, collect As Boolean, visible As Boolean
    On Error GoTo Unavailable
    sequenceId = "": binding = "": Set definition = Nothing
    notice = "Select a permitted saved recording before editing its expected conclusion."
    If context = "" Or context <> modActivity.CaptureContext() Then GoTo Unavailable
    If context <> mContext Or pathId = "" Or pathId <> mPathId Or mDefinition Is Nothing Then Exit Function
    Set target = modNasConnection.GetCurrentTarget(): Set policy = CreateObject("Scripting.Dictionary")
    If Not modActivityPolicy.ReadPolicy(target, "ADMIN_SETTINGS_SAVE_VALUE", version, collect, visible, notice, policy) Then GoTo Unavailable
    If context <> modActivity.CaptureContext() Then GoTo Unavailable
    sequenceId = mSequence: binding = mBinding
    Set definition = modTrainingJson.DecodeObject(modTrainingJson.EncodeObject(mDefinition))
    notice = mSource & ": " & CStr(definition("Steps").Count) & " expected step(s)."
    ReadDefinition = True
    Exit Function
Unavailable:
    ClearContext context
    notice = "Unavailable: the selected recording, invSys session or policy changed. Reopen the library."
End Function

Public Function Stage(ByVal context As String, ByVal pathId As String, ByVal binding As String, _
                      ByVal definition As Object, ByRef notice As String) As Boolean
    Dim sequenceId As String, currentBinding As String, current As Object
    If Not ReadDefinition(context, pathId, sequenceId, currentBinding, current, notice) Then Exit Function
    notice = "This selected recording changed. Reopen Expected conclusion."
    If binding = "" Or binding <> currentBinding Then Exit Function
    notice = "Choose valid expected steps and a conclusion, or remove all steps and choose None."
    If Not modExpectationModel.Validate(definition, modActivityCatalog.CATALOG_VERSION) Then Exit Function
    Set mDefinition = modTrainingJson.DecodeObject(modTrainingJson.EncodeObject(definition))
    Stage = Not mDefinition Is Nothing
    If Stage Then mSource = "This evaluation": notice = "Expected conclusion staged for this evaluation. The recording is unchanged."
End Function

Public Function Summary(ByVal context As String, ByVal pathId As String) As String
    Dim sequenceId As String, binding As String, definition As Object, notice As String, available As Boolean
    available = ReadDefinition(context, pathId, sequenceId, binding, definition, notice)
    Summary = notice
End Function

Public Function ReadSelected(ByVal context As String, ByVal pathId As String, ByRef header As Object, _
                             ByRef definition As Object, ByRef source As String, ByRef notice As String) As Boolean
    Dim sequenceId As String, binding As String
    Set header = Nothing
    If Not ReadDefinition(context, pathId, sequenceId, binding, definition, notice) Then Exit Function
    If mHeader Is Nothing Then Exit Function
    Set header = modTrainingJson.DecodeObject(modTrainingJson.EncodeObject(mHeader))
    source = mSource
    If source = "No expected conclusion selected" Then source = "No expectation"
    ReadSelected = Not header Is Nothing
End Function

Public Sub ClearContext(ByVal context As String)
    If context <> mContext Then Exit Sub
    modExpectationDraft.ClosePath mContext, mPathId
    mContext = "": mPathId = "": mSequence = "": mBinding = "": mSource = ""
    Set mDefinition = Nothing
    Set mHeader = Nothing
End Sub
