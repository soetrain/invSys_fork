Attribute VB_Name = "modExecutionBinding"
Option Explicit
Option Private Module

Public Function Read(ByVal context As String, ByVal key As String, ByRef target As WarehouseTarget, _
                     ByRef guide As Object, ByRef profile As Object, ByRef policyHash As String, _
                     ByRef builds As Collection, ByRef snapshot As String, ByRef notice As String) As Boolean
    Dim policy As Object, parts As Variant, version As Long, capture As Object, visible As Object, step As Object, definition As Object
    On Error GoTo Failed
    If Not modExecutionTarget.Read(context, target, snapshot, notice) Then Exit Function
    If Not modTrainingReadContext.Read(context, target, policy, notice) Then Exit Function
    notice = "Execution is unavailable: the selected guide is invalid, hidden or incompatible."
    parts = Split(key, "|")
    If UBound(parts) <> 2 Then Exit Function
    version = CLng(parts(1))
    If version < 1 Or CStr(version) <> parts(1) Then Exit Function
    Set guide = modGuideStore.ReadChain(target, CStr(parts(0)), version)
    If guide Is Nothing Then Exit Function
    If guide("ContentSha256") <> parts(2) Then Exit Function
    Set capture = modEvaluationMatches.PolicyControls(policy, True): Set visible = modEvaluationMatches.PolicyControls(policy, False)
    For Each step In guide("Steps")
        If Not modEvaluationMatches.Permitted(capture, CStr(step("ControlId"))) Or Not modEvaluationMatches.Permitted(visible, CStr(step("ControlId"))) Then Exit Function
        Set definition = modActivityCatalog.Control(CStr(step("ControlId")))
        If definition Is Nothing Then Exit Function
        If Not modAuth.CanPerform(CStr(definition("Capability")), modAuth.GetCurrentUserId(), target.WarehouseId, target.StationId) Then
            notice = "Run is unavailable: a required workflow permission is missing.": Exit Function
        End If
    Next step
    For Each step In guide("Observations")
        If Not modEvaluationMatches.Permitted(visible, CStr(step("ControlId"))) Then Exit Function
    Next step
    If Not modExecutionProfileStore.Latest(target, guide, profile, notice) Then Exit Function
    notice = "Execution not configured for this exact guide version."
    If profile Is Nothing Then Exit Function
    If profile("CatalogVersion") <> modActivityCatalog.CATALOG_VERSION Then Exit Function
    If profile("PackageSetVersion") <> CStr(ThisWorkbook.CustomDocumentProperties("invSysPackageSetVersion").Value) Then Exit Function
    policy("SavedCatalogVersion") = CLng(policy("SavedCatalogVersion"))
    policyHash = modTrainingWire.Sha256(modTrainingJson.EncodeObject(policy))
    Set builds = PackageBuilds()
    If builds Is Nothing Or context <> modActivity.CaptureContext() Then Exit Function
    notice = "": Read = True
Failed:
End Function

Public Function PackageBuilds() As Collection
    Dim result As New Collection, name As Variant, wb As Workbook, entry As Object
    On Error GoTo Failed
    For Each name In Array("invSys.Core.xlam", "invSys.Inventory.Domain.xlam", "invSys.Designs.Domain.xlam", "invSys.Operations.xlam", "invSys.Admin.xlam")
        Set wb = Nothing
        ' Loaded XLAMs are addressable by name but omitted from workbook enumeration.
        On Error Resume Next
        Set wb = Application.Workbooks(CStr(name))
        Err.Clear
        On Error GoTo Failed
        If wb Is Nothing Then
            If name = "invSys.Core.xlam" Or name = "invSys.Inventory.Domain.xlam" Or name = "invSys.Operations.xlam" Then Exit Function
            GoTo NextPackage
        End If
        If CStr(wb.CustomDocumentProperties("invSysPackageSetVersion").Value) <> CStr(ThisWorkbook.CustomDocumentProperties("invSysPackageSetVersion").Value) Then Exit Function
        Set entry = CreateObject("Scripting.Dictionary")
        entry.Add "PackageId", CStr(name): entry.Add "BuildIdentity", CStr(wb.CustomDocumentProperties("invSysBuildIdentity").Value)
        If entry("BuildIdentity") = "" Then Exit Function
        result.Add entry
NextPackage:
    Next name
    Set PackageBuilds = result
Failed:
End Function

Public Function BuildsHash(ByVal builds As Collection) As String
    Dim wrapper As Object
    Set wrapper = CreateObject("Scripting.Dictionary"): wrapper.Add "Packages", builds
    BuildsHash = modTrainingWire.Sha256(modTrainingJson.EncodeObject(wrapper))
End Function
