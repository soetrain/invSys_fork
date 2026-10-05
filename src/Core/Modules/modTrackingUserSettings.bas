Attribute VB_Name = "modTrackingUserSettings"
Option Explicit

' Primitive/serialized Admin boundary; no Auth credential fields are projected.
Public Function ReadUsers(ByVal context As String, ByVal request As String, _
                          ByRef projection As String, ByRef report As String) As Boolean
    Dim target As WarehouseTarget, model As Object, roster As Object, result As Object
    Dim rows As New Collection, row As Object, id As Variant, saved As Variant
    On Error GoTo Failed
    projection = ""
    If Not modTrackingPolicyModel.AuthorizeEditor(context, target, report) Then Exit Function
    Set model = modTrackingPolicyModel.Decode(request)
    If model Is Nothing Then Exit Function
    Set roster = modTrackingPolicyUsers.ReadRoster(target, report)
    If roster Is Nothing Then Exit Function
    For Each id In roster.Keys
        Set row = CreateObject("Scripting.Dictionary")
        row.Add "UserId", CStr(id): row.Add "Available", True
        rows.Add row
    Next id
    If model.Exists("Users") Then
        For Each saved In model("Users")
            If Not roster.Exists(CStr(saved("UserId"))) Then
                Set row = CreateObject("Scripting.Dictionary")
                row.Add "UserId", saved("UserId"): row.Add "Available", False
                rows.Add row
            End If
        Next saved
    End If
    If Not modTrackingPolicyModel.AuthorizeEditor(context, target, report) Then Exit Function
    Set result = CreateObject("Scripting.Dictionary"): result.Add "Users", rows
    projection = modTrainingJson.EncodeObject(result)
    report = ""
    ReadUsers = True
    Exit Function
Failed:
    projection = ""
    report = "Warehouse users are unavailable. Reload before saving."
End Function

Public Function UserCount(ByVal projection As String) As Long
    Dim model As Object
    On Error GoTo Invalid
    Set model = modTrainingJson.DecodeObject(projection)
    UserCount = model("Users").Count
Invalid:
End Function

Public Function UserField(ByVal projection As String, ByVal index As Long, ByVal field As String) As String
    Dim model As Object
    On Error GoTo Invalid
    If field <> "UserId" And field <> "Available" Then Exit Function
    Set model = modTrainingJson.DecodeObject(projection)
    UserField = CStr(model("Users")(index)(field))
Invalid:
End Function

Public Function RecordEnabled(ByVal request As String, ByVal userId As String) As Boolean
    Dim model As Object
    Set model = modTrackingPolicyModel.Decode(request)
    If model Is Nothing Then Exit Function
    RecordEnabled = True
    If model.Exists("Users") Then RecordEnabled = modTrackingPolicyUsers.Enabled(model("Users"), userId)
End Function

Public Function StageUser(ByVal request As String, ByVal userId As String, ByVal record As Boolean) As String
    Dim model As Object, rows As New Collection, row As Variant, added As Object
    On Error GoTo Invalid
    Set model = modTrackingPolicyModel.Decode(request)
    If model Is Nothing Then Exit Function
    If CLng(model("SchemaVersion")) <> 2 Then Exit Function
    For Each row In model("Users")
        If StrComp(row("UserId"), userId, vbTextCompare) <> 0 Then rows.Add row
    Next row
    If Not record Then
        Set added = CreateObject("Scripting.Dictionary")
        added.Add "UserId", userId: added.Add "Record", False
        rows.Add added
    End If
    Set model("Users") = rows
    If Not modTrackingPolicyUsers.ValidRows(rows) Then Exit Function
    StageUser = modTrainingJson.EncodeObject(model)
Invalid:
End Function
