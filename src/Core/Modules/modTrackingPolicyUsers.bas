Attribute VB_Name = "modTrackingPolicyUsers"
Option Explicit
Option Private Module

Public Function ValidRows(ByVal rows As Collection) As Boolean
    Dim seen As Object, row As Variant, id As String
    On Error GoTo Invalid
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
    For Each row In rows
        If Not IsObject(row) Then Exit Function
        If TypeName(row) <> "Dictionary" Then Exit Function
        If row.Count <> 2 Or Not row.Exists("UserId") Or Not row.Exists("Record") Then Exit Function
        If VarType(row("UserId")) <> vbString Or VarType(row("Record")) <> vbBoolean Then Exit Function
        id = row("UserId")
        If id = "" Or id <> Trim$(id) Or seen.Exists(id) Then Exit Function
        seen.Add id, True
    Next row
    ValidRows = True
Invalid:
End Function

Public Function Enabled(ByVal rows As Collection, ByVal userId As String) As Boolean
    Dim row As Variant
    Enabled = True
    For Each row In rows
        If StrComp(row("UserId"), userId, vbTextCompare) = 0 Then Enabled = CBool(row("Record")): Exit Function
    Next row
End Function

Public Function FindTable(ByVal wb As Workbook, ByVal name As String) As ListObject
    Dim sheet As Worksheet, table As ListObject
    For Each sheet In wb.Worksheets
        For Each table In sheet.ListObjects
            If StrComp(table.Name, name, vbTextCompare) = 0 Then Set FindTable = table: Exit Function
        Next table
    Next sheet
End Function

Public Function ReadStored(ByVal wb As Workbook, ByVal headers As ListObject, ByVal selected As Long, _
                           ByVal version As Long, ByVal versions As Object, ByRef users As Collection) As Boolean
    Dim table As ListObject, row As Object, index As Long, number As Variant, schema As Long, count As Variant
    On Error GoTo Invalid
    Set users = New Collection
    schema = CLng(modTrackingPolicyModel.TableValue(headers, selected, "SchemaVersion"))
    Set table = FindTable(wb, "tblEventTrackingUsers")
    If schema = 2 Then
        If table Is Nothing Then Exit Function
        count = modTrackingPolicyModel.TableValue(headers, selected, "UserCount")
        If Not WholeNumber(count, True) Then Exit Function
    ElseIf schema <> 1 Then
        Exit Function
    End If
    If Not table Is Nothing Then
        ' Validate headers even when the selected version has zero overrides.
        index = modTrackingPolicyModel.TableColumn(table, "PolicyVersion")
        index = modTrackingPolicyModel.TableColumn(table, "UserId")
        index = modTrackingPolicyModel.TableColumn(table, "Record")
        For index = 1 To table.ListRows.Count
            number = modTrackingPolicyModel.TableValue(table, index, "PolicyVersion")
            If Not WholeNumber(number, False) Then Exit Function
            If Not versions.Exists(CStr(CLng(number))) Then Exit Function
            If CLng(modTrackingPolicyModel.TableValue(headers, CLng(versions(CStr(CLng(number)))), "SchemaVersion")) <> 2 Then Exit Function
            If CLng(number) = version Then
                Set row = CreateObject("Scripting.Dictionary")
                row.Add "UserId", modTrackingPolicyModel.TableValue(table, index, "UserId")
                row.Add "Record", modTrackingPolicyModel.TableValue(table, index, "Record")
                users.Add row
            End If
        Next index
    End If
    If schema = 2 Then
        If users.Count <> CLng(count) Then Exit Function
    End If
    ReadStored = ValidRows(users)
Invalid:
End Function

Private Function WholeNumber(ByVal value As Variant, ByVal zeroAllowed As Boolean) As Boolean
    Select Case VarType(value)
        Case vbInteger, vbLong, vbDouble
            If value < IIf(zeroAllowed, 0, 1) Or value > 2147483647# Then Exit Function
            WholeNumber = (value = Fix(value))
    End Select
End Function

Public Function ReadRoster(ByVal target As WarehouseTarget, ByRef report As String) As Object
    Dim wb As Workbook, candidate As Workbook, table As ListObject, rows As Object
    Dim opened As Boolean, index As Long, value As Variant, id As String
    On Error GoTo Failed
    report = "Warehouse users are unavailable. No configuration was changed."
    If target.AuthPath = "" Or Len(Dir$(target.AuthPath)) = 0 Then Exit Function
    For Each candidate In Application.Workbooks
        If StrComp(candidate.FullName, target.AuthPath, vbTextCompare) = 0 Then Set wb = candidate: Exit For
    Next candidate
    If wb Is Nothing Then
        Set wb = Application.Workbooks.Open(target.AuthPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
        opened = True
    ElseIf Not wb.Saved Then
        GoTo CleanExit
    End If
    If Not modRuntimeWorkbooks.RuntimeWorkbookSchemaPresentForRead(wb, "AUTH") Then GoTo CleanExit
    Set table = FindTable(wb, "tblUsers")
    If table Is Nothing Then GoTo CleanExit
    Set rows = CreateObject("Scripting.Dictionary"): rows.CompareMode = vbTextCompare
    For index = 1 To table.ListRows.Count
        value = modTrackingPolicyModel.TableValue(table, index, "UserId")
        If VarType(value) <> vbString Then GoTo CleanExit
        id = Trim$(CStr(value))
        If id = "" Or rows.Exists(id) Then GoTo CleanExit
        rows.Add id, id
    Next index
    If Not rows.Exists(modAuth.GetCurrentUserId()) Then GoTo CleanExit
    Set ReadRoster = rows
    report = ""
CleanExit:
    On Error Resume Next
    If opened And Not wb Is Nothing Then wb.Close SaveChanges:=False
    Exit Function
Failed:
    Set ReadRoster = Nothing
    Resume CleanExit
End Function

Public Function ValidateSave(ByVal target As WarehouseTarget, ByVal model As Object, ByVal prior As Object, _
                             ByVal savedSchema As Long, ByRef report As String) As Boolean
    Dim roster As Object, row As Variant, old As Variant, retained As Boolean
    report = "Invalid user recording policy. Reload before saving."
    If CLng(model("SchemaVersion")) = 1 Then
        ValidateSave = (savedSchema < 2)
        Exit Function
    End If
    Set roster = ReadRoster(target, report)
    If roster Is Nothing Then Exit Function
    For Each row In model("Users")
        If Not roster.Exists(CStr(row("UserId"))) Then
            retained = False
            For Each old In prior("Users")
                If StrComp(old("UserId"), row("UserId"), vbTextCompare) = 0 Then
                    retained = (CBool(old("Record")) = CBool(row("Record")))
                    Exit For
                End If
            Next old
            If Not retained Then report = "Selected user is unavailable. Reload before saving.": Exit Function
        End If
    Next row
    ValidateSave = True
End Function
