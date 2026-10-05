Attribute VB_Name = "modReceivingExecution"
Option Explicit

' Typed Operations adapter. Inputs come from the validated Core execution profile.
Public Function Choices(ByVal snapshot As String) As Variant
    Dim wb As Workbook, candidate As Workbook, sheet As Worksheet, table As ListObject, opened As Boolean, window As Window
    On Error GoTo Done
    For Each candidate In Application.Workbooks
        If StrComp(candidate.FullName, snapshot, vbTextCompare) = 0 Then Set wb = candidate: Exit For
    Next candidate
    If wb Is Nothing Then
        Set wb = Application.Workbooks.Open(snapshot, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
        opened = True
        For Each window In wb.Windows: window.Visible = False: Next window
    ElseIf Not wb.Saved Then
        Exit Function
    End If
    For Each sheet In wb.Worksheets
        For Each table In sheet.ListObjects
            If table.Name = "tblInventorySnapshot" Then
                Choices = modReceivingInventoryChoices.FromTable(table, "", "SKU")
                GoTo Done
            End If
        Next table
    Next sheet
Done:
    If opened And Not wb Is Nothing Then wb.Close SaveChanges:=False
End Function

Public Function SelectEntity(ByVal list As MSForms.ListBox, ByVal keys As Collection, ByVal inputs As Collection, ByVal key As String) As Boolean
    Dim index As Long, found As Long, matches As Long, inputState As cReceivingSelectionInput, consumed As Boolean
    If key = "" Or keys Is Nothing Or inputs Is Nothing Then Exit Function
    For index = 1 To keys.Count
        If CStr(keys(index)) = key Then found = index - 1: matches = matches + 1
    Next index
    If matches <> 1 Then Exit Function
    Set inputState = inputs("lstReceiveItems")
    list.ListIndex = -1
    inputState.ArmExecutionInput
    list.ListIndex = found
    consumed = inputState.Consume()
    SelectEntity = (list.ListIndex = found And Not consumed)
End Function

Public Function Dispatch(ByVal context As String, ByVal controlId As String, ByVal inputText As String, ByRef captured As Workbook, ByRef workbookName As String, ByRef notice As String) As Boolean
    Dim form As frmReceiving, values As Variant
    On Error GoTo Failed
    If context = "" Or context <> modActivity.CaptureContext() Then Exit Function
    If controlId = "RECEIVING_OPEN" Then
        modTS_Received.ShowReceivingForm True, workbookName, context
        Set form = modTS_Received.ExecutionForm(context, workbookName)
        If form Is Nothing Then Exit Function
        Set captured = modOperationsInit.ResolveOpenWorkbookByName(workbookName)
        Dispatch = form.CanReuseFor(captured)
        Exit Function
    End If
    Set form = modTS_Received.ExecutionForm(context, workbookName)
    If form Is Nothing Then notice = "The captured Receiving workbook or form is unavailable.": Exit Function
    If captured Is Nothing Then Exit Function
    If Not form.CanReuseFor(captured) Then notice = "The captured Receiving workbook was replaced. Open a new setup.": Exit Function
    Select Case controlId
        Case "RECEIVING_REFRESH": form.Controls("btnRefresh").Value = True
        Case "RECEIVING_CLEAR": form.Controls("btnClear").Value = True
        Case "RECEIVING_SELECT_ITEM"
            Dispatch = form.SelectExecutionEntity(inputText): Exit Function
        Case "RECEIVING_ADD_SELECTED"
            values = Split(inputText, vbTab)
            If UBound(values) <> 4 Then Exit Function
            form.Controls("txtRef").Value = values(0): form.Controls("txtQty").Value = values(1)
            form.Controls("txtReceiveLocation").Value = values(2): form.Controls("txtLotNumber").Value = values(3)
            form.Controls("cboCondition").Value = values(4): form.Controls("btnAdd").Value = True
        Case "RECEIVING_CONFIRM_WRITES": form.Controls("btnConfirm").Value = True
        Case Else: notice = "This action has no registered Receiving adapter.": Exit Function
    End Select
    Dispatch = True: Exit Function
Failed:
    notice = "Receiving action did not return normally. Inspect its retained outcome before opening a new attempt."
End Function
