Attribute VB_Name = "modProductionOperatorBinding"
Option Explicit
Option Private Module

' The run form owns this session. Ordinary Production never opens or starts a runner.
Private mContext As String
Private mToken As String
Private mStep As String
Private mTicket As String
Private mDigest As String
Private mOwner As frmProduction
Private mWorkbook As Workbook
Private mRunner As frmActionPathRun
Private mClaimedActivity As String

Public Sub WaitForSelection(ByVal context As String, ByVal token As String, ByVal stepId As String, ByVal runner As frmActionPathRun)
    Clear mToken
    mContext = context: mToken = token: mStep = stepId
    Set mRunner = runner
End Sub

Public Sub Clear(ByVal token As String)
    If token <> mToken Then Exit Sub
    Set mOwner = Nothing: Set mWorkbook = Nothing: Set mRunner = Nothing: mTicket = "": mDigest = "": mStep = "": mToken = "": mContext = "": mClaimedActivity = ""
End Sub

Public Function BindSelected(ByVal context As String, ByVal token As String, ByRef notice As String) As Boolean
    Dim form As Object, owner As frmProduction, kind As String, digest As String, caption As String, ticket As String
    On Error GoTo Failed
    If mToken = "" Or token <> mToken Or context <> mContext Then Exit Function
    ticket = modExecutionRun.SelectionTicket(context, token, mStep, notice)
    If ticket = "" Then Exit Function
    notice = "Open the prepared Production form and select the exact checked-in Process or worksheet output."
    For Each form In VBA.UserForms
        If TypeOf form Is frmProduction Then
            If Not owner Is Nothing Then Exit Function
            Set owner = form
        End If
    Next form
    If owner Is Nothing Then Exit Function
    digest = owner.CompletionSelection(context, ticket, True, kind, caption)
    If digest = "" Then Exit Function
    If Not modExecutionRun.BindSelection(context, token, mStep, ticket, kind, digest, notice) Then Exit Function
    Set mOwner = owner: Set mWorkbook = owner.CompletionWorkbook: mTicket = ticket: mDigest = digest
    notice = caption & ". " & notice
    BindSelected = True
    Exit Function
Failed:
    notice = "This prepared selection could not be bound. Review it before continuing."
End Function

Public Function SelectionStillBound(ByVal token As String) As Boolean
    Dim kind As String, caption As String
    On Error GoTo Done
    If token <> mToken Or mOwner Is Nothing Or mDigest = "" Then Exit Function
    If Not modOperationsFormLifetime.IsLoaded(mOwner) Then Exit Function
    If Not mOwner.CompletionWorkbook Is mWorkbook Then Exit Function
    SelectionStillBound = (mOwner.CompletionSelection(mContext, mTicket, False, kind, caption) = mDigest)
Done:
End Function

Public Sub ObserveOwner(ByVal owner As frmProduction, ByVal activity As String, Optional ByVal atWrite As Boolean = False, _
                        Optional ByVal resolvedRow As Long = 0, Optional ByVal actual As String = "")
    Dim kind As String, caption As String, digest As String
    On Error GoTo Done
    If mOwner Is Nothing Or mTicket = "" Or activity = "" Then Exit Sub
    If Not owner Is mOwner Then Exit Sub
    If Not owner.CompletionWorkbook Is mWorkbook Then Exit Sub
    digest = owner.CompletionSelection(mContext, mTicket, False, kind, caption, resolvedRow, actual)
    If digest = "" Then Exit Sub
    If modExecutionRun.ClaimCompletion(mContext, mToken, mStep, mTicket, digest, activity, atWrite) Then mClaimedActivity = activity
Done:
End Sub

Public Sub OwnerExited(ByVal activity As String)
    On Error GoTo Done
    If activity = "" Or activity <> mClaimedActivity Then Exit Sub
    modExecutionRun.OperatorReturned mToken, activity
Done:
End Sub

Public Sub Returned(ByVal activity As String)
    On Error GoTo Done
    If activity = "" Or activity <> mClaimedActivity Then Exit Sub
    mClaimedActivity = ""
    If Not mRunner Is Nothing Then mRunner.OperatorOwnerReturned
Done:
End Sub

Public Function Framed(ByVal value As String) As String
    Framed = CStr(Len(value)) & ":" & value
End Function

' The hidden workbook name identifies the owned staging column. A pre-existing
' user column with that caption is never adopted. The marker is inert outside this run.
Public Function MarkedRow(ByVal book As Workbook, ByVal output As ListObject, ByVal row As Long, ByVal ticket As String, ByVal stamp As Boolean) As Boolean
    Const header As String = "invSys_TrainingDraft"
    Const markName As String = "_invSysTrainingOutputColumn"
    Dim col As ListColumn, marker As Name, anchor As Range, adjacent As Range
    On Error Resume Next
    Set col = output.ListColumns(header): Set marker = book.Names(markName)
    On Error GoTo Failed
    If col Is Nothing Then
        If Not stamp Or Not marker Is Nothing Then Exit Function
        Set adjacent = output.Range.Columns(output.Range.Columns.Count).Offset(0, 1)
        If Application.WorksheetFunction.CountA(adjacent) > 0 Then Exit Function
        Set col = output.ListColumns.Add: col.Name = header
        Set anchor = col.Range.Cells(1, 1)
        Set marker = book.Names.Add(Name:=markName, RefersTo:="=" & anchor.Address(External:=True), Visible:=False)
    End If
    If marker Is Nothing Then Exit Function
    Set anchor = marker.RefersToRange
    If Not anchor.Parent Is output.Parent Then Exit Function
    If anchor.Address <> col.Range.Cells(1, 1).Address Then Exit Function
    If stamp Then col.DataBodyRange.Cells(row, 1).Value2 = ticket
    If CStr(col.DataBodyRange.Cells(row, 1).Value2) <> ticket Then Exit Function
    MarkedRow = (Application.WorksheetFunction.CountIf(col.DataBodyRange, ticket) = 1)
Failed:
End Function
