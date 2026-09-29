Attribute VB_Name = "modProductionDesignerActions"
Option Explicit
Option Private Module

Public Enum ProductionInstructionAction
    InstructionAdd = 1
    InstructionUpdate
    InstructionRemove
    InstructionUp
    InstructionDown
End Enum

Public Function ContextIsCurrent(ByVal capturedContext As String, ByVal operatorWorkbook As Workbook) As Boolean
    Dim candidate As Workbook
    On Error GoTo Invalid
    If capturedContext = "" Or capturedContext <> modActivity.CaptureContext() Then Exit Function
    If operatorWorkbook Is Nothing Then Exit Function
    For Each candidate In Application.Workbooks
        If candidate Is operatorWorkbook Then ContextIsCurrent = True: Exit Function
    Next candidate
Invalid:
End Function

Public Function EditInstruction(ByVal action As ProductionInstructionAction, ByVal capturedContext As String, _
                                ByVal operatorWorkbook As Workbook, ByVal instructions As MSForms.ListBox, _
                                ByVal editor As MSForms.TextBox, ByVal loading As Boolean, ByRef busy As Boolean) As String
    Dim activityId As String, notice As String, report As String, outcome As String, permitted As Boolean, suffix As String
    If loading Or busy Then Exit Function
    On Error GoTo Failed
    Select Case action
        Case InstructionAdd: suffix = "ADD"
        Case InstructionUpdate: suffix = "UPDATE"
        Case InstructionRemove: suffix = "REMOVE"
        Case InstructionUp: suffix = "UP"
        Case InstructionDown: suffix = "DOWN"
        Case Else: Exit Function
    End Select
    report = "Session, warehouse, or captured workbook changed. Reopen Production before editing the draft."
    If Not ContextIsCurrent(capturedContext, operatorWorkbook) Then EditInstruction = report: Exit Function
    busy = True
    activityId = modActivity.BeginAction("PRODUCTION_PROCESS_INSTRUCTION_" & suffix, capturedContext, notice)
    If Not ContextIsCurrent(capturedContext, operatorWorkbook) Then GoTo Done
    permitted = modRoleUiAccess.CanCurrentUserPerformCapabilityCached("PROD_POST", report)
    If Not permitted Then permitted = modRoleUiAccess.CanCurrentUserPerformCapabilityCached("ADMIN_MAINT", report)
    If Not permitted Then
        outcome = "DENIED": report = "Production permission is required; the draft was not changed."
    Else
        outcome = EditLocalInstruction(action, instructions, editor)
        report = "Process instructions changed locally; saved definitions were not changed."
        If outcome = "REJECTED" Then report = "The instruction edit requires valid input or selection; the draft was not changed."
    End If
Done:
    If activityId <> "" Then modActivity.FinishAction activityId, outcome, notice
    busy = False
    If notice <> "" Then report = report & " " & notice
    EditInstruction = report
    Exit Function
Failed:
    outcome = "FAILED": report = "The instruction edit failed; verify the current draft before retrying."
    Resume Done
End Function

Private Function EditLocalInstruction(ByVal action As ProductionInstructionAction, ByVal instructions As MSForms.ListBox, _
                                      ByVal editor As MSForms.TextBox) As String
    Dim selected As Long, target As Long, column As Long, index As Long, value As Variant
    EditLocalInstruction = "REJECTED"
    selected = instructions.ListIndex
    Select Case action
        Case InstructionAdd
            If Trim$(editor.Text) = "" Then Exit Function
            instructions.AddItem CStr(instructions.ListCount + 1)
            instructions.List(instructions.ListCount - 1, 1) = Trim$(editor.Text)
        Case InstructionUpdate
            If selected < 0 Then Exit Function
            instructions.List(selected, 1) = Trim$(editor.Text)
        Case InstructionRemove
            If selected < 0 Then Exit Function
            instructions.RemoveItem selected
        Case InstructionUp, InstructionDown
            target = selected + IIf(action = InstructionUp, -1, 1)
            If selected < 0 Or target < 0 Or target >= instructions.ListCount Then Exit Function
            For column = 0 To instructions.ColumnCount - 1
                value = instructions.List(selected, column)
                instructions.List(selected, column) = instructions.List(target, column)
                instructions.List(target, column) = value
            Next column
            instructions.ListIndex = target
        Case Else: Exit Function
    End Select
    If action = InstructionRemove Or action = InstructionUp Or action = InstructionDown Then
        For index = 0 To instructions.ListCount - 1
            instructions.List(index, 0) = CStr(index + 1)
        Next index
    End If
    EditLocalInstruction = "STAGED"
End Function
