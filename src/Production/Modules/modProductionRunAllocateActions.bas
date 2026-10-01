Attribute VB_Name = "modProductionRunAllocateActions"
Option Explicit
Option Private Module

' Observe the existing local allocation owners; no inventory submission occurs.
Public Function Execute(ByVal owner As frmProduction, ByVal context As String, _
                        ByVal operatorBook As Workbook, ByRef loading As Boolean, ByRef busy As Boolean, _
                        ByVal isTree As Boolean, ByVal palette As MSForms.ListBox) As String
    Dim action As cProductionWorksheetAction, controlId As String, report As String
    Dim priorLoading As Boolean, number As Long, source As String, description As String
    If loading Or busy Then Exit Function
    On Error GoTo Failed
    priorLoading = loading: busy = True
    If isTree Then controlId = "PRODUCTION_RUN_TREE_ALLOCATE" Else controlId = "PRODUCTION_RUN_ALLOCATE"
    Set action = New cProductionWorksheetAction
    If Not action.Begin(controlId, context, operatorBook, report) Then GoTo Done
    If Not action.CanContinue(report) Then GoTo Done
    Set owner.RunActionContinuation = action
    If Not isTree Then
        If palette.ListCount = 0 Then
            owner.ResetInventoryCache
            owner.RefreshRunPaletteState
            If Not action.CanContinue(report) Then GoTo Done
            If palette.ListCount = 0 Then
                action.OutcomeCode = "REJECTED"
                report = "No acceptable inventory assignments were found for this recipe. Use Ingredients Assignment to select each USED ingredient, add acceptable inventory, and Save Assignment."
                GoTo Done
            End If
        End If
        If palette.ListIndex < 0 And palette.ListCount = 1 Then
            palette.ListIndex = 0
            owner.LoadSelectedRunPaletteRow
            If Not action.CanContinue(report) Then GoTo Done
        End If
    End If
    owner.ApplySelectedRunPaletteSplit action
    If Not action.CanContinue(report) Then GoTo Done
    report = owner.RunAllocationStatus()
Done:
    Set owner.RunActionContinuation = Nothing
    loading = priorLoading
    If Not action Is Nothing Then action.Finish report
    busy = False
    Execute = report
    If number <> 0 Then
        On Error GoTo 0
        Err.Raise number, source, description
    End If
    Exit Function
Failed:
    number = Err.Number: source = Err.Source: description = Err.Description
    If Not action Is Nothing Then action.OutcomeCode = "FAILED"
    Resume Done
End Function

Public Sub ApplyReusable(ByVal owner As frmProduction, ByVal action As cProductionWorksheetAction, _
                         ByVal palette As MSForms.ListBox, ByVal splitBox As MSForms.TextBox, _
                         ByVal qtyBox As MSForms.TextBox, ByVal runLocation As String)
    Dim idx As Long, splitText As String, qtyText As String, splitVal As Double, qtyVal As Double
    Dim requiredQty As Double, report As String, applied As Boolean
    If palette.ListIndex < 0 Then
        owner.ShowStatus "Select an acceptable inventory stock row first."
        Exit Sub
    End If
    idx = palette.ListIndex
    splitText = Trim$(splitBox.Text): qtyText = Trim$(qtyBox.Text)
    requiredQty = modProductionReusableRun.ReusableRunRequirementQty( _
        owner.NzStr(palette.List(idx, 0)), owner.NzStr(palette.List(idx, 1)))
    If qtyText <> "" Then
        If Not owner.TryParseNonNegativeRunNumber(qtyText, qtyVal, "Quantity") Then Exit Sub
        If requiredQty > 0 Then splitVal = qtyVal / requiredQty * 100#
    ElseIf splitText <> "" Then
        If Not owner.TryParseNonNegativeRunNumber(splitText, splitVal, "% of Requirement") Then Exit Sub
        qtyVal = requiredQty * splitVal / 100#
    Else
        owner.ShowStatus "Enter % of Requirement or Qty first."
        Exit Sub
    End If
    If runLocation = "" Then
        owner.ShowStatus "Choose a production run location before allocating inventory."
        Exit Sub
    End If
    If StrComp(runLocation, owner.NzStr(palette.List(idx, 9)), vbTextCompare) <> 0 Then
        owner.ShowStatus "Allocation rejected. Inventory is at " & owner.NzStr(palette.List(idx, 9)) & _
                         "; production run location is " & runLocation & "."
        Exit Sub
    End If
    If Not action.CanContinue(report) Then owner.ShowStatus report: Exit Sub
    applied = modProductionReusableRun.ApplyReusableRunStockAllocation( _
        CStr(palette.List(idx, 0)), CStr(palette.List(idx, 1)), CStr(palette.List(idx, 3)), qtyVal, report, action)
    If Not action.CanContinue(report) Then owner.ShowStatus report: Exit Sub
    If applied Then
        splitBox.Text = owner.FormatRunNumber(splitVal)
        qtyBox.Text = owner.FormatRunNumber(qtyVal)
        owner.RefreshReusableRunControls False, action
        If Not action.CanContinue(report) Then owner.ShowStatus report: Exit Sub
        action.OutcomeCode = "STAGED"
    End If
    owner.ShowStatus report
End Sub
