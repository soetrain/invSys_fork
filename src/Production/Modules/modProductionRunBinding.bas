Attribute VB_Name = "modProductionRunBinding"
Option Explicit

Public Function RequireCurrentContext(ByVal owner As frmProduction, ByVal context As String, ByVal operatorWb As Workbook) As Boolean
    RequireCurrentContext = modProductionDesignerActions.ContextIsCurrent(context, operatorWb)
    If Not RequireCurrentContext Then _
        owner.ShowStatus "Session, warehouse, or captured workbook changed. Reopen Production before editing the draft."
End Function

Public Function RequireWorksheetContext(ByVal owner As frmProduction, ByVal context As String, ByVal operatorWb As Workbook) As Boolean
    If Not RequireCurrentContext(owner, context, operatorWb) Then Exit Function
    RequireWorksheetContext = BindWorksheetOwner(operatorWb)
    If Not RequireWorksheetContext Then owner.ShowStatus "Production sheet not found."
End Function

Public Function BindWorksheetOwner(ByVal operatorWb As Workbook) As Boolean
    On Error GoTo Unavailable

    Dim openWb As Workbook
    Dim productionSheet As Worksheet

    If operatorWb Is Nothing Then Exit Function
    For Each openWb In Application.Workbooks
        If openWb Is operatorWb Then
            If openWb.IsAddin Then Exit Function
            Set productionSheet = openWb.Worksheets("Production")
            If productionSheet Is Nothing Then Exit Function
            mProduction.BindProductionOperatorWorkbook openWb
            BindWorksheetOwner = True
            Exit Function
        End If
    Next openWb
    Exit Function

Unavailable:
    BindWorksheetOwner = False
End Function
