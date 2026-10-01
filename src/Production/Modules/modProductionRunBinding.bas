Attribute VB_Name = "modProductionRunBinding"
Option Explicit

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
