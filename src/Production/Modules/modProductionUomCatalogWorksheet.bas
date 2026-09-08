Attribute VB_Name = "modProductionUomCatalog"
Option Explicit

Private Const SHEET_NAME As String = "invSys UOM Catalog"
Private Const TABLE_NAME As String = "tblInvSysUomCatalog"

Public Function SendUomCatalogToWorksheet(ByVal wb As Workbook, _
                                          Optional ByRef report As String = "") As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim rows As Variant
    Dim headers As Variant
    Dim rowCount As Long

    If wb Is Nothing Then
        report = "Production has no captured workbook for the UOM Catalog."
        Exit Function
    End If
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SHEET_NAME
    End If
    On Error Resume Next
    Set lo = ws.ListObjects(TABLE_NAME)
    On Error GoTo 0
    rows = modUomSettings.GetUomCatalogRows()
    headers = Array("UOM", "Dimension", "Base UOM", "Units Per Base UOM", "Convertible", "Enabled", "Notes")
    If Not lo Is Nothing Then lo.Unlist
    ws.Cells.Clear
    ws.Range("A1").Value2 = "invSys UOM Catalog"
    ws.Range("A2").Value2 = "Add same-dimension units here. CS and EA must remain nonconvertible. Select this table, then Retrieve UOM Catalog."
    ws.Range("A4").Resize(1, 7).Value = headers
    If IsArray(rows) Then
        rowCount = UBound(rows, 1)
        ws.Range("A5").Resize(rowCount, 7).Value = rows
    Else
        rowCount = 1
    End If
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A4").Resize(rowCount + 1, 7), , xlYes)
    lo.Name = TABLE_NAME
    lo.TableStyle = "TableStyleMedium2"
    ws.Columns("A:G").AutoFit
    SendUomCatalogToWorksheet = True
    report = "UOM Catalog table sent to the captured workbook. Edit the table, select it, then Retrieve UOM Catalog."
End Function

Public Function RetrieveUomCatalogFromWorksheet(ByVal wb As Workbook, _
                                                 Optional ByRef report As String = "", _
                                                 Optional ByRef outcome As String = "") As Boolean
    Dim lo As ListObject
    Dim values As Variant
    outcome = "REJECTED"

    If wb Is Nothing Then
        report = "Production has no captured workbook for UOM Catalog retrieval."
        Exit Function
    End If
    On Error Resume Next
    Set lo = Application.ActiveCell.ListObject
    On Error GoTo 0
    If lo Is Nothing Or Not lo.Parent.Parent Is wb Or StrComp(lo.Name, TABLE_NAME, vbTextCompare) <> 0 Then
        report = "Select a cell in the invSys UOM Catalog table in the captured Production workbook."
        Exit Function
    End If
    If lo.DataBodyRange Is Nothing Then
        report = "The selected UOM Catalog table has no rows."
        Exit Function
    End If
    values = lo.DataBodyRange.Value2
    If Not modUomSettings.PublishUomCatalogRows(values, report, outcome) Then Exit Function
    lo.Unlist
    report = report & " The staging table was retrieved and removed."
    RetrieveUomCatalogFromWorksheet = True
End Function
