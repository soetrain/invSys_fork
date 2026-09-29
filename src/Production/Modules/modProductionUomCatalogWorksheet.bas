Attribute VB_Name = "modProductionUomCatalog"
Option Explicit

Private Const SHEET_NAME As String = "invSys UOM Catalog"
Private Const TABLE_NAME As String = "tblInvSysUomCatalog"
Private Const EXTENT_NAME As String = "_invSysUomDraftExtent"
Private Const EXTENT_OWNER As String = "invSys.UomDraftExtent.v1"

Public Function SendUomCatalogToWorksheet(ByVal wb As Workbook, _
                                          Optional ByRef report As String = "", _
                                          Optional ByRef outcome As String = "") As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim rows As Variant
    Dim headers As Variant
    Dim rowCount As Long
    Dim stage As Range, indexes() As Long, otherSheet As Worksheet, otherTable As ListObject
    Dim marker As Name

    On Error GoTo Failed
    outcome = "REJECTED"
    If wb Is Nothing Then
        report = "Production has no captured workbook for the UOM Catalog."
        Exit Function
    End If
    For Each otherSheet In wb.Worksheets
        If StrComp(otherSheet.Name, SHEET_NAME, vbTextCompare) = 0 Then Set ws = otherSheet
        For Each otherTable In otherSheet.ListObjects
            If StrComp(otherTable.Name, TABLE_NAME, vbTextCompare) = 0 Then
                If StrComp(otherSheet.Name, SHEET_NAME, vbTextCompare) <> 0 Then
                    report = "The UOM staging table belongs to another worksheet. No cells were changed."
                    Exit Function
                End If
                Set lo = otherTable
            End If
        Next otherTable
    Next otherSheet
    If Not ws Is Nothing Then
        If Not ReadStagingExtent(ws, marker, stage) Then GoTo InvalidStage
    End If
    If Not lo Is Nothing Then
        If Not StagingColumns(lo.HeaderRowRange, indexes) Then GoTo InvalidStage
        GoTo Reused
    End If
    If Not ws Is Nothing Then
        If Application.WorksheetFunction.CountA(ws.UsedRange) > 0 Then
            If ws.ListObjects.Count <> 0 Then GoTo InvalidStage
            If stage Is Nothing Then Set stage = FindStagingRegion(ws)
            If stage Is Nothing Then GoTo InvalidStage
            Set lo = ws.ListObjects.Add(xlSrcRange, stage, , xlYes)
            lo.Name = TABLE_NAME
            GoTo Reused
        End If
    End If
    rows = modUomSettings.GetUomCatalogRows()
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SHEET_NAME
    End If
    headers = StagingHeaders()
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
    outcome = "OPENED"
    report = "UOM Catalog table sent to the captured workbook. Edit the table, select it, then Retrieve UOM Catalog."
    Exit Function
Reused:
    SendUomCatalogToWorksheet = True
    outcome = "REUSED"
    report = "UOM Catalog draft reopened in the captured workbook. Existing edits retained; saved catalog not reloaded."
    Exit Function
InvalidStage:
    report = "The UOM staging headers or ownership are ambiguous. No cells were changed."
    Exit Function
Failed:
    outcome = "FAILED"
    report = "The UOM workbench could not be opened. Verify the staging worksheet before retrying."
End Function

' The successful Retrieve boundary retains the full local extent across unlisting.
' A foreign, broken or off-sheet marker is ambiguity, never a fallback hint.
Private Function ReadStagingExtent(ByVal ws As Worksheet, ByRef marker As Name, _
                                   ByRef stage As Range) As Boolean
    Dim candidate As Name, localName As String, indexes() As Long
    On Error GoTo InvalidMarker
    For Each candidate In ws.Names
        localName = Mid$(candidate.Name, InStrRev(candidate.Name, "!") + 1)
        If StrComp(localName, EXTENT_NAME, vbTextCompare) = 0 Then
            If Not marker Is Nothing Then Exit Function
            Set marker = candidate
        End If
    Next candidate
    If marker Is Nothing Then ReadStagingExtent = True: Exit Function
    If marker.Visible Or marker.Comment <> EXTENT_OWNER Then Exit Function
    Set stage = marker.RefersToRange
    If Not stage.Parent Is ws Then Exit Function
    If stage.Areas.Count <> 1 Then Exit Function
    If Not StagingColumns(stage.Rows(1), indexes) Then Exit Function
    ReadStagingExtent = True
InvalidMarker:
End Function

Private Sub RememberStagingExtent(ByVal lo As ListObject, ByVal marker As Name)
    Dim sheetPrefix As String, reference As String
    sheetPrefix = "'" & Replace(lo.Parent.Name, "'", "''") & "'!"
    reference = "=" & sheetPrefix & lo.Range.Address(True, True, xlA1)
    If marker Is Nothing Then
        Set marker = lo.Parent.Names.Add(Name:=sheetPrefix & EXTENT_NAME, _
                                        RefersTo:=reference, Visible:=False)
        marker.Comment = EXTENT_OWNER
    Else
        marker.RefersTo = reference
    End If
End Sub

Private Function StagingHeaders() As Variant
    StagingHeaders = Array("UOM", "Dimension", "Base UOM", "Units Per Base UOM", "Convertible", "Enabled", "Notes")
End Function

' Resolve the required subset, rejecting duplicate normalized managed headers.
Private Function StagingColumns(ByVal headerRange As Range, ByRef indexes() As Long) As Boolean
    Dim headers As Variant, column As Long, field As Long, value As Variant, key As String
    headers = StagingHeaders()
    ReDim indexes(1 To 7)
    For column = 1 To headerRange.Columns.Count
        value = headerRange.Cells(1, column).Value2
        If IsError(value) Then Exit Function
        key = Trim$(CStr(value))
        If key = "" Then Exit Function
        For field = 1 To 7
            If StrComp(key, CStr(headers(field - 1)), vbTextCompare) = 0 Then
                If indexes(field) <> 0 Then Exit Function
                indexes(field) = column
            End If
        Next field
    Next column
    For field = 1 To 7
        If indexes(field) = 0 Then Exit Function
    Next field
    StagingColumns = True
End Function

' Only one complete, contiguous header-led region identifies an unlisted draft.
Private Function FindStagingRegion(ByVal ws As Worksheet) As Range
    Dim found As Range, region As Range, candidate As Range, firstAddress As String
    Dim indexes() As Long, key As Variant
    Set found = ws.UsedRange.Find(What:="UOM", LookIn:=xlValues, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    If found Is Nothing Then Exit Function
    firstAddress = found.Address
    Do
        key = found.Value2
        If Not IsError(key) Then
            If StrComp(Trim$(CStr(key)), "UOM", vbTextCompare) = 0 Then
                Set region = found.CurrentRegion
                If region.Row = found.Row Then
                    If StagingColumns(region.Rows(1), indexes) Then
                        If Not candidate Is Nothing Then Exit Function
                        Set candidate = region
                    End If
                End If
            End If
        End If
        Set found = ws.UsedRange.FindNext(found)
        If found Is Nothing Then Exit Do
    Loop Until found.Address = firstAddress
    Set FindStagingRegion = candidate
End Function

Public Function RetrieveUomCatalogFromWorksheet(ByVal wb As Workbook, _
                                                 Optional ByRef report As String = "", _
                                                 Optional ByRef outcome As String = "") As Boolean
    Dim lo As ListObject
    Dim values As Variant
    Dim indexes() As Long, row As Long, field As Long
    Dim marker As Name, stage As Range
    outcome = "REJECTED"

    If wb Is Nothing Then
        report = "Production has no captured workbook for UOM Catalog retrieval."
        Exit Function
    End If
    On Error Resume Next
    Set lo = Application.ActiveCell.ListObject
    On Error GoTo 0
    If lo Is Nothing Then
        report = "Select a cell in the invSys UOM Catalog table in the captured Production workbook."
        Exit Function
    End If
    If Not lo.Parent.Parent Is wb Or StrComp(lo.Name, TABLE_NAME, vbTextCompare) <> 0 Then
        report = "Select a cell in the invSys UOM Catalog table in the captured Production workbook."
        Exit Function
    End If
    If lo.DataBodyRange Is Nothing Then
        report = "The selected UOM Catalog table has no rows."
        Exit Function
    End If
    If Not StagingColumns(lo.HeaderRowRange, indexes) Then
        report = "The UOM staging table requires one of each managed header. No cells were changed."
        Exit Function
    End If
    If Not ReadStagingExtent(lo.Parent, marker, stage) Then
        report = "The UOM staging extent has ambiguous ownership or headers. No cells or catalog values were changed."
        Exit Function
    End If
    ReDim values(1 To lo.ListRows.Count, 1 To 7)
    For row = 1 To lo.ListRows.Count
        For field = 1 To 7
            values(row, field) = lo.DataBodyRange.Cells(row, indexes(field)).Value2
        Next field
    Next row
    If Not modUomSettings.PublishUomCatalogRows(values, report, outcome) Then Exit Function
    RememberStagingExtent lo, marker
    lo.Unlist
    report = report & " The staging table was retrieved and removed."
    RetrieveUomCatalogFromWorksheet = True
End Function
