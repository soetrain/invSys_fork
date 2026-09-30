Attribute VB_Name = "modProductionProcessWorksheet"
Option Explicit

Private Const EDITOR_SHEET As String = "invSys Process Editor"
Private Const TABLE_PREFIX As String = "invSys_Process_"
Private Const FIRST_TABLE_TOP_ROW As Long = 1
Private Const TABLE_HEADER_OFFSET As Long = 5
Private Const TABLE_GAP_ROWS As Long = 6

' Initial construction only; runtime fields resolve through normalized headers.
Private Const INITIAL_ACCEPTABLE_ITEM_COLUMN As Long = 14
Private Const INITIAL_ACCEPTED_SKU_COLUMN As Long = 15
Private Const FIRST_ALTERNATIVE_PAIR As Long = 1
Private Const DEFAULT_ALTERNATIVE_PAIRS As Long = 4

Private mApplyingManagedColumns As Boolean

Public Function SendProcessDraftToWorksheet(ByVal wb As Workbook, _
                                            ByVal processId As String, _
                                            ByVal processVersion As String, _
                                            ByVal processName As String, _
                                            ByVal description As String, _
                                            ByVal payloadJson As String, _
                                            ByRef tableName As String, _
                                            ByRef report As String) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim rows As Collection
    Dim rowCount As Long
    Dim tableRange As Range
    Dim rowIndex As Long
    Dim record As Object
    Dim tableTopRow As Long
    Dim tableHeaderRow As Long
    Dim alternativePairCount As Long

    On Error GoTo Failed
    If wb Is Nothing Then
        report = "The captured Production operator workbook is unavailable."
        Exit Function
    End If
    processId = UCase$(Trim$(processId))
    processVersion = Trim$(processVersion)
    processName = Trim$(processName)
    If Not mProduction.IsBase36Identifier(processId) Then
        report = "The Process needs a generated three-character Base-36 ID before worksheet editing."
        Exit Function
    End If
    If processVersion = "" Or Not IsNumeric(processVersion) Or CDbl(processVersion) <= 0 Then
        report = "The Process needs a positive generated version before worksheet editing."
        Exit Function
    End If
    Set rows = BuildWorksheetRows(payloadJson, report)
    If rows Is Nothing Then Exit Function
    AddWorksheetTemplateRows rows
    rowCount = rows.Count
    alternativePairCount = WorksheetAlternativePairCountForRows(rows)
    If rowCount = 0 Then
        report = "The Process worksheet could not create an editable row set."
        Exit Function
    End If

    Set ws = EnsureProcessEditorSheet(wb)
    tableTopRow = NextProcessTableTopRow(ws)
    tableHeaderRow = tableTopRow + TABLE_HEADER_OFFSET
    ws.Cells(tableTopRow + 1, 5).NumberFormat = "@"
    ws.Cells(tableTopRow, 1).Value2 = "invSys Process worksheet"
    ws.Cells(tableTopRow + 1, 1).Value2 = "Process Name"
    ws.Cells(tableTopRow + 1, 2).Value2 = processName
    ws.Cells(tableTopRow + 1, 4).Value2 = "Process ID"
    ws.Cells(tableTopRow + 1, 5).Value2 = processId
    ws.Cells(tableTopRow + 1, 7).Value2 = "Version"
    ws.Cells(tableTopRow + 1, 8).Value2 = processVersion
    ws.Cells(tableTopRow + 2, 1).Value2 = "Description"
    ws.Cells(tableTopRow + 2, 2).Value2 = description
    ws.Cells(tableTopRow + 3, 1).Value2 = _
        "FIXED INPUT quantities calculate by UOM. ACTUAL inputs/outputs stay blank until Check In/Actual Output; do not enter a planned Qty."

    WriteWorksheetHeaders ws, tableHeaderRow, alternativePairCount
    Set tableRange = ws.Range(ws.Cells(tableHeaderRow, 1), _
                              ws.Cells(tableHeaderRow + rowCount, _
                                  AlternativeSkuColumnIndex(alternativePairCount)))
    Set lo = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    tableName = BuildUniqueProcessTableName(wb, processId)
    lo.Name = tableName
    ApplyProcessWorksheetTextIdentityFormats lo

    rowIndex = 1
    For Each record In rows
        WriteWorksheetRecord lo, rowIndex, record
        rowIndex = rowIndex + 1
    Next record
    ApplyProcessWorksheetManagedColumns lo
    FormatProcessWorksheet ws, lo
    wb.Save
    ws.Visible = xlSheetVisible
    ws.Activate
    lo.Range.Cells(1, 1).Select
    report = "Created Process table " & tableName & ". Select a cell in any completed table, then choose Retrieve Selected Process."
    SendProcessDraftToWorksheet = True
    Exit Function

Failed:
    report = "Process worksheet creation failed: " & Err.Description
End Function

Public Function ReadProcessDraftFromWorksheet(ByVal wb As Workbook, _
                                              ByVal tableName As String, _
                                              ByRef processId As String, _
                                              ByRef processVersion As String, _
                                              ByRef processName As String, _
                                              ByRef description As String, _
                                              ByRef payloadJson As String, _
                                              ByRef report As String) As Boolean
    Dim lo As ListObject
    Dim records As New Collection
    Dim record As Object
    Dim usedRowIds As Object
    Dim inputUoms As Object
    Dim inputPercentTotals As Object
    Dim inputRowCount As Long
    Dim outputRowCount As Long
    Dim instructionOrdinal As Long
    Dim rowIndex As Long
    Dim recordType As String
    Dim rowId As String
    Dim rowName As String
    Dim acceptedSku As String
    Dim acceptableItem As String
    Dim outputSku As String
    Dim requirementId As String
    Dim designId As String
    Dim designVersion As String
    Dim uom As String
    Dim instructionText As String
    Dim qtyMode As String
    Dim qty As Double
    Dim percentValue As Double
    Dim basisQty As Double
    Dim hasQty As Boolean
    Dim hasPercent As Boolean
    Dim expectedProcessId As String
    Dim inputUom As Variant
    Dim inputPercentTotal As Double

    On Error GoTo Failed
    If wb Is Nothing Then
        report = "The captured Production operator workbook is unavailable."
        Exit Function
    End If
    Set lo = FindProcessTableByName(wb, tableName)
    If lo Is Nothing Then
        report = "The bound Process worksheet table is missing: " & tableName & "."
        Exit Function
    End If
    expectedProcessId = ProcessIdFromTableName(lo.Name)
    If Not mProduction.IsBase36Identifier(expectedProcessId) Then
        report = "The Process worksheet table identity is invalid."
        Exit Function
    End If
    processId = expectedProcessId
    ApplyProcessWorksheetManagedColumns lo
    Application.Calculate
    processVersion = Trim$(ProcessMetadataValue(lo, 1, 8))
    processName = Trim$(ProcessMetadataValue(lo, 1, 2))
    description = Trim$(ProcessMetadataValue(lo, 2, 2))
    If processName = "" Then
        report = "Process Name is required on the worksheet."
        Exit Function
    End If
    If processVersion = "" Or Not IsNumeric(processVersion) Or CDbl(processVersion) <= 0 Then
        report = "The Process worksheet version must be a positive number."
        Exit Function
    End If

    Set usedRowIds = CreateObject("Scripting.Dictionary")
    usedRowIds.CompareMode = vbTextCompare
    Set inputUoms = CreateObject("Scripting.Dictionary")
    inputUoms.CompareMode = vbTextCompare
    Set inputPercentTotals = CreateObject("Scripting.Dictionary")
    inputPercentTotals.CompareMode = vbTextCompare

    Set record = NewWorksheetRecord("PROCESS")
    record("ProcessName") = processName
    record("Description") = description
    records.Add record

    If lo.DataBodyRange Is Nothing Then
        report = "The Process worksheet table contains no rows."
        Exit Function
    End If
    For rowIndex = 1 To lo.ListRows.Count
        recordType = UCase$(Trim$(modProcessWorksheetColumns.Text(lo, rowIndex, "Record Type")))
        If Not WorksheetRowHasBusinessData(lo, rowIndex) Then GoTo ContinueRow
        rowId = UCase$(Trim$(modProcessWorksheetColumns.Text(lo, rowIndex, "ID")))
        rowName = Trim$(modProcessWorksheetColumns.Text(lo, rowIndex, "Name"))
        acceptedSku = Trim$(modProcessWorksheetColumns.Text(lo, rowIndex, "Accepted SKU 1"))
        acceptableItem = Trim$(modProcessWorksheetColumns.Text(lo, rowIndex, "Acceptable Managed Item 1"))
        outputSku = Trim$(modProcessWorksheetColumns.Text(lo, rowIndex, "Output SKU"))
        requirementId = UCase$(Trim$(modProcessWorksheetColumns.Text(lo, rowIndex, "Requirement ID")))
        uom = UCase$(Trim$(modProcessWorksheetColumns.Text(lo, rowIndex, "UOM")))
        instructionText = Trim$(modProcessWorksheetColumns.Text(lo, rowIndex, "Instruction"))
        qtyMode = UCase$(Trim$(modProcessWorksheetColumns.Text(lo, rowIndex, "Qty Mode")))
        If qtyMode = "" Then qtyMode = "FIXED"
        hasQty = TryPositiveWorksheetNumber(modProcessWorksheetColumns.Text(lo, rowIndex, "Qty"), qty)
        hasPercent = TryPositiveWorksheetNumber(modProcessWorksheetColumns.Text(lo, rowIndex, "Percent"), percentValue)
        Call TryPositiveWorksheetNumber(modProcessWorksheetColumns.Text(lo, rowIndex, "Basis Qty"), basisQty)

        Select Case recordType
            Case "INPUT", "REQUIREMENT"
                rowId = ResolveWorksheetRowId(rowId, usedRowIds)
                If rowName = "" Or uom = "" Then
                    report = "Each INPUT row needs a name and UOM."
                    Exit Function
                End If
                If Not IsConfiguredProcessUom(uom) Then
                    report = uom & " UOM is not in the Recipe UOM Catalog."
                    Exit Function
                End If
                If qtyMode <> "FIXED" And qtyMode <> "ACTUAL" Then
                    report = "Qty Mode must be FIXED or ACTUAL."
                    Exit Function
                End If
                If qtyMode = "ACTUAL" And (hasQty Or hasPercent Or basisQty > 0) Then
                    report = "Variable -- determined at Check In INPUT cannot carry Qty, Percent, or Batch basis."
                    Exit Function
                End If
                If qtyMode = "FIXED" And Not hasQty And Not hasPercent Then
                    report = "Each INPUT row needs a positive quantity or percentage."
                    Exit Function
                End If
                If qtyMode = "FIXED" And hasPercent And basisQty <= 0 Then
                    report = "Each percentage INPUT needs a positive Batch basis quantity."
                    Exit Function
                End If
                If usedRowIds.Exists(rowId) Then
                    report = "Duplicate Process row ID " & rowId & "."
                    Exit Function
                End If
                usedRowIds.Add rowId, True
                If qtyMode = "FIXED" Then inputUoms(uom) = True
                inputRowCount = inputRowCount + 1
                If qtyMode = "FIXED" Then
                    If Not inputPercentTotals.Exists(uom) Then inputPercentTotals.Add uom, 0#
                    If hasPercent Then _
                        inputPercentTotals(uom) = CDbl(inputPercentTotals(uom)) + percentValue
                End If
                Set record = NewWorksheetRecord("REQUIREMENT")
                record("RequirementId") = rowId
                record("RequirementName") = rowName
                If hasQty Then record("Qty") = qty
                If hasPercent Then record("Percent") = percentValue
                If basisQty > 0 Then record("YieldBasis") = basisQty
                record("UOM") = uom
                record("RequirementQtyMode") = qtyMode
                records.Add record
                AppendWorksheetAlternativeRecords lo, rowIndex, rowId, records
            Case "OUTPUT"
                If outputSku = "" Then outputSku = acceptedSku
                If rowName = "" Then rowName = acceptableItem
                rowId = ResolveWorksheetRowId(rowId, usedRowIds)
                If rowName = "" Or outputSku = "" Or uom = "" Then
                    report = "Each OUTPUT row needs a managed item selected from item search and a UOM."
                    Exit Function
                End If
                If Not IsConfiguredProcessUom(uom) Then
                    report = uom & " UOM is not in the Recipe UOM Catalog."
                    Exit Function
                End If
                If qtyMode <> "FIXED" And qtyMode <> "ACTUAL" Then
                    report = "Qty Mode must be FIXED or ACTUAL."
                    Exit Function
                End If
                If qtyMode = "ACTUAL" And (hasQty Or hasPercent Or basisQty > 0) Then
                    report = "Variable -- determined by Actual Output OUTPUT cannot carry Qty, Percent, or Yield basis."
                    Exit Function
                End If
                If qtyMode = "FIXED" And Not hasQty And Not hasPercent Then
                    report = "Each OUTPUT row needs a positive quantity or percentage."
                    Exit Function
                End If
                If qtyMode = "FIXED" And hasPercent And basisQty <= 0 Then
                    report = "Each percentage OUTPUT needs a positive Yield basis quantity."
                    Exit Function
                End If
                If usedRowIds.Exists(rowId) Then
                    report = "Duplicate Process row ID " & rowId & "."
                    Exit Function
                End If
                usedRowIds.Add rowId, True
                outputRowCount = outputRowCount + 1
                designId = GeneratedOutputDesignId(processId, rowId)
                designVersion = processVersion
                Set record = NewWorksheetRecord("OUTPUT")
                record("OutputId") = rowId
                record("OutputName") = rowName
                record("ITEM_CODE") = outputSku
                record("ComponentDesignId") = designId
                record("ComponentDesignVersion") = designVersion
                If hasQty Then record("Qty") = qty
                If hasPercent Then record("Percent") = percentValue
                If basisQty > 0 Then record("YieldBasis") = basisQty
                record("UOM") = uom
                record("OutputQtyMode") = qtyMode
                records.Add record
            Case "INSTRUCTION"
                rowId = ResolveWorksheetRowId(rowId, usedRowIds)
                If usedRowIds.Exists(rowId) Then
                    report = "Duplicate Process row ID " & rowId & "."
                    Exit Function
                End If
                usedRowIds.Add rowId, True
                If instructionText = "" Then instructionText = rowName
                If instructionText = "" Then
                    report = "Each INSTRUCTION row needs instruction text."
                    Exit Function
                End If
                instructionOrdinal = instructionOrdinal + 1
                Set record = NewWorksheetRecord("INSTRUCTION")
                record("InstructionOrdinal") = instructionOrdinal
                record("Instruction") = instructionText
                records.Add record
            Case "ALTERNATIVE"
                If requirementId = "" Then requirementId = rowId
                If requirementId = "" Or acceptedSku = "" Then
                    report = "Each ALTERNATIVE row needs its Requirement ID and an Acceptable Managed Item selected from item search."
                    Exit Function
                End If
                Set record = NewWorksheetRecord("ALTERNATIVE")
                record("RequirementId") = requirementId
                record("ITEM_CODE") = acceptedSku
                If acceptableItem <> "" Then record("ItemName") = acceptableItem
                records.Add record
            Case Else
                report = "Unknown Process worksheet row type: " & recordType & "."
                Exit Function
        End Select
ContinueRow:
    Next rowIndex

    If outputRowCount = 0 Then
        report = "Every Process must declare at least one OUTPUT."
        Exit Function
    End If
    For Each inputUom In inputUoms.Keys
        inputPercentTotal = 0#
        If inputPercentTotals.Exists(CStr(inputUom)) Then _
            inputPercentTotal = CDbl(inputPercentTotals(CStr(inputUom)))
        If Abs(inputPercentTotal - 100#) > 0.05 Then
            report = "INPUT formula percentages for " & CStr(inputUom) & _
                " UOM must total 100.0%; current total is " & _
                Format$(inputPercentTotal, "0.0") & "%."
            Exit Function
        End If
    Next inputUom

    payloadJson = modProductionJson.BuildJsonArray(records)
    report = "Process worksheet validated: " & CStr(inputRowCount) & _
        " input(s), " & CStr(outputRowCount) & " output(s), and " & _
        CStr(instructionOrdinal) & " instruction(s)."
    ReadProcessDraftFromWorksheet = True
    Exit Function

Failed:
    report = "Process worksheet retrieval failed: " & Err.Description
End Function

Public Function DeleteProcessWorksheetTable(ByVal wb As Workbook, _
                                            ByVal tableName As String, _
                                            ByRef report As String) As Boolean
    Dim lo As ListObject
    Dim ws As Worksheet

    On Error GoTo Failed
    Set lo = FindProcessTableByName(wb, tableName)
    If lo Is Nothing Then
        report = "The Process worksheet table is already absent."
        DeleteProcessWorksheetTable = True
        Exit Function
    End If
    Set ws = lo.Parent
    ClearProcessTableMetadata lo
    lo.Delete
    wb.Save
    report = "Retrieved Process and removed selected table " & tableName & "."
    DeleteProcessWorksheetTable = True
    Exit Function
Failed:
    report = "The Process was retrieved, but its temporary table could not be removed: " & Err.Description
End Function

Public Function FindOutstandingProcessWorksheetTable(ByVal wb As Workbook, _
                                                     ByRef tableName As String, _
                                                     ByRef report As String) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim foundCount As Long

    tableName = ""
    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If Left$(lo.Name, Len(TABLE_PREFIX)) = TABLE_PREFIX Then
                foundCount = foundCount + 1
                tableName = lo.Name
            End If
        Next lo
    Next ws
    If foundCount >= 1 Then
        report = CStr(foundCount) & " Process worksheet table(s) found; selected " & tableName & "."
        FindOutstandingProcessWorksheetTable = True
    Else
        report = "No Process worksheet table was found."
    End If
End Function

Public Function FindSelectedProcessWorksheetTable(ByVal wb As Workbook, _
                                                   ByRef tableName As String, _
                                                   ByRef report As String) As Boolean
    Dim tableNames As Collection

    tableName = ""
    Set tableNames = FindSelectedProcessWorksheetTables(wb, report)
    If tableNames Is Nothing Then Exit Function
    If tableNames.Count = 0 Then Exit Function
    tableName = CStr(tableNames(1))
    FindSelectedProcessWorksheetTable = True
End Function

Public Function FindSelectedProcessWorksheetTables(ByVal wb As Workbook, _
                                                    ByRef report As String) As Collection
    Dim selectedRange As Range
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim hit As Range
    Dim result As New Collection
    Dim selectedAreaCount As Long

    If wb Is Nothing Then
        report = "The captured Production operator workbook is unavailable."
        Exit Function
    End If
    On Error Resume Next
    Set selectedRange = Application.Selection
    On Error GoTo 0
    If selectedRange Is Nothing Then
        report = "Select one or Ctrl+click cells inside the Process tables to retrieve."
        Exit Function
    End If
    If Not selectedRange.Worksheet.Parent Is wb Then
        report = "Select a Process table cell in the captured Production workbook."
        Exit Function
    End If
    selectedAreaCount = selectedRange.Areas.Count
    For Each ws In wb.Worksheets
        If ws Is selectedRange.Worksheet Then
            For Each lo In ws.ListObjects
                If IsInvSysProcessTable(lo) Then
                    Set hit = Nothing
                    On Error Resume Next
                    Set hit = Application.Intersect(selectedRange, lo.Range)
                    On Error GoTo 0
                    If Not hit Is Nothing Then result.Add lo.Name
                End If
            Next lo
        End If
    Next ws
    If result.Count = 0 Then
        report = "The selection does not intersect an invSys Process table."
        Exit Function
    End If
    report = "Selected " & CStr(result.Count) & " Process table(s) across " & _
        CStr(selectedAreaCount) & " selection area(s)."
    Set FindSelectedProcessWorksheetTables = result
End Function

Public Function CountProcessWorksheetTables(ByVal wb As Workbook) As Long
    Dim ws As Worksheet
    Dim lo As ListObject

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If IsInvSysProcessTable(lo) Then _
                CountProcessWorksheetTables = CountProcessWorksheetTables + 1
        Next lo
    Next ws
End Function

Public Function SelectProcessWorksheetTableForTest(ByVal wb As Workbook, _
                                                    ByVal tableName As String) As Boolean
    Dim lo As ListObject

    Set lo = FindProcessTableByName(wb, tableName)
    If lo Is Nothing Then Exit Function
    lo.Parent.Activate
    lo.DataBodyRange.Cells(1, 1).Select
    SelectProcessWorksheetTableForTest = True
End Function

Public Function SelectProcessWorksheetTablesForTest(ByVal wb As Workbook, _
                                                     ByVal firstTableName As String, _
                                                     ByVal secondTableName As String) As Boolean
    Dim firstTable As ListObject
    Dim secondTable As ListObject
    Dim selectedCells As Range

    Set firstTable = FindProcessTableByName(wb, firstTableName)
    Set secondTable = FindProcessTableByName(wb, secondTableName)
    If firstTable Is Nothing Or secondTable Is Nothing Then Exit Function
    If Not firstTable.Parent Is secondTable.Parent Then Exit Function
    firstTable.Parent.Activate
    Set selectedCells = Application.Union(firstTable.DataBodyRange.Cells(1, 1), _
                                          secondTable.DataBodyRange.Cells(1, 1))
    selectedCells.Select
    SelectProcessWorksheetTablesForTest = (selectedCells.Areas.Count = 2)
End Function

Public Function AddAcceptableItemPairToSelectedTable(ByVal wb As Workbook, _
                                                     ByRef report As String) As Boolean
    Dim tableName As String
    Dim lo As ListObject
    Dim pairNumber As Long
    Dim itemColumn As ListColumn
    Dim skuColumn As ListColumn

    If Not FindSelectedProcessWorksheetTable(wb, tableName, report) Then Exit Function
    Set lo = FindProcessTableByName(wb, tableName)
    If lo Is Nothing Then Exit Function
    pairNumber = AlternativePairCount(lo) + 1
    Set itemColumn = lo.ListColumns.Add
    itemColumn.Name = "Acceptable Managed Item " & CStr(pairNumber)
    Set skuColumn = lo.ListColumns.Add
    skuColumn.Name = "Accepted SKU " & CStr(pairNumber)
    skuColumn.DataBodyRange.NumberFormat = "@"
    itemColumn.Range.ColumnWidth = 28
    skuColumn.Range.EntireColumn.Hidden = True
    wb.Save
    report = "Added Acceptable Managed Item " & CStr(pairNumber) & _
        " to Process table " & tableName & "."
    AddAcceptableItemPairToSelectedTable = True
End Function

Public Function IsProcessWorksheetTableTarget(ByVal target As Range) As Boolean
    Dim lo As ListObject

    If target Is Nothing Then Exit Function
    On Error Resume Next
    Set lo = target.ListObject
    On Error GoTo 0
    IsProcessWorksheetTableTarget = IsInvSysProcessTable(lo)
End Function

Public Function IsProcessWorksheetItemSearchTarget(ByVal target As Range) As Boolean
    IsProcessWorksheetItemSearchTarget = (ProcessManagedItemPairNumber(target) > 0)
End Function

Public Function ProcessManagedItemPairNumber(ByVal target As Range) As Long
    Dim lo As ListObject
    Dim recordTypeColumn As ListColumn
    Dim rowIndex As Long
    Dim pairNumber As Long
    Dim targetColumn As ListColumn

    If target Is Nothing Or target.Cells.CountLarge <> 1 Then Exit Function
    On Error Resume Next
    Set lo = target.ListObject
    On Error GoTo 0
    If Not IsInvSysProcessTable(lo) Then Exit Function
    On Error Resume Next
    Set recordTypeColumn = modProcessWorksheetColumns.Field(lo, "Record Type")
    On Error GoTo 0
    If recordTypeColumn Is Nothing Or target.Row <= lo.HeaderRowRange.Row Then Exit Function
    For pairNumber = FIRST_ALTERNATIVE_PAIR To AlternativePairCount(lo)
        Set targetColumn = Nothing
        On Error Resume Next
        Set targetColumn = modProcessWorksheetColumns.Field(lo, "Acceptable Managed Item " & CStr(pairNumber))
        On Error GoTo 0
        If Not targetColumn Is Nothing Then
            If target.Column = targetColumn.Range.Column Then Exit For
        End If
    Next pairNumber
    If pairNumber > AlternativePairCount(lo) Then Exit Function
    rowIndex = target.Row - lo.DataBodyRange.Row + 1
    If rowIndex < 1 Or rowIndex > lo.ListRows.Count Then Exit Function
    Select Case UCase$(Trim$(CellText(lo.DataBodyRange.Cells(rowIndex, _
            recordTypeColumn.Index).Value2)))
        Case "INPUT", "REQUIREMENT"
            ProcessManagedItemPairNumber = pairNumber
        Case "OUTPUT"
            If pairNumber = FIRST_ALTERNATIVE_PAIR Then _
                ProcessManagedItemPairNumber = pairNumber
    End Select
End Function

Public Sub RefreshProcessWorksheetManagedColumns(ByVal target As Range)
    Dim lo As ListObject

    If mApplyingManagedColumns Then Exit Sub
    If target Is Nothing Then Exit Sub
    On Error Resume Next
    Set lo = target.ListObject
    On Error GoTo 0
    If Not IsInvSysProcessTable(lo) Then Exit Sub
    ApplyProcessWorksheetManagedColumns lo
End Sub

Public Function PopulateFormulationExampleForTest(ByVal wb As Workbook, _
                                                  ByVal tableName As String, _
                                                  ByVal mixedUom As Boolean, _
                                                  ByRef report As String) As Boolean
    Dim lo As ListObject
    Dim names As Variant
    Dim quantities As Variant
    Dim uoms As Variant
    Dim rowIndex As Long

    On Error GoTo Failed
    Set lo = FindProcessTableByName(wb, tableName)
    If lo Is Nothing Then
        report = "Test Process worksheet table is missing."
        Exit Function
    End If
    If mixedUom Then
        names = Array("Chai Concentrate", "64oz Empty Bottle", _
            "48mm Bottle Cap", "")
        quantities = Array(4.5, 1#, 1#, "")
        uoms = Array("LB", "EA", "EA", "")
    Else
        names = Array("Sugar", "Flour", "Baking Powder", "Filtered Water")
        quantities = Array(100#, 200#, 11.2, 300#)
        uoms = Array("LB", "LB", "LB", "LB")
    End If
    For rowIndex = 1 To 4
        modProcessWorksheetColumns.Cell(lo, rowIndex, "Record Type").Value2 = "INPUT"
        modProcessWorksheetColumns.Cell(lo, rowIndex, "Name").Value2 = names(rowIndex - 1)
        modProcessWorksheetColumns.Cell(lo, rowIndex, "Qty").Value2 = quantities(rowIndex - 1)
        modProcessWorksheetColumns.Cell(lo, rowIndex, "UOM").Value2 = uoms(rowIndex - 1)
    Next rowIndex
    modProcessWorksheetColumns.Cell(lo, 7, "Record Type").Value2 = "OUTPUT"
    If mixedUom Then
        modProcessWorksheetColumns.Cell(lo, 7, "Name").Value2 = "64oz Classic Chai"
        modProcessWorksheetColumns.Cell(lo, 7, "Output SKU").Value2 = "SKU-CHAI-64OZ"
        modProcessWorksheetColumns.Cell(lo, 7, "Qty").Value2 = 1#
        modProcessWorksheetColumns.Cell(lo, 7, "UOM").Value2 = "EA"
        modProcessWorksheetColumns.Cell(lo, 1, "Acceptable Managed Item " & CStr(1)).Value2 = "Chai Concentrate"
        modProcessWorksheetColumns.Cell(lo, 1, "Accepted SKU " & CStr(1)).Value2 = "SKU-CHAI-CONCENTRATE"
        modProcessWorksheetColumns.Cell(lo, 2, "Acceptable Managed Item " & CStr(1)).Value2 = "64oz Empty Bottle"
        modProcessWorksheetColumns.Cell(lo, 2, "Accepted SKU " & CStr(1)).Value2 = "SKU-BOTTLE-64OZ"
        modProcessWorksheetColumns.Cell(lo, 3, "Acceptable Managed Item " & CStr(1)).Value2 = "48mm Bottle Cap"
        modProcessWorksheetColumns.Cell(lo, 3, "Accepted SKU " & CStr(1)).Value2 = "SKU-CAP-48MM"
    Else
        modProcessWorksheetColumns.Cell(lo, 7, "Name").Value2 = "Finished Formula"
        modProcessWorksheetColumns.Cell(lo, 7, "Output SKU").Value2 = "SKU-FINISHED"
        modProcessWorksheetColumns.Cell(lo, 7, "Qty").Value2 = 611.2
        modProcessWorksheetColumns.Cell(lo, 7, "UOM").Value2 = "LB"
        modProcessWorksheetColumns.Cell(lo, 1, "Acceptable Managed Item " & CStr(1)).Value2 = "Sugar Stock"
        modProcessWorksheetColumns.Cell(lo, 1, "Accepted SKU " & CStr(1)).Value2 = "SKU-SUGAR"
    End If
    ApplyProcessWorksheetManagedColumns lo
    Application.Calculate
    report = "Example populated."
    PopulateFormulationExampleForTest = True
    Exit Function
Failed:
    report = "Example population failed: " & Err.Description
End Function

Public Function ReadFormulaEvidenceForTest(ByVal wb As Workbook, _
                                           ByVal tableName As String) As String
    Dim lo As ListObject
    Dim totalPercent As Double
    Dim rowIndex As Long

    Set lo = FindProcessTableByName(wb, tableName)
    If lo Is Nothing Then
        ReadFormulaEvidenceForTest = "FAIL|TableMissing"
        Exit Function
    End If
    For rowIndex = 1 To 4
        totalPercent = totalPercent + CDbl(modProcessWorksheetColumns.Cell(lo, rowIndex, "Percent").Value2)
    Next rowIndex
    ReadFormulaEvidenceForTest = "OK|Basis=" & _
        Format$(CDbl(modProcessWorksheetColumns.Cell(lo, 1, "Basis Qty").Value2), "0.0") & _
        "|Sugar=" & Format$(CDbl(modProcessWorksheetColumns.Cell(lo, 1, "Percent").Value2), "0.0") & _
        "|Flour=" & Format$(CDbl(modProcessWorksheetColumns.Cell(lo, 2, "Percent").Value2), "0.0") & _
        "|BakingPowder=" & Format$(CDbl(modProcessWorksheetColumns.Cell(lo, 3, "Percent").Value2), "0.0") & _
        "|Water=" & Format$(CDbl(modProcessWorksheetColumns.Cell(lo, 4, "Percent").Value2), "0.0") & _
        "|Total=" & Format$(totalPercent, "0.0")
End Function

Private Function BuildWorksheetRows(ByVal payloadJson As String, _
                                    ByRef report As String) As Collection
    Dim sourceRecords As Collection
    Dim rows As New Collection
    Dim record As Object
    Dim rowRecord As Object
    Dim parseReport As String
    Dim recordType As String
    Dim itemNameBySku As Object

    Set itemNameBySku = BuildManagedItemNameBySku()

    Set sourceRecords = modProductionReusableDesigns.ParseReusableDefinitionRecords(payloadJson, parseReport)
    If sourceRecords Is Nothing Then
        report = "The Process draft could not be serialized for worksheet editing: " & parseReport
        Exit Function
    End If
    For Each record In sourceRecords
        recordType = UCase$(modProductionReusableDesigns.ReusableRecordText(record, "RecordType"))
        Select Case recordType
            Case "REQUIREMENT"
                Set rowRecord = NewWorksheetRecord("INPUT")
                rowRecord("Id") = modProductionReusableDesigns.ReusableRecordText(record, "RequirementId")
                rowRecord("Name") = modProductionReusableDesigns.ReusableRecordText(record, "RequirementName")
                rowRecord("Qty") = modProductionReusableDesigns.ReusableRecordValue(record, "Qty")
                rowRecord("Percent") = modProductionReusableDesigns.ReusableRecordValue(record, "Percent")
                rowRecord("BasisQty") = modProductionReusableDesigns.ReusableRecordValue(record, "YieldBasis")
                rowRecord("UOM") = modProductionReusableDesigns.ReusableRecordText(record, "UOM")
                rowRecord("QtyMode") = CellText(modProductionReusableDesigns.ReusableRecordValue(record, "RequirementQtyMode"))
                rows.Add rowRecord
            Case "ALTERNATIVE"
                Call AttachAlternativeToRequirementRow(rows, _
                    modProductionReusableDesigns.ReusableRecordText(record, "RequirementId"), _
                    modProductionReusableDesigns.ReusableRecordText(record, "ITEM_CODE"), _
                    itemNameBySku)
            Case "OUTPUT"
                Set rowRecord = NewWorksheetRecord("OUTPUT")
                rowRecord("Id") = modProductionReusableDesigns.ReusableRecordText(record, "OutputId")
                rowRecord("Name") = modProductionReusableDesigns.ReusableRecordText(record, "OutputName")
                rowRecord("OutputSku") = modProductionReusableDesigns.ReusableRecordText(record, "ITEM_CODE")
                rowRecord("AcceptedSku1") = DictionaryText(rowRecord, "OutputSku")
                If itemNameBySku.Exists(DictionaryText(rowRecord, "OutputSku")) Then
                    rowRecord("AcceptableItem1") = _
                        itemNameBySku(DictionaryText(rowRecord, "OutputSku"))
                Else
                    rowRecord("AcceptableItem1") = DictionaryText(rowRecord, "Name")
                End If
                rowRecord("Qty") = modProductionReusableDesigns.ReusableRecordValue(record, "Qty")
                rowRecord("Percent") = modProductionReusableDesigns.ReusableRecordValue(record, "Percent")
                rowRecord("BasisQty") = modProductionReusableDesigns.ReusableRecordValue(record, "YieldBasis")
                rowRecord("UOM") = modProductionReusableDesigns.ReusableRecordText(record, "UOM")
                rowRecord("QtyMode") = CellText(modProductionReusableDesigns.ReusableRecordValue(record, "OutputQtyMode"))
                rowRecord("DesignId") = modProductionReusableDesigns.ReusableRecordText(record, "ComponentDesignId")
                rowRecord("DesignVersion") = modProductionReusableDesigns.ReusableRecordText(record, "ComponentDesignVersion")
                rows.Add rowRecord
            Case "INSTRUCTION"
                Set rowRecord = NewWorksheetRecord("INSTRUCTION")
                rowRecord("Instruction") = modProductionReusableDesigns.ReusableRecordText(record, "Instruction")
                rows.Add rowRecord
        End Select
    Next record
    Set BuildWorksheetRows = rows
End Function

Private Sub AddWorksheetTemplateRows(ByVal rows As Collection)
    Dim inputCount As Long
    Dim outputCount As Long
    Dim instructionCount As Long
    Dim usedRowIds As Object
    Dim record As Object
    Dim rowRecord As Object
    Dim recordType As String
    Dim addCount As Long
    Dim index As Long

    Set usedRowIds = CreateObject("Scripting.Dictionary")
    usedRowIds.CompareMode = vbTextCompare
    For Each record In rows
        recordType = UCase$(CellText(record("RecordType")))
        Select Case recordType
            Case "INPUT"
                inputCount = inputCount + 1
                If Trim$(CellText(record("Id"))) <> "" Then usedRowIds(UCase$(Trim$(CellText(record("Id"))))) = True
            Case "OUTPUT"
                outputCount = outputCount + 1
                If Trim$(CellText(record("Id"))) <> "" Then usedRowIds(UCase$(Trim$(CellText(record("Id"))))) = True
            Case "INSTRUCTION"
                instructionCount = instructionCount + 1
                If Trim$(CellText(record("Id"))) <> "" Then usedRowIds(UCase$(Trim$(CellText(record("Id"))))) = True
        End Select
    Next record

    addCount = IIf(inputCount = 0, 6, 2)
    For index = 1 To addCount
        Set rowRecord = NewWorksheetRecord("INPUT")
        rowRecord("Id") = NextIdFromDictionary(usedRowIds)
        usedRowIds(rowRecord("Id")) = True
        rows.Add rowRecord
    Next index
    addCount = IIf(outputCount = 0, 2, 1)
    For index = 1 To addCount
        Set rowRecord = NewWorksheetRecord("OUTPUT")
        rowRecord("Id") = NextIdFromDictionary(usedRowIds)
        usedRowIds(rowRecord("Id")) = True
        rows.Add rowRecord
    Next index
    Set rowRecord = NewWorksheetRecord("INSTRUCTION")
    rowRecord("Id") = NextIdFromDictionary(usedRowIds)
    usedRowIds(rowRecord("Id")) = True
    rows.Add rowRecord
End Sub

Private Function NewWorksheetRecord(ByVal recordType As String) As Object
    Dim record As Object
    Dim pairNumber As Long

    Set record = CreateObject("Scripting.Dictionary")
    record.CompareMode = vbTextCompare
    record("RecordType") = recordType
    record("Id") = ""
    record("Name") = ""
    record("Qty") = ""
    record("Percent") = ""
    record("BasisQty") = ""
    record("UOM") = ""
    record("DesignId") = ""
    record("DesignVersion") = ""
    record("Instruction") = ""
    record("RequirementId") = ""
    record("QtyMode") = "FIXED"
    For pairNumber = FIRST_ALTERNATIVE_PAIR To DEFAULT_ALTERNATIVE_PAIRS
        record("AcceptableItem" & CStr(pairNumber)) = ""
        record("AcceptedSku" & CStr(pairNumber)) = ""
    Next pairNumber
    Set NewWorksheetRecord = record
End Function

Private Sub AttachAlternativeToRequirementRow(ByVal rows As Collection, _
                                              ByVal requirementId As String, _
                                              ByVal acceptedSku As String, _
                                              ByVal itemNameBySku As Object)
    Dim rowRecord As Object
    Dim pairNumber As Long
    Dim itemName As String

    requirementId = UCase$(Trim$(requirementId))
    acceptedSku = Trim$(acceptedSku)
    If requirementId = "" Or acceptedSku = "" Then Exit Sub
    If Not itemNameBySku Is Nothing Then
        If itemNameBySku.Exists(acceptedSku) Then itemName = itemNameBySku(acceptedSku)
    End If
    For Each rowRecord In rows
        If UCase$(Trim$(CellText(rowRecord("RecordType")))) = "INPUT" _
           And StrComp(Trim$(CellText(rowRecord("Id"))), requirementId, vbTextCompare) = 0 Then
            For pairNumber = FIRST_ALTERNATIVE_PAIR To 99
                If Not rowRecord.Exists("AcceptedSku" & CStr(pairNumber)) _
                   Or Trim$(CellText(rowRecord("AcceptedSku" & CStr(pairNumber)))) = "" Then
                    rowRecord("AcceptedSku" & CStr(pairNumber)) = acceptedSku
                    rowRecord("AcceptableItem" & CStr(pairNumber)) = itemName
                    Exit Sub
                End If
            Next pairNumber
            Exit Sub
        End If
    Next rowRecord
End Sub

Private Sub WriteWorksheetRecord(ByVal lo As ListObject, ByVal rowIndex As Long, _
                                 ByVal record As Object)
    Dim pairNumber As Long

    modProcessWorksheetColumns.Cell(lo, rowIndex, "Record Type").Value2 = record("RecordType")
    modProcessWorksheetColumns.Cell(lo, rowIndex, "ID").Value2 = record("Id")
    modProcessWorksheetColumns.Cell(lo, rowIndex, "Name").Value2 = record("Name")
    modProcessWorksheetColumns.Cell(lo, rowIndex, "Qty").Value2 = record("Qty")
    modProcessWorksheetColumns.Cell(lo, rowIndex, "Percent").Value2 = record("Percent")
    modProcessWorksheetColumns.Cell(lo, rowIndex, "Basis Qty").Value2 = record("BasisQty")
    modProcessWorksheetColumns.Cell(lo, rowIndex, "UOM").Value2 = record("UOM")
    modProcessWorksheetColumns.Cell(lo, rowIndex, "Design ID").Value2 = record("DesignId")
    modProcessWorksheetColumns.Cell(lo, rowIndex, "Design Version").Value2 = record("DesignVersion")
    modProcessWorksheetColumns.Cell(lo, rowIndex, "Instruction").Value2 = record("Instruction")
    modProcessWorksheetColumns.Cell(lo, rowIndex, "Requirement ID").Value2 = record("RequirementId")
    modProcessWorksheetColumns.Cell(lo, rowIndex, "Qty Mode").Value2 = record("QtyMode")
    modProcessWorksheetColumns.Cell(lo, rowIndex, "Output SKU").Value2 = DictionaryText(record, "OutputSku")
    For pairNumber = FIRST_ALTERNATIVE_PAIR To AlternativePairCount(lo)
        modProcessWorksheetColumns.Cell(lo, rowIndex, "Acceptable Managed Item " & CStr(pairNumber)).Value2 = _
            DictionaryText(record, "AcceptableItem" & CStr(pairNumber))
        modProcessWorksheetColumns.Cell(lo, rowIndex, "Accepted SKU " & CStr(pairNumber)).Value2 = _
            DictionaryText(record, "AcceptedSku" & CStr(pairNumber))
    Next pairNumber
End Sub

Private Sub ApplyProcessWorksheetManagedColumns(ByVal lo As ListObject)
    Dim rowIndex As Long
    Dim recordType As String
    Dim priorAutoFill As Boolean
    Dim priorEvents As Boolean
    Dim processIdCell As String
    Dim processVersionCell As String

    On Error GoTo CleanExit
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    If mApplyingManagedColumns Then Exit Sub
    mApplyingManagedColumns = True
    priorEvents = Application.EnableEvents
    Application.EnableEvents = False
    priorAutoFill = Application.AutoCorrect.AutoFillFormulasInLists
    Application.AutoCorrect.AutoFillFormulasInLists = True
    ApplyProcessWorksheetTextIdentityFormats lo
    ApplyRecordTypeValidation lo
    ApplyQtyModeValidation lo
    ApplyProcessWorksheetUomValidation lo
    EnsureWorksheetRowIds lo
    processIdCell = lo.Parent.Cells(lo.HeaderRowRange.Row - TABLE_HEADER_OFFSET + 1, 5).Address(True, True)
    processVersionCell = lo.Parent.Cells(lo.HeaderRowRange.Row - TABLE_HEADER_OFFSET + 1, 8).Address(True, True)

    modProcessWorksheetColumns.Field(lo, "Requirement ID").DataBodyRange.NumberFormat = "General"
    modProcessWorksheetColumns.Field(lo, "Design ID").DataBodyRange.NumberFormat = "General"
    modProcessWorksheetColumns.SetFormula lo, "Basis Qty", _
        "=IF(AND(UPPER([@[Record Type]])=""INPUT"",UPPER([@[Qty Mode]])<>""ACTUAL""),IFERROR(SUMIFS([Qty],[Record Type],""INPUT"",[UOM],[@UOM]),""""),"""")"
    modProcessWorksheetColumns.SetFormula lo, "Percent", _
        "=IF(AND(UPPER([@[Record Type]])=""INPUT"",UPPER([@[Qty Mode]])<>""ACTUAL""),IFERROR([@Qty]/[@[Basis Qty]]*100,""""),"""")"
    modProcessWorksheetColumns.SetFormula lo, "Design ID", _
        "=IF(AND(UPPER([@[Record Type]])=""OUTPUT"",[@ID]<>""""),""D-""&" & processIdCell & "&""-""&[@ID],"""")"
    modProcessWorksheetColumns.SetFormula lo, "Design Version", _
        "=IF(UPPER([@[Record Type]])=""OUTPUT""," & processVersionCell & ","""")"

    For rowIndex = 1 To lo.ListRows.Count
        recordType = UCase$(Trim$(modProcessWorksheetColumns.Text(lo, rowIndex, "Record Type")))
        If recordType = "INPUT" Or recordType = "REQUIREMENT" Then
            modProcessWorksheetColumns.Cell(lo, rowIndex, "Requirement ID").Formula = modProcessWorksheetColumns.FormulaText(lo, "=[@ID]")
        ElseIf recordType <> "ALTERNATIVE" Then
            modProcessWorksheetColumns.Cell(lo, rowIndex, "Requirement ID").ClearContents
        End If
    Next rowIndex
CleanExit:
    On Error Resume Next
    Application.AutoCorrect.AutoFillFormulasInLists = priorAutoFill
    Application.EnableEvents = priorEvents
    mApplyingManagedColumns = False
    On Error GoTo 0
End Sub

Private Sub ApplyQtyModeValidation(ByVal lo As ListObject)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    With modProcessWorksheetColumns.Field(lo, "Qty Mode").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="FIXED,ACTUAL"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
        .ErrorTitle = "Choose quantity mode"
        .ErrorMessage = "Select FIXED or ACTUAL."
    End With
End Sub

Private Sub ApplyRecordTypeValidation(ByVal lo As ListObject)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    With modProcessWorksheetColumns.Field(lo, "Record Type").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="INPUT,OUTPUT,INSTRUCTION,ALTERNATIVE"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
        .ErrorTitle = "Choose a Process record type"
        .ErrorMessage = "Select INPUT, OUTPUT, INSTRUCTION, or ALTERNATIVE."
    End With
End Sub

Private Sub ApplyProcessWorksheetTextIdentityFormats(ByVal lo As ListObject)
    Dim pairNumber As Long

    If lo Is Nothing Then Exit Sub
    If Not lo.DataBodyRange Is Nothing Then
        modProcessWorksheetColumns.Field(lo, "ID").DataBodyRange.NumberFormat = "@"
        modProcessWorksheetColumns.Field(lo, "Output SKU").DataBodyRange.NumberFormat = "@"
        For pairNumber = FIRST_ALTERNATIVE_PAIR To AlternativePairCount(lo)
            modProcessWorksheetColumns.Field(lo, "Accepted SKU " & CStr(pairNumber)).DataBodyRange.NumberFormat = "@"
        Next pairNumber
    End If
    lo.Parent.Cells(lo.HeaderRowRange.Row - TABLE_HEADER_OFFSET + 1, 5).NumberFormat = "@"
End Sub

Private Sub ApplyProcessWorksheetUomValidation(ByVal lo As ListObject)
    Dim catalogText As String

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    catalogText = Replace$(modUomSettings.GetConfiguredUomsPackedText(), "|", ",")
    If catalogText = "" Then Exit Sub
    With modProcessWorksheetColumns.Field(lo, "UOM").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=catalogText
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
        .ErrorTitle = "Choose a Recipe UOM"
        .ErrorMessage = "Select a UOM from Settings > Recipe UOM Catalog."
    End With
End Sub

Private Sub EnsureWorksheetRowIds(ByVal lo As ListObject)
    Dim usedRowIds As Object
    Dim rowsNeedingIds As Object
    Dim rowIndex As Long
    Dim recordType As String
    Dim rowId As String

    Set usedRowIds = CreateObject("Scripting.Dictionary")
    usedRowIds.CompareMode = vbTextCompare
    Set rowsNeedingIds = CreateObject("Scripting.Dictionary")
    For rowIndex = 1 To lo.ListRows.Count
        recordType = UCase$(Trim$(modProcessWorksheetColumns.Text(lo, rowIndex, "Record Type")))
        rowId = UCase$(Trim$(modProcessWorksheetColumns.Text(lo, rowIndex, "ID")))
        Select Case recordType
            Case "INPUT", "REQUIREMENT", "OUTPUT", "INSTRUCTION"
                If mProduction.IsBase36Identifier(rowId) _
                   And Not usedRowIds.Exists(rowId) Then
                    usedRowIds.Add rowId, True
                Else
                    rowsNeedingIds.Add CStr(rowIndex), True
                End If
        End Select
    Next rowIndex
    For rowIndex = 1 To lo.ListRows.Count
        If rowsNeedingIds.Exists(CStr(rowIndex)) Then
            rowId = NextIdFromDictionary(usedRowIds)
            modProcessWorksheetColumns.Cell(lo, rowIndex, "ID").NumberFormat = "@"
            modProcessWorksheetColumns.Cell(lo, rowIndex, "ID").Value2 = rowId
            usedRowIds.Add rowId, True
        End If
    Next rowIndex
End Sub

Private Sub WriteWorksheetHeaders(ByVal ws As Worksheet, ByVal headerRow As Long, _
                                  ByVal pairCount As Long)
    Dim headers As Variant
    Dim index As Long
    Dim pairNumber As Long

    headers = Array("Record Type", "ID", "Name", "Qty", "Percent", _
                    "Basis Qty", "UOM", "Design ID", "Design Version", "Instruction", _
                    "Requirement ID", "Qty Mode", "Output SKU")
    For index = LBound(headers) To UBound(headers)
        ws.Cells(headerRow, index + 1).Value2 = headers(index)
    Next index
    For pairNumber = FIRST_ALTERNATIVE_PAIR To pairCount
        ws.Cells(headerRow, AlternativeItemColumnIndex(pairNumber)).Value2 = _
            "Acceptable Managed Item " & CStr(pairNumber)
        ws.Cells(headerRow, AlternativeSkuColumnIndex(pairNumber)).Value2 = _
            "Accepted SKU " & CStr(pairNumber)
    Next pairNumber
End Sub

Private Sub FormatProcessWorksheet(ByVal ws As Worksheet, ByVal lo As ListObject)
    Dim pairNumber As Long

    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 10
    ws.Columns("C").ColumnWidth = 24
    ws.Columns("D:F").ColumnWidth = 13
    ws.Columns("G").ColumnWidth = 10
    ws.Columns("H:I").ColumnWidth = 16
    ws.Columns("J").ColumnWidth = 48
    ws.Columns("K").ColumnWidth = 14
    modProcessWorksheetColumns.Field(lo, "Output SKU").Range.EntireColumn.Hidden = True
    For pairNumber = FIRST_ALTERNATIVE_PAIR To AlternativePairCount(lo)
        modProcessWorksheetColumns.Field(lo, "Acceptable Managed Item " & CStr(pairNumber)).Range.ColumnWidth = 28
        modProcessWorksheetColumns.Field(lo, "Accepted SKU " & CStr(pairNumber)).Range.EntireColumn.Hidden = True
    Next pairNumber
    modProcessWorksheetColumns.Field(lo, "Qty").DataBodyRange.NumberFormat = "0.########"
    modProcessWorksheetColumns.Field(lo, "Percent").DataBodyRange.NumberFormat = "0.0\%"
    modProcessWorksheetColumns.Field(lo, "Basis Qty").DataBodyRange.NumberFormat = "0.########"
    lo.Range.VerticalAlignment = xlTop
End Sub

Private Function EnsureProcessEditorSheet(ByVal wb As Workbook) As Worksheet
    On Error Resume Next
    Set EnsureProcessEditorSheet = wb.Worksheets(EDITOR_SHEET)
    On Error GoTo 0
    If EnsureProcessEditorSheet Is Nothing Then
        Set EnsureProcessEditorSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureProcessEditorSheet.Name = EDITOR_SHEET
    End If
End Function

Private Function NextProcessTableTopRow(ByVal ws As Worksheet) As Long
    Dim lo As ListObject
    Dim nextRow As Long

    nextRow = FIRST_TABLE_TOP_ROW
    If ws Is Nothing Then
        NextProcessTableTopRow = nextRow
        Exit Function
    End If
    For Each lo In ws.ListObjects
        If IsInvSysProcessTable(lo) Then
            If lo.Range.Row + lo.Range.Rows.Count + TABLE_GAP_ROWS > nextRow Then _
                nextRow = lo.Range.Row + lo.Range.Rows.Count + TABLE_GAP_ROWS
        End If
    Next lo
    NextProcessTableTopRow = nextRow
End Function

Private Function ProcessMetadataValue(ByVal lo As ListObject, _
                                      ByVal rowOffset As Long, _
                                      ByVal columnIndex As Long) As String
    Dim topRow As Long

    If lo Is Nothing Then Exit Function
    topRow = lo.HeaderRowRange.Row - TABLE_HEADER_OFFSET
    ProcessMetadataValue = CellText(lo.Parent.Cells(topRow + rowOffset, columnIndex).Value2)
End Function

Private Sub ClearProcessTableMetadata(ByVal lo As ListObject)
    Dim topRow As Long

    If lo Is Nothing Then Exit Sub
    topRow = lo.HeaderRowRange.Row - TABLE_HEADER_OFFSET
    lo.Parent.Range(lo.Parent.Cells(topRow, 1), _
                    lo.Parent.Cells(lo.HeaderRowRange.Row - 1, lo.Range.Columns.Count)).Clear
End Sub

Private Function IsInvSysProcessTable(ByVal lo As ListObject) As Boolean
    If lo Is Nothing Then Exit Function
    IsInvSysProcessTable = _
        (LCase$(Left$(lo.Name, Len(TABLE_PREFIX))) = LCase$(TABLE_PREFIX))
End Function

Private Function GeneratedOutputDesignId(ByVal processId As String, _
                                         ByVal outputId As String) As String
    processId = UCase$(Trim$(processId))
    outputId = UCase$(Trim$(outputId))
    If processId = "" Or outputId = "" Then Exit Function
    GeneratedOutputDesignId = "D-" & processId & "-" & outputId
End Function

Private Function BuildManagedItemNameBySku() As Object
    Dim pickerRows As Variant
    Dim result As Object
    Dim rowIndex As Long
    Dim sku As String
    Dim itemName As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    pickerRows = mProduction.LoadProductionInventoryPickerItems("")
    If IsEmpty(pickerRows) Or Not IsArray(pickerRows) Then
        Set BuildManagedItemNameBySku = result
        Exit Function
    End If
    For rowIndex = LBound(pickerRows, 1) To UBound(pickerRows, 1)
        sku = Trim$(CellText(pickerRows(rowIndex, 7)))
        itemName = Trim$(CellText(pickerRows(rowIndex, 2)))
        If sku <> "" And itemName <> "" Then result(sku) = itemName
    Next rowIndex
    Set BuildManagedItemNameBySku = result
End Function

Private Function BuildUniqueProcessTableName(ByVal wb As Workbook, _
                                             ByVal processId As String) As String
    Dim token As String
    Dim candidate As String
    Dim existing As ListObject

    token = UCase$(Replace$(mProduction.CreateProductionGuid(), "-", ""))
    token = Left$(token, 8)
    candidate = TABLE_PREFIX & processId & "_" & token
    Set existing = FindProcessTableByName(wb, candidate)
    Do While Not existing Is Nothing
        token = UCase$(Replace$(mProduction.CreateProductionGuid(), "-", ""))
        candidate = TABLE_PREFIX & processId & "_" & Left$(token, 8)
        Set existing = FindProcessTableByName(wb, candidate)
    Loop
    BuildUniqueProcessTableName = candidate
End Function

Private Function FindProcessTableByName(ByVal wb As Workbook, _
                                        ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Or Trim$(tableName) = "" Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindProcessTableByName = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindProcessTableByName Is Nothing Then Exit Function
    Next ws
End Function

Private Function ProcessIdFromTableName(ByVal tableName As String) As String
    Dim suffix As String

    If Left$(tableName, Len(TABLE_PREFIX)) <> TABLE_PREFIX Then Exit Function
    suffix = Mid$(tableName, Len(TABLE_PREFIX) + 1)
    ProcessIdFromTableName = UCase$(Left$(suffix, 3))
End Function

Private Function AlternativeItemColumnIndex(ByVal pairNumber As Long) As Long
    AlternativeItemColumnIndex = INITIAL_ACCEPTABLE_ITEM_COLUMN + ((pairNumber - 1) * 2)
End Function

Private Function AlternativeSkuColumnIndex(ByVal pairNumber As Long) As Long
    AlternativeSkuColumnIndex = INITIAL_ACCEPTED_SKU_COLUMN + ((pairNumber - 1) * 2)
End Function

Private Function AlternativePairCount(ByVal lo As ListObject) As Long
    Dim pairNumber As Long
    Dim col As ListColumn

    If lo Is Nothing Then Exit Function
    For pairNumber = FIRST_ALTERNATIVE_PAIR To 99
        On Error Resume Next
        Err.Clear
        Set col = Nothing
        Set col = modProcessWorksheetColumns.Field(lo, "Acceptable Managed Item " & CStr(pairNumber))
        On Error GoTo 0
        If col Is Nothing Then Exit For
        AlternativePairCount = pairNumber
    Next pairNumber
End Function

Private Function WorksheetAlternativePairCountForRows(ByVal rows As Collection) As Long
    Dim rowRecord As Object
    Dim pairNumber As Long

    WorksheetAlternativePairCountForRows = DEFAULT_ALTERNATIVE_PAIRS
    For Each rowRecord In rows
        For pairNumber = DEFAULT_ALTERNATIVE_PAIRS + 1 To 99
            If rowRecord.Exists("AcceptedSku" & CStr(pairNumber)) Then
                If Trim$(DictionaryText(rowRecord, "AcceptedSku" & CStr(pairNumber))) <> "" Then _
                    WorksheetAlternativePairCountForRows = pairNumber
            End If
        Next pairNumber
    Next rowRecord
End Function

Private Function DictionaryText(ByVal record As Object, ByVal keyName As String) As String
    If record Is Nothing Then Exit Function
    If Not record.Exists(keyName) Then Exit Function
    DictionaryText = CellText(record(keyName))
End Function

Private Function IsConfiguredProcessUom(ByVal uom As String) As Boolean
    Dim packed As String
    Dim values As Variant
    Dim idx As Long

    uom = UCase$(Trim$(uom))
    If uom = "" Then Exit Function
    packed = modUomSettings.GetConfiguredUomsPackedText()
    values = Split(packed, "|")
    For idx = LBound(values) To UBound(values)
        If StrComp(Trim$(CStr(values(idx))), uom, vbTextCompare) = 0 Then
            IsConfiguredProcessUom = True
            Exit Function
        End If
    Next idx
End Function

Private Sub AppendWorksheetAlternativeRecords(ByVal lo As ListObject, _
                                              ByVal rowIndex As Long, _
                                              ByVal requirementId As String, _
                                              ByVal records As Collection)
    Dim pairNumber As Long
    Dim acceptedSku As String
    Dim acceptableItem As String
    Dim record As Object

    For pairNumber = FIRST_ALTERNATIVE_PAIR To AlternativePairCount(lo)
        acceptedSku = Trim$(modProcessWorksheetColumns.Text(lo, rowIndex, _
            "Accepted SKU " & CStr(pairNumber)))
        acceptableItem = Trim$(modProcessWorksheetColumns.Text(lo, rowIndex, _
            "Acceptable Managed Item " & CStr(pairNumber)))
        If acceptedSku <> "" Then
            Set record = NewWorksheetRecord("ALTERNATIVE")
            record("RequirementId") = requirementId
            record("ITEM_CODE") = acceptedSku
            If acceptableItem <> "" Then record("ItemName") = acceptableItem
            records.Add record
        End If
    Next pairNumber
End Sub

Private Function WorksheetRowHasBusinessData(ByVal lo As ListObject, _
                                             ByVal rowIndex As Long) As Boolean
    Dim pairNumber As Long

    WorksheetRowHasBusinessData = _
        (Trim$(modProcessWorksheetColumns.Text(lo, rowIndex, "Name")) <> "") Or _
        (Trim$(modProcessWorksheetColumns.Text(lo, rowIndex, "Qty")) <> "") Or _
        (Trim$(modProcessWorksheetColumns.Text(lo, rowIndex, "UOM")) <> "") Or _
        (Trim$(modProcessWorksheetColumns.Text(lo, rowIndex, "Instruction")) <> "")
    If WorksheetRowHasBusinessData Then Exit Function
    For pairNumber = FIRST_ALTERNATIVE_PAIR To AlternativePairCount(lo)
        If Trim$(modProcessWorksheetColumns.Text(lo, rowIndex, _
                "Acceptable Managed Item " & CStr(pairNumber))) <> "" _
           Or Trim$(modProcessWorksheetColumns.Text(lo, rowIndex, _
                "Accepted SKU " & CStr(pairNumber))) <> "" Then
            WorksheetRowHasBusinessData = True
            Exit Function
        End If
    Next pairNumber
End Function

Private Function TryPositiveWorksheetNumber(ByVal valueText As String, _
                                            ByRef numberValue As Double) As Boolean
    numberValue = 0
    If Trim$(valueText) = "" Or Not IsNumeric(valueText) Then Exit Function
    numberValue = CDbl(valueText)
    TryPositiveWorksheetNumber = (numberValue > 0)
End Function

Private Function ResolveWorksheetRowId(ByVal proposedId As String, _
                                       ByVal used As Object) As String
    proposedId = UCase$(Trim$(proposedId))
    If mProduction.IsBase36Identifier(proposedId) And Not used.Exists(proposedId) Then
        ResolveWorksheetRowId = proposedId
    Else
        ResolveWorksheetRowId = NextIdFromDictionary(used)
    End If
End Function

Private Function NextIdFromDictionary(ByVal used As Object) As String
    Dim keys As Variant

    If used Is Nothing Or used.Count = 0 Then
        keys = Array("")
    Else
        keys = used.Keys
    End If
    NextIdFromDictionary = mProduction.NextBase36Identifier(keys)
End Function

Private Function CellText(ByVal valueIn As Variant) As String
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    CellText = CStr(valueIn)
End Function
