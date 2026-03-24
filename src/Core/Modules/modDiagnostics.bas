Attribute VB_Name = "modDiagnostics"

Option Explicit

Public Sub DumpActiveWorkbookToImmediate()
    DumpWorkbookToImmediateCore Application.ActiveWorkbook, 0
End Sub

Public Sub DumpLikelyRuntimeWorkbooksToImmediate()
    Dim wb As Workbook

    Debug.Print String$(80, "=")
    Debug.Print "invSys Runtime Workbook Dump", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print String$(80, "=")

    For Each wb In Application.Workbooks
        If ShouldDumpWorkbook(wb) Then
            DumpWorkbookToImmediateCore wb, 0
        End If
    Next wb
End Sub

Public Sub DumpWorkbookByNameToImmediate(Optional ByVal workbookName As String = "", Optional ByVal maxRows As Long = 0)
    Dim wb As Workbook

    If Trim$(workbookName) = "" Then
        Set wb = Application.ActiveWorkbook
    Else
        Set wb = FindWorkbookByNameDiagnostics(workbookName)
    End If

    If wb Is Nothing Then
        Debug.Print "Workbook not found: " & workbookName
        Exit Sub
    End If

    DumpWorkbookToImmediateCore wb, maxRows
End Sub

Public Sub DumpAllOpenWorkbooksToImmediate(Optional ByVal maxRows As Long = 0)
    Dim wb As Workbook

    Debug.Print String$(80, "=")
    Debug.Print "All Open Workbook Dump", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print String$(80, "=")

    For Each wb In Application.Workbooks
        DumpWorkbookToImmediateCore wb, maxRows
    Next wb
End Sub

Public Function DumpLikelyRuntimeWorkbooksToFile(Optional ByVal outputPath As String = "", Optional ByVal maxRows As Long = 0) As String
    Dim wb As Workbook
    Dim lines As Collection

    Set lines = New Collection
    AddDumpLine lines, String$(80, "=")
    AddDumpLine lines, "invSys Runtime Workbook Dump " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    AddDumpLine lines, String$(80, "=")

    For Each wb In Application.Workbooks
        If ShouldDumpWorkbook(wb) Then
            AppendWorkbookDumpLines wb, maxRows, lines
        End If
    Next wb

    DumpLikelyRuntimeWorkbooksToFile = WriteDumpLinesToFile(lines, outputPath, "invSys_runtime_dump")
    NotifyDumpFileWritten DumpLikelyRuntimeWorkbooksToFile
End Function

Public Function DumpWorkbookByNameToFile(Optional ByVal workbookName As String = "", Optional ByVal outputPath As String = "", Optional ByVal maxRows As Long = 0) As String
    Dim wb As Workbook
    Dim lines As Collection

    If Trim$(workbookName) = "" Then
        Set wb = Application.ActiveWorkbook
    Else
        Set wb = FindWorkbookByNameDiagnostics(workbookName)
    End If

    Set lines = New Collection
    If wb Is Nothing Then
        AddDumpLine lines, "Workbook not found: " & workbookName
    Else
        AddDumpLine lines, String$(80, "=")
        AddDumpLine lines, "Workbook Dump " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
        AddDumpLine lines, String$(80, "=")
        AppendWorkbookDumpLines wb, maxRows, lines
    End If

    DumpWorkbookByNameToFile = WriteDumpLinesToFile(lines, outputPath, "invSys_workbook_dump")
    NotifyDumpFileWritten DumpWorkbookByNameToFile
End Function

Public Function DumpAllOpenWorkbooksToFile(Optional ByVal outputPath As String = "", Optional ByVal maxRows As Long = 0) As String
    Dim wb As Workbook
    Dim lines As Collection

    Set lines = New Collection
    AddDumpLine lines, String$(80, "=")
    AddDumpLine lines, "All Open Workbook Dump " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    AddDumpLine lines, String$(80, "=")

    For Each wb In Application.Workbooks
        AppendWorkbookDumpLines wb, maxRows, lines
    Next wb

    DumpAllOpenWorkbooksToFile = WriteDumpLinesToFile(lines, outputPath, "invSys_all_workbooks_dump")
    NotifyDumpFileWritten DumpAllOpenWorkbooksToFile
End Function

Private Sub DumpWorkbookToImmediateCore(ByVal wb As Workbook, ByVal maxRows As Long)
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Sub

    Debug.Print String$(80, "-")
    Debug.Print "Workbook: " & wb.Name
    Debug.Print "  FullName: " & SafeWorkbookFullName(wb)
    Debug.Print "  IsAddin=" & CStr(wb.IsAddin) & _
                "; ReadOnly=" & CStr(wb.ReadOnly) & _
                "; Saved=" & CStr(wb.Saved) & _
                "; Sheets=" & CStr(wb.Worksheets.Count)

    For Each ws In wb.Worksheets
        DumpWorksheetToImmediate ws, maxRows
    Next ws
End Sub

Private Sub AppendWorkbookDumpLines(ByVal wb As Workbook, ByVal maxRows As Long, ByVal lines As Collection)
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Sub

    AddDumpLine lines, String$(80, "-")
    AddDumpLine lines, "Workbook: " & wb.Name
    AddDumpLine lines, "  FullName: " & SafeWorkbookFullName(wb)
    AddDumpLine lines, "  IsAddin=" & CStr(wb.IsAddin) & _
                       "; ReadOnly=" & CStr(wb.ReadOnly) & _
                       "; Saved=" & CStr(wb.Saved) & _
                       "; Sheets=" & CStr(wb.Worksheets.Count)

    For Each ws In wb.Worksheets
        AppendWorksheetDumpLines ws, maxRows, lines
    Next ws
End Sub

Private Sub DumpWorksheetToImmediate(ByVal ws As Worksheet, ByVal maxRows As Long)
    Dim lo As ListObject

    If ws Is Nothing Then Exit Sub

    Debug.Print "  Sheet: " & ws.Name & _
                "; Visible=" & CStr(ws.Visible = xlSheetVisible) & _
                "; Protected=" & CStr(ws.ProtectContents) & _
                "; Tables=" & CStr(ws.ListObjects.Count)

    If ws.ListObjects.Count = 0 Then
        Debug.Print "    UsedRange: " & SafeRangeAddress(ws.UsedRange) & _
                    "; CellsWithValues=" & CStr(SafeCountA(ws))
        Exit Sub
    End If

    For Each lo In ws.ListObjects
        DumpListObjectToImmediate lo, maxRows
    Next lo
End Sub

Private Sub AppendWorksheetDumpLines(ByVal ws As Worksheet, ByVal maxRows As Long, ByVal lines As Collection)
    Dim lo As ListObject

    If ws Is Nothing Then Exit Sub

    AddDumpLine lines, "  Sheet: " & ws.Name & _
                       "; Visible=" & CStr(ws.Visible = xlSheetVisible) & _
                       "; Protected=" & CStr(ws.ProtectContents) & _
                       "; Tables=" & CStr(ws.ListObjects.Count)

    If ws.ListObjects.Count = 0 Then
        AddDumpLine lines, "    UsedRange: " & SafeRangeAddress(ws.UsedRange) & _
                           "; CellsWithValues=" & CStr(SafeCountA(ws))
        Exit Sub
    End If

    For Each lo In ws.ListObjects
        AppendListObjectDumpLines lo, maxRows, lines
    Next lo
End Sub

Private Sub DumpListObjectToImmediate(ByVal lo As ListObject, ByVal maxRows As Long)
    Dim rowCount As Long
    Dim colCount As Long
    Dim rowsToPrint As Long
    Dim r As Long
    Dim c As Long
    Dim lineOut As String

    If lo Is Nothing Then Exit Sub

    rowCount = ListObjectRowCountDiagnostics(lo)
    colCount = lo.ListColumns.Count
    rowsToPrint = rowCount
    If maxRows > 0 And rowsToPrint > maxRows Then rowsToPrint = maxRows

    Debug.Print "    Table: " & lo.Name & _
                "; Range=" & SafeRangeAddress(lo.Range) & _
                "; Rows=" & CStr(rowCount) & _
                "; Cols=" & CStr(colCount)

    lineOut = "      Headers: "
    For c = 1 To colCount
        If c > 1 Then lineOut = lineOut & " | "
        lineOut = lineOut & lo.ListColumns(c).Name
    Next c
    Debug.Print lineOut

    If rowCount = 0 Then
        Debug.Print "      <no data rows>"
        Exit Sub
    End If

    For r = 1 To rowsToPrint
        lineOut = "      Row " & CStr(r) & ": "
        For c = 1 To colCount
            If c > 1 Then lineOut = lineOut & " | "
            lineOut = lineOut & lo.ListColumns(c).Name & "=" & CellValueTextDiagnostics(lo.DataBodyRange.Cells(r, c).Value)
        Next c
        Debug.Print lineOut
    Next r

    If rowsToPrint < rowCount Then
        Debug.Print "      ... " & CStr(rowCount - rowsToPrint) & " more rows omitted"
    End If
End Sub

Private Sub AppendListObjectDumpLines(ByVal lo As ListObject, ByVal maxRows As Long, ByVal lines As Collection)
    Dim rowCount As Long
    Dim colCount As Long
    Dim rowsToPrint As Long
    Dim r As Long
    Dim c As Long
    Dim lineOut As String

    If lo Is Nothing Then Exit Sub

    rowCount = ListObjectRowCountDiagnostics(lo)
    colCount = lo.ListColumns.Count
    rowsToPrint = rowCount
    If maxRows > 0 And rowsToPrint > maxRows Then rowsToPrint = maxRows

    AddDumpLine lines, "    Table: " & lo.Name & _
                       "; Range=" & SafeRangeAddress(lo.Range) & _
                       "; Rows=" & CStr(rowCount) & _
                       "; Cols=" & CStr(colCount)

    lineOut = "      Headers: "
    For c = 1 To colCount
        If c > 1 Then lineOut = lineOut & " | "
        lineOut = lineOut & lo.ListColumns(c).Name
    Next c
    AddDumpLine lines, lineOut

    If rowCount = 0 Then
        AddDumpLine lines, "      <no data rows>"
        Exit Sub
    End If

    For r = 1 To rowsToPrint
        lineOut = "      Row " & CStr(r) & ": "
        For c = 1 To colCount
            If c > 1 Then lineOut = lineOut & " | "
            lineOut = lineOut & lo.ListColumns(c).Name & "=" & CellValueTextDiagnostics(lo.DataBodyRange.Cells(r, c).Value)
        Next c
        AddDumpLine lines, lineOut
    Next r

    If rowsToPrint < rowCount Then
        AddDumpLine lines, "      ... " & CStr(rowCount - rowsToPrint) & " more rows omitted"
    End If
End Sub

Private Function ShouldDumpWorkbook(ByVal wb As Workbook) As Boolean
    Dim wbName As String

    If wb Is Nothing Then Exit Function
    If wb.IsAddin Then Exit Function

    wbName = LCase$(Trim$(wb.Name))
    If wbName = "" Then Exit Function

    If wbName Like "wh*.invsys.*.xls*" _
       Or wbName Like "invsys.inbox.*.xls*" _
       Or wbName Like "*.outbox.events.xls*" _
       Or wbName Like "*.snapshot.inventory.xls*" Then
        ShouldDumpWorkbook = True
        Exit Function
    End If

    If wb Is Application.ActiveWorkbook Then ShouldDumpWorkbook = True
End Function

Private Function FindWorkbookByNameDiagnostics(ByVal workbookName As String) As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.Name, workbookName, vbTextCompare) = 0 Then
            Set FindWorkbookByNameDiagnostics = wb
            Exit Function
        End If
    Next wb
End Function

Private Sub AddDumpLine(ByVal lines As Collection, ByVal textOut As String)
    If lines Is Nothing Then Exit Sub
    lines.Add textOut
End Sub

Private Function WriteDumpLinesToFile(ByVal lines As Collection, ByVal outputPath As String, ByVal filePrefix As String) As String
    Dim fileNo As Integer
    Dim lineItem As Variant

    outputPath = ResolveDumpOutputPath(outputPath, filePrefix)
    EnsureParentFolderForFileDiagnostics outputPath
    fileNo = FreeFile
    Open outputPath For Output As #fileNo
    For Each lineItem In lines
        Print #fileNo, CStr(lineItem)
    Next lineItem
    Close #fileNo
    WriteDumpLinesToFile = outputPath
End Function

Private Function ResolveDumpOutputPath(ByVal outputPath As String, ByVal filePrefix As String) As String
    If Trim$(outputPath) <> "" Then
        ResolveDumpOutputPath = outputPath
    Else
        ResolveDumpOutputPath = Environ$("TEMP") & "\" & filePrefix & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"
    End If
End Function

Private Sub EnsureParentFolderForFileDiagnostics(ByVal filePath As String)
    Dim parentPath As String

    parentPath = GetParentFolderDiagnostics(filePath)
    If parentPath = "" Then Exit Sub
    EnsureFolderExistsDiagnostics parentPath
End Sub

Private Function GetParentFolderDiagnostics(ByVal filePath As String) As String
    Dim sepPos As Long

    sepPos = InStrRev(Trim$(filePath), "\")
    If sepPos > 1 Then GetParentFolderDiagnostics = Left$(filePath, sepPos - 1)
End Function

Private Sub EnsureFolderExistsDiagnostics(ByVal folderPath As String)
    Dim parentPath As String
    Dim sepPos As Long

    folderPath = Trim$(folderPath)
    If folderPath = "" Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) > 0 Then Exit Sub

    sepPos = InStrRev(folderPath, "\")
    If sepPos > 1 Then
        parentPath = Left$(folderPath, sepPos - 1)
        If Right$(parentPath, 1) = ":" Then parentPath = parentPath & "\"
        If parentPath <> "" And Len(Dir$(parentPath, vbDirectory)) = 0 Then
            EnsureFolderExistsDiagnostics parentPath
        End If
    End If

    If Len(Dir$(folderPath, vbDirectory)) = 0 Then MkDir folderPath
End Sub

Private Sub NotifyDumpFileWritten(ByVal outputPath As String)
    If Trim$(outputPath) = "" Then Exit Sub
    Debug.Print "Workbook dump written to: " & outputPath
    MsgBox "Workbook dump written to:" & vbCrLf & outputPath, vbInformation
End Sub

Private Function ListObjectRowCountDiagnostics(ByVal lo As ListObject) As Long
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    ListObjectRowCountDiagnostics = lo.ListRows.Count
End Function

Private Function SafeRangeAddress(ByVal target As Range) As String
    On Error Resume Next
    If target Is Nothing Then
        SafeRangeAddress = "<none>"
    Else
        SafeRangeAddress = target.Address(False, False)
    End If
    On Error GoTo 0
End Function

Private Function SafeCountA(ByVal ws As Worksheet) As Long
    On Error Resume Next
    SafeCountA = Application.WorksheetFunction.CountA(ws.Cells)
    On Error GoTo 0
End Function

Private Function SafeWorkbookFullName(ByVal wb As Workbook) As String
    On Error Resume Next
    SafeWorkbookFullName = wb.FullName
    On Error GoTo 0
End Function

Private Function CellValueTextDiagnostics(ByVal valueIn As Variant) As String
    If IsError(valueIn) Then
        CellValueTextDiagnostics = "#ERROR"
    ElseIf IsEmpty(valueIn) Or IsNull(valueIn) Then
        CellValueTextDiagnostics = "<blank>"
    ElseIf IsDate(valueIn) Then
        CellValueTextDiagnostics = Format$(CDate(valueIn), "yyyy-mm-dd hh:nn:ss")
    Else
        CellValueTextDiagnostics = Replace$(Replace$(Trim$(CStr(valueIn)), vbCr, "\r"), vbLf, "\n")
        If CellValueTextDiagnostics = "" Then CellValueTextDiagnostics = "<blank>"
    End If
End Function

