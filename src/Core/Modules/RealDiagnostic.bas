Attribute VB_Name = "RealDiagnostic"
Option Explicit

'===================================================================================
' Module: RealDiagnostic
' Purpose: Find why Worksheet_SelectionChange in ReceivedTally.cls isn't firing
'          when binding looks correct
'===================================================================================

Public Sub FullDiagnostic()
    Debug.Print "=== REAL DIAGNOSTIC: CHECKING WHY EVENTS AREN'T FIRING ==="
    Debug.Print ""
    
    Call DiagApplicationState
    Debug.Print ""
    
    Call DiagReceivedTallySheet
    Debug.Print ""
    
    Call DiagEventHandlerExistence
    Debug.Print ""
    
    Call DiagListObjectState
    Debug.Print ""
    
    Call DiagMissingCode
    Debug.Print ""
    
    Debug.Print "=== END DIAGNOSTIC ==="
End Sub

'--- CHECK 1: Application State
Public Sub DiagApplicationState()
    Debug.Print "[1] APPLICATION STATE"
    Debug.Print "  Application.EnableEvents: " & Application.EnableEvents
    Debug.Print "  Application.ScreenUpdating: " & Application.ScreenUpdating
    Debug.Print "  ActiveSheet.Name: " & ActiveSheet.Name
    Debug.Print "  ActiveSheet.CodeName: " & ActiveSheet.CodeName
    
    If Application.EnableEvents = False Then
        Debug.Print "  WARNING: Application.EnableEvents is FALSE!"
        Debug.Print "          Events will NOT fire until this is set to True"
    End If
End Sub

'--- CHECK 2: ReceivedTally Sheet Exists and Accessible
Public Sub DiagReceivedTallySheet()
    Debug.Print "[2] RECEIVEDTALLY SHEET STATE"
    
    Dim wsRT As Worksheet
    On Error Resume Next
    Set wsRT = ActiveWorkbook.Sheets("ReceivedTally")
    On Error GoTo 0
    
    If wsRT Is Nothing Then
        Debug.Print "  ERROR: Cannot access ReceivedTally sheet by name!"
        Exit Sub
    End If
    
    Debug.Print "  Sheet found: " & wsRT.Name
    Debug.Print "  Sheet CodeName: " & wsRT.CodeName
    
    ' Check if sheet is visible
    Debug.Print "  Sheet Visible: " & (wsRT.Visible = xlSheetVisible)
    If wsRT.Visible <> xlSheetVisible Then
        Debug.Print "  WARNING: Sheet is hidden! Events won't trigger."
    End If
End Sub

'--- CHECK 3: Event Handler Code Exists in ReceivedTally Class
Public Sub DiagEventHandlerExistence()
    Debug.Print "[3] EVENT HANDLER CODE EXISTENCE"
    
    Dim vbp As Object, vbc As Object
    Set vbp = ActiveWorkbook.VBProject
    
    On Error Resume Next
    Set vbc = vbp.VBComponents("ReceivedTally")
    On Error GoTo 0
    
    If vbc Is Nothing Then
        Debug.Print "  ERROR: ReceivedTally class not found in VBProject!"
        Exit Sub
    End If
    
    Debug.Print "  ReceivedTally class exists: Yes"
    Debug.Print "  Class type: " & ComponentTypeName(vbc.Type)
    
    ' Try to find the Worksheet_SelectionChange code
    Dim hasSelectionChange As Boolean
    Dim hasWorksheetChange As Boolean
    
    On Error Resume Next
    Dim codeModule As Object
    Set codeModule = vbc.CodeModule
    
    If Not codeModule Is Nothing Then
        Dim i As Long
        For i = 1 To codeModule.CountOfLines
            Dim line As String
            line = codeModule.Lines(i, 1)
            If InStr(line, "Worksheet_SelectionChange") > 0 Then
                hasSelectionChange = True
            End If
            If InStr(line, "Worksheet_Change") > 0 Then
                hasWorksheetChange = True
            End If
        Next i
    End If
    On Error GoTo 0
    
    Debug.Print "  Worksheet_SelectionChange handler exists: " & hasSelectionChange
    Debug.Print "  Worksheet_Change handler exists: " & hasWorksheetChange
    
    If Not hasSelectionChange Then
        Debug.Print "  ERROR: Worksheet_SelectionChange code not found!"
        Debug.Print "         The class exists but has no event handler!"
    End If
End Sub

'--- CHECK 4: ListObject ReceivedTally Exists
Public Sub DiagListObjectState()
    Debug.Print "[4] RECEIVEDTALLY TABLE (LISTOBJECT) STATE"
    
    Dim wsRT As Worksheet
    Set wsRT = ActiveWorkbook.Sheets("ReceivedTally")
    
    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsRT.ListObjects("ReceivedTally")
    On Error GoTo 0
    
    If lo Is Nothing Then
        Debug.Print "  ERROR: ReceivedTally table not found!"
        Debug.Print "         Cannot trigger ITEMS column events without it"
        Exit Sub
    End If
    
    Debug.Print "  ReceivedTally table exists: Yes"
    Debug.Print "  Table name: " & lo.Name
    Debug.Print "  Table range: " & lo.Range.Address
    
    ' Check DataBodyRange
    If lo.DataBodyRange Is Nothing Then
        Debug.Print "  WARNING: Table has no DataBodyRange (no data rows)!"
    Else
        Debug.Print "  DataBodyRange exists: " & lo.DataBodyRange.Address
        Debug.Print "  Number of data rows: " & lo.DataBodyRange.Rows.Count
    End If
    
    ' Check ITEMS column
    Dim itemsCol As ListColumn
    On Error Resume Next
    Set itemsCol = lo.ListColumns("ITEMS")
    On Error GoTo 0
    
    If itemsCol Is Nothing Then
        Debug.Print "  ERROR: ITEMS column not found in table!"
        Debug.Print "         Check column name (exact spelling)."
        Exit Sub
    End If
    
    Debug.Print "  ITEMS column exists: Yes"
    Debug.Print "  ITEMS column number: " & itemsCol.Index
    Debug.Print "  ITEMS column range: " & itemsCol.Range.Address
End Sub

'--- CHECK 5: Code That Might Be Preventing Events
Public Sub DiagMissingCode()
    Debug.Print "[5] POTENTIAL CODE BLOCKERS"
    
    ' Check if gSelectedCell global exists
    On Error Resume Next
    Dim dummy As Range
    Set dummy = gSelectedCell
    Dim globalExists As Boolean
    globalExists = (Err.Number = 0)
    On Error GoTo 0
    
    Debug.Print "  Global gSelectedCell exists: " & globalExists
    
    ' Check if modTS_Received exists
    Dim vbp As Object
    Set vbp = ActiveWorkbook.VBProject
    
    Dim modComp As Object
    Dim modExists As Boolean
    On Error Resume Next
    Set modComp = vbp.VBComponents("modTS_Received")
    modExists = (modComp Is Nothing = False)
    On Error GoTo 0
    
    Debug.Print "  modTS_Received module exists: " & modExists
    
    ' Check cDynItemSearch
    Dim dynExists As Boolean
    On Error Resume Next
    Set modComp = vbp.VBComponents("cDynItemSearch")
    dynExists = (modComp Is Nothing = False)
    On Error GoTo 0
    
    Debug.Print "  cDynItemSearch class exists: " & dynExists
End Sub

'--- MANUAL TEST: Force the Event
Public Sub TestManualSelection()
    Debug.Print ""
    Debug.Print "=== MANUAL SELECTION TEST ==="
    
    Dim wsRT As Worksheet
    Set wsRT = ActiveWorkbook.Sheets("ReceivedTally")
    
    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsRT.ListObjects("ReceivedTally")
    On Error GoTo 0
    
    If lo Is Nothing Then
        Debug.Print "ERROR: ReceivedTally table not found"
        Exit Sub
    End If
    
    If lo.DataBodyRange Is Nothing Then
        Debug.Print "ERROR: Table has no data rows"
        Exit Sub
    End If
    
    ' Find ITEMS column
    Dim itemsCol As ListColumn
    On Error Resume Next
    Set itemsCol = lo.ListColumns("ITEMS")
    On Error GoTo 0
    
    If itemsCol Is Nothing Then
        Debug.Print "ERROR: ITEMS column not found"
        Exit Sub
    End If
    
    ' Select first data cell in ITEMS column
    Dim testCell As Range
    Set testCell = lo.DataBodyRange.Cells(1, itemsCol.Index)
    
    Debug.Print "Selecting cell: " & testCell.Address
    testCell.Select
    
    Debug.Print "Waiting 2 seconds for events to fire..."
    Application.Wait (Now + TimeValue("0:00:02"))
    
    Debug.Print "Selection test complete. Check Immediate Window above for SelectionChange output."
End Sub

'--- HELPER: Component Type Name
Private Function ComponentTypeName(compType As Long) As String
    Select Case compType
        Case 1: ComponentTypeName = "Worksheet"
        Case 2: ComponentTypeName = "Module"
        Case 3: ComponentTypeName = "Class"
        Case 4: ComponentTypeName = "Form"
        Case Else: ComponentTypeName = "Unknown(" & compType & ")"
    End Select
End Function

'--- EMERGENCY: Force Re-enable Events
Public Sub EmergencyEnableEvents()
    Debug.Print ""
    Debug.Print "=== EMERGENCY: RE-ENABLING EVENTS ==="
    
    Dim oldSetting As Boolean
    oldSetting = Application.EnableEvents
    
    Application.EnableEvents = True
    
    Debug.Print "Application.EnableEvents was: " & oldSetting
    Debug.Print "Application.EnableEvents is now: " & Application.EnableEvents
    
    If oldSetting = False Then
        Debug.Print "Events were disabled! They should now be active."
        Debug.Print "Try clicking an ITEMS cell now."
    End If
End Sub
