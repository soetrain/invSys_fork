Attribute VB_Name = "modPhase6DemoData"
Option Explicit

Private Const DEMO_PREFIX As String = "DEMO-"
Private Const DEMO_RECIPE_ID As String = "DEMO-RECIPE-CLASSIC-CHAI"

Public Sub SeedActiveWorkbookDemoData()
    Dim wb As Workbook
    Dim report As String

    Set wb = ResolveDemoWorkbook()
    If wb Is Nothing Then
        MsgBox "Open a non-addin operational workbook before seeding demo data.", vbExclamation
        Exit Sub
    End If

    modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface wb, report
    modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface wb, report
    modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface wb, report

    SeedDemoInventory wb
    SeedDemoRecipes wb
    SeedDemoIngredientPalette wb

    MsgBox "Phase 6 demo data seeded into '" & wb.Name & "'.", vbInformation
End Sub

Public Sub ClearActiveWorkbookDemoData()
    Dim wb As Workbook

    Set wb = ResolveDemoWorkbook()
    If wb Is Nothing Then Exit Sub

    RemoveDemoRowsByColumn FindTableByName(wb, "invSys"), "ITEM_CODE"
    RemoveDemoRowsByColumn FindTableByName(wb, "Recipes"), "RECIPE_ID"
    RemoveDemoRowsByColumn FindTableByName(wb, "IngredientPalette"), "RECIPE_ID"
End Sub

Private Function ResolveDemoWorkbook() As Workbook
    If Not Application.ActiveWorkbook Is Nothing Then
        If Not Application.ActiveWorkbook.IsAddin Then
            Set ResolveDemoWorkbook = Application.ActiveWorkbook
        End If
    End If
End Function

Private Sub SeedDemoInventory(ByVal wb As Workbook)
    Dim lo As ListObject
    Dim items As Variant
    Dim i As Long

    Set lo = FindTableByName(wb, "invSys")
    If lo Is Nothing Then Exit Sub

    RemoveDemoRowsByColumn lo, "ITEM_CODE"

    items = DemoInventoryRows()
    For i = LBound(items) To UBound(items)
        AppendInventoryRow lo, items(i)
    Next i
End Sub

Private Sub SeedDemoRecipes(ByVal wb As Workbook)
    Dim lo As ListObject
    Dim rows As Variant
    Dim i As Long

    Set lo = FindTableByName(wb, "Recipes")
    If lo Is Nothing Then Exit Sub

    RemoveDemoRowsByColumn lo, "RECIPE_ID"

    rows = DemoRecipeRows()
    For i = LBound(rows) To UBound(rows)
        AppendRecipeRow lo, rows(i)
    Next i
End Sub

Private Sub SeedDemoIngredientPalette(ByVal wb As Workbook)
    Dim lo As ListObject
    Dim rows As Variant
    Dim i As Long

    Set lo = FindTableByName(wb, "IngredientPalette")
    If lo Is Nothing Then Exit Sub

    RemoveDemoRowsByColumn lo, "RECIPE_ID"

    rows = DemoIngredientPaletteRows()
    For i = LBound(rows) To UBound(rows)
        AppendIngredientPaletteRow lo, rows(i)
    Next i
End Sub

Private Sub RemoveDemoRowsByColumn(ByVal lo As ListObject, ByVal columnName As String)
    Dim colIdx As Long
    Dim r As Long
    Dim cellValue As String

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    colIdx = ColumnIndexDemo(lo, columnName)
    If colIdx = 0 Then Exit Sub

    For r = lo.ListRows.Count To 1 Step -1
        cellValue = NzStrDemo(lo.DataBodyRange.Cells(r, colIdx).Value)
        If Left$(cellValue, Len(DEMO_PREFIX)) = DEMO_PREFIX Then
            lo.ListRows(r).Delete
        End If
    Next r
End Sub

Private Sub AppendInventoryRow(ByVal lo As ListObject, ByVal rowData As Variant)
    Dim lr As ListRow

    Set lr = lo.ListRows.Add
    WriteField lr, "ROW", rowData(0)
    WriteField lr, "ITEM_CODE", rowData(1)
    WriteField lr, "ITEM", rowData(2)
    WriteField lr, "UOM", rowData(3)
    WriteField lr, "LOCATION", rowData(4)
    WriteField lr, "DESCRIPTION", rowData(5)
    WriteField lr, "VENDOR(s)", rowData(6)
    WriteField lr, "VENDOR_CODE", rowData(7)
    WriteField lr, "CATEGORY", rowData(8)
    WriteField lr, "RECEIVED", 0
    WriteField lr, "USED", 0
    WriteField lr, "MADE", 0
    WriteField lr, "SHIPMENTS", 0
    WriteField lr, "TOTAL INV", rowData(9)
    WriteField lr, "LAST EDITED", Now
    WriteField lr, "TOTAL INV LAST EDIT", Now
    WriteField lr, "TIMESTAMP", Now
End Sub

Private Sub AppendRecipeRow(ByVal lo As ListObject, ByVal rowData As Variant)
    Dim lr As ListRow

    Set lr = lo.ListRows.Add
    WriteField lr, "RECIPE", rowData(0)
    WriteField lr, "RECIPE_ID", rowData(1)
    WriteField lr, "DESCRIPTION", rowData(2)
    WriteField lr, "DEPARTMENT", rowData(3)
    WriteField lr, "PROCESS", rowData(4)
    WriteField lr, "DIAGRAM_ID", rowData(5)
    WriteField lr, "INPUT/OUTPUT", rowData(6)
    WriteField lr, "INGREDIENT", rowData(7)
    WriteField lr, "PERCENT", rowData(8)
    WriteField lr, "UOM", rowData(9)
    WriteField lr, "AMOUNT", rowData(10)
    WriteField lr, "RECIPE_LIST_ROW", rowData(11)
    WriteField lr, "INGREDIENT_ID", rowData(12)
    WriteField lr, "GUID", rowData(13)
End Sub

Private Sub AppendIngredientPaletteRow(ByVal lo As ListObject, ByVal rowData As Variant)
    Dim lr As ListRow

    Set lr = lo.ListRows.Add
    WriteField lr, "RECIPE_ID", rowData(0)
    WriteField lr, "INGREDIENT_ID", rowData(1)
    WriteField lr, "INPUT/OUTPUT", rowData(2)
    WriteField lr, "ITEM", rowData(3)
    WriteField lr, "PERCENT", rowData(4)
    WriteField lr, "UOM", rowData(5)
    WriteField lr, "AMOUNT", rowData(6)
    WriteField lr, "ROW", rowData(7)
    WriteField lr, "GUID", rowData(8)
End Sub

Private Sub WriteField(ByVal lr As ListRow, ByVal columnName As String, ByVal valueIn As Variant)
    Dim colIdx As Long

    If lr Is Nothing Then Exit Sub
    colIdx = ColumnIndexDemo(lr.Parent, columnName)
    If colIdx = 0 Then Exit Sub
    lr.Range.Cells(1, colIdx).Value = valueIn
End Sub

Private Function ColumnIndexDemo(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long

    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            ColumnIndexDemo = i
            Exit Function
        End If
    Next i
End Function

Private Function FindTableByName(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindTableByName = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindTableByName Is Nothing Then Exit Function
    Next ws
End Function

Private Function DemoInventoryRows() As Variant
    DemoInventoryRows = Array( _
        Array(9001, "DEMO-RAW-BLACK-TEA", "Black Tea", "lbs", "CLEARVIEW", "Loose black tea for brewing.", "Tea Importers", "TEA-001", "raw", 5000), _
        Array(9002, "DEMO-RAW-FILTERED-WATER", "Filtered Water", "lbs", "CLEARVIEW", "Filtered brewing water.", "Municipal Supply", "WATER-001", "raw", 20000), _
        Array(9003, "DEMO-RAW-CARDAMOM", "Cardamom (Decorticated)", "lbs", "CLEARVIEW", "Cardamom for chai blend.", "Spice House", "SPICE-001", "raw", 500), _
        Array(9004, "DEMO-RAW-BLACK-PEPPER", "Black Pepper (Whole)", "lbs", "CLEARVIEW", "Black pepper for chai blend.", "Spice House", "SPICE-002", "raw", 300), _
        Array(9005, "DEMO-RAW-NUTMEG", "Nutmeg (Ground)", "lbs", "CLEARVIEW", "Ground nutmeg for chai blend.", "Spice House", "SPICE-003", "raw", 250), _
        Array(9006, "DEMO-RAW-GINGER", "Ginger (Ground)", "lbs", "CLEARVIEW", "Ground ginger for chai blend.", "Spice House", "SPICE-004", "raw", 250), _
        Array(9007, "DEMO-RAW-CITRIC-ACID", "Citric Acid", "lbs", "CLEARVIEW", "Citric acid ingredient.", "Acid Supply", "ACID-001", "raw", 120), _
        Array(9008, "DEMO-RAW-CASSIA-OIL", "Cassia Oil 340139", "lbs", "CLEARVIEW", "Cassia oil for chai blend.", "Flavor House", "OIL-001", "raw", 80), _
        Array(9009, "DEMO-RAW-LEMON-OIL", "Lemon Oil (5x) 34013", "lbs", "CLEARVIEW", "Lemon oil for chai blend.", "Flavor House", "OIL-002", "raw", 80), _
        Array(9010, "DEMO-RAW-ORANGE-OIL", "Orange Oil (Cold Press)", "lbs", "CLEARVIEW", "Orange oil for chai blend.", "Flavor House", "OIL-003", "raw", 80), _
        Array(9011, "DEMO-RAW-SUGAR-WHITE", "Pure Cane Sugar White Granulated", "lbs", "CLEARVIEW", "White granulated cane sugar.", "Sugar Co", "SUGAR-001", "raw", 8000), _
        Array(9012, "DEMO-RAW-SUGAR-CLOUDY", "Pure Cane Sugar Cloudy White Granulated", "lbs", "CLEARVIEW", "Cloudy white granulated cane sugar.", "Sugar Co", "SUGAR-002", "raw", 6000), _
        Array(9013, "DEMO-WIP-BREWED-BLACK-TEA", "Brewed Black Tea", "lbs", "CLEARVIEW", "Intermediate brewed tea.", "Internal", "WIP-001", "wip", 1200), _
        Array(9014, "DEMO-WIP-CHAI-SPICE-BLEND", "Classic Chai Spice Blend", "lbs", "CLEARVIEW", "Intermediate chai spice blend.", "Internal", "WIP-002", "wip", 600), _
        Array(9015, "DEMO-RAW-BROWN-COLOR", "Brown Color 10.5g", "lbs", "CLEARVIEW", "Brown color ingredient.", "Color Lab", "COLOR-001", "raw", 100), _
        Array(9016, "DEMO-FG-CLASSIC-CHAI", "Black Scottie Chai Classic Concentrate", "gal", "CLEARVIEW", "Finished good concentrate for shipping.", "Internal", "FG-001", "shippable", 400), _
        Array(9017, "DEMO-FG-12PACK-CASE", "Classic Chai 12-Pack Case", "each", "CLEARVIEW", "Finished case pack.", "Internal", "FG-002", "shippable", 120), _
        Array(9018, "DEMO-FG-SAMPLE-BOX", "Black Scottie Sample Box", "each", "CLEARVIEW", "Sample assortment box.", "Internal", "FG-003", "shippable", 80))
End Function

Private Function DemoRecipeRows() As Variant
    DemoRecipeRows = Array( _
        Array("Black Scottie Chai Classic", DEMO_RECIPE_ID, "Classic chai concentrate demo recipe.", "PRODUCTION", "1-BREWING", "DGM-001", "USED", "Black Tea", 0.065, "lbs", 65, 1, "DEMO-ING-BLACK-TEA", "DEMO-RCP-001"), _
        Array("Black Scottie Chai Classic", DEMO_RECIPE_ID, "Classic chai concentrate demo recipe.", "PRODUCTION", "1-BREWING", "DGM-001", "USED", "Filtered Water", 1, "lbs", 1000, 2, "DEMO-ING-FILTERED-WATER", "DEMO-RCP-002"), _
        Array("Black Scottie Chai Classic", DEMO_RECIPE_ID, "Classic chai concentrate demo recipe.", "PRODUCTION", "1-BREWING", "DGM-001", "MADE", "Brewed Black Tea", 0.6834, "lbs", 683.4, 3, "DEMO-INT-BREWED-BLACK-TEA", "DEMO-RCP-003"), _
        Array("Black Scottie Chai Classic", DEMO_RECIPE_ID, "Classic chai concentrate demo recipe.", "PRODUCTION", "2-SPICE BLEND", "DGM-002", "USED", "Cardamom (Decorticated)", 0.4495, "lbs", 53.94, 4, "DEMO-ING-CARDAMOM", "DEMO-RCP-004"), _
        Array("Black Scottie Chai Classic", DEMO_RECIPE_ID, "Classic chai concentrate demo recipe.", "PRODUCTION", "2-SPICE BLEND", "DGM-002", "USED", "Black Pepper (Whole)", 0.1492, "lbs", 17.904, 5, "DEMO-ING-BLACK-PEPPER", "DEMO-RCP-005"), _
        Array("Black Scottie Chai Classic", DEMO_RECIPE_ID, "Classic chai concentrate demo recipe.", "PRODUCTION", "2-SPICE BLEND", "DGM-002", "USED", "Nutmeg (Ground)", 0.1071, "lbs", 12.852, 6, "DEMO-ING-NUTMEG", "DEMO-RCP-006"), _
        Array("Black Scottie Chai Classic", DEMO_RECIPE_ID, "Classic chai concentrate demo recipe.", "PRODUCTION", "2-SPICE BLEND", "DGM-002", "USED", "Ginger (Ground)", 0.0331, "lbs", 3.972, 7, "DEMO-ING-GINGER", "DEMO-RCP-007"), _
        Array("Black Scottie Chai Classic", DEMO_RECIPE_ID, "Classic chai concentrate demo recipe.", "PRODUCTION", "2-SPICE BLEND", "DGM-002", "USED", "Citric Acid", 0.0147, "lbs", 1.764, 8, "DEMO-ING-CITRIC-ACID", "DEMO-RCP-008"), _
        Array("Black Scottie Chai Classic", DEMO_RECIPE_ID, "Classic chai concentrate demo recipe.", "PRODUCTION", "2-SPICE BLEND", "DGM-002", "USED", "Cassia Oil 340139", 0.0146, "lbs", 1.752, 9, "DEMO-ING-CASSIA-OIL", "DEMO-RCP-009"), _
        Array("Black Scottie Chai Classic", DEMO_RECIPE_ID, "Classic chai concentrate demo recipe.", "PRODUCTION", "2-SPICE BLEND", "DGM-002", "USED", "Lemon Oil (5x) 34013", 0.0055, "lbs", 0.66, 10, "DEMO-ING-LEMON-OIL", "DEMO-RCP-010"), _
        Array("Black Scottie Chai Classic", DEMO_RECIPE_ID, "Classic chai concentrate demo recipe.", "PRODUCTION", "2-SPICE BLEND", "DGM-002", "USED", "Orange Oil (Cold Press)", 0.0024, "lbs", 0.288, 11, "DEMO-ING-ORANGE-OIL", "DEMO-RCP-011"), _
        Array("Black Scottie Chai Classic", DEMO_RECIPE_ID, "Classic chai concentrate demo recipe.", "PRODUCTION", "2-SPICE BLEND", "DGM-002", "MADE", "Classic Chai Spice Blend", 1, "lbs", 120, 12, "DEMO-INT-CHAI-SPICE-BLEND", "DEMO-RCP-012"), _
        Array("Black Scottie Chai Classic", DEMO_RECIPE_ID, "Classic chai concentrate demo recipe.", "PRODUCTION", "3-COOK CHAI", "DGM-003", "USED", "Brewed Black Tea", 0.6834, "lbs", 683.4, 13, "DEMO-INT-BREWED-BLACK-TEA", "DEMO-RCP-013"), _
        Array("Black Scottie Chai Classic", DEMO_RECIPE_ID, "Classic chai concentrate demo recipe.", "PRODUCTION", "3-COOK CHAI", "DGM-003", "USED", "Pure Cane Sugar", 0.187, "lbs", 187, 14, "DEMO-ING-PURE-CANE-SUGAR", "DEMO-RCP-014"), _
        Array("Black Scottie Chai Classic", DEMO_RECIPE_ID, "Classic chai concentrate demo recipe.", "PRODUCTION", "3-COOK CHAI", "DGM-003", "USED", "Brown Color 10.5g", 0.002484, "lbs", 2.484, 15, "DEMO-ING-BROWN-COLOR", "DEMO-RCP-015"), _
        Array("Black Scottie Chai Classic", DEMO_RECIPE_ID, "Classic chai concentrate demo recipe.", "PRODUCTION", "3-COOK CHAI", "DGM-003", "USED", "Classic Chai Spice Blend", 0.009515, "lbs", 9.515, 16, "DEMO-INT-CHAI-SPICE-BLEND", "DEMO-RCP-016"), _
        Array("Black Scottie Chai Classic", DEMO_RECIPE_ID, "Classic chai concentrate demo recipe.", "PRODUCTION", "3-COOK CHAI", "DGM-003", "MADE", "Black Scottie Chai Classic Concentrate", 1, "gal", 240, 17, "DEMO-FG-CLASSIC-CHAI", "DEMO-RCP-017"))
End Function

Private Function DemoIngredientPaletteRows() As Variant
    DemoIngredientPaletteRows = Array( _
        Array(DEMO_RECIPE_ID, "DEMO-ING-BLACK-TEA", "USED", "Black Tea", 0.065, "lbs", 65, 9001, "DEMO-PAL-001"), _
        Array(DEMO_RECIPE_ID, "DEMO-ING-FILTERED-WATER", "USED", "Filtered Water", 1, "lbs", 1000, 9002, "DEMO-PAL-002"), _
        Array(DEMO_RECIPE_ID, "DEMO-INT-BREWED-BLACK-TEA", "MADE", "Brewed Black Tea", 0.6834, "lbs", 683.4, 9013, "DEMO-PAL-003"), _
        Array(DEMO_RECIPE_ID, "DEMO-ING-CARDAMOM", "USED", "Cardamom (Decorticated)", 0.4495, "lbs", 53.94, 9003, "DEMO-PAL-004"), _
        Array(DEMO_RECIPE_ID, "DEMO-ING-BLACK-PEPPER", "USED", "Black Pepper (Whole)", 0.1492, "lbs", 17.904, 9004, "DEMO-PAL-005"), _
        Array(DEMO_RECIPE_ID, "DEMO-ING-NUTMEG", "USED", "Nutmeg (Ground)", 0.1071, "lbs", 12.852, 9005, "DEMO-PAL-006"), _
        Array(DEMO_RECIPE_ID, "DEMO-ING-GINGER", "USED", "Ginger (Ground)", 0.0331, "lbs", 3.972, 9006, "DEMO-PAL-007"), _
        Array(DEMO_RECIPE_ID, "DEMO-ING-CITRIC-ACID", "USED", "Citric Acid", 0.0147, "lbs", 1.764, 9007, "DEMO-PAL-008"), _
        Array(DEMO_RECIPE_ID, "DEMO-ING-CASSIA-OIL", "USED", "Cassia Oil 340139", 0.0146, "lbs", 1.752, 9008, "DEMO-PAL-009"), _
        Array(DEMO_RECIPE_ID, "DEMO-ING-LEMON-OIL", "USED", "Lemon Oil (5x) 34013", 0.0055, "lbs", 0.66, 9009, "DEMO-PAL-010"), _
        Array(DEMO_RECIPE_ID, "DEMO-ING-ORANGE-OIL", "USED", "Orange Oil (Cold Press)", 0.0024, "lbs", 0.288, 9010, "DEMO-PAL-011"), _
        Array(DEMO_RECIPE_ID, "DEMO-INT-CHAI-SPICE-BLEND", "MADE", "Classic Chai Spice Blend", 1, "lbs", 120, 9014, "DEMO-PAL-012"), _
        Array(DEMO_RECIPE_ID, "DEMO-ING-PURE-CANE-SUGAR", "USED", "Pure Cane Sugar White Granulated", 0.187, "lbs", 187, 9011, "DEMO-PAL-013"), _
        Array(DEMO_RECIPE_ID, "DEMO-ING-PURE-CANE-SUGAR", "USED", "Pure Cane Sugar Cloudy White Granulated", 0.187, "lbs", 187, 9012, "DEMO-PAL-014"), _
        Array(DEMO_RECIPE_ID, "DEMO-ING-BROWN-COLOR", "USED", "Brown Color 10.5g", 0.002484, "lbs", 2.484, 9015, "DEMO-PAL-015"), _
        Array(DEMO_RECIPE_ID, "DEMO-FG-CLASSIC-CHAI", "MADE", "Black Scottie Chai Classic Concentrate", 1, "gal", 240, 9016, "DEMO-PAL-016"))
End Function

Private Function NzStrDemo(ByVal valueIn As Variant) As String
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then
        NzStrDemo = ""
    Else
        NzStrDemo = CStr(valueIn)
    End If
End Function
