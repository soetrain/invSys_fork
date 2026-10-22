VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "cInventoryCommandInputs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Transient Core-owned input binding. This object never crosses an XLAM boundary.
Private mWarehouse As String, mPath As String, mIdentity As String, mHash As String
Private mCatalog As cInventoryTableInput, mEntities As cInventoryTableInput
Private mStaging As cInventoryTableInput, mStage As ListObject
Private mCreatedStaging As cInventoryTableInput, mCreatedStage As ListObject
Private mCreatedBook As Workbook, mCreatedRange As String, mCreatedName As String
Private mSelectedSku As String, mPhase As String, mSaving As Boolean
Private mInitiallyBorrowed As Boolean
Private mCreate As cInventoryCreatePlan
Private mRetirement As cInventoryRetirementPlan
Private mQuantityKey As String, mQuantityBefore As Double, mQuantityTarget As Double
Private mEditReceipts As cInventoryTableInput, mEditReceiptAbsent As Boolean

Public Property Get SourcePath() As String
    SourcePath = mPath
End Property

Public Function CaptureSelection(ByVal operatorBook As Workbook) As Boolean
    Dim selected As Range, candidate As ListObject
    On Error Resume Next
    Set selected = Application.Selection
    On Error GoTo Invalid
    If Not selected Is Nothing Then
        If selected.Parent.Parent Is operatorBook Then
            For Each candidate In selected.Parent.ListObjects
                If Not Application.Intersect(selected, candidate.Range) Is Nothing Then
                    If Left$(candidate.Name, 17) = "invSys_Inventory_" Then
                        Set mStage = candidate
                        Set mStaging = New cInventoryTableInput
                        If Not mStaging.Capture(candidate, "", True) Then Exit Function
                    End If
                    Exit For
                End If
            Next candidate
        End If
    End If
    CaptureSelection = True
Invalid:
End Function

Public Function CaptureSource(ByVal warehouse As String) As Boolean
    Dim owner As cSnapshotInventorySource, book As Workbook, borrowed As Boolean
    On Error GoTo Invalid
    mWarehouse = warehouse
    mPath = modProcessor.ResolveInventoryWorkbookPathForAutomation(warehouse)
    If mPath = "" Then Exit Function
    mIdentity = modAggregationFiles.Identity(mPath): mHash = modTrainingWire.FileSha256(mPath)
    If mIdentity = "" Or mHash = "" Then Exit Function
    Set book = ReadSource(owner, borrowed)
    If book Is Nothing Then Exit Function
    If borrowed And (book.ReadOnly Or Not book.Saved) Then Exit Function
    mInitiallyBorrowed = borrowed
    Set mCatalog = New cInventoryTableInput: Set mEntities = New cInventoryTableInput
    If Not mCatalog.Capture(modExecutionTarget.Table(book, "tblSkuCatalog"), "SKU", borrowed) Then Exit Function
    If Not mEntities.Capture(modExecutionTarget.Table(book, "tblInventoryEntities"), "System_Key", borrowed) Then Exit Function
    Set book = Nothing: Set owner = Nothing
    CaptureSource = Unchanged()
Invalid:
End Function

Public Function Unchanged() As Boolean
    Dim owner As cSnapshotInventorySource, book As Workbook, borrowed As Boolean
    On Error GoTo Invalid
    If mCatalog Is Nothing Or mEntities Is Nothing Then Exit Function
    If Not mSaving Then
        If modAggregationFiles.Identity(mPath) <> mIdentity Or modTrainingWire.FileSha256(mPath) <> mHash Then Exit Function
    End If
    Set book = ReadSource(owner, borrowed)
    If book Is Nothing Then Exit Function
    If borrowed And book.ReadOnly Then Exit Function
    If mPhase = "" And Not book.Saved Then Exit Function
    If Not mCatalog.Matches(modExecutionTarget.Table(book, "tblSkuCatalog")) Then Exit Function
    If Not mEntities.Matches(modExecutionTarget.Table(book, "tblInventoryEntities"), _
                            mSelectedSku, mSelectedSku = "") Then Exit Function
    If Not EditReceiptsMatch(book) Then Exit Function
    If Not mStaging Is Nothing Then
        If Not mStaging.Matches(mStage) Then Exit Function
    End If
    If Not mCreatedStaging Is Nothing Then
        If Not (mCreatedStage.Parent.Parent Is mCreatedBook) Then Exit Function
        If StrComp(mCreatedStage.Name, mCreatedName, vbBinaryCompare) <> 0 Then Exit Function
        If mCreatedStage.Parent.Name <> "invSys Inventory Editor" Then Exit Function
        If mCreatedStage.Range.Address <> mCreatedRange Or mCreatedStage.ListRows.Count <> 20 Then Exit Function
        If Not mCreatedStaging.Matches(mCreatedStage) Then Exit Function
    End If
    Unchanged = True
Invalid:
End Function

Public Function BindSku(ByVal sku As String) As Boolean
    Dim owner As cSnapshotInventorySource, book As Workbook, borrowed As Boolean
    On Error GoTo Invalid
    If Not Unchanged() Or mPhase <> "" Then Exit Function
    Set book = ReadSource(owner, borrowed)
    If book Is Nothing Then Exit Function
    If sku <> "" Then
        If Not mEntities.Matches(modExecutionTarget.Table(book, "tblInventoryEntities"), sku) Then Exit Function
    End If
    mSelectedSku = sku: BindSku = True
Invalid:
End Function

Public Property Get StagingName() As String
    On Error GoTo Invalid
    If Not mStage Is Nothing Then StagingName = mStage.Name
Invalid:
End Property

Public Function PlanStageCell(ByVal row As Long, ByVal header As String, ByVal value As String) As Boolean
    If Not Unchanged() Or mStaging Is Nothing Then Exit Function
    Select Case UCase$(Trim$(header))
        Case "ITEM CODE", "UPLOAD STATUS", "UPLOAD RESULT"
            mStaging.PlanValue CStr(row), header, value
            PlanStageCell = True
    End Select
End Function

Public Function ConfirmStageCell() As Boolean
    If mStaging Is Nothing Then Exit Function
    If Not Unchanged() Then Exit Function
    ConfirmStageCell = mStaging.Confirm(mStage)
End Function

Public Function BindCreatedStage(ByVal operatorBook As Workbook, ByVal tableName As String) As Boolean
    Dim sheet As Worksheet, table As ListObject, headers As Variant, i As Long, cell As Range
    On Error GoTo Invalid
    If Not mCreatedStage Is Nothing Or Not Unchanged() Then Exit Function
    If Left$(tableName, 17) <> "invSys_Inventory_" Then Exit Function
    Set sheet = operatorBook.Worksheets("invSys Inventory Editor")
    Set table = sheet.ListObjects(tableName)
    If table Is mStage Then Exit Function
    headers = Split("Action|Item Code|Item Name|UOM|Qty Mode|Quantity|Default Location|Category|Description|Vendor(s)|Vendor Code|External Code|Picture Path/URL|Edit Reason|Upload Status|Upload Result", "|")
    If table.ListColumns.Count <> UBound(headers) + 1 Or table.ListRows.Count <> 20 Then Exit Function
    For i = 0 To UBound(headers)
        If StrComp(table.ListColumns(i + 1).Name, CStr(headers(i)), vbBinaryCompare) <> 0 Then Exit Function
    Next i
    For Each cell In table.DataBodyRange.Cells
        If cell.HasFormula Or Not IsEmpty(cell.Value2) Then Exit Function
    Next cell
    Set mCreatedStage = table: Set mCreatedStaging = New cInventoryTableInput
    Set mCreatedBook = operatorBook: mCreatedRange = table.Range.Address: mCreatedName = tableName
    If Not mCreatedStaging.Capture(table, "", True) Then Exit Function
    BindCreatedStage = Unchanged()
Invalid:
End Function

Public Function ConfirmStageSaved(Optional ByVal created As Boolean = False) As Boolean
    Dim source As Workbook, copy As Workbook, copyPath As String, hash As String
    Dim sheet As Worksheet, table As ListObject, persisted As ListObject
    Dim selected As ListObject, state As cInventoryTableInput
    On Error GoTo Failed
    If created Then
        Set selected = mCreatedStage: Set state = mCreatedStaging
    Else
        Set selected = mStage: Set state = mStaging
    End If
    If state Is Nothing Or selected Is Nothing Then Exit Function
    If Not Unchanged() Then Exit Function
    Set source = selected.Parent.Parent
    If Not source.Saved Then Exit Function
    hash = modTrainingWire.FileSha256(source.FullName)
    If hash = "" Then Exit Function
    Set copy = modAggregationFiles.OpenCopy(source.FullName, copyPath, True)
    If copy Is Nothing Then GoTo CleanExit
    If modTrainingWire.FileSha256(copyPath) <> hash Then GoTo CleanExit
    For Each sheet In copy.Worksheets
        For Each table In sheet.ListObjects
            If table.Name = selected.Name And sheet.Name = selected.Parent.Name Then Set persisted = table
        Next table
    Next sheet
    If persisted Is Nothing Then GoTo CleanExit
    If created Then
        If persisted.ListRows.Count <> 20 Or persisted.Range.Address <> mCreatedRange Then GoTo CleanExit
    End If
    If Not state.MatchesPersistedCopy(persisted) Then GoTo CleanExit
    If hash <> modTrainingWire.FileSha256(source.FullName) Then GoTo CleanExit
    ConfirmStageSaved = Unchanged()
CleanExit:
    If copyPath <> "" Then modAggregationFiles.CloseCopy copy, copyPath
    Exit Function
Failed:
    Resume CleanExit
End Function

' The coordinator declares exact intended fields before the corresponding write.
' Partial values may match their old or declared value; unrelated changes refuse.
Public Function BeginCatalog(ByVal sku As String) As Boolean
    If Not Unchanged() Or mPhase <> "" Then Exit Function
    If sku = "" Then Exit Function
    If Not BindSku(sku) Then Exit Function
    mPhase = "Catalog": BeginCatalog = True
End Function

Public Function PlanEditReceipt(ByVal headers As Variant, ByVal values As Variant) As Boolean
    Dim owner As cSnapshotInventorySource, book As Workbook, borrowed As Boolean
    Dim table As ListObject, i As Long
    On Error GoTo Invalid
    If mPhase <> "Catalog" Or Not Unchanged() Then Exit Function
    If LBound(headers) <> 0 Or LBound(values) <> 0 Then Exit Function
    If UBound(headers) <> UBound(values) Then Exit Function
    If CStr(headers(0)) <> "CommandId" Or CStr(values(0)) = "" Then Exit Function
    Set book = ReadSource(owner, borrowed)
    If book Is Nothing Then Exit Function
    Set table = modExecutionTarget.Table(book, "tblInventoryEditReceipts")
    Set mEditReceipts = New cInventoryTableInput
    mEditReceiptAbsent = (table Is Nothing)
    If mEditReceiptAbsent Then
        If Not mEditReceipts.CaptureAbsent(headers, "CommandId") Then Exit Function
    Else
        If Not mEditReceipts.Capture(table, "CommandId", borrowed) Then Exit Function
    End If
    If Not mEditReceipts.PlanNewRow(CStr(values(0))) Then Exit Function
    For i = 0 To UBound(headers)
        mEditReceipts.PlanValue CStr(values(0)), CStr(headers(i)), values(i)
    Next i
    ' Release our own read-only source before another source check classifies
    ' open workbooks. It must not be mistaken for a borrowed read-only owner.
    Set table = Nothing: Set book = Nothing: Set owner = Nothing
    PlanEditReceipt = Unchanged()
Invalid:
End Function

Private Function EditReceiptsMatch(ByVal book As Workbook, Optional ByVal saved As Boolean = False) As Boolean
    Dim table As ListObject
    If mEditReceipts Is Nothing Then EditReceiptsMatch = True: Exit Function
    Set table = modExecutionTarget.Table(book, "tblInventoryEditReceipts")
    If table Is Nothing Then
        EditReceiptsMatch = (mEditReceiptAbsent And Not saved)
    ElseIf saved Then
        EditReceiptsMatch = mEditReceipts.MatchesPersistedCopy(table)
    Else
        EditReceiptsMatch = mEditReceipts.Matches(table)
    End If
End Function

Public Sub PlanCatalogField(ByVal sku As String, ByVal header As String, ByVal value As Variant)
    If mPhase <> "Catalog" Or sku <> mSelectedSku Then Err.Raise 5, , "No matching catalog command."
    mCatalog.PlanValue sku, header, value
End Sub

Public Function BindCatalogWriter(ByVal book As Workbook) As Boolean
    On Error GoTo Invalid
    If mPhase <> "Catalog" Or Not Unchanged() Then Exit Function
    If Not modAggregationFiles.SameFile(book.FullName, mPath) Or book.ReadOnly Then Exit Function
    If Not book.Saved Then Exit Function
    If Not mCatalog.BindObject(modExecutionTarget.Table(book, "tblSkuCatalog")) Then Exit Function
    If Not mEntities.BindObject(modExecutionTarget.Table(book, "tblInventoryEntities")) Then Exit Function
    If Not mEditReceipts Is Nothing And Not mEditReceiptAbsent Then
        If Not mEditReceipts.BindObject(modExecutionTarget.Table(book, "tblInventoryEditReceipts")) Then Exit Function
    End If
    BindCatalogWriter = True
Invalid:
End Function

Public Function BeginQuantity(ByVal systemKey As String, ByVal target As Double) As Boolean
    Dim before As Variant
    If mPhase <> "" Or Not Unchanged() Or target < 0 Then Exit Function
    If CStr(mEntities.Scalar(systemKey, "SKU")) <> mSelectedSku Then Exit Function
    before = mEntities.Scalar(systemKey, "QtyOnHand")
    If IsEmpty(before) Or Not IsNumeric(before) Then Exit Function
    mQuantityKey = systemKey: mQuantityBefore = CDbl(before): mQuantityTarget = target
    mEntities.PlanValue systemKey, "QtyOnHand", target
    mPhase = "Quantity": BeginQuantity = True
End Function

Public Function BeginRetirement(ByVal sku As String, ByVal reason As String, _
        ByRef payload As String, ByRef count As Long) As Boolean
    If mPhase <> "" Or Not BindSku(sku) Then Exit Function
    Set mRetirement = New cInventoryRetirementPlan
    If Not mRetirement.Plan(mEntities, sku, reason, payload, count) Then Exit Function
    If count > 0 Then mPhase = "Retirement"
    BeginRetirement = True
End Function

Public Function BeginCreate(ByVal payload As String) As Boolean
    If mPhase <> "" Or Not Unchanged() Then Exit Function
    Set mCreate = New cInventoryCreatePlan
    If Not mCreate.Plan(payload, mCatalog, mEntities, mSelectedSku) Then Exit Function
    mPhase = "Create": BeginCreate = True
End Function

Public Function BeginSave(Optional ByVal expectedPhase As String = "Catalog") As Boolean
    Dim owner As cSnapshotInventorySource, book As Workbook, borrowed As Boolean
    If mPhase <> expectedPhase Or Not Unchanged() Then Exit Function
    Set book = ReadSource(owner, borrowed)
    If book Is Nothing Then Exit Function
    If mPhase <> "Create" Then
        If Not mCatalog.Matches(modExecutionTarget.Table(book, "tblSkuCatalog"), "", False, True) Then Exit Function
    End If
    If Not EditReceiptsMatch(book, True) Then Exit Function
    mSaving = True: BeginSave = True
End Function

Public Function ConfirmCatalog() As Boolean
    If mPhase = "Catalog" Then ConfirmCatalog = ConfirmSavedInputs()
End Function

Public Function ConfirmEventResult(ByVal eventId As String, ByVal expectedPhase As String) As Boolean
    If eventId = "" Or mPhase <> expectedPhase Then Exit Function
    If mPhase <> "Quantity" And mPhase <> "Create" And mPhase <> "Retirement" Then Exit Function
    ConfirmEventResult = ConfirmSavedInputs(eventId)
End Function

Private Function ConfirmSavedInputs(Optional ByVal eventId As String = "") As Boolean
    Dim owner As cSnapshotInventorySource, book As Workbook, borrowed As Boolean
    Dim persisted As Workbook, copyPath As String, savedHash As String, savedIdentity As String
    On Error GoTo Invalid
    If Not mSaving Then Exit Function
    If mPhase = "Catalog" Then
        If Not Unchanged() Then Exit Function
    ElseIf mPhase <> "Quantity" And mPhase <> "Create" And mPhase <> "Retirement" Then
        Exit Function
    End If
    Set book = ReadSource(owner, borrowed)
    If book Is Nothing Then Exit Function
    If Not book.Saved Then Exit Function
    savedIdentity = modAggregationFiles.Identity(mPath): savedHash = modTrainingWire.FileSha256(mPath)
    If savedIdentity = "" Or savedHash = "" Then Exit Function
    Set persisted = modAggregationFiles.OpenCopy(mPath, copyPath, True)
    If persisted Is Nothing Then GoTo CleanExit
    If modTrainingWire.FileSha256(copyPath) <> savedHash Then GoTo CleanExit
    If mPhase = "Quantity" Then
        If Not PlanSavedQuantity(persisted, eventId) Then GoTo CleanExit
    ElseIf mPhase = "Create" Then
        If Not mCreate.PlanSaved(persisted, eventId, mEntities) Then GoTo CleanExit
    ElseIf mPhase = "Retirement" Then
        If Not mRetirement.PlanSaved(persisted, eventId, mEntities) Then GoTo CleanExit
    End If
    If Not mCatalog.MatchesPersistedCopy(modExecutionTarget.Table(persisted, "tblSkuCatalog")) Then GoTo CleanExit
    If Not mEntities.MatchesPersistedCopy(modExecutionTarget.Table(persisted, "tblInventoryEntities")) Then GoTo CleanExit
    If Not EditReceiptsMatch(persisted, True) Then GoTo CleanExit
    If modAggregationFiles.Identity(mPath) <> savedIdentity Or modTrainingWire.FileSha256(mPath) <> savedHash Then GoTo CleanExit
    If borrowed Then
        If Not Unchanged() Then GoTo CleanExit
    ElseIf Not mStaging Is Nothing Then
        If Not mStaging.Matches(mStage) Then GoTo CleanExit
    End If
    If Not mCatalog.Confirm(modExecutionTarget.Table(book, "tblSkuCatalog")) Then GoTo CleanExit
    If Not mEntities.Confirm(modExecutionTarget.Table(book, "tblInventoryEntities")) Then GoTo CleanExit
    If Not mEditReceipts Is Nothing Then
        If Not mEditReceipts.Confirm(modExecutionTarget.Table(book, "tblInventoryEditReceipts")) Then GoTo CleanExit
        mEditReceiptAbsent = False
    End If
    If Not mInitiallyBorrowed Then
        mCatalog.ReleaseObjectReference
        mEntities.ReleaseObjectReference
        If Not mEditReceipts Is Nothing Then mEditReceipts.ReleaseObjectReference
    End If
    ' Release our read-only source before another read classifies open books.
    Set book = Nothing: Set owner = Nothing
    mIdentity = savedIdentity: mHash = savedHash
    mPhase = "": mSaving = False
    ConfirmSavedInputs = Unchanged()
CleanExit:
    If copyPath <> "" Then modAggregationFiles.CloseCopy persisted, copyPath
    Exit Function
Invalid:
    Resume CleanExit
End Function

Private Function PlanSavedQuantity(ByVal book As Workbook, ByVal eventId As String) As Boolean
    Dim table As ListObject, fields As Object, row As Long, receipts As Long, lines As Long
    Dim appliedAt As Variant, delta As Variant
    On Error GoTo Invalid
    Set table = modExecutionTarget.Table(book, "tblAppliedEvents")
    Set fields = modAggregationSources.RequiredColumns(table, Array("EventID", "Status"))
    If fields Is Nothing Then Exit Function
    For row = 1 To table.ListRows.Count
        If CStr(table.DataBodyRange.Cells(row, fields("EventID")).Value2) = eventId Then
            receipts = receipts + 1
            If CStr(table.DataBodyRange.Cells(row, fields("Status")).Value2) <> "APPLIED" Then Exit Function
        End If
    Next row
    If receipts <> 1 Then Exit Function
    Set table = modExecutionTarget.Table(book, "tblInventoryLog")
    Set fields = modAggregationSources.RequiredColumns(table, Array("EventID", "System_Key", "SKU", "QtyDelta", "AppliedAtUTC"))
    If fields Is Nothing Then Exit Function
    For row = 1 To table.ListRows.Count
        If CStr(table.DataBodyRange.Cells(row, fields("EventID")).Value2) = eventId Then
            lines = lines + 1
            If CStr(table.DataBodyRange.Cells(row, fields("System_Key")).Value2) <> mQuantityKey Then Exit Function
            If CStr(table.DataBodyRange.Cells(row, fields("SKU")).Value2) <> mSelectedSku Then Exit Function
            delta = table.DataBodyRange.Cells(row, fields("QtyDelta")).Value2
            If Not IsNumeric(delta) Then Exit Function
            If Abs(CDbl(delta) - (mQuantityTarget - mQuantityBefore)) > 0.0000001 Then Exit Function
            appliedAt = table.DataBodyRange.Cells(row, fields("AppliedAtUTC")).Value2
            If IsNumeric(appliedAt) Then appliedAt = CDate(appliedAt)
            If Not IsDate(appliedAt) Then Exit Function
        End If
    Next row
    If lines <> 1 Then Exit Function
    mEntities.PlanValue mQuantityKey, "LastAppliedUTC", CDbl(CDate(appliedAt))
    PlanSavedQuantity = True
Invalid:
End Function

Private Function ReadSource(ByRef owner As cSnapshotInventorySource, ByRef borrowed As Boolean) As Workbook
    Dim candidate As Workbook, book As Workbook
    For Each candidate In Application.Workbooks
        If modAggregationFiles.SameFile(candidate.FullName, mPath) Then borrowed = True: Exit For
    Next candidate
    Set owner = New cSnapshotInventorySource
    Set book = owner.Read(mWarehouse, Nothing)
    If book Is Nothing Then Exit Function
    If Not modAggregationFiles.SameFile(book.FullName, mPath) Then Exit Function
    Set ReadSource = book
End Function
