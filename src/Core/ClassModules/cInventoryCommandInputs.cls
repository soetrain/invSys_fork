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
Private mSelectedSku As String, mPhase As String, mSaving As Boolean
Private mInitiallyBorrowed As Boolean
Private mCreate As cInventoryCreatePlan
Private mRetirement As cInventoryRetirementPlan
Private mQuantityKey As String, mQuantityBefore As Double, mQuantityTarget As Double

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
    If Not mStaging Is Nothing Then
        If Not mStaging.Matches(mStage) Then Exit Function
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

Public Function ConfirmStageSaved() As Boolean
    Dim source As Workbook, copy As Workbook, copyPath As String, hash As String
    Dim sheet As Worksheet, table As ListObject, persisted As ListObject
    On Error GoTo Failed
    If mStaging Is Nothing Or mStage Is Nothing Then Exit Function
    If Not Unchanged() Then Exit Function
    Set source = mStage.Parent.Parent
    If Not source.Saved Then Exit Function
    hash = modTrainingWire.FileSha256(source.FullName)
    If hash = "" Then Exit Function
    Set copy = modAggregationFiles.OpenCopy(source.FullName, copyPath, True)
    If copy Is Nothing Then GoTo CleanExit
    If modTrainingWire.FileSha256(copyPath) <> hash Then GoTo CleanExit
    For Each sheet In copy.Worksheets
        For Each table In sheet.ListObjects
            If table.Name = mStage.Name And sheet.Name = mStage.Parent.Name Then Set persisted = table
        Next table
    Next sheet
    If persisted Is Nothing Then GoTo CleanExit
    If Not mStaging.MatchesPersistedCopy(persisted) Then GoTo CleanExit
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
    If modAggregationFiles.Identity(mPath) <> savedIdentity Or modTrainingWire.FileSha256(mPath) <> savedHash Then GoTo CleanExit
    If borrowed Then
        If Not Unchanged() Then GoTo CleanExit
    ElseIf Not mStaging Is Nothing Then
        If Not mStaging.Matches(mStage) Then GoTo CleanExit
    End If
    If Not mCatalog.Confirm(modExecutionTarget.Table(book, "tblSkuCatalog")) Then GoTo CleanExit
    If Not mEntities.Confirm(modExecutionTarget.Table(book, "tblInventoryEntities")) Then GoTo CleanExit
    If Not mInitiallyBorrowed Then
        mCatalog.ReleaseObjectReference
        mEntities.ReleaseObjectReference
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
