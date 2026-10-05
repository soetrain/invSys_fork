# Disposable worksheet staging backed by Admin-seeded exact inventory entities.
. (Join-Path $PSScriptRoot 'Slice4beProductionCompleteWorksheetPermission.ps1')
. (Join-Path $PSScriptRoot 'Slice4beProductionOutputIdentity.ps1')
function Install-ProductionCompleteWorksheetProbe {
    Install-ProductionOutputIdentityProbe
    $project=$packages['invSys.Operations.xlam'].VBProject
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $form.InsertLines(1,'Private mCompleteWorksheetPhase As String, mCompleteWorksheetError As Long, mCompleteWorksheetInputBefore As Double')
    $form.AddFromString(@'
Public Function CompleteWorksheetPrepareForTest(ByVal canary As String) As Boolean
    Dim ws As Worksheet, logSheet As Worksheet, lo As ListObject, headers As Variant, col As Long
    Dim priorLoading As Boolean
    On Error GoTo Failed
    priorLoading = mLoading: mCompleteWorksheetError = 0
    mCompleteWorksheetPhase = "Stock"
    If Not RunWorksheetPrepareForTest() Then Exit Function
    mCompleteWorksheetInputBefore = CompleteWorksheetExactQtyForTest(mRunSheetKey)
    mCompleteWorksheetPhase = "Staging"
    If Not CheckBaselineWorksheetStageForTest(mRunSheetKey, canary) Then Exit Function
    mLoading = True
    Set ws = mOperatorWorkbook.Worksheets("Production")
    headers = Array("Operator Extra", "System_Key", "PROCESS", "OUTPUT", "ITEM_CODE", "UOM", _
        "REAL OUTPUT", "LOCATION", "Condition", "BATCH", "RECALL CODE", "GUID", "Operator Formula")
    For col = 0 To UBound(headers): ws.Cells(1, 31 + col).Value2 = headers(col): Next col
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("AE1:AQ2"), , xlYes)
    lo.Name = "ProductionOutput"
    SetCellByHeader lo, 1, "Operator Extra", canary
    SetCellByHeader lo, 1, "PROCESS", canary
    SetCellByHeader lo, 1, "OUTPUT", "Black Tea"
    SetCellByHeader lo, 1, "ITEM_CODE", "DEMO-RAW-BLACK-TEA"
    SetCellByHeader lo, 1, "UOM", mRunSheetUom
    SetCellByHeader lo, 1, "REAL OUTPUT", 1#
    SetCellByHeader lo, 1, "LOCATION", mRunSheetLocation
    SetCellByHeader lo, 1, "Condition", "GOOD"
    SetCellByHeader lo, 1, "BATCH", 1#
    SetCellByHeader lo, 1, "GUID", canary
    lo.DataBodyRange.Cells(1, ProductionColumnIndex(lo, "Operator Formula")).Formula = "=1+2"
    mCompleteWorksheetPhase = "LogSurface"
    Set logSheet = mOperatorWorkbook.Worksheets.Add
    logSheet.Name = "ProductionLog"
    headers = Array("Operator Extra", "System_Key", "PROCESS", "OUTPUT", "ITEM_CODE", "UOM", _
        "REAL OUTPUT", "QUANTITY", "LOCATION", "BATCH", "TIMESTAMP", "RECALL CODE", "GUID", "Operator Formula")
    For col = 0 To UBound(headers): logSheet.Cells(1, col + 1).Value2 = headers(col): Next col
    Set lo = logSheet.ListObjects.Add(xlSrcRange, logSheet.Range("A1:N2"), , xlYes)
    lo.Name = "ProductionLog"
    SetCellByHeader lo, 1, "Operator Extra", canary
    lo.DataBodyRange.Cells(1, ProductionColumnIndex(lo, "Operator Formula")).Formula = "=1+2"
    mLoading = priorLoading
    mCompleteWorksheetPhase = "CheckIn"
    mBtnManagerCheckIn_Click
    If Not CheckBaselineWorksheetFactForTest("ReachedCheckRows") Then Exit Function
    mCompleteWorksheetPhase = "SelectOutput"
    mLoading = True
    RefreshManagerState
    If mLstManagerOutput.ListCount <> 1 Then GoTo Failed
    mLstManagerOutput.ListIndex = 0
    LoadSelectedProductionOutput
    mTxtOutputReal.Text = "1"
    mLoading = priorLoading
    mCompleteWorksheetPhase = "Ready"
    CompleteWorksheetPrepareForTest = Not modProductionReusableRun.ReusableRunIsLoaded() And _
        HasProductionCheckRows() And SelectedProductionOutputTableRow() = 1 And mCompleteWorksheetInputBefore >= 10#
    Exit Function
Failed:
    mCompleteWorksheetError = Err.Number
    mLoading = priorLoading
End Function
Private Function CompleteWorksheetExactQtyForTest(ByVal key As String) As Double
    Dim entities As Variant, row As Long, matches As Long
    entities = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge("")
    If Not IsArray(entities) Then Err.Raise vbObjectError + 270, , "Exact worksheet inventory unavailable."
    For row = LBound(entities, 1) To UBound(entities, 1)
        If StrComp(CStr(entities(row, 1)), key, vbBinaryCompare) = 0 Then
            matches = matches + 1
            CompleteWorksheetExactQtyForTest = CDbl(entities(row, 6))
        End If
    Next row
    If matches > 1 Then Err.Raise vbObjectError + 271, , "Exact worksheet entity is ambiguous."
End Function
Public Function CompleteWorksheetPhaseForTest() As String
    CompleteWorksheetPhaseForTest = mCompleteWorksheetPhase & "|" & CStr(mCompleteWorksheetError)
End Function
Public Function CompleteWorksheetFactForTest(ByVal fact As String, ByVal canary As String) As Boolean
    Dim lo As ListObject, key As String, row As Long, matches As Long
    Set lo = ProductionTable(TABLE_MANAGER_OUTPUT)
    Select Case fact
        Case "InputUnchanged"
            CompleteWorksheetFactForTest = Abs(CompleteWorksheetExactQtyForTest(mRunSheetKey) - mCompleteWorksheetInputBefore) < 0.0000001
        Case "Unsubmitted"
            CompleteWorksheetFactForTest = CellByHeader(lo, 1, "System_Key") = "" And HasProductionCheckRows()
        Case "PermissionMessage"
            CompleteWorksheetFactForTest = mTxtStatus.Text = "Current user lacks PROD_POST capability."
        Case "ExactInput"
            CompleteWorksheetFactForTest = Abs(CompleteWorksheetExactQtyForTest(mRunSheetKey) - (mCompleteWorksheetInputBefore - 10#)) < 0.0000001
        Case "FreshOutput"
            key = CellByHeader(lo, 1, "System_Key")
            If key = "" Or StrComp(key, mRunSheetKey, vbBinaryCompare) = 0 Then Exit Function
            CompleteWorksheetFactForTest = Abs(CompleteWorksheetExactQtyForTest(key) - 1#) < 0.0000001
        Case "OutputCustom"
            CompleteWorksheetFactForTest = CellByHeader(lo, 1, "Operator Extra") = canary And _
                lo.DataBodyRange.Cells(1, ProductionColumnIndex(lo, "Operator Formula")).Formula = "=1+2"
        Case "CheckCustom"
            Set lo = ProductionTable(TABLE_MANAGER_CHECK)
            CompleteWorksheetFactForTest = CellByHeader(lo, 1, "Operator Extra") = canary And _
                lo.DataBodyRange.Cells(1, ProductionColumnIndex(lo, "Operator Formula")).Formula = "=1+2"
        Case "Log"
            key = CellByHeader(lo, 1, "System_Key")
            Set lo = mOperatorWorkbook.Worksheets("ProductionLog").ListObjects("ProductionLog")
            For row = 1 To lo.ListRows.Count
                If key <> "" And StrComp(CellByHeader(lo, row, "System_Key"), key, vbBinaryCompare) = 0 Then
                    If CellByHeader(lo, row, "REAL OUTPUT") = "1" Then matches = matches + 1
                End If
            Next row
            CompleteWorksheetFactForTest = (matches = 1)
    End Select
End Function
'@)
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,'Private mCompleteWorksheetEvents As String')
    $adapter.AddFromString(@'
Public Function CompleteWorksheetPrepare(ByVal canary As String) As Boolean
    mCompleteWorksheetEvents = ""
    CompleteWorksheetPrepare = mForm.CompleteWorksheetPrepareForTest(canary)
End Function
Public Function CompleteWorksheetPhase() As String
    CompleteWorksheetPhase = mForm.CompleteWorksheetPhaseForTest()
End Function
Public Function CompleteWorksheetFact(ByVal fact As String, ByVal canary As String) As Boolean
    CompleteWorksheetFact = mForm.CompleteWorksheetFactForTest(fact, canary)
End Function
Public Sub CompleteWorksheetWriterReturned(ByVal eventId As String)
    If mCompleteWorksheetEvents <> "" Then mCompleteWorksheetEvents = mCompleteWorksheetEvents & vbLf
    mCompleteWorksheetEvents = mCompleteWorksheetEvents & eventId
End Sub
Public Function CompleteWorksheetEvents() As String
    CompleteWorksheetEvents = mCompleteWorksheetEvents
End Function
'@)
    $owner=$project.VBComponents.Item('modProductionCompletionService').CodeModule
    $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestShippingCatalog').CodeModule.AddFromString(@'
Public Function CompleteWorksheetInventoryForTest(ByVal workbookName As String) As Boolean
    Dim wb As Workbook, report As String
    Set wb = Application.Workbooks(workbookName)
    If Not modRoleWorkbookSurfaces.EnsureInventoryManagementSurface(wb, report) Then Exit Function
    CompleteWorksheetInventoryForTest = modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wb, "", "LOCAL", report)
End Function
Public Function CompleteWorksheetAdminOnlyForTest() As Boolean
    CompleteWorksheetAdminOnlyForTest = modRoleUiAccess.CanCurrentUserPerformCapability("ADMIN_MAINT") And _
        Not modRoleUiAccess.CanCurrentUserPerformCapability("PROD_POST")
End Function
'@)
    foreach($point in @(@('session.MarkConsumeQueued','consumeEventId'),@('session.MarkCompleteQueued','completeEventId'))){
        $start=$owner.ProcStartLine('QueueProductionSessionEvents',0);$end=$start+$owner.ProcCountLines('QueueProductionSessionEvents',0)
        $hits=@(for($line=$start;$line -lt $end;$line++){if($owner.Lines($line,1).Trim() -ieq $point[0]){$line}})
        if($hits.Count -ne 1){throw 'Worksheet writer acknowledgment anchor changed; not product RED.'}
        $code='    TestProductionDesigner.CompleteWorksheetWriterReturned '+$point[1]
        $owner.InsertLines($hits[0]+1,$code)
        if($owner.Lines($hits[0]+1,1) -ine $code){throw 'Worksheet writer hook misplaced; not product RED.'}
    }
}

function Test-ProductionCompleteWorksheet($Fixture,$Book,$Decoy,[string]$Canary) {
    if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.CompleteWorksheetInventoryForTest' @($Book.Name))){throw 'Core-owned worksheet inventory projection unavailable; not product RED.'}
    Check 'CompleteActivity.Worksheet.CoreInventoryProjectionPrepared' $true
    $ready=[bool](Probe 'CompleteWorksheetPrepare' @($Canary))
    if(-not $ready){
        $phase=([string](Probe 'CompleteWorksheetPhase')).Split('|')
        if($phase.Count -ne 2 -or $phase[0] -cnotin @('Stock','Staging','LogSurface','CheckIn','SelectOutput','Ready')){throw 'Unknown worksheet preparation boundary.'}
        [pscustomobject]@{Phase=$phase[0];ErrorNumber=[int]$phase[1]}|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'complete-worksheet-prerequisite.json')
        throw 'Actual worksheet Check In/output prerequisite unavailable; not product RED.'
    }
    [void](Probe 'RunLocalShowAndCapture' @($Book.Name,'CHECK_IN'));$Decoy.Activate()
    Check 'CompleteActivity.Worksheet.Identity.BlankDraftHasNoInventedKey' ([bool](Probe 'OutputIdentity' @('Blank',$Canary)))
    Test-ProductionCompleteWorksheetPermission $Fixture $Book $Decoy $Canary
    $before=@(Get-Slice4beActivityFiles $Fixture);$label='CompleteActivity.Worksheet.Normal'
    Check ($label+'.ActualHandlerReturned') ([bool](Probe 'CompleteEntryAct' @('')))
    Check ($label+'.GuardsRestored') ([bool](Probe 'CompleteEntryFact' @('GuardsRestored')))
    $ids=([string](Probe 'CompleteWorksheetEvents')).Split("`n")
    if($ids.Count -ne 2 -or $ids[0] -ceq '' -or $ids[1] -ceq '' -or $ids[0] -ceq $ids[1]){throw 'Actual worksheet writer acknowledgments unavailable; not product RED.'}
    foreach($fact in @('ExactInput','FreshOutput','Log','OutputCustom','CheckCustom')){
        Check ($label+'.'+$fact) ([bool](Probe 'CompleteWorksheetFact' @($fact,$Canary)))
    }
    Pair $before 'CONFIRMED' $label $ids
    CaptureOwnedFormByCaptionEvidence 'Production' 'complete-activity-worksheet-confirmed.png'
    Check ($label+'.DecoyPreserved') ($Decoy.Worksheets.Count -eq 1 -and $Decoy.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq $Canary)
    Test-ProductionOutputIdentity $Book $Decoy $Canary
}
