# D14/D15 owner-boundary baseline. No new activity catalog contract is assumed.
# Adapters stage disposable projections and invoke the unchanged operator handler.
function Install-ProductionCheckInBaselineProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $start=$form.ProcStartLine('CheckInProductionRun',0);$end=$start+$form.ProcCountLines('CheckInProductionRun',0)
    $hits=@(for($line=$start;$line -lt $end;$line++){if($form.Lines($line,1).Trim() -match '^checkInStage = "[A-Za-z]+"$'){$line}})
    if($hits.Count -ne 9){throw 'Check In diagnostic stage anchors changed; not product RED.'}
    foreach($line in @($hits|Sort-Object -Descending)){$form.InsertLines($line+1,'    TestProductionDesigner.CheckBaselineStageHit checkInStage')}
    $form.InsertLines(1,'Private mCheckBaselineKey As String, mCheckBaselineCanary As String, mCheckBaselineGuardsRestored As Boolean')
    $form.AddFromString(@'
Public Function CheckBaselineReusableStageForTest(ByVal mode As String) As Boolean
    Dim report As String
    If RunLocalStageForTest("ALLOCATE", "Normal") <> "READY" Then Exit Function
    If mode <> "Insufficient" Then
        If Not modProductionReusableRun.ApplyReusableRunStockAllocation( _
            NzStr(mLstRunPalette.List(0, 0)), NzStr(mLstRunPalette.List(0, 1)), _
            NzStr(mLstRunPalette.List(0, 3)), 2#, report) Then Exit Function
    End If
    RefreshReusableRunControls False
    mLoading = True
    If mode = "NoProcess" Then
        mCmbRunProcess.ListIndex = 0: mCmbTreeRunProcess.ListIndex = 0
    Else
        If mCmbRunProcess.ListCount <> 2 Or mCmbTreeRunProcess.ListCount <> 2 Then GoTo Finished
        mCmbRunProcess.ListIndex = 1: mCmbTreeRunProcess.ListIndex = 1
    End If
    CheckBaselineReusableStageForTest = Not modProductionReusableRun.ReusableRunIsCheckedIn() And _
        Not modProductionReusableRun.ReusableRunBatchNoteFrozen() And _
        ((mode = "NoProcess" And ActiveRunProcess() = "") Or _
         (mode <> "NoProcess" And ActiveRunProcess() <> ""))
Finished:
    mLoading = False
End Function
Public Function CheckBaselineActForTest(ByVal guard As String) As Boolean
    Dim priorLoading As Boolean, priorBusy As Boolean, entryLoading As Boolean, entryBusy As Boolean
    On Error GoTo Failed
    priorLoading = mLoading: priorBusy = mDesignerActionInProgress
    mCheckBaselineGuardsRestored = False
    If guard = "Loading" Then mLoading = True
    If guard = "Busy" Then mDesignerActionInProgress = True
    entryLoading = mLoading: entryBusy = mDesignerActionInProgress
    mBtnManagerCheckIn_Click
    mCheckBaselineGuardsRestored = (mLoading = entryLoading And mDesignerActionInProgress = entryBusy)
    CheckBaselineActForTest = True
Failed:
    mLoading = priorLoading: mDesignerActionInProgress = priorBusy
End Function
Public Function CheckBaselineGuardsForTest() As Boolean
    CheckBaselineGuardsForTest = mCheckBaselineGuardsRestored
End Function
Public Function CheckBaselineContextRefusedForTest() As Boolean
    CheckBaselineContextRefusedForTest = InStr(1, mTxtStatus.Text, _
        "Session, warehouse, or captured workbook changed.", vbBinaryCompare) = 1
End Function
Public Function CheckBaselinePermissionRefusedForTest() As Boolean
    CheckBaselinePermissionRefusedForTest = InStr(1, mTxtStatus.Text, "Production permission", vbBinaryCompare) = 1
End Function
Public Function CheckBaselineWorksheetSourceForTest() As Variant
    CheckBaselineWorksheetSourceForTest = Array(mRunSheetKey, mRunSheetLocation, mRunSheetUom, mRunSheetAvailable)
End Function
Public Sub CheckBaselineRestoreWorksheetSourceForTest(ByVal fixture As Variant)
    mRunSheetKey = fixture(0): mRunSheetLocation = fixture(1)
    mRunSheetUom = fixture(2): mRunSheetAvailable = fixture(3)
End Sub
Public Function CheckBaselineReusableResultForTest(ByVal checked As Boolean) As Boolean
    CheckBaselineReusableResultForTest = _
        modProductionReusableRun.ReusableRunIsCheckedIn() = checked And _
        modProductionReusableRun.ReusableRunBatchNoteFrozen() = checked And _
        Not modProductionReusableRun.ReusableRunIsCompleted() And _
        modProductionReusableRun.ReusableRunBatchNote() = mReadCanaryForTest
End Function
Public Function CheckBaselineReusableDisplayForTest(ByVal keyA As String, ByVal keyB As String) As Boolean
    Dim key As String
    If mLstManagerCheck.ListCount <> 1 Or mLstManagerCheck.ColumnCount <> 9 Then Exit Function
    key = NzStr(mLstManagerCheck.List(0, 3))
    CheckBaselineReusableDisplayForTest = (key = keyA Or key = keyB) And _
        NzStr(mLstManagerCheck.List(0, 0)) = "EXTERNAL" And _
        NzStr(mLstManagerCheck.List(0, 4)) = "DEMO-RAW-BLACK-TEA" And _
        NzStr(mLstManagerCheck.List(0, 7)) = "2" And _
        mPages.Pages(3).Controls("hdrManagerCheck4").Caption = "System_Key"
End Function
Public Function CheckBaselineWorksheetStageForTest(ByVal selectedKey As String, ByVal canary As String) As Boolean
    Dim ws As Worksheet, lo As ListObject, headers As Variant, col As Long
    Set lo = ProductionTable(TABLE_MANAGER_CHECK)
    If Not lo Is Nothing Then lo.Unlist
    mOperatorWorkbook.Worksheets("Production").Range("U1:AB40").Clear
    If RunWorksheetStageForTest("ALLOCATE", "Quantity", canary) <> "READY" Then Exit Function
    mCheckBaselineKey = selectedKey: mCheckBaselineCanary = canary
    mLoading = True
    mLstLoaderLines.ListIndex = -1
    mLstRunPalette.List(0, 3) = selectedKey
    mLstRunPalette.List(0, 5) = "100": mLstRunPalette.List(0, 6) = "10"
    StoreRunItemCode mLstRunPalette, 0, "DEMO-RAW-BLACK-TEA"
    StoreRunProcess mLstRunPalette, 0, canary
    StoreRunBaseQty mLstRunPalette, 0, "10"
    StoreRunAllocationOverride mLstRunPalette, 0, "100", "10"
    Set ws = mOperatorWorkbook.Worksheets("Production")
    ws.Range("B2").Value2 = 10#: ws.Range("C2").Value2 = selectedKey: ws.Range("F2").Value2 = 100#
    headers = Array("Operator Extra", "USED", "System_Key", "Operator Formula", "ITEM_CODE", "ITEM", "UOM", "TOTAL INV")
    For col = 0 To UBound(headers)
        ws.Cells(1, 21 + col).Value2 = headers(col)
    Next col
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("U1:AB2"), , xlYes)
    lo.Name = TABLE_MANAGER_CHECK
    lo.DataBodyRange.Cells(1, 1).Value2 = canary
    lo.DataBodyRange.Cells(1, 3).Value2 = selectedKey
    lo.DataBodyRange.Cells(1, 4).Formula = "=1+2"
    mLoading = False
    CheckBaselineWorksheetStageForTest = Not modProductionReusableRun.ReusableRunIsLoaded() And _
        ValidateRunAllocationsComplete() And ValidateRunAllocationLocations() And _
        RunItemCodeFromList(mLstRunPalette, 0) = "DEMO-RAW-BLACK-TEA" And _
        RunBaseQtyFromList(mLstRunPalette, 0) = 10# And _
        StrComp(NzStr(mLstRunPalette.List(0, 3)), selectedKey, vbBinaryCompare) = 0
End Function
Public Sub CheckBaselineWorksheetInvalidForTest(ByVal mode As String)
    Dim lo As ListObject
    mLoading = True
    Set lo = ProductionTable(TABLE_MANAGER_CHECK)
    SetCellByHeader lo, 1, "USED", 7#
    Select Case mode
        Case "MissingKey": mLstRunPalette.List(0, 3) = ""
        Case "UnknownKey": mLstRunPalette.List(0, 3) = "UNRESOLVED-ENTITY-FOR-NEGATIVE-TEST"
        Case "MissingUsed": lo.ListColumns("USED").Name = "Operator Used"
        Case "MissingIdentityHeader": lo.ListColumns("System_Key").Name = "Operator Identity"
    End Select
    StoreRunItemCode mLstRunPalette, 0, "DEMO-RAW-BLACK-TEA"
    StoreRunProcess mLstRunPalette, 0, mCheckBaselineCanary
    StoreRunBaseQty mLstRunPalette, 0, "10"
    StoreRunAllocationOverride mLstRunPalette, 0, "100", "10"
    mLoading = False
End Sub
Public Function CheckBaselineWorksheetStateForTest() As String
    CheckBaselineWorksheetStateForTest = CheckBaselineTableStateForTest(ProductionTable(TABLE_MANAGER_CHECK))
End Function
Private Function CheckBaselineTableStateForTest(ByVal lo As ListObject) As String
    Dim cell As Range
    If lo Is Nothing Then Exit Function
    For Each cell In lo.Range.Cells
        CheckBaselineTableStateForTest = CheckBaselineTableStateForTest & _
            CStr(Len(CStr(cell.Formula))) & ":" & CStr(cell.Formula) & "|"
    Next cell
End Function
Public Function CheckBaselineInventoryPrepareForTest(ByVal firstKey As String, ByVal selectedKey As String) As Boolean
    Dim ws As Worksheet, lo As ListObject, headers As Variant, col As Long, row As Long
    If WorkbookHasSheet(mOperatorWorkbook, "InventoryManagement") Then Exit Function
    Set ws = mOperatorWorkbook.Worksheets.Add
    ws.Name = "InventoryManagement"
    headers = Array("Operator Extra", "ITEM_CODE", "System_Key", "ITEM", "LOCATION", "UOM", "QUANTITY", "Operator Formula")
    For col = 0 To UBound(headers)
        ws.Cells(1, col + 1).Value2 = headers(col)
    Next col
    For row = 2 To 3
        ws.Cells(row, 1).Value2 = mCheckBaselineCanary
        ws.Cells(row, 2).Value2 = "DEMO-RAW-BLACK-TEA"
        ws.Cells(row, 3).Value2 = IIf(row = 2, firstKey, selectedKey)
        ws.Cells(row, 4).Value2 = mCheckBaselineCanary
        ws.Cells(row, 5).Value2 = mRunSheetLocation
        ws.Cells(row, 6).Value2 = mRunSheetUom
        ws.Cells(row, 7).Value2 = mRunSheetAvailable
        ws.Cells(row, 8).Formula = "=1+2"
    Next row
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:H3"), , xlYes)
    lo.Name = "invSys"
    CheckBaselineInventoryPrepareForTest = StrComp(firstKey, selectedKey, vbBinaryCompare) <> 0
End Function
Public Function CheckBaselineInventoryStateForTest() As String
    CheckBaselineInventoryStateForTest = CheckBaselineTableStateForTest(InventoryTable())
End Function
Public Function CheckBaselineWorksheetFactForTest(ByVal fact As String) As Boolean
    Dim lo As ListObject, ws As Worksheet, headers As Variant, col As Long
    Set ws = mOperatorWorkbook.Worksheets("Production")
    Set lo = ProductionTable(TABLE_MANAGER_CHECK)
    If lo Is Nothing Then Exit Function
    Select Case fact
        Case "ReachedCheckRows"
            CheckBaselineWorksheetFactForTest = CellByHeader(lo, 1, "USED") = "10" And _
                InStr(1, mTxtStatus.Text, "Checked in ", vbBinaryCompare) = 1
        Case "ExactSelectedKey"
            CheckBaselineWorksheetFactForTest = StrComp(CellByHeader(lo, 1, "System_Key"), mCheckBaselineKey, vbBinaryCompare) = 0
        Case "CustomValue"
            CheckBaselineWorksheetFactForTest = CellByHeader(lo, 1, "Operator Extra") = mCheckBaselineCanary
        Case "CustomFormula"
            CheckBaselineWorksheetFactForTest = lo.DataBodyRange.Cells(1, ProductionColumnIndex(lo, "Operator Formula")).Formula = "=1+2"
        Case "PalettePreserved"
            CheckBaselineWorksheetFactForTest = ws.Range("A2").Value2 = mCheckBaselineCanary And _
                ws.Range("D2").Formula = "=1+2" And _
                StrComp(CStr(ws.Range("C2").Value2), mCheckBaselineKey, vbBinaryCompare) = 0
        Case "HeadersPreserved"
            CheckBaselineWorksheetFactForTest = lo.ListColumns.Count = 8 And _
                lo.ListColumns(1).Name = "Operator Extra" And lo.ListColumns(3).Name = "System_Key" And _
                lo.ListColumns(4).Name = "Operator Formula"
        Case "NoSuccessStatus"
            CheckBaselineWorksheetFactForTest = InStr(1, mTxtStatus.Text, "Checked in ", vbBinaryCompare) <> 1
        Case "DisplayColumns"
            If mLstManagerCheck.ListCount <> 1 Or mLstManagerCheck.ColumnCount <> 9 Then Exit Function
            If mPages.Pages(3).Controls("hdrManagerCheck4").Caption <> "System_Key" Then Exit Function
            For col = 0 To 2
                If NzStr(mLstManagerCheck.List(0, col)) <> "" Then Exit Function
            Next col
            headers = Array("System_Key", "ITEM_CODE", "ITEM", "UOM", "USED", "TOTAL INV")
            For col = 0 To UBound(headers)
                If StrComp(NzStr(mLstManagerCheck.List(0, col + 3)), CellByHeader(lo, 1, CStr(headers(col))), vbBinaryCompare) <> 0 Then Exit Function
            Next col
            CheckBaselineWorksheetFactForTest = True
    End Select
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Sub CheckBaselineStageHit(ByVal stage As String)
    mCheckBaselineStage = stage
    If stage = "ReusableState" Then
        mCheckBaselineEntries = mCheckBaselineEntries + 1
        If mCheckBaselineNest Then
            mCheckBaselineNest = False
            Call mForm.CheckBaselineActForTest("")
        End If
    End If
End Sub
Public Sub CheckBaselineArmNested()
    mCheckBaselineEntries = 0: mCheckBaselineNest = True
End Sub
Public Sub CheckBaselineResetOwnerEntries()
    mCheckBaselineEntries = 0: mCheckBaselineNest = False: mCheckBaselineStage = ""
End Sub
Public Sub CheckBaselineReopen(ByVal workbookName As String)
    Dim fixture As Variant
    fixture = mForm.CheckBaselineWorksheetSourceForTest()
    RunLocalReopen workbookName
    mForm.CheckBaselineRestoreWorksheetSourceForTest fixture
End Sub
Public Function CheckBaselinePermissionRefused() As Boolean
    CheckBaselinePermissionRefused = mForm.CheckBaselinePermissionRefusedForTest()
End Function
Public Function CheckBaselineOwnerEntries() As Long
    CheckBaselineOwnerEntries = mCheckBaselineEntries
End Function
Public Function CheckBaselineGuards() As Boolean
    CheckBaselineGuards = mForm.CheckBaselineGuardsForTest()
End Function
Public Function CheckBaselineContextRefused() As Boolean
    CheckBaselineContextRefused = mForm.CheckBaselineContextRefusedForTest()
End Function
Public Function CheckBaselineStage() As String
    CheckBaselineStage = mCheckBaselineStage
End Function
Public Function CheckBaselineReusableStage(ByVal mode As String) As Boolean
    CheckBaselineReusableStage = mForm.CheckBaselineReusableStageForTest(mode)
End Function
Public Function CheckBaselineAct(ByVal guard As String) As Boolean
    CheckBaselineAct = mForm.CheckBaselineActForTest(guard)
End Function
Public Function CheckBaselineReusableResult(ByVal checked As Boolean) As Boolean
    CheckBaselineReusableResult = mForm.CheckBaselineReusableResultForTest(checked)
End Function
Public Function CheckBaselineReusableDisplay(ByVal keyA As String, ByVal keyB As String) As Boolean
    CheckBaselineReusableDisplay = mForm.CheckBaselineReusableDisplayForTest(keyA, keyB)
End Function
Public Function CheckBaselineInventoryPrepare(ByVal firstKey As String, ByVal selectedKey As String) As Boolean
    CheckBaselineInventoryPrepare = mForm.CheckBaselineInventoryPrepareForTest(firstKey, selectedKey)
End Function
Public Function CheckBaselineInventoryState() As String
    CheckBaselineInventoryState = mForm.CheckBaselineInventoryStateForTest()
End Function
Public Function CheckBaselineWorksheetStage(ByVal selectedKey As String, ByVal canary As String) As Boolean
    CheckBaselineWorksheetStage = mForm.CheckBaselineWorksheetStageForTest(selectedKey, canary)
End Function
Public Sub CheckBaselineWorksheetInvalid(ByVal mode As String)
    mForm.CheckBaselineWorksheetInvalidForTest mode
End Sub
Public Function CheckBaselineWorksheetState() As String
    CheckBaselineWorksheetState = mForm.CheckBaselineWorksheetStateForTest()
End Function
Public Function CheckBaselineWorksheetFact(ByVal fact As String) As Boolean
    CheckBaselineWorksheetFact = mForm.CheckBaselineWorksheetFactForTest(fact)
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.InsertLines(1,'Private mCheckBaselineStage As String, mCheckBaselineEntries As Long, mCheckBaselineNest As Boolean')
}

function Test-ProductionCheckInBaseline($Fixture,$Other) {
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Hash([string]$Path){$s=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $s).Hash}finally{$s.Dispose()}}
    $canary='CHECKIN'+[guid]::NewGuid().ToString('N');$book=$null;$decoy=$null;$pins=@{}
    SelectTarget $Fixture
    $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($Fixture.Warehouse,'S1','config-admin'))
    if(-not $seed.StartsWith('OK|')){throw 'Admin Seed unavailable; not product RED.'}
    $ready=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.StockPrepareForTest' @($Fixture.Warehouse))
    if($ready -cne 'READY'){throw 'Two real received stock entities unavailable; not product RED.'}
    $keys=([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.StockKeysForTest')).Split("`t")
    if($keys.Count -ne 2 -or $keys[0] -ceq $keys[1]){throw 'Two distinct owner-generated identities required.'}
    try{
        SelectTarget $Fixture 'config-producer'
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$canary;$sheet.Cells.Item(2,2).Formula='=1+2'
        $path=Join-Path $runRoot 'check-in-baseline.xlsb';$book.SaveAs($path,50);$book.Close($false)
        $bookPin=Hash $path;$book=$excel.Workbooks.Open($path,0,$false);$sheet=$book.Worksheets.Item(1)
        $decoy=$excel.Workbooks.Add();$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name))
        if(-not [bool](Probe 'ReadPrepare' @($canary))){throw 'Saved/released reusable prerequisite unavailable; not product RED.'}
        [void](Probe 'RunLocalRememberFixture')
        Check 'CheckInBaseline.RealSeedReceivingAndReleasedDefinitions' $true
        foreach($root in @($Fixture.Root,$Other.Root)){
            foreach($file in Get-ChildItem -LiteralPath $root -Recurse -File|Where-Object{$_.Extension -in '.xlsb','.xlsm' -and $_.Name -notlike '~$*'}){$pins[$file.FullName]=Hash $file.FullName}
        }
        foreach($case in @(
            @{Mode='Selected';Guard='';Checked=$true},
            @{Mode='Insufficient';Guard='';Checked=$false},
            @{Mode='NoProcess';Guard='';Checked=$false},
            @{Mode='Selected';Guard='Loading';Checked=$false},
            @{Mode='Selected';Guard='Busy';Checked=$false}
        )){
            if(-not [bool](Probe 'CheckBaselineReusableStage' @($case.Mode))){throw 'Reusable Check In prerequisite unavailable; not product RED.'}
            $label='CheckInBaseline.Reusable.'+$case.Mode+$(if($case.Guard){'.'+$case.Guard}else{''})
            $visible=$CaptureEvidence -and -not $case.Guard -and $case.Mode -cin @('Selected','NoProcess')
            if($visible){[void](Probe 'RunLocalShowAndCapture' @($book.Name,'CHECK_IN'))}
            Check ($label+'.ActualHandlerReturned') ([bool](Probe 'CheckBaselineAct' @($case.Guard)))
            Check ($label+'.CheckedStateNoteFreezeAndNoCompletion') ([bool](Probe 'CheckBaselineReusableResult' @($case.Checked)))
            Check ($label+'.GuardsRestoredWithoutAdapterReset') ([bool](Probe 'CheckBaselineGuards'))
            if($case.Mode -ceq 'Selected' -and -not $case.Guard){Check ($label+'.IdentityUnderMatchingHeading') ([bool](Probe 'CheckBaselineReusableDisplay' @($keys[0],$keys[1])))}
            if($visible){CaptureOwnedFormByCaptionEvidence 'Production' ('check-in-'+$case.Mode.ToLowerInvariant()+'.png')}
        }
        if(-not [bool](Probe 'CheckBaselineReusableStage' @('Selected'))){throw 'Nested Check In prerequisite unavailable; not product RED.'}
        [void](Probe 'CheckBaselineArmNested')
        Check 'CheckInBaseline.Nested.ActualHandlerReturned' ([bool](Probe 'CheckBaselineAct' @('')))
        Check 'CheckInBaseline.Nested.OnlyOneOwnerEntry' ([int](Probe 'CheckBaselineOwnerEntries') -eq 1)
        Check 'CheckInBaseline.Nested.ExpectedOwnerResult' ([bool](Probe 'CheckBaselineReusableResult' @($true)))
        Check 'CheckInBaseline.Nested.GuardsRestoredWithoutAdapterReset' ([bool](Probe 'CheckBaselineGuards'))
        if(-not [bool](Probe 'RunWorksheetPrepare')){throw 'Worksheet Check In prerequisite unavailable; not product RED.'}
        # Select the other real entity from the one the current bridge returns first.
        # No replacement identity is created or inferred from SKU.
        $firstKey=[string](Probe 'RunWorksheetKey')
        if($firstKey -cnotin $keys){throw 'Worksheet source identity missing from real receiving fixture.'}
        $selectedKey=@($keys|Where-Object{$_ -cne $firstKey})[0]
        if(-not [bool](Probe 'CheckBaselineWorksheetStage' @($selectedKey,$canary))){throw 'Complete exact-key worksheet allocation unavailable; not product RED.'}
        if($CaptureEvidence){[void](Probe 'RunLocalShowAndCapture' @($book.Name,'CHECK_IN'))}
        Check 'CheckInBaseline.Worksheet.ActualHandlerReturned' ([bool](Probe 'CheckBaselineAct' @('')))
        $stage=[string](Probe 'CheckBaselineStage')
        if($stage -cnotin @('ReusableState','LoaderSelection','PalettePresence','LocationCleanup','AllocationCompleteness','AllocationLocations','BuildPayload','WriteCheckRows','RefreshManager')){throw 'Unknown diagnostic stage; raw output suppressed.'}
        [pscustomobject]@{Stage=$stage;SelectedKeyDiffersFromFirst=$selectedKey -cne $firstKey}|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'check-in-boundary.json')
        foreach($fact in @('ReachedCheckRows','ExactSelectedKey','CustomValue','CustomFormula','PalettePreserved','HeadersPreserved')){
            Check ('CheckInBaseline.Worksheet.'+$fact) ([bool](Probe 'CheckBaselineWorksheetFact' @($fact)))
        }
        Check 'CheckInBaseline.Worksheet.ValuesUnderMatchingHeadings' ([bool](Probe 'CheckBaselineWorksheetFact' @('DisplayColumns')))
        if($CaptureEvidence){CaptureOwnedFormByCaptionEvidence 'Production' 'check-in-worksheet.png'}
        foreach($mode in @('MissingKey','UnknownKey','MissingUsed','MissingIdentityHeader')){
            if(-not [bool](Probe 'CheckBaselineWorksheetStage' @($selectedKey,$canary))){throw 'Negative worksheet prerequisite unavailable; not product RED.'}
            [void](Probe 'CheckBaselineWorksheetInvalid' @($mode))
            $before=[string](Probe 'CheckBaselineWorksheetState');$label='CheckInBaseline.Worksheet.'+$mode
            Check ($label+'.ActualHandlerReturned') ([bool](Probe 'CheckBaselineAct' @('')))
            Check ($label+'.CheckTablePreserved') ([string](Probe 'CheckBaselineWorksheetState') -ceq $before)
            Check ($label+'.NoSuccessStatus') ([bool](Probe 'CheckBaselineWorksheetFact' @('NoSuccessStatus')))
        }
        if(-not [bool](Probe 'CheckBaselineInventoryPrepare' @($firstKey,$selectedKey))){throw 'Two-entity local inventory projection unavailable; not product RED.'}
        if(-not [bool](Probe 'CheckBaselineWorksheetStage' @($selectedKey,$canary))){throw 'Local-source Check In prerequisite unavailable; not product RED.'}
        $inventoryBefore=[string](Probe 'CheckBaselineInventoryState')
        Check 'CheckInBaseline.LocalProjection.ActualHandlerReturned' ([bool](Probe 'CheckBaselineAct' @('')))
        foreach($fact in @('ReachedCheckRows','ExactSelectedKey','CustomValue','CustomFormula','DisplayColumns')){
            Check ('CheckInBaseline.LocalProjection.'+$fact) ([bool](Probe 'CheckBaselineWorksheetFact' @($fact)))
        }
        Check 'CheckInBaseline.LocalProjection.InventoryValuesAndFormulasPreserved' ([string](Probe 'CheckBaselineInventoryState') -ceq $inventoryBefore)
        SelectTarget $Fixture 'config-reader'
        [void](Probe 'CheckBaselineReopen' @($book.Name))
        foreach($mode in @('Reusable','Worksheet')){
            $ready=if($mode -ceq 'Reusable'){[bool](Probe 'CheckBaselineReusableStage' @('Selected'))}else{[bool](Probe 'CheckBaselineWorksheetStage' @($selectedKey,$canary))}
            if(-not $ready){throw 'Permission Check In prerequisite unavailable; not product RED.'}
            $ownerBefore=[string](Probe 'RunLocalOwnerState')
            $projectionBefore=[string](Probe 'RunLocalState')+'|'+[string](Probe 'CheckBaselineWorksheetState')
            [void](Probe 'CheckBaselineResetOwnerEntries')
            $label='CheckInBaseline.Permission.'+$mode
            Check ($label+'.ActualHandlerReturned') ([bool](Probe 'CheckBaselineAct' @('')))
            Check ($label+'.OwnerNotEntered') ([int](Probe 'CheckBaselineOwnerEntries') -eq 0)
            Check ($label+'.OwnerStatePreserved') ([string](Probe 'RunLocalOwnerState') -ceq $ownerBefore)
            Check ($label+'.ProjectionPreserved') (([string](Probe 'RunLocalState')+'|'+[string](Probe 'CheckBaselineWorksheetState')) -ceq $projectionBefore)
            Check ($label+'.VisiblePermissionRefusal') ([bool](Probe 'CheckBaselinePermissionRefused'))
            Check ($label+'.GuardsRestored') ([bool](Probe 'CheckBaselineGuards'))
        }
        Test-ProductionCheckInYield $Fixture $Other $book $selectedKey $canary
        foreach($guard in @('Target','Session','SignedOut')){
            SelectTarget $Fixture 'config-producer'
            [void](Probe 'RunLocalReopen' @($book.Name))
            if(-not [bool](Probe 'CheckBaselineReusableStage' @('Selected'))){throw 'Captured-context Check In prerequisite unavailable; not product RED.'}
            $before=[string](Probe 'RunLocalOwnerState');$label='CheckInBaseline.Context.'+$guard
            if($guard -ceq 'Target'){SelectTarget $Other 'config-producer'}
            if($guard -ceq 'Session'){SelectTarget $Fixture 'config-producer'}
            if($guard -ceq 'SignedOut'){[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
            Check ($label+'.ActualHandlerReturned') ([bool](Probe 'CheckBaselineAct' @('')))
            Check ($label+'.OwnerPreserved') ([string](Probe 'RunLocalOwnerState') -ceq $before)
            Check ($label+'.VisibleContextRefusal') ([bool](Probe 'CheckBaselineContextRefused'))
        }
        SelectTarget $Fixture 'config-producer'
        Check 'CheckInBaseline.CanonicalExactEntitiesUnchanged' ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.StockSourcePreservedForTest'))
        $same=$true;foreach($file in $pins.Keys){$same=$same -and (Hash $file) -ceq $pins[$file]}
        Check 'CheckInBaseline.SavedAuthorityPreserved' $same
        Check 'CheckInBaseline.CapturedBookExtraValues' ($sheet.Cells.Item(2,1).Value2 -ceq $canary -and $sheet.Cells.Item(2,2).Formula -ceq '=1+2')
        [void](Probe 'CloseDesigner');$book.Close($false);$book=$null
        Check 'CheckInBaseline.OperatorBytesPreserved' ((Hash $path) -ceq $bookPin)
    }finally{
        [void](Probe 'CloseDesigner')
        if($null -ne $book){$book.Close($false)}
        if($null -ne $decoy){$decoy.Close($false)}
        SelectTarget $Fixture
    }
}
