# D14/D15 owner-boundary baseline. No new activity catalog contract is assumed.
# Adapters stage disposable projections and invoke the unchanged operator handler.
function Install-ProductionCheckInBaselineProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $start=$form.ProcStartLine('CheckInProductionRun',0);$end=$start+$form.ProcCountLines('CheckInProductionRun',0)
    $hits=@(for($line=$start;$line -lt $end;$line++){if($form.Lines($line,1).Trim() -match '^checkInStage = "[A-Za-z]+"$'){$line}})
    if($hits.Count -ne 9){throw 'Check In diagnostic stage anchors changed; not product RED.'}
    foreach($line in @($hits|Sort-Object -Descending)){$form.InsertLines($line+1,'    TestProductionDesigner.CheckBaselineStageHit checkInStage')}
    $form.InsertLines(1,'Private mCheckBaselineKey As String, mCheckBaselineCanary As String')
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
    Dim priorLoading As Boolean, priorBusy As Boolean
    On Error GoTo Failed
    priorLoading = mLoading: priorBusy = mDesignerActionInProgress
    If guard = "Loading" Then mLoading = True
    If guard = "Busy" Then mDesignerActionInProgress = True
    mBtnManagerCheckIn_Click
    CheckBaselineActForTest = True
Failed:
    mLoading = priorLoading: mDesignerActionInProgress = priorBusy
End Function
Public Function CheckBaselineReusableResultForTest(ByVal checked As Boolean) As Boolean
    CheckBaselineReusableResultForTest = _
        modProductionReusableRun.ReusableRunIsCheckedIn() = checked And _
        modProductionReusableRun.ReusableRunBatchNoteFrozen() = checked And _
        Not modProductionReusableRun.ReusableRunIsCompleted() And _
        modProductionReusableRun.ReusableRunBatchNote() = mReadCanaryForTest
End Function
Public Function CheckBaselineWorksheetStageForTest(ByVal selectedKey As String, ByVal canary As String) As Boolean
    Dim ws As Worksheet, lo As ListObject, headers As Variant, col As Long
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
Public Function CheckBaselineWorksheetFactForTest(ByVal fact As String) As Boolean
    Dim lo As ListObject, ws As Worksheet
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
    End Select
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Sub CheckBaselineStageHit(ByVal stage As String)
    mCheckBaselineStage = stage
End Sub
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
Public Function CheckBaselineWorksheetStage(ByVal selectedKey As String, ByVal canary As String) As Boolean
    CheckBaselineWorksheetStage = mForm.CheckBaselineWorksheetStageForTest(selectedKey, canary)
End Function
Public Function CheckBaselineWorksheetFact(ByVal fact As String) As Boolean
    CheckBaselineWorksheetFact = mForm.CheckBaselineWorksheetFactForTest(fact)
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.InsertLines(1,'Private mCheckBaselineStage As String')
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
            Check ($label+'.ActualHandlerReturned') ([bool](Probe 'CheckBaselineAct' @($case.Guard)))
            Check ($label+'.CheckedStateNoteFreezeAndNoCompletion') ([bool](Probe 'CheckBaselineReusableResult' @($case.Checked)))
        }
        if(-not [bool](Probe 'RunWorksheetPrepare')){throw 'Worksheet Check In prerequisite unavailable; not product RED.'}
        # Select the other real entity from the one the current bridge returns first.
        # No replacement identity is created or inferred from SKU.
        $firstKey=[string](Probe 'RunWorksheetKey')
        if($firstKey -cnotin $keys){throw 'Worksheet source identity missing from real receiving fixture.'}
        $selectedKey=@($keys|Where-Object{$_ -cne $firstKey})[0]
        if(-not [bool](Probe 'CheckBaselineWorksheetStage' @($selectedKey,$canary))){throw 'Complete exact-key worksheet allocation unavailable; not product RED.'}
        Check 'CheckInBaseline.Worksheet.ActualHandlerReturned' ([bool](Probe 'CheckBaselineAct' @('')))
        $stage=[string](Probe 'CheckBaselineStage')
        if($stage -cnotin @('ReusableState','LoaderSelection','PalettePresence','LocationCleanup','AllocationCompleteness','AllocationLocations','BuildPayload','WriteCheckRows','RefreshManager')){throw 'Unknown diagnostic stage; raw output suppressed.'}
        [pscustomobject]@{Stage=$stage;SelectedKeyDiffersFromFirst=$selectedKey -cne $firstKey}|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'check-in-boundary.json')
        foreach($fact in @('ReachedCheckRows','ExactSelectedKey','CustomValue','CustomFormula','PalettePreserved','HeadersPreserved')){
            Check ('CheckInBaseline.Worksheet.'+$fact) ([bool](Probe 'CheckBaselineWorksheetFact' @($fact)))
        }
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
