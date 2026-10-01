# Owned operator staging uses an Admin-created identity. Original Apply handlers remain intact.
function Install-ProductionRunWorksheetProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $form.InsertLines(1,@'
Private mRunSheetKey As String, mRunSheetLocation As String, mRunSheetUom As String
Private mRunSheetAvailable As Double, mRunSheetCanary As String
'@)
    $form.AddFromString(@'
Public Function RunWorksheetPrepareForTest() As Boolean
    Dim entities As Variant, row As Long, ws As Worksheet
    entities = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge("DEMO-RAW-BLACK-TEA")
    If Not IsArray(entities) Then Exit Function
    For row = LBound(entities, 1) To UBound(entities, 1)
        If CStr(entities(row, 3)) = "DEMO-RAW-BLACK-TEA" Then
            mRunSheetKey = CStr(entities(row, 1)): mRunSheetLocation = CStr(entities(row, 7))
            mRunSheetUom = CStr(entities(row, 5)): mRunSheetAvailable = CDbl(entities(row, 6))
            Exit For
        End If
    Next row
    If mRunSheetKey = "" Or mRunSheetLocation = "" Or mRunSheetAvailable < 10# Then Exit Function
    If WorkbookHasSheet(mOperatorWorkbook, "Production") Then Exit Function
    Set ws = mOperatorWorkbook.Worksheets.Add
    ws.Name = "Production"
    RunWorksheetPrepareForTest = True
End Function
Public Function RunWorksheetStageForTest(ByVal action As String, ByVal mode As String, ByVal canary As String) As String
    Dim ws As Worksheet, lo As ListObject, headers As Variant, col As Long, idx As Long, phase As String
    Dim splitBox As MSForms.TextBox, qtyBox As MSForms.TextBox
    On Error GoTo Failed
    phase = "ClearReusable"
    modProductionReusableRun.ClearReusableRun
    mLoading = True: mRunSheetCanary = canary
    If action = "TREE_ALLOCATE" Then mPages.Value = 4 Else mPages.Value = 3
    phase = "OwnedStaging"
    Set ws = mOperatorWorkbook.Worksheets("Production")
    If ws.ListObjects.Count > 0 Then ws.ListObjects("InventoryPalette_RunFixture").Unlist
    ws.Range("A1:S2").Clear
    headers = Array("Operator Extra", "QUANTITY", "System_Key", "Operator Formula", "ITEM_CODE", _
                    "SPLIT %", "BASE QUANTITY", "LOCATION", "UOM", "INGREDIENT", "INGREDIENT_ID", _
                    "VENDORS", "VENDOR_CODE", "DESCRIPTION", "ITEM", "PERCENT", "PROCESS", "INPUT/OUTPUT")
    For col = 0 To UBound(headers)
        ws.Cells(1, col + 1).Value2 = headers(col)
    Next col
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:R2"), , xlYes)
    lo.Name = "InventoryPalette_RunFixture"
    ws.Range("A2").Value2 = canary: ws.Range("B2").Value2 = 2.5
    ws.Range("C2").Value2 = mRunSheetKey: ws.Range("D2").Formula = "=1+2"
    ws.Range("E2").Value2 = "DEMO-RAW-BLACK-TEA": ws.Range("F2").Value2 = 25#
    ws.Range("G2").Value2 = 10#: ws.Range("H2").Value2 = mRunSheetLocation
    ws.Range("I2").Value2 = mRunSheetUom: ws.Range("J2").Value2 = canary
    ws.Range("K2").Value2 = "A02": ws.Range("O2").Value2 = canary
    ws.Range("Q2").Value2 = canary: ws.Range("R2").Value2 = "INPUT"
    phase = "Palette"
    EnsureRunSplitOverrides
    mRunSplitOverrides.RemoveAll
    EnsureRunBaseQtyMap
    mRunBaseQtyByKey.RemoveAll
    EnsureRunProcessMap
    mRunProcessByKey.RemoveAll
    EnsureRunTreeState
    mRunTreeCollapsed.RemoveAll
    mLstRunPalette.Clear: mLstRunTree.Clear: mLstManagerOutput.Clear
    mCmbRunProcess.Clear: mCmbTreeRunProcess.Clear
    mCmbRunLocation.Clear: mCmbTreeRunLocation.Clear
    mCmbRunLocation.AddItem mRunSheetLocation: mCmbTreeRunLocation.AddItem mRunSheetLocation
    mCmbRunLocation.ListIndex = 0: mCmbTreeRunLocation.ListIndex = 0
    If mode = "WrongLocation" Then
        mCmbRunLocation.AddItem "RUN-OTHER": mCmbTreeRunLocation.AddItem "RUN-OTHER"
        mCmbRunLocation.ListIndex = 1: mCmbTreeRunLocation.ListIndex = 1
    End If
    mLstRunPalette.AddItem lo.Name
    mLstRunPalette.List(0, 1) = "1": mLstRunPalette.List(0, 2) = canary
    mLstRunPalette.List(0, 3) = mRunSheetKey: mLstRunPalette.List(0, 4) = canary
    mLstRunPalette.List(0, 5) = "25": mLstRunPalette.List(0, 6) = "2.5"
    mLstRunPalette.List(0, 7) = mRunSheetUom: mLstRunPalette.List(0, 8) = CStr(mRunSheetAvailable)
    mLstRunPalette.List(0, 9) = mRunSheetLocation
    StoreRunProcess mLstRunPalette, 0, canary
    StoreRunBaseQty mLstRunPalette, 0, "10"
    StoreRunItemCode mLstRunPalette, 0, "DEMO-RAW-BLACK-TEA"
    StoreRunAllocationOverride mLstRunPalette, 0, "25", "2.5"
    BuildRunTreeFromPaletteList
    mLstRunPalette.ListIndex = 0
    If mLstRunTree.ListCount <> 3 Then GoTo Unavailable
    mLstRunTree.ListIndex = 2
    phase = "Inputs"
    Set splitBox = ActiveRunSplitTextBox(): Set qtyBox = ActiveRunQtyTextBox()
    mUpdatingPaletteInputs = True
    splitBox.Text = "75": qtyBox.Text = "5": mPaletteInputSource = "QTY"
    If mode = "Percent" Then mPaletteInputSource = "SPLIT"
    If mode = "QtyOnly" Then splitBox.Text = "": mPaletteInputSource = ""
    If mode = "Zero" Then qtyBox.Text = "0"
    If mode = "OverTotal" Then qtyBox.Text = "11"
    ' Inactive inputs differ deliberately, protecting use of the visible Run page.
    If action = "TREE_ALLOCATE" Then
        mTxtPaletteSplit.Text = "10": mTxtPaletteQty.Text = "1"
    Else
        mTxtTreePaletteSplit.Text = "10": mTxtTreePaletteQty.Text = "1"
    End If
    mUpdatingPaletteInputs = False
    If mode = "MissingTable" Then lo.Unlist
    If mode = "MissingQuantity" Then lo.ListColumns("QUANTITY").Name = "Operator Quantity"
    mLoading = False
    RunWorksheetStageForTest = "READY"
    Exit Function
Unavailable:
    RunWorksheetStageForTest = "FIXTURE_FAILED|" & phase
    GoTo Finished
Failed:
    RunWorksheetStageForTest = "FIXTURE_FAILED|" & phase & "|" & CStr(Err.Number)
Finished:
    mUpdatingPaletteInputs = False: mLoading = False
End Function
Public Function RunWorksheetBranchForTest(ByVal action As String) As Boolean
    RunWorksheetBranchForTest = Not modProductionReusableRun.ReusableRunIsLoaded() And _
        mPages.Value = IIf(action = "TREE_ALLOCATE", 4, 3)
End Function
Public Function RunWorksheetPreservedForTest(ByVal mode As String) As Boolean
    Dim ws As Worksheet
    Set ws = mOperatorWorkbook.Worksheets("Production")
    RunWorksheetPreservedForTest = ws.Range("A1").Value2 = "Operator Extra" And _
        ws.Range("A2").Value2 = mRunSheetCanary And ws.Range("D1").Value2 = "Operator Formula" And _
        ws.Range("D2").Formula = "=1+2" And ws.Range("C1").Value2 = "System_Key" And _
        StrComp(CStr(ws.Range("C2").Value2), mRunSheetKey, vbBinaryCompare) = 0
    If mode = "MissingQuantity" Then RunWorksheetPreservedForTest = RunWorksheetPreservedForTest And _
        ws.Range("B1").Value2 = "Operator Quantity" And ws.Range("B2").Value2 = 2.5
    If mode = "MissingTable" Then RunWorksheetPreservedForTest = RunWorksheetPreservedForTest And ws.ListObjects.Count = 0
End Function
Public Function RunWorksheetResultForTest(ByVal mode As String) As Boolean
    Dim ws As Worksheet, qty As Double, pct As Double, row As Long, key As String, values As Variant
    Dim splitBox As MSForms.TextBox, qtyBox As MSForms.TextBox
    If mode = "MissingTable" Or mode = "MissingQuantity" Then
        RunWorksheetResultForTest = (InStr(1, mTxtStatus.Text, "allocation updated", vbTextCompare) = 0)
        Exit Function
    End If
    Set ws = mOperatorWorkbook.Worksheets("Production")
    Set splitBox = ActiveRunSplitTextBox(): Set qtyBox = ActiveRunQtyTextBox()
    qty = 5#: pct = 50#
    If mode = "Percent" Then qty = 7.5: pct = 75#
    If mode = "Zero" Then qty = 0#: pct = 0#
    If mode = "OverTotal" Then qty = 2.5: pct = 25#
    If mode = "WrongLocation" Then
        If CStr(ws.Range("B2").Value2) <> "" Or CStr(ws.Range("F2").Value2) <> "" Then Exit Function
        If NzStr(mLstRunPalette.List(0, 5)) <> "" Or NzStr(mLstRunPalette.List(0, 6)) <> "" Then Exit Function
    Else
        If ws.Range("B2").Value2 <> qty Or ws.Range("F2").Value2 <> pct Then Exit Function
        If CDbl(mLstRunPalette.List(0, 5)) <> pct Or CDbl(mLstRunPalette.List(0, 6)) <> qty Then Exit Function
    End If
    If mode = "OverTotal" Then
        If CDbl(splitBox.Text) <> 110# Or CDbl(qtyBox.Text) <> 11# Then Exit Function
    ElseIf mode = "WrongLocation" Then
        If CDbl(splitBox.Text) <> 50# Or CDbl(qtyBox.Text) <> 5# Then Exit Function
    Else
        If CDbl(splitBox.Text) <> pct Or CDbl(qtyBox.Text) <> qty Then Exit Function
    End If
    key = RunAllocationKeyFromList(mLstRunPalette, 0)
    If Not mRunSplitOverrides.Exists(key) Then Exit Function
    values = mRunSplitOverrides(key)
    If CStr(values(0)) <> NzStr(mLstRunPalette.List(0, 5)) Or CStr(values(1)) <> NzStr(mLstRunPalette.List(0, 6)) Then Exit Function
    For row = 0 To mLstRunTree.ListCount - 1
        If NzStr(mLstRunTree.List(row, 0)) = "InventoryPalette_RunFixture" Then
            If StrComp(NzStr(mLstRunTree.List(row, 3)), mRunSheetKey, vbBinaryCompare) <> 0 Then Exit Function
            If NzStr(mLstRunTree.List(row, 5)) <> NzStr(mLstRunPalette.List(0, 5)) Or _
                NzStr(mLstRunTree.List(row, 6)) <> NzStr(mLstRunPalette.List(0, 6)) Then Exit Function
            RunWorksheetResultForTest = True
            Exit Function
        End If
    Next row
End Function
Public Function RunWorksheetSourceForTest() As Boolean
    Dim entities As Variant, row As Long
    entities = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge("DEMO-RAW-BLACK-TEA")
    If Not IsArray(entities) Then Exit Function
    For row = LBound(entities, 1) To UBound(entities, 1)
        If StrComp(CStr(entities(row, 1)), mRunSheetKey, vbBinaryCompare) = 0 Then
            RunWorksheetSourceForTest = CDbl(entities(row, 6)) = mRunSheetAvailable And _
                CStr(entities(row, 7)) = mRunSheetLocation And CStr(entities(row, 5)) = mRunSheetUom
            Exit Function
        End If
    Next row
End Function
Public Function RunWorksheetKeyForTest() As String
    RunWorksheetKeyForTest = mRunSheetKey
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function RunWorksheetPrepare() As Boolean
    RunWorksheetPrepare = mForm.RunWorksheetPrepareForTest()
End Function
Public Function RunWorksheetStage(ByVal action As String, ByVal mode As String, ByVal canary As String) As String
    RunWorksheetStage = mForm.RunWorksheetStageForTest(action, mode, canary)
End Function
Public Function RunWorksheetBranch(ByVal action As String) As Boolean
    RunWorksheetBranch = mForm.RunWorksheetBranchForTest(action)
End Function
Public Function RunWorksheetPreserved(ByVal mode As String) As Boolean
    RunWorksheetPreserved = mForm.RunWorksheetPreservedForTest(mode)
End Function
Public Function RunWorksheetResult(ByVal mode As String) As Boolean
    RunWorksheetResult = mForm.RunWorksheetResultForTest(mode)
End Function
Public Function RunWorksheetSource() As Boolean
    RunWorksheetSource = mForm.RunWorksheetSourceForTest()
End Function
Public Function RunWorksheetKey() As String
    RunWorksheetKey = mForm.RunWorksheetKeyForTest()
End Function
'@)
}

function Test-ProductionRunWorksheet($Fixture) {
    if(-not [bool](Probe 'RunWorksheetPrepare')){throw 'Admin-seeded worksheet allocation fixture unavailable; not product RED.'}
    Check 'RunWorksheet.OwningSeedIdentity' $true
    $key=[string](Probe 'RunWorksheetKey')
    foreach($action in @('ALLOCATE','TREE_ALLOCATE')){
        foreach($mode in @('Quantity','Percent','QtyOnly','Zero','OverTotal','WrongLocation','MissingTable','MissingQuantity')){
            $before=@(Files);$ready=[string](Probe 'RunWorksheetStage' @($action,$mode,$canary))
            if($ready -cne 'READY'){
                if($ready -notmatch '^FIXTURE_FAILED\|[A-Za-z]+(\|-?[0-9]+)?$'){$ready='Unavailable'}
                throw ('Worksheet staging unavailable: '+$ready+'; not product RED.')
            }
            $decoy.Activate();$label='RunWorksheet.'+$action+'.'+$mode
            Check ($label+'.SetupNotUserAction') (@(Files).Count -eq $before.Count)
            Check ($label+'.ActualWorksheetBranchAndPage') ([bool](Probe 'RunWorksheetBranch' @($action)))
            $notice=[string](Probe 'RunFaultAct' @($action))
            Check ($label+'.NoUnhandledError') (-not $notice.StartsWith('HANDLER_ERROR|'))
            Check ($label+'.GuardsRestoredWithoutAdapterReset') ([bool](Probe 'RunFaultGuards'))
            $name=if($mode.StartsWith('Missing')){'UnavailableSurfaceNotReportedComplete'}else{'ExistingWritesMirroringAndOverrides'}
            Check ($label+'.'+$name) ([bool](Probe 'RunWorksheetResult' @($mode)))
            Check ($label+'.CustomColumnsFormulaAndExactKey') ([bool](Probe 'RunWorksheetPreserved' @($mode)))
            Check ($label+'.CanonicalSourcePreserved') ([bool](Probe 'RunWorksheetSource'))
            $outcome=if($mode.StartsWith('Missing')){'FAILED'}elseif($mode -cin @('OverTotal','WrongLocation')){'REJECTED'}else{'STAGED'}
            Pair $before $action $outcome $label
            $records=@(Files|Where-Object{$_ -cnotin $before}|ForEach-Object{[IO.File]::ReadAllText($_)})
            Check ($label+'.NoExactKeyInObservations') ($records.Count -eq 2 -and @($records|Where-Object{$_.Contains($key)}).Count -eq 0)
        }
    }
}
