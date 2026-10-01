# Real receiving creates a second immutable entity before measured local planning.
function Install-ProductionRunStockProbe {
    $core=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestShippingCatalog').CodeModule
    $core.InsertLines(1,'Private mRunStockKeyA As String, mRunStockKeyB As String, mRunStockLocation As String')
    $core.InsertLines(1,'Private mRunStockQty As Double')
    $core.AddFromString(@'
Public Function StockPrepareForTest(ByVal warehouseId As String) As String
    Dim entities As Variant, row As Long, count As Long, eventId As String, report As String
    Dim target As WarehouseTarget, phase As String
    On Error GoTo Failed
    phase = "Target": Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then GoTo Unavailable
    If target.WarehouseId <> warehouseId Then GoTo Unavailable
    phase = "SeedRead": entities = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge("DEMO-RAW-BLACK-TEA")
    If Not IsArray(entities) Then GoTo Unavailable
    phase = "SeedShape"
    For row = LBound(entities, 1) To UBound(entities, 1)
        If CStr(entities(row, 3)) = "DEMO-RAW-BLACK-TEA" Then
            count = count + 1: mRunStockKeyA = CStr(entities(row, 1)): mRunStockLocation = CStr(entities(row, 7))
            phase = "SeedQuantity"
            mRunStockQty = CDbl(entities(row, 6))
            If mRunStockQty <= 3# Then GoTo Unavailable
        End If
    Next row
    phase = "SeedCount": If count <> 1 Then GoTo Unavailable
    phase = "SeedIdentity": If mRunStockKeyA = "" Then GoTo Unavailable
    phase = "SeedLocation": If mRunStockLocation = "" Then GoTo Unavailable
    phase = "QueueReceive": mRunStockKeyB = ""
    If Not modRoleEventWriter.QueueReceiveEventServer(warehouseId, "S1", "config-admin", _
        "DEMO-RAW-BLACK-TEA", mRunStockQty, mRunStockLocation, "Disposable Run stock fixture", _
        eventId, report, "", mRunStockKeyB, "GOOD", "") Then GoTo Unavailable
    If mRunStockKeyB = "" Or StrComp(mRunStockKeyA, mRunStockKeyB, vbBinaryCompare) = 0 Then GoTo Unavailable
    phase = "ProcessReceive"
    If modProcessor.RunBatch(warehouseId, 0, report) <> 1 Then GoTo Unavailable
    phase = "TwoExactKeys"
    If Not StockSourcePreservedForTest() Then GoTo Unavailable
    StockPrepareForTest = "READY"
    Exit Function
Unavailable:
    StockPrepareForTest = "FIXTURE_FAILED|" & phase
    Exit Function
Failed:
    StockPrepareForTest = "FIXTURE_FAILED|" & phase & "|" & CStr(Err.Number)
End Function
Public Function StockKeysForTest() As String
    StockKeysForTest = mRunStockKeyA & vbTab & mRunStockKeyB
End Function
Public Function StockQtyForTest() As Double
    StockQtyForTest = mRunStockQty
End Function
Public Function StockSourcePreservedForTest() As Boolean
    Dim entities As Variant, row As Long, seen As Object, key As String
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbBinaryCompare
    entities = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge("DEMO-RAW-BLACK-TEA")
    If Not IsArray(entities) Then Exit Function
    For row = LBound(entities, 1) To UBound(entities, 1)
        If CStr(entities(row, 3)) = "DEMO-RAW-BLACK-TEA" Then
            key = CStr(entities(row, 1))
            If key <> mRunStockKeyA And key <> mRunStockKeyB Then Exit Function
            If seen.Exists(key) Or CDbl(entities(row, 6)) <> mRunStockQty Or CStr(entities(row, 7)) <> mRunStockLocation Then Exit Function
            seen.Add key, True
        End If
    Next row
    StockSourcePreservedForTest = (seen.Count = 2)
End Function
'@)
    $project=$packages['invSys.Operations.xlam'].VBProject
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $form.InsertLines(1,'Private mRunStockQtyForTest As Double')
    $start=$form.ProcStartLine('DesignerReleasedProcessForTest',0);$end=$start+$form.ProcCountLines('DesignerReleasedProcessForTest',0)
    $matches=@(for($line=$start;$line -lt $end;$line++){if($form.Lines($line,1).Trim() -ceq 'mLstProcessRequirements.List(0, 2) = "2"'){$line}})
    if($matches.Count -ne 1){throw 'Stock requirement fixture anchor changed; not product RED.'}
    $form.ReplaceLine($matches[0],'    mLstProcessRequirements.List(0, 2) = CStr(2# * mRunStockQtyForTest + 2#)')
    $form.AddFromString(@'
Public Sub RunStockConfigureForTest(ByVal sourceQty As Double)
    mRunStockQtyForTest = sourceQty
End Sub
Public Function RunStockStageForTest(ByVal action As String, ByVal mode As String) As Boolean
    Dim report As String
    If RunLocalStageForTest(action, "Normal") <> "READY" Then Exit Function
    If mode = "Zero" Or mode = "OverBucket" Then
        If Not modProductionReusableRun.ApplyReusableRunStockAllocation( _
            NzStr(mLstRunPalette.List(0, 0)), NzStr(mLstRunPalette.List(0, 1)), _
            NzStr(mLstRunPalette.List(0, 3)), mRunStockQtyForTest + 1#, report) Then Exit Function
        RefreshReusableRunControls False
        mLoading = True: mLstRunPalette.ListIndex = 0: mLoading = False
    End If
    mUpdatingPaletteInputs = True
    mTxtPaletteSplit.Text = "75": mTxtPaletteQty.Text = CStr(mRunStockQtyForTest + 1#)
    If mode = "Percent" Then mTxtPaletteQty.Text = ""
    If mode = "Zero" Then mTxtPaletteQty.Text = "0"
    If mode = "OverBucket" Then mTxtPaletteQty.Text = CStr(2# * mRunStockQtyForTest + 1#)
    mUpdatingPaletteInputs = False
    RunStockStageForTest = (mLstRunPalette.ListCount = 1)
End Function
'@)
    $project.VBComponents.Item('modProductionReusableRun').CodeModule.AddFromString(@'
Public Function RunStockPlanForTest(ByVal keyA As String, ByVal keyB As String, _
                                    ByVal expectedQty As Double, ByVal expectedCount As Long, ByVal sourceQty As Double) As Boolean
    Dim allocation As Variant, key As String, total As Double, qty As Double
    If mAllocations.Count <> expectedCount Or mCheckedIn Or mCompleted Or mEventIds <> "" Then Exit Function
    For Each allocation In mAllocations.Keys
        key = AllocationSystemKey(CStr(allocation)): qty = CDbl(mAllocations(allocation))
        If StrComp(key, keyA, vbBinaryCompare) <> 0 And StrComp(key, keyB, vbBinaryCompare) <> 0 Then Exit Function
        If qty <= 0# Or qty > sourceQty Then Exit Function
        total = total + qty
    Next allocation
    RunStockPlanForTest = (Abs(total - expectedQty) < 0.0000001)
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Sub RunStockConfigure(ByVal sourceQty As Double)
    mForm.RunStockConfigureForTest sourceQty
End Sub
Public Function RunStockStage(ByVal action As String, ByVal mode As String) As Boolean
    RunStockStage = mForm.RunStockStageForTest(action, mode)
End Function
Public Function RunStockPlan(ByVal keyA As String, ByVal keyB As String, ByVal qty As Double, ByVal count As Long, ByVal sourceQty As Double) As Boolean
    RunStockPlan = modProductionReusableRun.RunStockPlanForTest(keyA, keyB, qty, count, sourceQty)
End Function
'@)
}

function Test-ProductionRunStock($Fixture) {
    # Parent scope supplies the actual adapter, pair assertions and preservation pins.
    $keys=([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.StockKeysForTest')).Split("`t")
    $sourceQty=[double](Run 'invSys.Core.xlam' 'TestShippingCatalog.StockQtyForTest')
    if($keys.Count -ne 2 -or -not $keys[0] -or -not $keys[1] -or $keys[0] -ceq $keys[1]){throw 'Two owner-created stock identities unavailable; not product RED.'}
    foreach($action in @('ALLOCATE','TREE_ALLOCATE')){
        foreach($mode in @('Quantity','Percent','Zero','OverBucket')){
            $before=@(Files)
            if(-not [bool](Probe 'RunStockStage' @($action,$mode))){throw 'Released multi-key stock fixture unavailable; not product RED.'}
            $label='RunStock.'+$action+'.'+$mode
            Check ($label+'.SetupNotUserAction') (@(Files).Count -eq $before.Count)
            $state=State $action
            $notice=[string](Probe 'RunFaultAct' @($action))
            Check ($label+'.NoUnhandledError') (-not $notice.StartsWith('HANDLER_ERROR|'))
            $qty=if($mode -ceq 'Percent'){0.75*(2*$sourceQty+2)}elseif($mode -ceq 'Zero'){0.0}else{$sourceQty+1}
            $count=if($mode -ceq 'Zero'){0}else{2}
            Check ($label+'.ExactKeyPlanWithoutSubmission') ([bool](Probe 'RunStockPlan' @($keys[0],$keys[1],$qty,$count,$sourceQty)))
            if($mode -ceq 'OverBucket'){Check ($label+'.RejectedPlanAndInputsPreserved') ((State $action) -ceq $state)}
            Check ($label+'.OriginalKeysAndOnHandPreserved') ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.StockSourcePreservedForTest'))
            Pair $before $action $(if($mode -ceq 'OverBucket'){'REJECTED'}else{'STAGED'}) $label
            $records=@(Files|Where-Object{$_ -cnotin $before}|ForEach-Object{[IO.File]::ReadAllText($_)})
            $redacted=$records.Count -eq 2
            foreach($record in $records){foreach($key in $keys){if($record.Contains($key)){$redacted=$false}}}
            Check ($label+'.NoExactKeysInObservations') $redacted
        }
    }
}
