# D15 selected-Process owner boundary, before any Complete Run activity contract.
# Unsaved probes observe real owners; the operator callback remains unchanged.
function Install-ProductionCompleteBaselineProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $owner=$project.VBComponents.Item('modProductionReusableRun').CodeModule
    foreach($entry in @(
        @{Name='CompleteReusableRun';Anchor='If Not mCheckedIn Then';Kind='Whole'},
        @{Name='CompleteReusableProcess';Anchor='Set node = FindNodeByProcessName(processName)';Kind='Selected'}
    )){
        $start=$owner.ProcStartLine($entry.Name,0)
        $body=$owner.Lines($start,$owner.ProcCountLines($entry.Name,0)) -split '\r?\n'
        $hits=@(for($i=0;$i -lt $body.Count;$i++){if($body[$i].Trim() -ieq $entry.Anchor){$start+$i}})
        if($hits.Count -ne 1){throw 'Complete owner entry anchor changed; not behavioral RED.'}
        $owner.InsertLines($hits[0],('    TestProductionDesigner.CompleteBaselineOwnerHit "'+$entry.Kind+'"'))
    }
    $owner.InsertLines(1,'Private mCompleteBaselineBalances As Object')
    $owner.AddFromString(@'
Public Sub CompleteBaselineRememberBalancesForTest()
    Dim key As Variant, entity As String
    Set mCompleteBaselineBalances = CreateObject("Scripting.Dictionary")
    mCompleteBaselineBalances.CompareMode = vbBinaryCompare
    For Each key In mAllocations.Keys
        entity = AllocationSystemKey(CStr(key))
        If Not mCompleteBaselineBalances.Exists(entity) Then _
            mCompleteBaselineBalances.Add entity, ReusableRunExactEntityQty(entity)
    Next key
End Sub
Public Function CompleteBaselineBalancesForTest(ByVal consumed As Boolean) As Boolean
    Dim entity As Variant, key As Variant, used As Double
    If mCompleteBaselineBalances Is Nothing Then Exit Function
    If mCompleteBaselineBalances.Count = 0 Then Exit Function
    For Each entity In mCompleteBaselineBalances.Keys
        used = 0
        If consumed Then
            For Each key In mAllocations.Keys
                If StrComp(AllocationSystemKey(CStr(key)), CStr(entity), vbBinaryCompare) = 0 Then _
                    used = used + CDbl(mAllocations(key))
            Next key
        End If
        If Abs(ReusableRunExactEntityQty(CStr(entity)) - _
            (CDbl(mCompleteBaselineBalances(entity)) - used)) > 0.0000001 Then Exit Function
    Next entity
    CompleteBaselineBalancesForTest = True
End Function
Public Function CompleteBaselineOutputForTest() As Boolean
    Dim output As Object, entity As String
    If mOutputs.Count <> 1 Or mOutputKeys.Count <> 1 Then Exit Function
    Set output = mOutputs(1)
    entity = ReusableRunOutputSystemKey(RunRecordText(output, "ProcessNodeId"), RunRecordText(output, "OutputId"))
    If entity = "" Or mCompleteBaselineBalances.Exists(entity) Then Exit Function
    CompleteBaselineOutputForTest = Abs(ReusableRunExactEntityQty(entity) - 1#) < 0.0000001
End Function
'@)
    $project.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Function CompleteBaselinePrepareForTest(ByVal selected As Boolean) As Boolean
    If Not CheckBaselineReusableStageForTest("Selected") Then Exit Function
    mBtnManagerCheckIn_Click
    If Not CheckBaselineReusableResultForTest(True) Then Exit Function
    mLoading = True
    mLstManagerOutput.ListIndex = 0
    mTxtOutputReal.Text = "1"
    If Not selected Then
        mCmbRunProcess.ListIndex = 0: mCmbTreeRunProcess.ListIndex = 0
        mTxtOutputReal.Text = "9"
    End If
    mLoading = False
    modProductionReusableRun.CompleteBaselineRememberBalancesForTest
    CompleteBaselinePrepareForTest = ((ActiveRunProcess() <> "") = selected)
End Function
Public Function CompleteBaselineActForTest() As Boolean
    On Error GoTo Failed
    mBtnManagerApplyOutput_Click
    CompleteBaselineActForTest = True
Failed:
End Function
Public Function CompleteBaselineSelectionMessageForTest() As Boolean
    CompleteBaselineSelectionMessageForTest = _
        StrComp(mTxtStatus.Text, "Choose one Process before Complete Run.", vbBinaryCompare) = 0
End Function
Public Function CompleteBaselineCompletedForTest() As Boolean
    CompleteBaselineCompletedForTest = modProductionReusableRun.ReusableRunIsCompleted() And _
        modProductionReusableRun.ReusableRunIsProcessComplete(ActiveRunProcess()) And _
        Not modProductionReusableRun.ReusableRunIsCheckedIn()
End Function
'@)
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,'Private mCompleteBaselineSelected As Long, mCompleteBaselineWhole As Long')
    $adapter.AddFromString(@'
Public Sub CompleteBaselineOwnerHit(ByVal kind As String)
    If kind = "Whole" Then mCompleteBaselineWhole = mCompleteBaselineWhole + 1
    If kind = "Selected" Then mCompleteBaselineSelected = mCompleteBaselineSelected + 1
End Sub
Public Function CompleteBaselinePrepare(ByVal selected As Boolean) As Boolean
    mCompleteBaselineWhole = 0: mCompleteBaselineSelected = 0
    CompleteBaselinePrepare = mForm.CompleteBaselinePrepareForTest(selected)
End Function
Public Function CompleteBaselineAct() As Boolean
    CompleteBaselineAct = mForm.CompleteBaselineActForTest()
End Function
Public Function CompleteBaselineOwners() As String
    CompleteBaselineOwners = CStr(mCompleteBaselineSelected) & "|" & CStr(mCompleteBaselineWhole)
End Function
Public Function CompleteBaselineSelectionMessage() As Boolean
    CompleteBaselineSelectionMessage = mForm.CompleteBaselineSelectionMessageForTest()
End Function
Public Function CompleteBaselineCompleted() As Boolean
    CompleteBaselineCompleted = mForm.CompleteBaselineCompletedForTest()
End Function
'@)
}

function Test-ProductionCompleteBaseline($Fixture,$Other) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Owner([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('modProductionReusableRun.'+$Method) $Values}
    function Hash([string]$Path){$s=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $s).Hash}finally{$s.Dispose()}}
    $canary='COMPLETE'+[guid]::NewGuid().ToString('N');$book=$null;$decoy=$null
    SelectTarget $Fixture
    $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($Fixture.Warehouse,'S1','config-admin'))
    if(-not $seed.StartsWith('OK|')){throw 'Admin seed prerequisite unavailable; not product RED.'}
    $otherPins=RestartPins $Other.Root
    try{
        SelectTarget $Fixture 'config-producer'
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$canary;$sheet.Cells.Item(2,2).Formula='=1+2'
        $path=Join-Path $runRoot 'complete-baseline.xlsb';$book.SaveAs($path,50);$book.Close($false)
        $bookPin=Hash $path;$book=$excel.Workbooks.Open($path,0,$false);$sheet=$book.Worksheets.Item(1)
        $decoy=$excel.Workbooks.Add();$decoySheet=$decoy.Worksheets.Item(1)
        $decoySheet.Cells.Item(1,1).Value2=$canary;$decoy.Activate()
        [void](Probe 'OpenDesigner' @($book.Name))
        if(-not [bool](Probe 'ReadPrepare' @($canary))){throw 'Real released definitions unavailable; not product RED.'}
        [void](Probe 'RunLocalRememberFixture')
        Check 'CompleteBaseline.RealSeedAndReleasedDefinitions' $true
        foreach($selected in @($true,$false)){
            if(-not [bool](Probe 'CompleteBaselinePrepare' @($selected))){throw 'Actual Check In prerequisite unavailable; not product RED.'}
            $label='CompleteBaseline.'+$(if($selected){'Selected'}else{'NoProcess'})
            $before=[string](Probe 'RunLocalOwnerState')
            [void](Probe 'RunLocalShowAndCapture' @($book.Name,'CHECK_IN'));$decoy.Activate()
            Check ($label+'.ActualHandlerReturned') ([bool](Probe 'CompleteBaselineAct'))
            if($selected){
                Check ($label+'.OnlySelectedOwnerEntered') ([string](Probe 'CompleteBaselineOwners') -ceq '1|0')
                Check ($label+'.ProcessAndBatchCompleted') ([bool](Probe 'CompleteBaselineCompleted'))
                Check ($label+'.ExactAllocatedEntitiesConsumed') ([bool](Owner 'CompleteBaselineBalancesForTest' @($true)))
                Check ($label+'.FreshOutputEntityQuantity') ([bool](Owner 'CompleteBaselineOutputForTest'))
            }else{
                Check ($label+'.NoCompletionOwnerEntered') ([string](Probe 'CompleteBaselineOwners') -ceq '0|0')
                Check ($label+'.OwnerStagingPreserved') ([string](Probe 'RunLocalOwnerState') -ceq $before)
                Check ($label+'.ExactInputBalancesPreserved') ([bool](Owner 'CompleteBaselineBalancesForTest' @($false)))
                Check ($label+'.ChooseProcessMessage') ([bool](Probe 'CompleteBaselineSelectionMessage'))
            }
            Check ($label+'.CapturedBookCustomValueAndFormula') ($sheet.Cells.Item(2,1).Value2 -ceq $canary -and $sheet.Cells.Item(2,2).Formula -ceq '=1+2')
            Check ($label+'.DecoyPreserved') ($decoySheet.Cells.Item(1,1).Value2 -ceq $canary -and $decoy.Worksheets.Count -eq 1)
            CaptureOwnedFormByCaptionEvidence 'Production' ('complete-baseline-'+$(if($selected){'selected'}else{'no-process'})+'.png')
        }
        [void](Probe 'CloseDesigner');$book.Close($false);$book=$null
        Check 'CompleteBaseline.OperatorBytesPreserved' ((Hash $path) -ceq $bookPin)
        Check 'CompleteBaseline.OtherWarehousePreserved' (RestartPinsEqual $otherPins $Other.Root)
    }finally{
        [void](Probe 'CloseDesigner')
        if($null -ne $book){$book.Close($false)}
        if($null -ne $decoy){$decoy.Close($false)}
    }
}
