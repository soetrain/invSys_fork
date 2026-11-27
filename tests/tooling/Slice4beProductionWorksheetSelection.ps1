# The ordinary worksheet owner must preserve an explicit output selection within one Process.
function Install-WorksheetSelectionProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $project.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Function PickSecondPreparedOutputForTest(ByVal process As String) As Boolean
    Dim output As ListObject, report As String
    BindOperatorWorkbookForRun
    If Not mProduction.PrepareProductionOutputForCurrentRecipe(report) Then Exit Function
    Set output = ProductionTable(TABLE_MANAGER_OUTPUT)
    If output.ListRows.Count <> 2 Then Exit Function
    If CellByHeader(output, 1, "ITEM_CODE") <> "DEMO-RAW-BLACK-TEA" Then Exit Function
    If CellByHeader(output, 2, "ITEM_CODE") <> "DEMO-RAW-FILTERED-WATER" Then Exit Function
    RefreshLoaderState
    RefreshManagerState
    If mCmbRunProcess.ListCount <> 2 Then Exit Function
    mCmbRunProcess.ListIndex = 1
    If ActiveRunProcess() <> process Or mLstManagerOutput.ListCount <> 2 Then Exit Function
    mLstManagerOutput.ListIndex = 1
    mLstManagerOutput_Click
    mTxtOutputReal.Text = "1"
    PickSecondPreparedOutputForTest = (mLstManagerOutput.ListIndex = 1 And SelectedProductionOutputTableRow() = 2 And HasProductionCheckRows())
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function PickSecondPreparedOutput(ByVal process As String) As Boolean
    PickSecondPreparedOutput = mForm.PickSecondPreparedOutputForTest(process)
End Function
'@)
}

function Test-PreparedWorksheetSelection($Work,$Fixture,[string]$ProcessName,$Decoy){
    $sheet=$Work.Worksheets.Item('Production')
    # Only stage recipe lines on the disposable operator surface. The normal
    # Prepare command resolves both SKUs and populates the output picker.
    $headers=@('PROCESS','DIAGRAM_ID','INPUT/OUTPUT','INGREDIENT','PERCENT','UOM','AMOUNT NEEDED','INGREDIENT_ID')
    for($col=0;$col -lt $headers.Count;$col++){$sheet.Cells.Item(1,52+$col).Value2=$headers[$col]}
    $loader=$sheet.ListObjects.Add(1,$sheet.Range('AZ1:BG3'),$null,1)
    $loader.Name='RecipeChooser_generated'
    for($i=1;$i -le 2;$i++){
        $loader.DataBodyRange.Cells.Item($i,1).Value2=$ProcessName
        $loader.DataBodyRange.Cells.Item($i,3).Value2='OUTPUT'
        $loader.DataBodyRange.Cells.Item($i,4).Value2=@('Black Tea','Filtered Water')[$i-1]
        $loader.DataBodyRange.Cells.Item($i,6).Value2='lbs'
        $loader.DataBodyRange.Cells.Item($i,7).Value2='1'
    }
    if(-not [bool](Probe 'PickSecondPreparedOutput' @($ProcessName))){throw 'Two prepared outputs and the explicit second selection are required.'}
    $output=Table $Work 'ProductionOutput'
    $keys=$output.ListColumns.Item('System_Key').DataBodyRange
    if($keys.Cells.Item(1,1).Value2 -or $keys.Cells.Item(2,1).Value2){throw 'Expected two fresh output drafts.'}
    Pass 'TwoDistinctPreparedOutputsSecondSelected' $true
    $Decoy.Activate()
    [void](Probe 'CompleteSubmissionArm' @('','','ObserveOnly'))
    if(-not [bool](Probe 'CompleteEntryAct' @(''))){throw 'Ordinary Complete Run handler did not return.'}
    $ids=@(([string](Probe 'CompleteWorksheetEvents')).Split("`n"))
    if($ids.Count -ne 2 -or !$ids[0] -or !$ids[1] -or $ids[0] -ceq $ids[1]){throw 'Expected one consumption event and one output event.'}
    $keys=(Table $Work 'ProductionOutput').ListColumns.Item('System_Key').DataBodyRange
    $first=[string]$keys.Cells.Item(1,1).Value2
    $second=[string]$keys.Cells.Item(2,1).Value2
    if([bool]$first -eq [bool]$second){throw 'Expected one owner-created output identity.'}
    $inputKey=(Probe 'CompletePathWorksheetKeys')[0]
    $created=if($first){$first}else{$second}
    $createdSku=if($first){'DEMO-RAW-BLACK-TEA'}else{'DEMO-RAW-FILTERED-WATER'}
    Test-CompletePathAuthority $Fixture $ids @($inputKey,$created) @(-10.0,1.0) 'WorksheetSelection' $createdSku|Out-Host
    Pass 'SingleOrdinaryClickAndOwner' ((Run 'invSys.Operations.xlam' 'TestCompleteReplay.Counts') -ceq '1|1')
    Pass 'OwnerCompletedExplicitlySelectedOutput' ([bool]$second -and !$first)
}
