# Reuse the established packaged fixture setup; keep the actual edit handler intact.
function Install-ProductionRecipeStructureReleasedProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $source=$form.Lines($form.ProcStartLine('ExerciseReusableProductionFormActions',0),$form.ProcCountLines('ExerciseReusableProductionFormActions',0))
    $marker='        mBtnRecipeUpdateConnection_Click'
    if(([regex]::Matches($source,[regex]::Escape($marker))).Count -ne 1){throw 'Reusable fixture Update boundary changed; not product RED.'}
    $prefix=$source.Substring(0,$source.IndexOf($marker,[StringComparison]::Ordinal))
    $declaration='Private Function ExerciseReusableProductionFormActions(ByVal boundWorkbookName As String) As String'
    if(-not $prefix.Contains($declaration)){throw 'Reusable fixture declaration changed.'}
    $prefix=$prefix.Replace($declaration,'Public Function StructureReleasedSetupForTest() As String')
    if($prefix.Contains('boundWorkbookName')){throw 'Unexpected fixture dependency before Update.'}
    $form.AddFromString($prefix+@'
        If Not recipeConnectionSelected Or Not processReleased Then GoTo ActionFailed
        StructureReleasedSetupForTest = "READY"
        mReusableActionTestInProgress = False
        Exit Function
    End If
ActionFailed:
    StructureReleasedSetupForTest = "SETUP_FAILED"
    mReusableActionTestInProgress = False
    Exit Function
Failed:
    StructureReleasedSetupForTest = "SETUP_ERROR|" & CStr(Err.Number)
    mReusableActionTestInProgress = False
End Function
Public Sub StructureReleasedEditForTest()
    mPages.Value = 1
    mTxtConnectionQty.Text = " 4 "
    mTxtConnectionPercent.Text = " 75 "
End Sub
Public Function StructureReleasedNodesForTest() As Boolean
    Dim i As Long, record As Object, records As Collection, report As String, released As Boolean
    If mLstRecipeNodes.ListCount <> 2 Then Exit Function
    For i = 0 To 1
        released = False
        Set records = ProcessRecordsForRecipeNode(i, report)
        If records Is Nothing Then Exit Function
        For Each record In records
            If modProductionReusableDesigns.ReusableRecordText(record, "RecordType") = "PROCESS" Then
                released = (modProductionReusableDesigns.ReusableRecordText(record, "Status") = "RELEASED")
            End If
        Next record
        If Not released Then Exit Function
    Next i
    StructureReleasedNodesForTest = True
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function StructureReleasedSetup() As String
    StructureReleasedSetup = mForm.StructureReleasedSetupForTest()
End Function
Public Sub StructureReleasedEdit()
    mForm.StructureReleasedEditForTest
End Sub
Public Function StructureReleasedNodes() As Boolean
    StructureReleasedNodes = mForm.StructureReleasedNodesForTest()
End Function
'@)
}

function Test-ProductionRecipeStructureReleased($Fixture){
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Hash([string]$Path){
        $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
        try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}
    }
    $book=$null;$decoy=$null
    try{
        SelectTarget $Fixture 'config-producer'
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2='Retain local column'
        $path=Join-Path $runRoot 'recipe-structure-released-operator.xlsb'
        $book.SaveAs($path,50);$book.Close($false);$bookPin=Hash $path
        $book=$excel.Workbooks.Open($path,0,$false);$decoy=$excel.Workbooks.Add();$decoy.Activate()
        [void](Probe 'OpenDesigner' @($book.Name))
        $ready=[string](Probe 'StructureReleasedSetup')
        Check 'RecipeStructureReleased.PackagedFixtureSetup' ($ready -ceq 'READY')
        if($ready -cne 'READY'){throw 'Released Process fixture setup failed; not product RED.'}
        $released=[bool](Probe 'StructureReleasedNodes')
        Check 'RecipeStructureReleased.BothNodesHaveReleasedDomainRecords' $released
        if(-not $released){throw 'Released Domain records unavailable; not product RED.'}
        $pins=@{}
        foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File|Where-Object{$_.Extension -in '.xlsb','.xlsm' -and $_.Name -notlike '~$*'}){$pins[$file.FullName]=Hash $file.FullName}
        $original=[string](Probe 'StructureRows' @('Connections'))
        $nodes=[string](Probe 'StructureRows' @('Nodes'))
        $instructions=[string](Probe 'StructureRows' @('Instructions'))
        [void](Probe 'StructureReleasedEdit')
        $expected=[string](Probe 'StructureEditor');$fields=$expected.Split([char]9)
        $valid=$fields.Count -eq 7 -and $fields[4] -ceq '4' -and $fields[5] -ceq '75' -and $fields[6] -ceq 'LB' -and $original -cne $expected -and @($fields[0..3]|Where-Object{$_ -ceq ''}).Count -eq 0
        Check 'RecipeStructureReleased.EditedFieldsPresentBeforeHandler' $valid
        if(-not $valid){throw 'Changed connection editor fixture unavailable; not product RED.'}
        $notice=[string](Probe 'StructureAct' @('UPDATE_CONNECTION'))
        $actual=[string](Probe 'StructureRows' @('Connections'));$observed=$actual.Split([char]9)
        Check 'RecipeStructureReleased.ActualUpdateReturnsStagedStatus' ($notice.StartsWith('Recipe connection staged.'))
        Check 'RecipeStructureReleased.AllSevenValidatedFieldsWritten' ($actual -ceq $expected)
        Check 'RecipeStructureReleased.ChangedQuantityWritten' ($observed.Count -eq 7 -and $observed[4] -ceq '4')
        Check 'RecipeStructureReleased.ChangedPercentageWritten' ($observed.Count -eq 7 -and $observed[5] -ceq '75')
        Check 'RecipeStructureReleased.RoutingAndUomPreserved' (($observed[0..3] -join '|') -ceq ($fields[0..3] -join '|') -and $observed[6] -ceq 'LB')
        Check 'RecipeStructureReleased.NodesAndInstructionsPreserved' ([string](Probe 'StructureRows' @('Nodes')) -ceq $nodes -and [string](Probe 'StructureRows' @('Instructions')) -ceq $instructions)
        $unchanged=$actual;[void](Probe 'StructureAct' @('UPDATE_CONNECTION'))
        Check 'RecipeStructureReleased.UnchangedUpdatePreservesRow' ([string](Probe 'StructureRows' @('Connections')) -ceq $unchanged)
        $same=$true;foreach($file in $pins.Keys){$same=$same -and (Hash $file) -ceq $pins[$file]}
        Check 'RecipeStructureReleased.UpdateDoesNotSaveAuthority' $same
        $sheet=$book.Worksheets.Item(1)
        Check 'RecipeStructureReleased.UnknownColumnsPreserved' ($sheet.Cells.Item(1,1).Value2 -ceq 'Operator Extra' -and $sheet.Cells.Item(2,1).Value2 -ceq 'Retain local column')
        if($CaptureEvidence){[void](Probe 'StructureShow');CaptureOwnedFormByCaptionEvidence 'Production' 'recipe-structure-released-update.png'}
        $book.Close($false);$book=$null
        Check 'RecipeStructureReleased.OperatorWorkbookBytesPreserved' ((Hash $path) -ceq $bookPin)
    }finally{[void](Probe 'CloseDesigner');if($null -ne $book){$book.Close($false)};if($null -ne $decoy){$decoy.Close($false)};SelectTarget $Fixture}
}
