# The released reusable Recipe selected by the actual form is not a Design/BOM record.
# Observe real owner reads and worksheet effects; never replace their results.
function Install-ProductionRunWorksheetScaleProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $owner=$project.VBComponents.Item('mProduction').CodeModule
    function Scale-Seam([string]$Anchor,[string]$Code,[bool]$After=$false){
        $start=$owner.ProcStartLine('LoadRecipeChooser',0);$end=$start+$owner.ProcCountLines('LoadRecipeChooser',0)
        # These anchors contain only VBA identifiers, with no string literals.
        $hits=@(for($i=$start;$i -lt $end;$i++){if($owner.Lines($i,1).Trim() -ieq $Anchor){$i}})
        if($hits.Count -ne 1){throw ('Worksheet Scale owner seam changed: '+$Anchor+'; not product RED.')}
        $owner.InsertLines($hits[0]+[int]$After,$Code)
    }
    Scale-Seam 'If wsProd Is Nothing Then Exit Sub' '    TestProductionDesigner.RunScaleOwnerResolved wsProd'
    Scale-Seam 'Set stagingWb = BuildReleasedDesignRecipeStagingWorkbook(recipeId, syncReport)' '        TestProductionDesigner.RunScaleReadReturned (Not stagingWb Is Nothing)' $true
    Scale-Seam 'If Not LocalProductionRecipeRowsExist(wsProd.Parent, recipeId) Then' '        TestProductionDesigner.RunScaleLegacyEntered'
    $start=$owner.ProcStartLine('LoadRecipeChooser',0);$end=$start+$owner.ProcCountLines('LoadRecipeChooser',0)
    $notices=0
    for($line=$start;$line -lt $end;$line++){
        $text=$owner.Lines($line,1)
        if($text.Trim().StartsWith('MsgBox ')){
            $owner.ReplaceLine($line,$text.Replace('MsgBox ','TestProductionDesigner.RunSheetNotice '))
            $notices++
        }
    }
    if($notices -ne 4){throw 'Worksheet Scale notification boundaries changed; not product RED.'}
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $form.AddFromString(@'
Public Function RunScaleSheetStageForTest(ByVal mode As String, ByVal canary As String) As String
    Dim ws As Worksheet, lo As ListObject, name As Variant, row As Long, phase As String
    On Error GoTo Failed
    phase = "Restore"
    RunSheetOwnerRestoreForTest
    modProductionReusableRun.ClearReusableRun
    mLoading = True: mPages.Value = 3
    phase = "Selection"
    RefreshRecipeLists
    row = FindIdentityListRow(mLstLoaderRecipes, mReadRecipeIdForTest, mReadRecipeVersionForTest)
    If row < 0 Then GoTo Unavailable
    mLstLoaderRecipes.ListIndex = row
    mTxtBatchScalePercent.Text = "200"
    Select Case mode
        Case "NoSelection": mLstLoaderRecipes.ListIndex = -1
        Case "BadScale": mTxtBatchScalePercent.Text = "not a number"
        Case "BelowMinimum": mTxtBatchScalePercent.Text = "0"
        Case "AboveMaximum": mTxtBatchScalePercent.Text = "1001"
        Case "Minimum": mTxtBatchScalePercent.Text = "0.001"
        Case "Maximum": mTxtBatchScalePercent.Text = "1000"
    End Select
    Set ws = mOperatorWorkbook.Worksheets("Production")
    phase = "Staging"
    For Each name In Array("RecipeChooser_generated", "InventoryPalette_generated")
        Set lo = ws.ListObjects(CStr(name))
        If lo.ListRows.Count = 0 Then lo.ListRows.Add AlwaysInsert:=False
        If ProductionColumnIndex(lo, "Operator Extra") = 0 Then lo.ListColumns.Add.Name = "Operator Extra"
        If ProductionColumnIndex(lo, "Operator Formula") = 0 Then lo.ListColumns.Add.Name = "Operator Formula"
        lo.ListColumns("Operator Extra").DataBodyRange.Cells(1, 1).Value2 = canary
        lo.ListColumns("Operator Formula").DataBodyRange.Cells(1, 1).Formula = "=1+2"
        If CStr(name) = "RecipeChooser_generated" Then
            lo.ListColumns("AMOUNT NEEDED").DataBodyRange.Cells(1, 1).Value2 = 2#
        Else
            lo.ListColumns("BASE QUANTITY").DataBodyRange.Cells(1, 1).Value2 = 10#
            lo.ListColumns("QUANTITY").DataBodyRange.Cells(1, 1).Value2 = 2.5
            lo.ListColumns("System_Key").DataBodyRange.Cells(1, 1).Value2 = mRunSheetKey
        End If
    Next name
    If mode = "MissingSheet" Or mode = "DecoyMissing" Then ws.Name = "RunOwnerUnavailable"
    TestProductionDesigner.RunSheetReset mOperatorWorkbook.Name
    TestProductionDesigner.RunScaleReset
    mLoading = False
    RunScaleSheetStageForTest = "READY"
    Exit Function
Unavailable:
    RunScaleSheetStageForTest = "FIXTURE_FAILED|" & phase
    GoTo Finished
Failed:
    RunScaleSheetStageForTest = "FIXTURE_FAILED|" & phase & "|" & CStr(Err.Number)
Finished:
    mLoading = False
End Function
Public Function RunScaleSheetValuesForTest(ByVal mode As String, ByVal canary As String) As Boolean
    On Error GoTo Unavailable
    Dim ws As Worksheet, lo As ListObject, name As Variant, factor As Double
    If mode = "MissingSheet" Or mode = "DecoyMissing" Then
        Set ws = mOperatorWorkbook.Worksheets("RunOwnerUnavailable")
    Else
        Set ws = mOperatorWorkbook.Worksheets("Production")
    End If
    factor = 1#
    Select Case mode
        Case "UnavailableDesign": factor = 2#
        Case "Minimum": factor = 0.00001
        Case "Maximum": factor = 10#
    End Select
    For Each name In Array("RecipeChooser_generated", "InventoryPalette_generated")
        Set lo = ws.ListObjects(CStr(name))
        If lo.ListColumns("Operator Extra").DataBodyRange.Cells(1, 1).Value2 <> canary Or _
           lo.ListColumns("Operator Formula").DataBodyRange.Cells(1, 1).Formula <> "=1+2" Then Exit Function
        If CStr(name) = "RecipeChooser_generated" Then
            If Abs(CDbl(lo.ListColumns("AMOUNT NEEDED").DataBodyRange.Cells(1, 1).Value2) - 2# * factor) > 0.00000001 Then Exit Function
        Else
            If Abs(CDbl(lo.ListColumns("BASE QUANTITY").DataBodyRange.Cells(1, 1).Value2) - 10# * factor) > 0.00000001 Then Exit Function
            If Abs(CDbl(lo.ListColumns("QUANTITY").DataBodyRange.Cells(1, 1).Value2) - 2.5 * factor) > 0.00000001 Then Exit Function
            If StrComp(CStr(lo.ListColumns("System_Key").DataBodyRange.Cells(1, 1).Value2), mRunSheetKey, vbBinaryCompare) <> 0 Then Exit Function
        End If
    Next name
    RunScaleSheetValuesForTest = True
Unavailable:
End Function
Public Function RunScaleSheetBranchForTest() As Boolean
    RunScaleSheetBranchForTest = Not modProductionReusableRun.ReusableRunIsLoaded() And mPages.Value = 3
End Function
'@)
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,@'
Private mScaleLoadCalls As Long, mScaleReadCalls As Long, mScaleReadOK As Boolean, mScaleLegacyCalls As Long
Private mScaleOwnerTarget As String
'@)
    $adapter.AddFromString(@'
Public Sub RunScaleReset()
    mScaleLoadCalls = 0: mScaleReadCalls = 0: mScaleReadOK = False: mScaleLegacyCalls = 0
    mScaleOwnerTarget = "NotEntered"
End Sub
Public Sub RunScaleOwnerResolved(ByVal ws As Worksheet)
    mScaleLoadCalls = mScaleLoadCalls + 1
    If ws Is Nothing Then
        mScaleOwnerTarget = "Unavailable"
    ElseIf ws.Parent.Name = mRunSheetExpectedBook Then
        mScaleOwnerTarget = "Captured"
    ElseIf ws.Parent.IsAddin Then
        mScaleOwnerTarget = "OtherAddin"
    Else
        mScaleOwnerTarget = "OtherWorkbook"
    End If
End Sub
Public Sub RunScaleReadReturned(ByVal succeeded As Boolean)
    mScaleReadCalls = mScaleReadCalls + 1: mScaleReadOK = succeeded
End Sub
Public Sub RunScaleLegacyEntered()
    mScaleLegacyCalls = mScaleLegacyCalls + 1
End Sub
Public Function RunScaleEvidence() As String
    RunScaleEvidence = CStr(mScaleLoadCalls) & "|" & CStr(mScaleReadCalls) & "|" & CStr(mScaleReadOK) & "|" & CStr(mScaleLegacyCalls) & "|" & mScaleOwnerTarget & "|" & CStr(mRunSheetNotices)
End Function
Public Function RunScaleSheetStage(ByVal mode As String, ByVal canary As String) As String
    RunScaleSheetStage = mForm.RunScaleSheetStageForTest(mode, canary)
End Function
Public Function RunScaleSheetValues(ByVal mode As String, ByVal canary As String) As Boolean
    RunScaleSheetValues = mForm.RunScaleSheetValuesForTest(mode, canary)
End Function
Public Function RunScaleSheetBranch() As Boolean
    RunScaleSheetBranch = mForm.RunScaleSheetBranchForTest()
End Function
'@)
}

function Test-ProductionRunWorksheetScale($Fixture) {
    if(-not [bool](Probe 'RunSheetOwnerPrepare')){throw 'Supported Scale surface fixture unavailable; not product RED.'}
    Check 'RunWorksheetScale.SupportedSurfaceAndSeedIdentity' $true
    $key=[string](Probe 'RunWorksheetKey')
    foreach($mode in @('NoSelection','BadScale','BelowMinimum','AboveMaximum','UnavailableDesign','Minimum','Maximum','MissingSheet','DecoyMissing')){
        if($mode -ceq 'DecoyMissing'){
            $decoySheet=$decoy.Worksheets.Add();$decoySheet.Name='Production'
            $decoySheet.Range('A1').Value2='Operator Extra';$decoySheet.Range('A2').Value2=$canary;$decoySheet.Range('B2').Formula='=1+2'
        }
        $before=@(Files);$ready=[string](Probe 'RunScaleSheetStage' @($mode,$canary))
        if($ready -cne 'READY'){
            if($ready -notmatch '^FIXTURE_FAILED\|[A-Za-z]+(\|-?[0-9]+)?$'){$ready='Unavailable'}
            throw ('Worksheet Scale staging unavailable: '+$ready+'; not product RED.')
        }
        try{
            $decoy.Activate();$label='RunWorksheetScale.'+$mode
            Check ($label+'.SetupNotUserAction') (@(Files).Count -eq $before.Count)
            Check ($label+'.ActualWorksheetBranch') ([bool](Probe 'RunScaleSheetBranch'))
            $notice=[string](Probe 'RunFaultAct' @('SCALE'))
            $evidence=([string](Probe 'RunScaleEvidence')).Split('|')
            if($evidence.Count -ne 6){throw 'Worksheet Scale owner evidence unavailable.'}
            $refused=$mode -cin @('NoSelection','BadScale','BelowMinimum','AboveMaximum')
            $missing=$mode -cin @('MissingSheet','DecoyMissing')
            Check ($label+'.NoUnhandledError') (-not $notice.StartsWith('HANDLER_ERROR|'))
            Check ($label+'.GuardsRestoredWithoutAdapterReset') ([bool](Probe 'RunFaultGuards'))
            Check ($label+'.CurrentWorksheetValuesAndUnknownColumns') ([bool](Probe 'RunScaleSheetValues' @($mode,$canary)))
            Check ($label+'.CanonicalSourcePreserved') ([bool](Probe 'RunWorksheetSource'))
            Check ($label+'.NoLegacyFallback') ([int]$evidence[3] -eq 0)
            if($refused){
                Check ($label+'.RefusalBeforeOwner') ([int]$evidence[0] -eq 0 -and [int]$evidence[1] -eq 0 -and [int]$evidence[5] -eq 0)
            }elseif($missing){
                Write-Output ('Run Scale '+$mode+' resolved target: '+$evidence[4])
                Check ($label+'.NoOtherWorkbookOwnerEntry') ($evidence[4] -cin @('NotEntered','Unavailable'))
            }else{
                Check ($label+'.RealReleasedDesignReadUnavailable') ([int]$evidence[0] -eq 1 -and [int]$evidence[1] -eq 1 -and $evidence[2] -ceq 'False' -and $evidence[4] -ceq 'Captured' -and [int]$evidence[5] -eq 1)
            }
            if($mode -ceq 'DecoyMissing'){Check ($label+'.DecoyUnknownValuesPreserved') ($decoySheet.Range('A2').Value2 -ceq $canary -and $decoySheet.Range('B2').Formula -ceq '=1+2')}
            Pair $before 'SCALE' $(if($refused){'REJECTED'}else{'FAILED'}) $label
            $records=@(Files|Where-Object{$_ -cnotin $before}|ForEach-Object{[IO.File]::ReadAllText($_)})
            Check ($label+'.NoExactKeyInObservations') ($records.Count -eq 2 -and @($records|Where-Object{$_.Contains($key)}).Count -eq 0)
        }finally{[void](Probe 'RunSheetOwnerRestore')}
    }
}
