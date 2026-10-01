# Disposable released graph and real upstream completions protect D15/D18.
# No canonical identities or completion states are fabricated by the adapter.
function Install-ProductionCheckInRoutedProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $form.InsertLines(1,'Private mCheckRoutedSetupStage As String, mCheckRoutedSetupError As Long')
    $form.AddFromString(@'
Public Function CheckRoutedPrepareForTest(ByVal token As String) As Variant
    Dim names As Variant, ids As Variant, outputs As Variant, nodes(2) As String
    Dim keys(1) As String, outputIds(1) As String, recipeId As String, row As Long, i As Long
    Dim location As String, rawKey As String, rawQty As Double, report As String
    Dim stage As String, priorTest As Boolean
    On Error GoTo Failed
    priorTest = mReusableActionTestInProgress: mReusableActionTestInProgress = True
    names = Array("Check Source A " & token, "Check Source B " & token, "Check Sink " & token)
    ids = Array("CI-A-" & token, "CI-B-" & token, "CI-C-" & token)
    outputs = Array("CI-OUT-A-" & token, "CI-OUT-B-" & token)
    For i = 0 To 1
        stage = "SourceDefinition" & CStr(i)
        If Not CreateChaiProcessForTest(CStr(ids(i)), CStr(names(i)), _
                Array(Array("Black Tea", "DEMO-RAW-BLACK-TEA", "lbs", "1")), _
                CStr(outputs(i)), CStr(outputs(i)), 1#, "EA") Then GoTo Failed
        If mLstProcessOutputs.ListCount <> 1 Then GoTo Failed
        outputIds(i) = NzStr(mLstProcessOutputs.List(0, 0))
        If outputIds(i) = "" Then GoTo Failed
    Next i
    stage = "SinkDefinition"
    If Not CreateChaiProcessForTest(CStr(ids(2)), CStr(names(2)), _
            Array(Array(CStr(outputs(0)), CStr(outputs(0)), "EA", "1"), _
                  Array(CStr(outputs(1)), CStr(outputs(1)), "EA", "1")), _
            "Check Final " & token, "CI-FINAL-" & token, 1#, "EA") Then GoTo Failed
    stage = "RecipeNodes"
    ClearRecipeDraft False
    mBtnRecipeNew_Click
    recipeId = Trim$(mTxtReusableRecipeId.Text)
    mTxtReusableRecipeName.Text = "Check Routed " & token
    RefreshReusableDesignLists
    For i = 0 To 2
        row = FindIdentityListRow(mLstReleasedProcesses, CStr(ids(i)), "1")
        If row < 0 Then GoTo Failed
        mLstReleasedProcesses.ListIndex = row
        mBtnRecipeAddProcess_Click
    Next i
    If mLstRecipeNodes.ListCount <> 3 Then GoTo Failed
    For i = 0 To 2
        nodes(i) = NzStr(mLstRecipeNodes.List(i, 0))
    Next i
    stage = "Connections"
    If Not ConnectChaiTestNodes(nodes(0), outputIds(0)) Then GoTo Failed
    If Not ConnectChaiTestNodes(nodes(1), outputIds(1)) Then GoTo Failed
    If mLstRecipeConnections.ListCount <> 2 Then GoTo Failed
    stage = "ReleaseRecipe"
    mBtnRecipeAutoOrder_Click
    mBtnRecipeSave_Click
    If InStr(1, TestStatusText(), " is DRAFT", vbTextCompare) = 0 Then GoTo Failed
    mBtnRecipeRelease_Click
    If InStr(1, TestStatusText(), " is RELEASED", vbTextCompare) = 0 Then GoTo Failed
    stage = "ResolveRaw"
    If Not ResolveReusableRunFixtureEntity("DEMO-RAW-BLACK-TEA", location, rawKey, rawQty, report) Then GoTo Failed
    If rawQty < 2# Then GoTo Failed
    stage = "LoadRecipe"
    RefreshReusableDesignLists
    row = FindIdentityListRow(mLstLoaderRecipes, recipeId, "1")
    If row < 0 Then GoTo Failed
    mLstLoaderRecipes.ListIndex = row
    mBtnLoaderLoad_Click
    If Not modProductionReusableRun.ReusableRunIsLoaded() Then GoTo Failed
    SelectComboText mCmbRunLocation, location
    mCmbRunLocation_Change
    mTxtRunBatchNote.Text = "Check routed fixture " & token
    For i = 0 To 2
        nodes(i) = ""
        For row = 0 To mLstLoaderLines.ListCount - 1
            If NzStr(mLstLoaderLines.List(row, 0)) = CStr(names(i)) Then nodes(i) = NzStr(mLstLoaderLines.List(row, 1))
        Next row
        If nodes(i) = "" Then GoTo Failed
    Next i
    For i = 0 To 1
        stage = "SelectSource" & CStr(i)
        SelectComboText mCmbRunProcess, CStr(names(i))
        mCmbRunProcess_Change
        stage = "PaletteSource" & CStr(i)
        row = -1
        For row = 0 To mLstRunPalette.ListCount - 1
            If NzStr(mLstRunPalette.List(row, 3)) = rawKey Then Exit For
        Next row
        If row < 0 Or row >= mLstRunPalette.ListCount Then GoTo Failed
        stage = "AllocateSource" & CStr(i)
        mLstRunPalette.ListIndex = row
        mTxtPaletteSplit.Text = "100": mTxtPaletteQty.Text = "1"
        mBtnRunApplyPalette_Click
        stage = "CheckInSource" & CStr(i)
        mBtnManagerCheckIn_Click
        If Not modProductionReusableRun.ReusableRunIsCheckedIn() Then GoTo Failed
        stage = "ActualOutput" & CStr(i)
        If Not StageReusableTestActualOutputs(1#) Then GoTo Failed
        stage = "ApplyOutput" & CStr(i)
        mBtnManagerApplyOutput_Click
        If Not modProductionReusableRun.ReusableRunIsProcessComplete(CStr(names(i))) Then GoTo Failed
        stage = "SourceIdentity" & CStr(i)
        keys(i) = modProductionReusableRun.ReusableRunOutputSystemKey(nodes(i), outputIds(i))
        If keys(i) = "" Then GoTo Failed
        stage = "SourceBalance" & CStr(i)
        If Abs(modProductionReusableRun.ReusableRunExactEntityQty(keys(i)) - 1#) > 0.0000001 Then GoTo Failed
    Next i
    stage = "DistinctOutputs"
    If StrComp(keys(0), keys(1), vbBinaryCompare) = 0 Then GoTo Failed
    CheckRoutedPrepareForTest = Array(names(2), names(0), names(1), outputs(0), outputs(1), keys(0), keys(1), location)
    mReusableActionTestInProgress = priorTest
    Exit Function
Failed:
    mCheckRoutedSetupStage = stage: mCheckRoutedSetupError = Err.Number
    mReusableActionTestInProgress = priorTest
End Function
Public Function CheckRoutedSetupForTest() As String
    CheckRoutedSetupForTest = mCheckRoutedSetupStage & "|" & CStr(mCheckRoutedSetupError)
End Function
Public Function CheckRoutedStageForTest(ByVal fixture As Variant) As Boolean
    RefreshReusableRunControls False
    SelectComboText mCmbRunLocation, CStr(fixture(7))
    mCmbRunLocation_Change
    SelectComboText mCmbRunProcess, CStr(fixture(0))
    mCmbRunProcess_Change
    CheckRoutedStageForTest = ActiveRunProcess() = CStr(fixture(0)) And _
        modProductionReusableRun.ReusableRunIsProcessComplete(CStr(fixture(1))) And _
        modProductionReusableRun.ReusableRunIsProcessComplete(CStr(fixture(2))) And _
        Not modProductionReusableRun.ReusableRunIsProcessComplete(CStr(fixture(0)))
End Function
Public Function CheckRoutedResultForTest(ByVal fixture As Variant) As Boolean
    CheckRoutedResultForTest = modProductionReusableRun.ReusableRunIsCheckedIn() And _
        Not modProductionReusableRun.ReusableRunIsProcessComplete(CStr(fixture(0))) And _
        ManagerCheckHasRoutedInputForTest(CStr(fixture(5)), CStr(fixture(1)), CStr(fixture(3)), CStr(fixture(3))) And _
        ManagerCheckHasRoutedInputForTest(CStr(fixture(6)), CStr(fixture(2)), CStr(fixture(4)), CStr(fixture(4)))
End Function
'@)
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,'Private mCheckRoutedFixture As Variant')
    $adapter.AddFromString(@'
Public Function CheckRoutedPrepare(ByVal token As String) As Boolean
    mCheckRoutedFixture = mForm.CheckRoutedPrepareForTest(token)
    CheckRoutedPrepare = IsArray(mCheckRoutedFixture)
End Function
Public Function CheckRoutedStage() As Boolean
    CheckRoutedStage = mForm.CheckRoutedStageForTest(mCheckRoutedFixture)
End Function
Public Function CheckRoutedSetup() As String
    CheckRoutedSetup = mForm.CheckRoutedSetupForTest()
End Function
Public Function CheckRoutedResult() As Boolean
    CheckRoutedResult = mForm.CheckRoutedResultForTest(mCheckRoutedFixture)
End Function
'@)
}

function Test-ProductionCheckInRouted($Fixture,$Other){
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Hash([string]$Path){$s=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $s).Hash}finally{$s.Dispose()}}
    $book=$null;$decoy=$null;$pins=@{};$token=[guid]::NewGuid().ToString('N').Substring(0,10).ToUpperInvariant()
    SelectTarget $Fixture
    $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($Fixture.Warehouse,'S1','config-admin'))
    if(-not $seed.StartsWith('OK|')){throw 'Admin Seed unavailable; not product RED.'}
    $ready=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.StockPrepareForTest' @($Fixture.Warehouse))
    if($ready -cne 'READY'){throw 'Real received stock unavailable; not product RED.'}
    if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadPolicyForTest' @($true))){throw 'Authorized tracking policy unavailable; not product RED.'}
    $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $authBytes=[IO.File]::ReadAllBytes($authPath);$authHash=Hash $authPath
    try{
        SelectTarget $Fixture 'config-producer'
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$token;$sheet.Cells.Item(2,2).Formula='=1+2'
        $path=Join-Path $runRoot 'check-in-routed.xlsb';$book.SaveAs($path,50);$book.Close($false)
        $bookPin=Hash $path;$book=$excel.Workbooks.Open($path,0,$false);$sheet=$book.Worksheets.Item(1)
        $decoy=$excel.Workbooks.Add();$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name))
        if(-not [bool](Probe 'CheckRoutedPrepare' @($token))){
            $setup=[string](Probe 'CheckRoutedSetup')
            if($setup -notmatch '^[A-Za-z]+[0-2]?\|-?[0-9]+$'){throw 'Unknown routed fixture stage; raw output suppressed.'}
            throw ('Released graph fixture unavailable: '+$setup+'; not product RED.')
        }
        Check 'CheckInRouted.RealReleasedGraphAndTwoCompletedSources' $true
        foreach($root in @($Fixture.Root,$Other.Root)){
            foreach($file in Get-ChildItem -LiteralPath $root -Recurse -File|Where-Object{$_.Extension -in '.xlsb','.xlsm' -and $_.Name -notlike '~$*'}){$pins[$file.FullName]=Hash $file.FullName}
        }
        if(-not [bool](Probe 'CheckRoutedStage')){throw 'Selected routed sink prerequisite unavailable; not product RED.'}
        $before=@(Get-Slice4beActivityFiles $Fixture);$otherBefore=@(Get-Slice4beActivityFiles $Other)
        Check 'CheckInRouted.Positive.ActualHandlerReturned' ([bool](Probe 'CheckBaselineAct' @('')))
        Check 'CheckInRouted.Positive.ExactRoutedKeysAndNoSinkCompletion' ([bool](Probe 'CheckRoutedResult'))
        Check 'CheckInRouted.Positive.GuardsRestored' ([bool](Probe 'CheckBaselineGuards'))
        Test-ProductionCheckInJournal $Fixture $Other $before $otherBefore 'STAGED' 'CheckInRouted.Positive' $token
        if($CaptureEvidence){[void](Probe 'RunLocalShowAndCapture' @($book.Name,'CHECK_IN'));CaptureOwnedFormByCaptionEvidence 'Production' 'check-in-routed.png'}
        foreach($interruption in @('SignedOut','Permission')){foreach($ordinal in 1,2){
          try{
            [void](Probe 'CheckYieldReset')
            SelectTarget $Fixture 'config-producer';[void](Probe 'OpenDesigner' @($book.Name))
            if(-not [bool](Probe 'CheckRoutedStage')){throw 'Routed state did not survive form reopen; not product RED.'}
            $before=@(Get-Slice4beActivityFiles $Fixture);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            [void](Probe 'CheckYieldArm' @('AvailableQuantity',$ordinal,$interruption,$authPath))
            $returned=[bool](Probe 'CheckBaselineAct' @(''))
            $evidence=([string](Probe 'CheckYieldEvidence')).Split('|')
            if($evidence.Count -ne 4 -or $evidence[0] -cne 'True' -or $evidence[1] -cne 'True' -or $evidence[3] -cne 'True'){throw 'Routed real read-return interruption unavailable; not product RED.'}
            $label='CheckInRouted.'+$interruption+'.UpstreamRead.'+$ordinal
            Check ($label+'.ActualHandlerReturned') $returned
            Check ($label+'.RealReadReturned') $true
            Check ($label+'.InterruptionApplied') $true
            Check ($label+'.NoLaterReadAttempts') ([int]$evidence[2] -eq 0)
            Check ($label+'.OwnerAtBoundaryPreserved') ([bool](Probe 'CheckYieldOwnerPreserved'))
            Check ($label+'.ProjectionAtBoundaryPreserved') ([bool](Probe 'CheckYieldProjectionPreserved'))
            $refused=if($interruption -ceq 'Permission'){[bool](Probe 'CheckBaselinePermissionRefused')}else{[bool](Probe 'CheckBaselineContextRefused')}
            Check ($label+'.VisibleContextRefusal') $refused
            Check ($label+'.GuardsRestored') ([bool](Probe 'CheckBaselineGuards'))
            $outcome=if($interruption -ceq 'SignedOut'){'REQUESTED'}else{'FAILED'}
            Test-ProductionCheckInJournal $Fixture $Other $before $otherBefore $outcome $label $token
          }finally{
            if($interruption -ceq 'Permission'){
                [IO.File]::WriteAllBytes($authPath,$authBytes)
                [void](Run 'invSys.Core.xlam' 'modAuth.ReloadAuth' @($Fixture.Warehouse))
            }
          }
        }}
        [void](Probe 'CheckYieldReset');SelectTarget $Fixture 'config-producer'
        Check 'CheckInRouted.FixtureAuthorizationRestored' ((Hash $authPath) -ceq $authHash)
        $same=$true;foreach($file in $pins.Keys){$same=$same -and (Hash $file) -ceq $pins[$file]}
        Check 'CheckInRouted.SavedAuthorityPreserved' $same
        Check 'CheckInRouted.CapturedBookExtrasPreserved' ($sheet.Cells.Item(2,1).Value2 -ceq $token -and $sheet.Cells.Item(2,2).Formula -ceq '=1+2')
        [void](Probe 'CloseDesigner');$book.Close($false);$book=$null
        Check 'CheckInRouted.OperatorBytesPreserved' ((Hash $path) -ceq $bookPin)
        Test-ProductionCheckInClosedYield $Fixture $Other $path $decoy $token -Routed
        $same=$true;foreach($file in $pins.Keys){$same=$same -and (Hash $file) -ceq $pins[$file]}
        Check 'CheckInClosedYield.Routed.SavedAuthorityPreserved' $same
    }finally{
        [void](Probe 'CheckYieldReset');[void](Probe 'CloseDesigner')
        if($null -ne $book){$book.Close($false)}
        if($null -ne $decoy){$decoy.Close($false)}
        SelectTarget $Fixture
    }
}
