# Existing D12/D18 lifecycle through the public launcher and its actual form.
function Install-ProductionUomPublicCloseProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $project.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Function UomBoundWorkbookForTest() As String
    If mOperatorWorkbook Is Nothing Then Exit Function
    UomBoundWorkbookForTest = mOperatorWorkbook.Name
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function LoadedProductionFormsForTest() As Long
    Dim loaded As Object
    For Each loaded In VBA.UserForms
        If TypeName(loaded) = "frmProduction" Then LoadedProductionFormsForTest = LoadedProductionFormsForTest + 1
    Next loaded
End Function
Public Function BindLaunchedUomForTest() As String
    Dim loaded As Object
    If LoadedProductionFormsForTest() <> 1 Then Err.Raise 5, , "One launched Production form is required."
    Set mForm = Nothing
    For Each loaded In VBA.UserForms
        If TypeName(loaded) = "frmProduction" Then
            Set mForm = loaded
            BindLaunchedUomForTest = mForm.UomBoundWorkbookForTest()
            Exit Function
        End If
    Next loaded
End Function
Public Sub ForgetLaunchedUomForTest()
    Set mForm = Nothing
End Sub
'@)
}

function Test-ProductionUomPublicClose($Fixture) {
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function UomRecords {
        @(Get-Slice4beActivityFiles $Fixture|ForEach-Object{[IO.File]::ReadAllText($_)|ConvertFrom-Json}|Where-Object{$_.ControlId -ceq 'PRODUCTION_UOM_EDIT'})
    }
    function CheckPair([object[]]$Before,[string]$Outcome,[string]$Case){
        $priorIds=@($Before|ForEach-Object{[string]$_.RecordId})
        $fresh=@(UomRecords|Where-Object{$_.RecordId -cnotin $priorIds})
        $attempt=@($fresh|Where-Object OutcomeCode -CEQ 'REQUESTED')
        $terminal=@($fresh|Where-Object OutcomeCode -CEQ $Outcome)
        $pair=$fresh.Count -eq 2 -and $attempt.Count -eq 1 -and $terminal.Count -eq 1
        if($pair){$pair=$attempt[0].ActivityId -ceq $terminal[0].ActivityId -and $attempt[0].RecordId -cne $terminal[0].RecordId -and $terminal[0].DataEffect -ceq 'Unchanged' -and $terminal[0].OwnerId -ceq 'PRODUCTION_UOM_STAGING' -and @($terminal[0].SourceEventRefs).Count -eq 0}
        Check $Case $pair
    }
    $book=$null;$decoy=$null
    $operatorRoot=Join-Path $runRoot 'uom-public-operators'
    if(-not [IO.Path]::GetFullPath($operatorRoot).StartsWith([IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Operator fixture root escaped this run.'}
    SelectTarget $Fixture
    if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.UomPolicyForTest' @($true))){throw 'UOM policy fixture unavailable; not product RED.'}
    SelectTarget $Fixture 'config-producer'
    if(-not [bool](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetLocalOperatorRootOverrideForAutomation' @($operatorRoot))){throw 'Isolated operator root unavailable; do not invoke launcher.'}
    $configPin=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    [void](Run 'invSys.Operations.xlam' 'modOperationsInit.Auto_Open')
    $canary='UOMPUBLIC'+[guid]::NewGuid().ToString('N')
    try {
        $decoy=$excel.Workbooks.Add();$decoy.Worksheets.Item(1).Range('A1').Value2='Operator Column'
        $decoy.Worksheets.Item(1).Range('A2').Value2=$canary
        $decoyCount=$decoy.Worksheets.Count
        $decoyBefore=$decoy.Worksheets.Item(1).UsedRange.Formula|ConvertTo-Json -Compress -Depth 5
        $decoy.Activate();$excel.Visible=$true
        [void](Run 'invSys.Operations.xlam' 'mProduction.BtnOpenProductionForm')
        $oneForm=[long](Probe 'LoadedProductionFormsForTest') -eq 1
        Check 'UomPublicClose.Initial.OneLaunchedForm' $oneForm
        if(-not $oneForm){return}
        $name=[string](Probe 'BindLaunchedUomForTest');$book=$excel.Workbooks.Item($name)
        $bookPath=[string]$book.FullName
        $isolated=[IO.Path]::GetFullPath($bookPath).StartsWith([IO.Path]::GetFullPath($operatorRoot).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)
        Check 'UomPublicClose.Initial.OwnedLocalWorkbook' ($isolated -and $book.Name -cne $decoy.Name)
        if(-not $isolated){throw 'Public launcher returned a workbook outside the disposable operator root.'}
        $before=@(UomRecords);$decoy.Activate();$notice=[string](Probe 'SendUom')
        $sheet=$null;$table=$null
        try{$sheet=$book.Worksheets.Item('invSys UOM Catalog');$table=$sheet.ListObjects.Item('tblInvSysUomCatalog')}catch{}
        $opened=$null -ne $table -and $notice -notlike '*ERROR*' -and $notice -notlike '*failed*'
        Check 'UomPublicClose.Open.CapturedWorkbook' $opened
        if(-not $opened){return}
        CheckPair $before 'OPENED' 'UomPublicClose.Open.ExactOriginalPair'
        $extra=$table.ListColumns.Add();$extra.Name='Operator Annotation';$extra.DataBodyRange.Value2=$canary
        $table.ListColumns.Item('Notes').DataBodyRange.Cells.Item(1,1).Value2=$canary
        $book.Save()
        [void](Probe 'ShowUom')
        CaptureOwnedFormByCaptionEvidence 'Production' 'uom-public-open.png'
        $beforeClose=@(UomRecords)
        $activityPins=@{};foreach($file in @(Get-Slice4beActivityFiles $Fixture)){$activityPins[$file]=(Get-FileHash -LiteralPath $file).Hash}
        $book.Close($false);$book=$null
        Check 'UomPublicClose.Close.FormDisposed' ([long](Probe 'LoadedProductionFormsForTest') -eq 0)
        Check 'UomPublicClose.Close.NoOwnedProductionWindow' ([InvSysSettingsCapture]::OwnedVisibleForm('Production',[IntPtr]$excel.Hwnd) -eq [IntPtr]::Zero)
        Check 'UomPublicClose.Close.NoUomActionInvented' (@(UomRecords).Count -eq $beforeClose.Count)
        [void](Probe 'ForgetLaunchedUomForTest')
        $decoy.Activate()
        [InvSysSettingsCapture]::SaveVisibleWindow([IntPtr]$excel.Hwnd,(Join-Path $reportRoot 'uom-public-owner-closed.png'))
        [void](Run 'invSys.Operations.xlam' 'mProduction.BtnOpenProductionForm')
        $oneForm=[long](Probe 'LoadedProductionFormsForTest') -eq 1
        Check 'UomPublicClose.Reopen.OneLaunchedForm' $oneForm
        if(-not $oneForm){return}
        $name=[string](Probe 'BindLaunchedUomForTest');$book=$excel.Workbooks.Item($name)
        Check 'UomPublicClose.Reopen.ReusesSavedOwner' ([string]$book.FullName -ceq $bookPath -and $book.Name -cne $decoy.Name)
        Check 'UomPublicClose.Reopen.NoUomActionInvented' (@(UomRecords).Count -eq $beforeClose.Count)
        $retained=$false
        try{
            $sheet=$book.Worksheets.Item('invSys UOM Catalog');$table=$sheet.ListObjects.Item('tblInvSysUomCatalog')
            $retained=[string]$table.ListColumns.Item('Operator Annotation').DataBodyRange.Cells.Item(1,1).Value2 -ceq $canary -and [string]$table.ListColumns.Item('Notes').DataBodyRange.Cells.Item(1,1).Value2 -ceq $canary
        }catch{}
        Check 'UomPublicClose.Reopen.SavedDraftAndUnknownColumn' $retained
        if(-not $retained){return}
        $before=@(UomRecords);$book.Saved=$false;$decoy.Activate();$notice=[string](Probe 'SendUom')
        Check 'UomPublicClose.Reuse.VisibleRetainedDraftNotice' ($notice.Contains('Existing edits retained') -and $notice.Contains('saved catalog not reloaded'))
        Check 'UomPublicClose.Reuse.NoImplicitSave' (-not $book.Saved)
        CheckPair $before 'REUSED' 'UomPublicClose.Reuse.ExactOriginalPair'
        Check 'UomPublicClose.Reuse.UnknownColumnPreserved' ([string]$table.ListColumns.Item('Operator Annotation').DataBodyRange.Cells.Item(1,1).Value2 -ceq $canary)
        [void](Probe 'ShowUom');CaptureOwnedFormByCaptionEvidence 'Production' 'uom-public-reopened.png'
        Check 'UomPublicClose.UnrelatedWorkbookPreserved' ($decoy.Worksheets.Count -eq $decoyCount -and ($decoy.Worksheets.Item(1).UsedRange.Formula|ConvertTo-Json -Compress -Depth 5) -ceq $decoyBefore)
        $same=$true;foreach($file in $activityPins.Keys){$same=$same -and (Get-FileHash -LiteralPath $file).Hash -ceq $activityPins[$file]}
        Check 'UomPublicClose.OriginalActivityImmutable' $same
        Check 'UomPublicClose.SavedCatalogAuthorityPreserved' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configPin)
        $book.Close($false);$book=$null
        Check 'UomPublicClose.FinalClose.FormDisposed' ([long](Probe 'LoadedProductionFormsForTest') -eq 0)
    }finally{
        [void](Probe 'ForgetLaunchedUomForTest')
        if($null -ne $book){$book.Close($false)}
        if($null -ne $decoy){$decoy.Close($false)}
        SelectTarget $Fixture
    }
}
