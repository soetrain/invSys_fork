# D18 captured context through the real Next Batch handler; no runtime replacement.
function Install-ProductionNextBaselineProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $owner=$project.VBComponents.Item('modProductionReusableRun').CodeModule
    $start=$owner.ProcStartLine('BeginNextReusableBatch',0)
    $body=$owner.Lines($start,$owner.ProcCountLines('BeginNextReusableBatch',0)) -split '\r?\n'
    $entries=@(for($i=0;$i -lt $body.Count;$i++){if($body[$i].Trim() -like 'Public Function BeginNextReusableBatch(*) As Boolean'){$start+$i+1}})
    if($entries.Count -ne 1){throw 'Next Batch owner entry anchor changed; not product RED.'}
    $owner.InsertLines($entries[0],'    TestProductionDesigner.NextBaselineOwnerHit')
    $project.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Function NextBaselineActForTest() As Boolean
    mBtnManagerNext_Click
    NextBaselineActForTest = True
End Function
Public Function NextBaselineReadyForTest() As Boolean
    NextBaselineReadyForTest = modProductionReusableRun.ReusableRunIsLoaded() And _
        Not modProductionReusableRun.ReusableRunIsCompleted() And _
        Not modProductionReusableRun.ReusableRunIsCheckedIn() And _
        modProductionReusableRun.RunLocalAllocatedForTest() = 0 And _
        InStr(1, mTxtStatus.Text, "Next Batch ", vbBinaryCompare) = 1
End Function
'@)
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,'Private mNextBaselineEntries As Long')
    $adapter.AddFromString(@'
Public Sub NextBaselineOwnerHit()
    mNextBaselineEntries = mNextBaselineEntries + 1
End Sub
Public Function NextBaselineAct() As Boolean
    mNextBaselineEntries = 0
    NextBaselineAct = mForm.NextBaselineActForTest()
End Function
Public Function NextBaselineEntries() As Long
    NextBaselineEntries = mNextBaselineEntries
End Function
Public Function NextBaselineReady() As Boolean
    NextBaselineReady = mForm.NextBaselineReadyForTest()
End Function
'@)
}

function Test-ProductionNextBaseline($Fixture,$Other) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Hash([string]$Path){$s=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $s).Hash}finally{$s.Dispose()}}
    $canary='NEXT'+[guid]::NewGuid().ToString('N');$book=$null;$decoy=$null
    SelectTarget $Fixture
    $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($Fixture.Warehouse,'S1','config-admin'))
    if(-not $seed.StartsWith('OK|')){throw 'Admin seed prerequisite unavailable; not product RED.'}
    $otherPins=RestartPins $Other.Root
    try {
        SelectTarget $Fixture 'config-producer'
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$canary;$sheet.Cells.Item(2,2).Formula='=1+2'
        $path=Join-Path $runRoot 'next-baseline.xlsb';$book.SaveAs($path,50);$book.Close($false)
        $bookPin=Hash $path;$book=$excel.Workbooks.Open($path,0,$false);$sheet=$book.Worksheets.Item(1)
        $decoy=$excel.Workbooks.Add();$decoySheet=$decoy.Worksheets.Item(1);$decoySheet.Cells.Item(1,1).Value2=$canary
        [void](Probe 'OpenDesigner' @($book.Name))
        if(-not [bool](Probe 'ReadPrepare' @($canary))){throw 'Released definitions unavailable; not product RED.'}
        [void](Probe 'RunLocalRememberFixture')
        foreach($guard in @('Current','Target','Session','SignedOut')) {
            SelectTarget $Fixture 'config-producer'
            [void](Probe 'RunLocalReopen' @($book.Name))
            if(-not [bool](Probe 'CompleteBaselinePrepare' @($true)) -or
               -not [bool](Probe 'CompleteBaselineAct') -or
               -not [bool](Probe 'CompleteBaselineCompleted')){throw 'Actual completed batch prerequisite unavailable; not product RED.'}
            $before=[string](Probe 'RunLocalOwnerState');$projection=[string](Probe 'RunLocalState')
            [void](Probe 'RunLocalShowAndCapture' @($book.Name,'CHECK_IN'));$decoy.Activate()
            if($guard -ceq 'Target'){SelectTarget $Other 'config-producer'}
            if($guard -ceq 'Session'){SelectTarget $Fixture 'config-producer'}
            if($guard -ceq 'SignedOut'){[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
            if([bool](Probe 'RunLocalClosedBindingCurrent') -ne ($guard -ceq 'Current')){throw 'Context prerequisite mismatch; not product RED.'}
            $label='NextBaseline.'+$guard
            Check ($label+'.ActualHandlerReturned') ([bool](Probe 'NextBaselineAct'))
            if($guard -ceq 'Current') {
                Check ($label+'.OneOwnerEntry') ([int](Probe 'NextBaselineEntries') -eq 1)
                Check ($label+'.NextBatchReady') ([bool](Probe 'NextBaselineReady'))
            } else {
                Check ($label+'.NoOwnerEntry') ([int](Probe 'NextBaselineEntries') -eq 0)
                Check ($label+'.OwnerPreserved') ([string](Probe 'RunLocalOwnerState') -ceq $before)
                Check ($label+'.ProjectionPreserved') ([string](Probe 'RunLocalState') -ceq $projection)
                Check ($label+'.VisibleContextRefusal') ([bool](Probe 'CheckBaselineContextRefused'))
            }
            Check ($label+'.CapturedCustomValueAndFormula') ($sheet.Cells.Item(2,1).Value2 -ceq $canary -and $sheet.Cells.Item(2,2).Formula -ceq '=1+2')
            Check ($label+'.DecoyPreserved') ($decoySheet.Cells.Item(1,1).Value2 -ceq $canary -and $decoy.Worksheets.Count -eq 1)
            CaptureOwnedFormByCaptionEvidence 'Production' ('next-baseline-'+$guard.ToLowerInvariant()+'.png')
        }
        if($RunNextActivityOnly){Test-ProductionNextActivity $Fixture $Other $book $decoy $canary}
        if($RunNextClosedOnly){Test-ProductionNextClosed $Fixture $Other $book $decoy $canary}
        elseif($RunNextYieldOnly){Test-ProductionNextYield $Fixture $Other $book $decoy $canary}
        [void](Probe 'CloseDesigner');$book.Close($false);$book=$null
        Check 'NextBaseline.OperatorBytesPreserved' ((Hash $path) -ceq $bookPin)
        Check 'NextBaseline.OtherWarehousePreserved' (RestartPinsEqual $otherPins $Other.Root)
    } finally {
        [void](Probe 'CloseDesigner')
        if($null -ne $book){$book.Close($false)}
        if($null -ne $decoy){$decoy.Close($false)}
    }
}
