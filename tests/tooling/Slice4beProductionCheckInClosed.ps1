# Native form/workbook lifetime at Check In entry, using the actual Click handler.
function Install-ProductionCheckInClosedProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $start=$form.ProcStartLine('mBtnManagerCheckIn_Click',0);$end=$start+$form.ProcCountLines('mBtnManagerCheckIn_Click',0)
    $hits=@(for($line=$start;$line -lt $end;$line++){if($form.Lines($line,1).Trim().StartsWith('modProductionCheckInActions.Execute ',[StringComparison]::Ordinal)){$line}})
    if($hits.Count -ne 1){throw 'Check In Click entry anchor changed; not product RED.'}
    $form.InsertLines($hits[0],'    TestProductionDesigner.CheckClosedHandlerHit')
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,'Private mCheckClosedEntries As Long')
    $adapter.AddFromString(@'
Public Sub CheckClosedHandlerHit()
    mCheckClosedEntries = mCheckClosedEntries + 1
End Sub
Public Sub CheckClosedReset()
    mCheckClosedEntries = 0
End Sub
Public Function CheckClosedEntries() As Long
    CheckClosedEntries = mCheckClosedEntries
End Function
'@)
}

function Test-ProductionCheckInClosed($Fixture,$Other){
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Hash([string]$Path){$s=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $s).Hash}finally{$s.Dispose()}}
    $book=$null;$decoy=$null;$pins=@{};$canary='CLOSED'+[guid]::NewGuid().ToString('N')
    SelectTarget $Fixture
    $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($Fixture.Warehouse,'S1','config-admin'))
    if(-not $seed.StartsWith('OK|')){throw 'Admin Seed unavailable; not product RED.'}
    if([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.StockPrepareForTest' @($Fixture.Warehouse)) -cne 'READY'){throw 'Real stock fixture unavailable; not product RED.'}
    try{
        SelectTarget $Fixture 'config-producer'
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$canary;$sheet.Cells.Item(2,2).Formula='=1+2'
        $path=Join-Path $runRoot 'check-in-closed.xlsb';$book.SaveAs($path,50);$book.Close($false)
        $bookPin=Hash $path;$book=$excel.Workbooks.Open($path,0,$false)
        $decoy=$excel.Workbooks.Add();$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name))
        if(-not [bool](Probe 'ReadPrepare' @($canary))){throw 'Released reusable fixture unavailable; not product RED.'}
        [void](Probe 'RunLocalRememberFixture')
        foreach($root in @($Fixture.Root,$Other.Root)){
            foreach($file in Get-ChildItem -LiteralPath $root -Recurse -File|Where-Object{$_.Extension -in '.xlsb','.xlsm' -and $_.Name -notlike '~$*'}){$pins[$file.FullName]=Hash $file.FullName}
        }
        foreach($mode in @('Reusable','Worksheet')){
            SelectTarget $Fixture 'config-producer'
            if($null -eq $book){$book=$excel.Workbooks.Open($path,0,$false)}
            [void](Probe 'RunLocalReopen' @($book.Name))
            if($mode -ceq 'Worksheet'){
                # The first close discarded the prior workbook's unsaved staging.
                # Build this fixture on the newly opened disposable workbook.
                if(-not [bool](Probe 'RunWorksheetPrepare')){throw 'Reopened worksheet fixture unavailable; not product RED.'}
                $selectedKey=[string](Probe 'RunWorksheetKey')
                if(-not $selectedKey){throw 'Real worksheet identity unavailable; not product RED.'}
            }
            $ready=if($mode -ceq 'Reusable'){[bool](Probe 'CheckBaselineReusableStage' @('Selected'))}else{[bool](Probe 'CheckBaselineWorksheetStage' @($selectedKey,$canary))}
            if(-not $ready){throw 'Closed Check In staging unavailable; not product RED.'}
            Initialize-SettingsCapture
            $page=[int](Probe 'RunLocalShowAndCapture' @($book.Name,'CHECK_IN'))
            [void](Probe 'CheckClosedReset')
            $ownerBefore=[string](Probe 'RunLocalOwnerState');$projectionBefore=[string](Probe 'RunLocalState')
            $before=@(Get-Slice4beActivityFiles $Fixture);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            $capturedName=$book.Name;$decoyName=$decoy.Name
            $receipt=[ordered]@{Mode=$mode;PageBefore=$page;BeforeUTC=[DateTimeOffset]::UtcNow.ToString('o');VisibleBefore=([InvSysSettingsCapture]::OwnedVisibleForm('Production',[IntPtr]$excel.Hwnd) -ne [IntPtr]::Zero);WorkbooksBefore=$excel.Workbooks.Count}
            $decoy.Activate();$book.Close($false);$book=$null
            $names=@(foreach($openBook in $excel.Workbooks){[string]$openBook.Name})
            $receipt.AfterUTC=[DateTimeOffset]::UtcNow.ToString('o')
            $receipt.VisibleAfter=([InvSysSettingsCapture]::OwnedVisibleForm('Production',[IntPtr]$excel.Hwnd) -ne [IntPtr]::Zero)
            $receipt.WorkbooksAfter=$excel.Workbooks.Count
            $receipt.CapturedStillOpen=($capturedName -cin $names);$receipt.DecoyStillOpen=($decoyName -cin $names)
            if(-not $receipt.VisibleBefore -or $receipt.CapturedStillOpen -or -not $receipt.DecoyStillOpen){throw 'Native closed workbook prerequisite unavailable; not product RED.'}
            $receipt.HandlerInvoked=[bool]$receipt.VisibleAfter;$receipt.HandlerEntries=0
            $protected=-not $receipt.VisibleAfter;$guards=$true;$projectionPreserved=$true
            if($receipt.VisibleAfter){
                $returned=[bool](Probe 'CheckBaselineAct' @(''))
                $receipt.HandlerEntries=[int](Probe 'CheckClosedEntries')
                if($receipt.HandlerEntries -ne 1){throw 'Surviving Check In Click handler was not exercised; not product RED.'}
                $protected=$returned -and [bool](Probe 'CheckBaselineContextRefused')
                $guards=[bool](Probe 'CheckBaselineGuards')
                $projectionPreserved=([string](Probe 'RunLocalState') -ceq $projectionBefore)
                if($CaptureEvidence){CaptureOwnedFormByCaptionEvidence 'Production' ('check-in-closed-'+$mode.ToLowerInvariant()+'.png')}
            }
            $label='CheckInClosed.Entry.'+$mode
            Check ($label+'.NativeSurfaceLifetimeEstablished') ($page -eq 3 -and $receipt.WorkbooksBefore -eq $receipt.WorkbooksAfter+1)
            Check ($label+'.NativeActionOpportunityProtected') $protected
            Check ($label+'.CapturedBindingRejected') (-not [bool](Probe 'RunLocalClosedBindingCurrent'))
            Check ($label+'.OwnerPreserved') ([string](Probe 'RunLocalOwnerState') -ceq $ownerBefore)
            Check ($label+'.SurvivingProjectionPreserved') $projectionPreserved
            Check ($label+'.GuardsRestored') $guards
            Check ($label+'.SavedOperatorBytesPreserved') ((Hash $path) -ceq $bookPin)
            Check ($label+'.NoActivityOrRedirectedRecords') ((@(Get-Slice4beActivityFiles $Fixture) -join '|') -ceq ($before -join '|') -and (@(Get-Slice4beActivityFiles $Other) -join '|') -ceq ($otherBefore -join '|'))
            $receipt|ConvertTo-Json|Set-Content (Join-Path $reportRoot ('check-in-closed-'+$mode.ToLowerInvariant()+'.json'))
            [void](Probe 'RunLocalSafeClose')
        }
        $same=$true;foreach($file in $pins.Keys){$same=$same -and (Hash $file) -ceq $pins[$file]}
        Check 'CheckInClosed.SavedAuthorityPreserved' $same
    }finally{
        [void](Probe 'RunLocalSafeClose')
        if($null -ne $book){$book.Close($false)}
        if($null -ne $decoy){$decoy.Close($false)}
        SelectTarget $Fixture
    }
}
