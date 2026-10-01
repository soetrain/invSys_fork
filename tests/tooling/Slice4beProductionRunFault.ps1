# Exceptions and nested actual clicks at existing owner boundaries. Adapters are unsaved.
function Install-ProductionRunFaultProbe {
    function Seam($Module,[string]$Procedure,[string]$Anchor,[string]$Boundary,[bool]$After=$false){
        $start=$Module.ProcStartLine($Procedure,0);$end=$start+$Module.ProcCountLines($Procedure,0)
        $lines=@(for($i=$start;$i -lt $end;$i++){if($Module.Lines($i,1).Trim() -ceq $Anchor){$i}})
        if($lines.Count -ne 1){throw ('Run fault fixture seam changed: '+$Procedure+'; not product RED.')}
        $Module.InsertLines($lines[0]+[int]$After,('    TestProductionDesigner.RunFaultBoundary "'+$Boundary+'"'))
    }
    $project=$packages['invSys.Operations.xlam'].VBProject
    $owner=$project.VBComponents.Item('modProductionReusableRun').CodeModule
    Seam $owner 'LoadReleasedReusableRecipe' 'ClearReusableRun' 'LOAD' $true
    Seam $owner 'ApplyReusableRunScale' 'mScalePercent = scalePercent' 'SCALE'
    Seam $owner 'ClearReusableRun' 'mLoaded = False' 'CLEAR'
    Seam $owner 'ApplyReusableRunStockAllocation' 'If Not mLoaded Then' 'ALLOCATE'
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    Seam $form 'RefreshReusableRunControls' 'selectedProcess = ActiveRunProcess()' 'REFRESH'
    Seam $form 'SetAllRunTreeGroupsCollapsed' 'EnsureRunTreeState' 'PRESENTATION'
    $form.AddFromString(@'
Public Function RunFaultActForTest(ByVal action As String) As String
    On Error GoTo Failed
    Select Case action
        Case "LOAD": mBtnLoaderLoad_Click
        Case "SCALE": mBtnApplyBatchScale_Click
        Case "CLEAR": mBtnLoaderClear_Click
        Case "LOADER_REFRESH": mBtnLoaderRefresh_Click
        Case "MANAGER_REFRESH": mBtnManagerRefresh_Click
        Case "ALLOCATE": mBtnRunApplyPalette_Click
        Case "TREE_ALLOCATE": mBtnRunTreeApplyPalette_Click
        Case "TREE_EXPAND": mBtnRunTreeExpandAll_Click
        Case "TREE_COLLAPSE": mBtnRunTreeCollapseAll_Click
        Case Else: Err.Raise 5
    End Select
    RunFaultActForTest = mTxtStatus.Text
    Exit Function
Failed:
    RunFaultActForTest = "HANDLER_ERROR|" & CStr(Err.Number)
End Function
Public Function RunFaultGuardsForTest() As Boolean
    RunFaultGuardsForTest = Not mLoading And Not mDesignerActionInProgress
End Function
Public Sub RunFaultResetFixtureForTest()
    mLoading = False: mDesignerActionInProgress = False
End Sub
'@)
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,@'
Private mRunFaultMode As String, mRunFaultAction As String, mRunFaultBoundary As String
Private mRunFaultCanary As String, mRunFaultArmed As Boolean, mRunFaultCalls As Long
Private mRunFaultExceptionReached As Boolean, mRunFaultNestedReached As Boolean
'@)
    $adapter.AddFromString(@'
Public Sub RunFaultArm(ByVal mode As String, ByVal action As String, ByVal boundary As String, ByVal canary As String)
    mRunFaultMode = mode: mRunFaultAction = action: mRunFaultBoundary = boundary
    mRunFaultCanary = canary: mRunFaultCalls = 0: mRunFaultArmed = True
    mRunFaultExceptionReached = False: mRunFaultNestedReached = False
End Sub
Public Sub RunFaultBoundary(ByVal boundary As String)
    If boundary <> mRunFaultBoundary Then Exit Sub
    mRunFaultCalls = mRunFaultCalls + 1
    If Not mRunFaultArmed Then Exit Sub
    mRunFaultArmed = False
    If mRunFaultMode = "Exception" Then
        mRunFaultExceptionReached = True
        Err.Raise 5432, , "Injected Run owner interruption " & mRunFaultCanary
    ElseIf mRunFaultMode = "Nested" Then
        mRunFaultNestedReached = True
        Call mForm.RunFaultActForTest(mRunFaultAction)
    End If
End Sub
Public Function RunFaultAct(ByVal action As String) As String
    RunFaultAct = mForm.RunFaultActForTest(action)
End Function
Public Function RunFaultEvidence() As String
    RunFaultEvidence = CStr(mRunFaultCalls) & "|" & CStr(mRunFaultExceptionReached) & "|" & CStr(mRunFaultNestedReached)
End Function
Public Function RunFaultGuards() As Boolean
    RunFaultGuards = mForm.RunFaultGuardsForTest()
End Function
Public Sub RunFaultResetFixture()
    mRunFaultArmed = False: mRunFaultBoundary = ""
    mForm.RunFaultResetFixtureForTest
End Sub
'@)
}

function Test-ProductionRunFault($Fixture,$Other,[switch]$YieldOnly,[switch]$StockOnly,[switch]$WorksheetOnly,[switch]$WorksheetOwnerOnly) {
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    function Hash([string]$Path){$s=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $s).Hash}finally{$s.Dispose()}}
    function Navigation([string]$Action){$Action -cin @('TREE_EXPAND','TREE_COLLAPSE')}
    function State([string]$Action){
        $value=[string](Probe 'RunLocalState')
        if(Navigation $Action){$value+='|'+[string](Probe 'RunPresentationState' @($false))}
        $value
    }
    function Pair([string[]]$Before,[string]$Action,[string]$Outcome,[string]$Label){
        $raw=@(Files|Where-Object{$_ -cnotin $Before}|ForEach-Object{[IO.File]::ReadAllText($_)})
        $rows=@($raw|ForEach-Object{$_|ConvertFrom-Json})
        $first=@($rows|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($rows|Where-Object OutcomeCode -CEQ $Outcome)
        $pair=$rows.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
        $context=$pair;$safe=$pair;$linked=$false;$facts=$false;$terminal=$false
        foreach($row in $rows){$context=$context -and $row.ControlId -ceq ('PRODUCTION_RUN_'+$Action) -and $row.OwnerId -ceq 'PRODUCTION_RUN_LOCAL' -and $row.UserId -ceq 'config-producer' -and $row.WarehouseId -ceq $Fixture.Warehouse -and $row.CatalogVersion -eq 24;$safe=$safe -and @($row.SourceEventRefs).Count -eq 0}
        foreach($wire in $raw){foreach($value in @($canary,$Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'Injected Run owner interruption','HANDLER_ERROR','mBtn','RUN-KEY-','RUN-LOCATION','DEMO-RAW-BLACK-TEA')){
            $encoded=ConvertTo-Json -InputObject $value -Compress
            if($wire.Contains($value) -or $wire.Contains($encoded.Substring(1,$encoded.Length-2))){$safe=$false}
        }}
        if($pair){
            $linked=$first[0].ActivityId -cne '' -and $first[0].ActivityId -ceq $last[0].ActivityId -and $first[0].RecordId -cne $last[0].RecordId
            $severity=if($Outcome -ceq 'FAILED'){'Error'}elseif($Outcome -ceq 'REJECTED'){'Warning'}else{'Info'};$effect=if($Outcome -ceq 'FAILED'){'Unknown'}else{'Unchanged'}
            $facts=$first[0].Severity -ceq 'Info' -and $first[0].DataEffect -ceq 'Unknown' -and $last[0].Severity -ceq $severity -and $last[0].DataEffect -ceq $effect -and $last[0].EventCode -ceq ('PRODUCTION_RUN_'+$Action+'_'+$Outcome)
            $terminal=-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @(($first[0]|ConvertTo-Json -Depth 20 -Compress))) -and [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @(($last[0]|ConvertTo-Json -Depth 20 -Compress))) -eq ($Outcome -cin @('STAGED','REFRESHED','PRESENTED'))
        }
        Check ($Label+'.OneExactPair') $pair
        Check ($Label+'.CapturedContext') $context
        Check ($Label+'.NoEnteredDataOrRawFailure') $safe
        Check ($Label+'.DistinctCorrelatedRecords') $linked
        Check ($Label+'.OwnerFactWithoutRollbackClaim') $facts
        Check ($Label+'.ExactTerminal') $terminal
    }
    $actions=@('LOAD','SCALE','CLEAR','LOADER_REFRESH','MANAGER_REFRESH','ALLOCATE','TREE_ALLOCATE','TREE_EXPAND','TREE_COLLAPSE')
    $canary='RUNFAULT'+[guid]::NewGuid().ToString('N');$book=$null;$decoy=$null;$pins=@{};$recordPins=@{}
    try{
        SelectTarget $Fixture
        $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($Fixture.Warehouse,'S1','config-admin'))
        if(-not $seed.StartsWith('OK|')){throw 'Admin Seed Run fault fixture unavailable; not product RED.'}
        if($StockOnly){
            $ready=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.StockPrepareForTest' @($Fixture.Warehouse))
            if($ready -cne 'READY'){
                if($ready -notmatch '^FIXTURE_FAILED\|[A-Za-z]+(\|-?[0-9]+)?$'){$ready='Unavailable'}
                throw ('Owning stock fixture unavailable: '+$ready+'; not product RED.')
            }
            Check 'RunStock.OwningReceiveCreatesTwoExactEntities' $true
        }
        if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.RunPresentationPolicy' @($true))){throw 'Authorized Run fault policy unavailable; not product RED.'}
        SelectTarget $Fixture 'config-producer'
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$canary;$sheet.Cells.Item(2,2).Formula='=1+2'
        $path=Join-Path $runRoot 'run-fault-operator.xlsb';$book.SaveAs($path,50);$book.Close($false)
        $bookPin=Hash $path;$book=$excel.Workbooks.Open($path,0,$false);$sheet=$book.Worksheets.Item(1)
        $decoy=$excel.Workbooks.Add();$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name))
        if($StockOnly){[void](Probe 'RunStockConfigure' @([double](Run 'invSys.Core.xlam' 'TestShippingCatalog.StockQtyForTest')))}
        if(-not [bool](Probe 'ReadPrepare' @($canary))){throw 'Released Run fault definitions unavailable; not product RED.'}
        if($YieldOnly){[void](Probe 'RunLocalRememberFixture')}
        foreach($root in @($Fixture.Root,$Other.Root)){
            foreach($file in Get-ChildItem -LiteralPath $root -Recurse -File|Where-Object{$_.Extension -in '.xlsb','.xlsm' -and $_.Name -notlike '~$*'}){$pins[$file.FullName]=Hash $file.FullName}
        }
        foreach($file in Files){$recordPins[$file]=Hash $file};$otherBefore=@(Get-Slice4beActivityFiles $Other)
        $modes=if($YieldOnly -or $StockOnly -or $WorksheetOnly -or $WorksheetOwnerOnly){@()}else{@('Exception','Nested')}
        foreach($mode in $modes){
            foreach($action in $actions){
                [void](Probe 'RunFaultResetFixture')
                if([string](Probe 'RunLocalStage' @($action,'Normal')) -cne 'READY'){throw 'Released Run fault staging unavailable; not product RED.'}
                if(Navigation $action){
                    $shapeMode=if($action -ceq 'TREE_EXPAND'){'Collapsed'}else{'Expanded'}
                    if([string](Probe 'RunPresentationStage' @($shapeMode,$canary)) -cne 'READY'){throw 'Run fault Tree fixture unavailable; not product RED.'}
                }
                $state=State $action;$owner=[string](Probe 'RunLocalOwnerState')
                $palette=if(Navigation $action){[string](Probe 'RunPresentationState' @($true))}else{''}
                $boundary=if(Navigation $action){'PRESENTATION'}elseif($action.EndsWith('REFRESH')){'REFRESH'}elseif($action.EndsWith('ALLOCATE')){'ALLOCATE'}else{$action}
                [void](Probe 'RunFaultArm' @($mode,$action,$boundary,$canary));$before=@(Files)
                $notice=[string](Probe 'RunFaultAct' @($action));$evidence=([string](Probe 'RunFaultEvidence')).Split('|');$label='RunFault.'+$mode+'.'+$action
                $reached=$evidence.Count -eq 3 -and [int]$evidence[0] -ge 1 -and $evidence[$(if($mode -ceq 'Exception'){1}else{2})] -ceq 'True'
                if(-not $reached){throw ('Run owner fault boundary unavailable: '+$action+'; not behavioral RED.')}
                Check ($label+'.ActualBoundaryReached') $reached
                Check ($label+'.OneOwnerEntry') ([int]$evidence[0] -eq 1)
                Check ($label+'.GuardsRestoredWithoutAdapterReset') ([bool](Probe 'RunFaultGuards'))
                if($mode -ceq 'Exception'){
                    $propagates=$action -cnotin @('LOAD','ALLOCATE','TREE_ALLOCATE')
                    Check ($label+'.ExistingErrorPropagation') ($notice.StartsWith('HANDLER_ERROR|5432') -eq $propagates)
                    $preserved=if($action -ceq 'LOAD'){[bool](Probe 'RunLocalPreserved' @('LOAD','UnavailableRecipe'))}
                        elseif($action -ceq 'LOADER_REFRESH'){[bool](Probe 'RunLocalPreserved' @($action,'Normal'))}
                        else{(State $action) -ceq $state}
                    Check ($label+'.ExistingPartialLocalState') $preserved
                    $outcome='FAILED'
                }else{
                    if(Navigation $action){
                        $shape=if($action -ceq 'TREE_EXPAND'){'2|4|0'}else{'2|2|1'}
                        $preserved=[string](Probe 'RunPresentationShape') -ceq $shape -and [string](Probe 'RunPresentationState' @($true)) -ceq $palette -and [string](Probe 'RunLocalOwnerState') -ceq $owner
                    }else{$preserved=[bool](Probe 'RunLocalPreserved' @($action,'Normal'))}
                    Check ($label+'.ExistingOwnerResult') $preserved
                    Check ($label+'.NoUnhandledError') (-not $notice.StartsWith('HANDLER_ERROR|'))
                    $outcome=if(Navigation $action){'PRESENTED'}elseif($action.EndsWith('REFRESH')){'REFRESHED'}else{'STAGED'}
                }
                Pair $before $action $outcome $label
            }
        }
        if($YieldOnly){Test-ProductionRunYield $Fixture $Other $book}
        if($StockOnly){Test-ProductionRunStock $Fixture}
        if($WorksheetOnly){Test-ProductionRunWorksheet $Fixture}
        if($WorksheetOwnerOnly){Test-ProductionRunWorksheetOwner $Fixture}
        [void](Probe 'RunFaultResetFixture')
        Check 'RunFault.UnknownValuesAndFormula' ($sheet.Cells.Item(1,1).Value2 -ceq 'Operator Extra' -and $sheet.Cells.Item(2,1).Value2 -ceq $canary -and $sheet.Cells.Item(2,2).Formula -ceq '=1+2')
        $same=$true;foreach($file in $pins.Keys){$same=$same -and (Hash $file) -ceq $pins[$file]};Check 'RunFault.SavedAuthorityPreserved' $same
        $same=$true;foreach($file in $recordPins.Keys){$same=$same -and (Hash $file) -ceq $recordPins[$file]};Check 'RunFault.OlderRecordsImmutable' $same
        Check 'RunFault.OtherWarehouseActivityPreserved' ((@(Get-Slice4beActivityFiles $Other) -join '|') -ceq ($otherBefore -join '|'))
        [void](Probe 'CloseDesigner');Check 'RunFault.OperatorBytesPreserved' ((Hash $path) -ceq $bookPin)
    }finally{[void](Probe 'CloseDesigner');if($null -ne $book){$book.Close($false)};if($null -ne $decoy){$decoy.Close($false)};SelectTarget $Fixture}
}
