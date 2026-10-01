# Real handlers and disposable policy/storage fixtures. No owner handler is replaced.
function Install-ProductionRunPolicyProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $owner=$project.VBComponents.Item('modProductionReusableRun').CodeModule
    $owner.InsertLines($owner.CountOfDeclarationLines+1,'Public RunPolicyOwnerCallsForTest As Long')
    foreach($name in @('ClearReusableRun','LoadReleasedReusableRecipe','ApplyReusableRunScale','ReusableRunLoaderRows','ReusableRunPaletteRows','ReusableRunManagerCheckRows','ReusableRunOutputRows','ApplyReusableRunStockAllocation')){
        $start=$owner.ProcBodyLine($name,0);$end=$start+$owner.ProcCountLines($name,0)
        while($owner.Lines($start,1).TrimEnd().EndsWith('_')){$start++;if($start -ge $end){throw 'Owner signature unavailable; not product RED.'}}
        $owner.InsertLines($start+1,'    RunPolicyOwnerCallsForTest = RunPolicyOwnerCallsForTest + 1')
    }
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $start=$form.ProcBodyLine('SetAllRunTreeGroupsCollapsed',0)
    $form.InsertLines($start+1,'    modProductionReusableRun.RunPolicyOwnerCallsForTest = modProductionReusableRun.RunPolicyOwnerCallsForTest + 1')
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Sub RunPolicyResetOwnerCalls()
    modProductionReusableRun.RunPolicyOwnerCallsForTest = 0
End Sub
Public Function RunPolicyOwnerCalls() As Long
    RunPolicyOwnerCalls = modProductionReusableRun.RunPolicyOwnerCallsForTest
End Function
'@)
}
function Test-ProductionRunPolicy($Fixture,$Other) {
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    function Hash([string]$Path){$s=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $s).Hash}finally{$s.Dispose()}}
    function Navigation([string]$Action){$Action -cin @('TREE_EXPAND','TREE_COLLAPSE')}
    function Stage([string]$Action){
        $ready=[string](Probe 'RunLocalStage' @($Action,'Normal'))
        if($ready -cne 'READY'){throw 'Released Run policy fixture unavailable; not product RED.'}
        if(Navigation $Action){
            $mode=if($Action -ceq 'TREE_EXPAND'){'Collapsed'}else{'Expanded'}
            if([string](Probe 'RunPresentationStage' @($mode,$canary)) -cne 'READY'){throw 'Run presentation policy fixture unavailable; not product RED.'}
        }
    }
    function State([string]$Action){
        $value=[string](Probe 'RunLocalState')
        if(Navigation $Action){$value+='|'+[string](Probe 'RunPresentationState' @($false))}
        $value
    }
    function Act([string]$Action){
        if(Navigation $Action){[string](Probe 'RunPresentationAct' @($Action,''))}else{[string](Probe 'RunLocalAct' @($Action,''))}
    }
    function Preserved([string]$Action,[string]$Palette,[string]$Owner){
        if(Navigation $Action){
            $shape=if($Action -ceq 'TREE_EXPAND'){'2|4|0'}else{'2|2|1'}
            ([string](Probe 'RunPresentationShape') -ceq $shape -and [string](Probe 'RunPresentationState' @($true)) -ceq $Palette -and [string](Probe 'RunLocalOwnerState') -ceq $Owner)
        }else{[bool](Probe 'RunLocalPreserved' @($Action,'Normal'))}
    }
    $actions=@('LOAD','SCALE','CLEAR','LOADER_REFRESH','MANAGER_REFRESH','ALLOCATE','TREE_ALLOCATE','TREE_EXPAND','TREE_COLLAPSE')
    $canary='RUNPOLICY'+[guid]::NewGuid().ToString('N');$book=$null;$decoy=$null;$pins=@{};$recordPins=@{}
    $configBytes=$null
    try{
        SelectTarget $Fixture
        $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($Fixture.Warehouse,'S1','config-admin'))
        if(-not $seed.StartsWith('OK|')){throw 'Admin Seed Run policy fixture unavailable; not product RED.'}
        if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.RunPresentationPolicy' @($true))){throw 'Authorized Run tracking policy unavailable; not product RED.'}
        $configBytes=[IO.File]::ReadAllBytes($Fixture.Config)
        SelectTarget $Fixture 'config-producer'
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$canary;$sheet.Cells.Item(2,2).Formula='=1+2'
        $path=Join-Path $runRoot 'run-policy-operator.xlsb';$book.SaveAs($path,50);$book.Close($false)
        $bookPin=Hash $path;$book=$excel.Workbooks.Open($path,0,$false);$sheet=$book.Worksheets.Item(1)
        $decoy=$excel.Workbooks.Add();$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name))
        if(-not [bool](Probe 'ReadPrepare' @($canary))){throw 'Released Run policy definitions unavailable; not product RED.'}
        [void](Probe 'RunLocalRememberFixture')
        foreach($root in @($Fixture.Root,$Other.Root)){
            foreach($file in Get-ChildItem -LiteralPath $root -Recurse -File|Where-Object{$_.Extension -in '.xlsb','.xlsm' -and $_.Name -notlike '~$*'}){$pins[$file.FullName]=Hash $file.FullName}
        }
        foreach($file in Files){$recordPins[$file]=Hash $file}
        $otherBefore=@(Get-Slice4beActivityFiles $Other)
        SelectTarget $Fixture 'config-reader';[void](Probe 'RunLocalReopen' @($book.Name))
        foreach($action in $actions){
            Stage $action;$state=State $action;$before=@(Files);[void](Probe 'RunPolicyResetOwnerCalls')
            $notice=Act $action;$label='RunPolicy.Denied.'+$action
            Check ($label+'.OwnerNotInvoked') ([int](Probe 'RunPolicyOwnerCalls') -eq 0)
            Check ($label+'.NoReadsOrMutation') ((State $action) -ceq $state -and [int](Probe 'ReadCalls') -eq 0)
            Check ($label+'.NoUnhandledError') (-not $notice.StartsWith('HANDLER_ERROR|'))
            $rows=@(Files|Where-Object{$_ -cnotin $before}|ForEach-Object{[IO.File]::ReadAllText($_)|ConvertFrom-Json})
            $first=@($rows|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($rows|Where-Object OutcomeCode -CEQ 'DENIED')
            $ok=$rows.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
            if($ok){$ok=$first[0].ActivityId -cne '' -and $first[0].ActivityId -ceq $last[0].ActivityId -and $first[0].RecordId -cne $last[0].RecordId}
            foreach($row in $rows){$ok=$ok -and $row.ControlId -ceq ('PRODUCTION_RUN_'+$action) -and $row.OwnerId -ceq 'PRODUCTION_RUN_LOCAL' -and $row.UserId -ceq 'config-reader' -and $row.WarehouseId -ceq $Fixture.Warehouse -and @($row.SourceEventRefs).Count -eq 0}
            Check ($label+'.ExactDeniedPair') $ok
            Check ($label+'.PreMutationDenial') ($ok -and $last[0].Severity -ceq 'Blocked' -and $last[0].DataEffect -ceq 'Unchanged')
        }
        foreach($mode in @('Off','Older','NavigationOff','Unavailable')){
            $blocked=Join-Path (Join-Path $Fixture.Root 'Training\Activity') $Fixture.Warehouse;$held=$blocked+'-run-held';$moved=$false
            foreach($candidate in @($blocked,$held)){if(-not [IO.Path]::GetFullPath($candidate).StartsWith([IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Policy path escaped disposable root.'}}
            if(Test-Path -LiteralPath $held){throw 'Preserve existing held fixture.'}
            try{
                SelectTarget $Fixture
                if($mode -ceq 'Off'){
                    if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadPolicyForTest' @($false))){throw 'Disabled policy fixture unavailable.'}
                }
                if($mode -ceq 'NavigationOff'){
                    if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.RunPresentationPolicy' @($false))){throw 'Navigation policy fixture unavailable.'}
                }
                if($mode -ceq 'Older'){
                    $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
                    try{
                        (Table $cfg 'tblEventTrackingPolicies').ListColumns.Item('CatalogVersion').DataBodyRange.Value2=23.0
                        $controls=Table $cfg 'tblEventTrackingControls'
                        for($i=$controls.ListRows.Count;$i -ge 1;$i--){if(([string]$controls.ListRows.Item($i).Range.Cells.Item(1,$controls.ListColumns.Item('ControlId').Index).Value2).StartsWith('PRODUCTION_RUN_')){$controls.ListRows.Item($i).Delete()}}
                        $cfg.Save()
                    }finally{$cfg.Close($false)}
                }
                SelectTarget $Fixture 'config-producer';[void](Probe 'RunLocalReopen' @($book.Name))
                if($mode -ceq 'Older'){Check 'RunPolicy.Older.ExistingControlReadable' (([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @('PRODUCTION_CLOSE'))).StartsWith('True|True|'))}
                if($mode -ceq 'Unavailable'){
                    if(Test-Path -LiteralPath $blocked){Move-Item -LiteralPath $blocked -Destination $held;$moved=$true}
                    [IO.File]::WriteAllText($blocked,'Blocked disposable Run activity path')
                }
                $selected=if($mode -ceq 'NavigationOff'){@('TREE_EXPAND','TREE_COLLAPSE')}else{$actions}
                foreach($action in $selected){
                    $before=@(Files);Stage $action
                    Check ('RunPolicy.'+$mode+'.'+$action+'.SetupNotUserAction') (@(Files).Count -eq $before.Count)
                    $owner=[string](Probe 'RunLocalOwnerState');$palette=if(Navigation $action){[string](Probe 'RunPresentationState' @($true))}else{''}
                    $notice=Act $action;$label='RunPolicy.'+$mode+'.'+$action
                    Check ($label+'.AuthorizedActionContinues') (Preserved $action $palette $owner)
                    Check ($label+'.NoUnhandledError') (-not $notice.StartsWith('HANDLER_ERROR|'))
                    Check ($label+'.NoActivityOrFallback') (@(Files).Count -eq $before.Count)
                    if($mode -ceq 'Unavailable'){Check ($label+'.NoticeVisible') $notice.Contains('Tracking unavailable')}
                }
            }finally{
                if($mode -ceq 'Unavailable' -and (Test-Path -LiteralPath $blocked -PathType Leaf)){Remove-Item -LiteralPath $blocked}
                if($moved){Move-Item -LiteralPath $held -Destination $blocked}
                [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
            }
        }
        Check 'RunPolicy.UnknownValuesAndFormula' ($sheet.Cells.Item(1,1).Value2 -ceq 'Operator Extra' -and $sheet.Cells.Item(2,1).Value2 -ceq $canary -and $sheet.Cells.Item(2,2).Formula -ceq '=1+2')
        $same=$true;foreach($file in $pins.Keys){$same=$same -and (Hash $file) -ceq $pins[$file]};Check 'RunPolicy.SavedAuthorityPreserved' $same
        $same=$true;foreach($file in $recordPins.Keys){$same=$same -and (Hash $file) -ceq $recordPins[$file]};Check 'RunPolicy.OlderRecordsImmutable' $same
        Check 'RunPolicy.OtherWarehouseActivityPreserved' ((@(Get-Slice4beActivityFiles $Other) -join '|') -ceq ($otherBefore -join '|'))
        [void](Probe 'CloseDesigner');Check 'RunPolicy.OperatorBytesPreserved' ((Hash $path) -ceq $bookPin)
    }finally{
        [void](Probe 'CloseDesigner');if($null -ne $book){$book.Close($false)};if($null -ne $decoy){$decoy.Close($false)}
        if($null -ne $configBytes){[IO.File]::WriteAllBytes($Fixture.Config,$configBytes)}
        SelectTarget $Fixture
    }
}
