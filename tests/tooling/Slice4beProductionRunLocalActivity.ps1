# Initial reusable-owner baseline. Worksheet branches, policy/fault companions and
# recording/publication/independent-reader evidence remain separate required gates.
function Test-ProductionRunLocalActivity($Fixture,$Other) {
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    function Hash([string]$Path){$s=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $s).Hash}finally{$s.Dispose()}}
    function State {[string](Probe 'RunLocalState')}
    function Stage([string]$Action,[string]$Mode='Normal'){
        $ready=[string](Probe 'RunLocalStage' @($Action,$Mode))
        if($ready -cne 'READY'){
            if($ready -notmatch '^FIXTURE_FAILED\|[A-Za-z]+(\|-?[0-9]+)?$'){$ready='Unavailable'}
            throw ('Run local fixture unavailable: '+$ready+'; not product RED.')
        }
    }
    function Positive([string]$Action){if($Action.EndsWith('REFRESH')){'REFRESHED'}else{'STAGED'}}
    function Pair([string[]]$Before,[string]$Action,[string]$Outcome,[string]$Case){
        $raw=@(Files|Where-Object{$_ -cnotin $Before}|ForEach-Object{[IO.File]::ReadAllText($_)})
        $records=@($raw|ForEach-Object{$_|ConvertFrom-Json})
        $first=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($records|Where-Object OutcomeCode -CEQ $Outcome)
        $paired=$records.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
        $context=$paired;$safe=$paired;$integrity=$paired;$linked=$false;$facts=$false;$terminal=$false;$distinct=$false
        $control='PRODUCTION_RUN_'+$Action
        foreach($r in $records){
            $context=$context -and $r.ControlId -ceq $control -and $r.OwnerId -ceq 'PRODUCTION_RUN_LOCAL' -and $r.UserId -ceq 'config-producer' -and $r.WarehouseId -ceq $Fixture.Warehouse -and $r.StationId -ceq 'S1' -and $r.CatalogVersion -eq 24
            $safe=$safe -and @($r.SourceEventRefs).Count -eq 0
        }
        foreach($value in $raw){
            foreach($secret in @($canary,$Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'mBtn','PinHash','DEMO-RAW-BLACK-TEA','RUN-OTHER')){
                $encoded=ConvertTo-Json -InputObject $secret -Compress
                if($value.Contains($secret) -or $value.Contains($encoded.Substring(1,$encoded.Length-2))){$safe=$false}
            }
            $match=[regex]::Match($value,',"ContentSha256":"([a-f0-9]{64})"\}$')
            if(-not $match.Success){$integrity=$false;continue}
            $sha=[Security.Cryptography.SHA256]::Create()
            try{$hash=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($value.Substring(0,$match.Index)+'}'))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
            $integrity=$integrity -and $hash -ceq $match.Groups[1].Value
        }
        if($paired){
            $linked=$first[0].ActivityId -cne '' -and $first[0].ActivityId -ceq $last[0].ActivityId -and $first[0].RecordId -cne $last[0].RecordId
            $distinct=$activityIds.Add([string]$first[0].ActivityId)
            $severity=if($Outcome -ceq 'REJECTED'){'Warning'}elseif($Outcome -ceq 'FAILED'){'Error'}else{'Info'}
            $effect=if($Outcome -ceq 'FAILED'){'Unknown'}else{'Unchanged'}
            $facts=$first[0].DataEffect -ceq 'Unknown' -and $first[0].Severity -ceq 'Info' -and $last[0].DataEffect -ceq $effect -and $last[0].Severity -ceq $severity -and $last[0].EventCode -ceq ($control+'_'+$Outcome)
            $attempt=[bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @(($first[0]|ConvertTo-Json -Depth 20 -Compress)))
            $completed=[bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @(($last[0]|ConvertTo-Json -Depth 20 -Compress)))
            $terminal=-not $attempt -and $completed -eq ($Outcome -ceq (Positive $Action))
        }
        Check ($Case+'.AttemptAndOutcome') $paired
        Check ($Case+'.ExactContext') $context
        Check ($Case+'.NoEnteredDataOrSources') $safe
        Check ($Case+'.Integrity') $integrity
        Check ($Case+'.LinkedDistinctRecords') $linked
        Check ($Case+'.DistinctActionIdentity') $distinct
        Check ($Case+'.OwnerFact') $facts
        Check ($Case+'.ExactCommandTerminal') $terminal
    }
    $canary='RUNLOCAL'+[guid]::NewGuid().ToString('N');$book=$null;$decoy=$null;$pins=@{};$recordPins=@{}
    $activityIds=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $actions=@('LOAD','SCALE','CLEAR','LOADER_REFRESH','MANAGER_REFRESH','ALLOCATE','TREE_ALLOCATE')
    if($RunLocalRefillDiagnostic){$actions=@()}
    SelectTarget $Fixture
    $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($Fixture.Warehouse,'S1','config-admin'))
    if(-not $seed.StartsWith('OK|')){throw 'Admin Seed fixture unavailable; not product RED.'}
    if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadPolicyForTest' @($true))){throw 'Authorized Run policy fixture unavailable; not product RED.'}
    try{
        SelectTarget $Fixture 'config-producer'
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$canary;$sheet.Cells.Item(2,2).Formula='=1+2'
        $path=Join-Path $runRoot 'run-local-operator.xlsb';$book.SaveAs($path,50);$book.Close($false)
        $bookPin=Hash $path;$book=$excel.Workbooks.Open($path,0,$false);$sheet=$book.Worksheets.Item(1)
        $decoy=$excel.Workbooks.Add();$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name))
        $ready=[bool](Probe 'ReadPrepare' @($canary))
        if(-not $ready){
            $stage=[string](Probe 'ReadSetupStage');if($stage -cnotin @('ReleasedProcess','RecipeDraft','RecipeSave','RecipeRelease','ReleasedStatus','Ready')){$stage='Unavailable'}
            throw ('Real released Run fixture unavailable at '+$stage+'; not product RED.')
        }
        Check 'RunLocal.AdminSeedAndReleasedDefinitions' $ready
        [void](Probe 'RunLocalRememberFixture')
        foreach($root in @($Fixture.Root,$Other.Root)){
            foreach($file in Get-ChildItem -LiteralPath $root -Recurse -File|Where-Object{$_.Extension -in '.xlsb','.xlsm' -and $_.Name -notlike '~$*'}){$pins[$file.FullName]=Hash $file.FullName}
        }
        foreach($file in Files){$recordPins[$file]=Hash $file}
        $measuredActions=if($RunLocalClosedDiagnostic){@()}else{$actions}
        foreach($action in $measuredActions){
            $json=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @(('PRODUCTION_RUN_'+$action),24));$def=if($json){$json|ConvertFrom-Json}else{$null}
            $caption=switch($action){'LOAD'{'Load Recipe'} 'SCALE'{'Apply Scale'} 'CLEAR'{'Clear Run'} 'LOADER_REFRESH'{'Refresh'} 'MANAGER_REFRESH'{'Refresh'} default{'Apply'}}
            $surface=if($action -ceq 'TREE_ALLOCATE'){'Operations > Production > Production Run - Tree'}else{'Operations > Production > Production Run - List'}
            Check ('RunLocal.'+$action+'.FixedMetadata') ($null -ne $def -and $def.Caption -ceq $caption -and $def.OwnerId -ceq 'PRODUCTION_RUN_LOCAL' -and $def.Class -ceq 'Command' -and $def.Role -ceq 'Production' -and $def.Capability -ceq 'PROD_POST' -and $def.Surface -ceq $surface)
            Stage $action;$before=@(Files);$notice=[string](Probe 'RunLocalAct' @($action,''))
            Check ('RunLocal.'+$action+'.ExistingOwnerBehavior') ([bool](Probe 'RunLocalPreserved' @($action,'Normal')))
            Check ('RunLocal.'+$action+'.NoUnhandledError') (-not $notice.StartsWith('HANDLER_ERROR|'))
            Pair $before $action (Positive $action) ('RunLocal.'+$action+'.Normal')
            foreach($guard in @('Loading','Busy')){
                Stage $action;$state=State;$before=@(Files);[void](Probe 'RunLocalAct' @($action,$guard))
                Check ('RunLocal.'+$action+'.'+$guard+'.NoMutationOrActivity') ((State) -ceq $state -and @(Files).Count -eq $before.Count -and [int](Probe 'ReadCalls') -eq 0)
            }
        }
        $cases=@(
            @{Action='LOAD';Mode='NoSelection';Outcome='REJECTED'},
            @{Action='LOAD';Mode='BadScale';Outcome='REJECTED'},
            @{Action='LOAD';Mode='UnavailableRecipe';Outcome='FAILED'},
            @{Action='SCALE';Mode='Minimum';Outcome='STAGED'},
            @{Action='SCALE';Mode='Maximum';Outcome='STAGED'},
            @{Action='SCALE';Mode='BelowMinimum';Outcome='REJECTED'},
            @{Action='SCALE';Mode='AboveMaximum';Outcome='REJECTED'},
            @{Action='SCALE';Mode='BadScale';Outcome='REJECTED'},
            @{Action='SCALE';Mode='SyntheticCompleted';Outcome='REJECTED'},
            @{Action='ALLOCATE';Mode='NoSelection';Outcome='STAGED'},
            @{Action='ALLOCATE';Mode='RefillPalette';Outcome='REJECTED'},
            @{Action='TREE_ALLOCATE';Mode='NoSelection';Outcome='REJECTED'}
        )
        foreach($action in @('ALLOCATE','TREE_ALLOCATE')){
            foreach($mode in @('NoQuantity','BadQuantity','NegativeQuantity','OverRequirement','WrongLocation')){$cases+=@{Action=$action;Mode=$mode;Outcome='REJECTED'}}
            foreach($mode in @('PercentOnly','ZeroQuantity')){$cases+=@{Action=$action;Mode=$mode;Outcome='STAGED'}}
        }
        if($RunLocalClosedDiagnostic){$cases=@()}
        if($RunLocalRefillDiagnostic){$cases=@(@{Action='ALLOCATE';Mode='RefillPalette';Outcome='REJECTED'})}
        foreach($case in $cases){
            Stage $case.Action $case.Mode;$state=State;$before=@(Files)
            if($case.Mode -ceq 'RefillPalette'){
                $refillBefore=([string](Probe 'RunRefillState')).Split('|')
                $refillOwner=[string](Probe 'RunLocalOwnerState');[void](Probe 'RunRefillReset')
            }
            $notice=[string](Probe 'RunLocalAct' @($case.Action,''));$label='RunLocal.'+$case.Action+'.'+$case.Mode
            if($case.Mode -ceq 'RefillPalette'){
                $refillAfter=([string](Probe 'RunRefillState')).Split('|')
                $refillRecords=@(Files|Where-Object{$_ -cnotin $before}|ForEach-Object{[IO.File]::ReadAllText($_)|ConvertFrom-Json})
                $refillReceipt=[pscustomobject]@{ReusableBefore=($refillBefore[0] -ceq 'True');ReusableAfter=($refillAfter[0] -ceq 'True');PaletteBefore=[int]$refillBefore[1];PaletteAfter=[int]$refillAfter[1];OwnerCalls=[int](Probe 'RunRefillCalls');OwnerStatePreserved=([string](Probe 'RunLocalOwnerState') -ceq $refillOwner);NoChoicesRefusal=($refillAfter[2] -ceq 'True');Records=$refillRecords.Count;Requested=@($refillRecords|Where-Object OutcomeCode -ceq 'REQUESTED').Count;Rejected=@($refillRecords|Where-Object OutcomeCode -ceq 'REJECTED').Count;Staged=@($refillRecords|Where-Object OutcomeCode -ceq 'STAGED').Count}
                $refillReceipt|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'run-refill-boundary.json')
                if($RunLocalRefillDiagnostic){
                    Check 'RunRefill.LoadedRunEmptyPalette' ($refillReceipt.ReusableBefore -and $refillReceipt.ReusableAfter -and $refillReceipt.PaletteBefore -eq 0 -and $refillReceipt.PaletteAfter -eq 0)
                    Check 'RunRefill.AllocationOwnerNotInvoked' ($refillReceipt.OwnerCalls -eq 0)
                    Check 'RunRefill.OwnerStatePreserved' $refillReceipt.OwnerStatePreserved
                    Check 'RunRefill.ExistingNoChoicesRefusal' $refillReceipt.NoChoicesRefusal
                    Check 'RunRefill.NoStagedObservation' ($refillReceipt.Staged -eq 0)
                }
            }
            # Refill can clear local presentation while preserving the owning allocation.
            # Retain its original owner check; REJECTED does not imply no local effects.
            $preserved=if($case.Outcome -ceq 'REJECTED' -and $case.Mode -cne 'RefillPalette'){(State) -ceq $state}else{[bool](Probe 'RunLocalPreserved' @($case.Action,$case.Mode))}
            Check ($label+'.ExistingOwnerBehavior') $preserved
            Check ($label+'.NoUnhandledError') (-not $notice.StartsWith('HANDLER_ERROR|'))
            Pair $before $case.Action $case.Outcome $label
        }
        Check 'RunLocal.UnknownValuesAndFormula' ($sheet.Cells.Item(1,1).Value2 -ceq 'Operator Extra' -and $sheet.Cells.Item(2,1).Value2 -ceq $canary -and $sheet.Cells.Item(2,2).Formula -ceq '=1+2')
        $guards=if($RunLocalClosedDiagnostic){@('ClosedWorkbook')}else{@('Target','Session','SignedOut','ClosedWorkbook')}
        foreach($guard in $guards){
            foreach($action in $actions){
                SelectTarget $Fixture 'config-producer'
                if($null -eq $book){$book=$excel.Workbooks.Open($path,0,$false)}
                [void](Probe 'RunLocalReopen' @($book.Name));$decoy.Activate();Stage $action;$state=State
                if($guard -ceq 'ClosedWorkbook'){
                    Initialize-SettingsCapture
                    $page=[int](Probe 'RunLocalShowAndCapture' @($book.Name,$action));$state=State
                    $capturedName=$book.Name;$decoyName=$decoy.Name
                    $ownerBefore=[string](Probe 'RunLocalOwnerState');$before=@(Files);$otherBefore=@(Get-Slice4beActivityFiles $Other)
                    $boundary=[ordered]@{Action=$action;PageBefore=$page;BeforeUTC=[DateTimeOffset]::UtcNow.ToString('o');VisibleBefore=([InvSysSettingsCapture]::OwnedVisibleForm('Production',[IntPtr]$excel.Hwnd) -ne [IntPtr]::Zero);WorkbooksBefore=$excel.Workbooks.Count}
                    $decoy.Activate();$book.Close($false);$book=$null
                    $names=@(foreach($openBook in $excel.Workbooks){[string]$openBook.Name})
                    $boundary.AfterUTC=[DateTimeOffset]::UtcNow.ToString('o')
                    $boundary.VisibleAfter=([InvSysSettingsCapture]::OwnedVisibleForm('Production',[IntPtr]$excel.Hwnd) -ne [IntPtr]::Zero)
                    $boundary.WorkbooksAfter=$excel.Workbooks.Count
                    $boundary.CapturedStillOpen=($capturedName -cin $names);$boundary.DecoyStillOpen=($decoyName -cin $names)
                    if(-not $boundary.VisibleBefore -or $boundary.CapturedStillOpen -or -not $boundary.DecoyStillOpen){throw 'Closed Run operator fixture unavailable; not product RED.'}
                    $boundary.HandlerInvoked=[bool]$boundary.VisibleAfter;$protected=-not $boundary.VisibleAfter
                    if($boundary.VisibleAfter){
                        $notice=[string](Probe 'RunLocalClosedAct' @($action));$entry=([string](Probe 'RunLocalClosedStatus')).Split('|')
                        $boundary.HandlerEntered=($entry[0] -ceq 'True');$boundary.AdapterError=[long]$entry[1]
                        if($entry[0] -cne 'True'){throw 'Visible closed Run fixture did not enter its handler; not product RED.'}
                        $protected=[long]$entry[1] -eq 0 -and $notice.StartsWith('Session, warehouse, or captured workbook changed.') -and (State) -ceq $state
                    }
                    $label='RunLocal.ClosedBoundary.'+$action
                    Check ($label+'.SurfaceLifetimeEstablished') ($boundary.WorkbooksBefore -eq $boundary.WorkbooksAfter+1 -and $page -eq $(if($action -ceq 'TREE_ALLOCATE'){4}else{3}))
                    Check ($label+'.UserActionProtected') $protected
                    Check ($label+'.BindingGuardRejectsClosedBook') (-not [bool](Probe 'RunLocalClosedBindingCurrent'))
                    Check ($label+'.OwnerStatePreserved') ([string](Probe 'RunLocalOwnerState') -ceq $ownerBefore)
                    Check ($label+'.SavedBytesPreserved') ((Hash $path) -ceq $bookPin)
                    Check ($label+'.NoRedirectedActivity') (@(Files).Count -eq $before.Count -and @(Get-Slice4beActivityFiles $Other).Count -eq $otherBefore.Count)
                    $boundary|ConvertTo-Json|Set-Content (Join-Path $reportRoot ('run-closed-boundary-'+$action+'.json'))
                    [void](Probe 'RunLocalSafeClose')
                    continue
                }
                if($guard -ceq 'Target'){SelectTarget $Other 'config-producer'}
                if($guard -ceq 'Session'){SelectTarget $Fixture 'config-producer'}
                if($guard -ceq 'SignedOut'){[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
                $before=@(Files);$otherBefore=@(Get-Slice4beActivityFiles $Other)
                $notice=[string](Probe 'RunLocalAct' @($action,''));$label='RunLocal.Guard.'+$guard+'.'+$action
                Check ($label+'.NoMutation') ((State) -ceq $state)
                Check ($label+'.VisibleRefusal') ($notice.StartsWith('Session, warehouse, or captured workbook changed.') -and $notice.Contains('Reopen Production'))
                Check ($label+'.NoRedirectedActivity') (@(Files).Count -eq $before.Count -and @(Get-Slice4beActivityFiles $Other).Count -eq $otherBefore.Count)
            }
        }
        $same=$true;foreach($file in $pins.Keys){$same=$same -and (Hash $file) -ceq $pins[$file]};Check 'RunLocal.SavedAuthorityPreserved' $same
        [void](Probe 'CloseDesigner');Check 'RunLocal.OperatorBytesPreserved' ((Hash $path) -ceq $bookPin)
        $same=$true;foreach($file in $recordPins.Keys){$same=$same -and (Hash $file) -ceq $recordPins[$file]};Check 'RunLocal.OlderRecordsImmutable' $same
    }finally{[void](Probe 'CloseDesigner');if($null -ne $book){$book.Close($false)};if($null -ne $decoy){$decoy.Close($false)};SelectTarget $Fixture}
}
