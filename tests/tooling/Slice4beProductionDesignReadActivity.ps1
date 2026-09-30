# Actual-handler gate; paired original-recording/publication/view evidence is separate.
function Test-ProductionDesignReadActivity($Fixture,$Other) {
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    function Hash([string]$Path){$s=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $s).Hash}finally{$s.Dispose()}}
    function Stage([string]$Mode='Normal'){[void](Probe 'ReadStage' @($Mode))}
    function State {[string](Probe 'ReadState')}
    function Positive([string]$Action){if($Action.EndsWith('REFRESH')){'REFRESHED'}elseif($Action -ceq 'PROCESS_REUSE'){'STAGED'}else{'PRESENTED'}}
    function Pair([string[]]$Before,[string]$Action,[string]$Outcome,[string]$Case,[string]$Actor='config-producer'){
        $raw=@(Files|Where-Object{$_ -cnotin $Before}|ForEach-Object{[IO.File]::ReadAllText($_)})
        $records=@($raw|ForEach-Object{$_|ConvertFrom-Json})
        $first=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($records|Where-Object OutcomeCode -CEQ $Outcome)
        $paired=$records.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
        $context=$paired;$redacted=$paired;$integrity=$paired;$linked=$false;$facts=$false;$terminal=$false
        $control='PRODUCTION_'+$Action
        foreach($r in $records){
            $context=$context -and $r.ControlId -ceq $control -and $r.OwnerId -ceq 'PRODUCTION_DESIGNER' -and $r.UserId -ceq $Actor -and $r.WarehouseId -ceq $Fixture.Warehouse -and $r.StationId -ceq 'S1' -and $r.CatalogVersion -eq 19
            $redacted=$redacted -and @($r.SourceEventRefs).Count -eq 0
        }
        foreach($value in $raw){
            foreach($secret in @($canary,$Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'mBtn','PinHash')){
                $encoded=ConvertTo-Json -InputObject $secret -Compress
                if($value.Contains($secret) -or $value.Contains($encoded.Substring(1,$encoded.Length-2))){$redacted=$false}
            }
            $match=[regex]::Match($value,',"ContentSha256":"([a-f0-9]{64})"\}$')
            if(-not $match.Success){$integrity=$false;continue}
            $sha=[Security.Cryptography.SHA256]::Create()
            try{$hash=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($value.Substring(0,$match.Index)+'}'))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
            $integrity=$integrity -and $hash -ceq $match.Groups[1].Value
        }
        if($paired){
            $linked=$first[0].ActivityId -cne '' -and $first[0].ActivityId -ceq $last[0].ActivityId -and $first[0].RecordId -cne $last[0].RecordId
            $severity=switch($Outcome){'REJECTED'{'Warning'} 'DENIED'{'Blocked'} 'FAILED'{'Error'} default{'Info'}}
            $effect=if($Outcome -ceq 'FAILED'){'Unknown'}else{'Unchanged'}
            $facts=$first[0].DataEffect -ceq 'Unknown' -and $first[0].Severity -ceq 'Info' -and $last[0].Severity -ceq $severity -and $last[0].DataEffect -ceq $effect -and $last[0].EventCode -ceq ($control+'_'+$Outcome)
            $attempt=[bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @(($first[0]|ConvertTo-Json -Depth 20 -Compress)))
            $completed=[bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @(($last[0]|ConvertTo-Json -Depth 20 -Compress)))
            $terminal=-not $attempt -and $completed -eq ($Outcome -ceq (Positive $Action))
        }
        Check ($Case+'.AttemptAndOutcome') $paired
        Check ($Case+'.ExactContext') $context
        Check ($Case+'.NoEnteredDataOrSources') $redacted
        Check ($Case+'.ContentIntegrity') $integrity
        Check ($Case+'.LinkedDistinctRecords') $linked
        Check ($Case+'.OwnerFact') $facts
        Check ($Case+'.ExactLocalCommandTerminal') $terminal
    }
    $canary='READ'+[guid]::NewGuid().ToString('N');$book=$null;$decoy=$null;$pins=@{};$recordPins=@{}
    $actions=@('PROCESS_REFRESH','PROCESS_LOAD','PROCESS_REUSE','RECIPE_REFRESH','RECIPE_LOAD')
    SelectTarget $Fixture
    if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadPolicyForTest' @($true))){throw 'Authorized read policy fixture unavailable; not product RED.'}
    try {
        SelectTarget $Fixture 'config-producer'
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$canary
        $path=Join-Path $runRoot 'designer-read-operator.xlsb';$book.SaveAs($path,50);$book.Close($false)
        $bookPin=Hash $path;$book=$excel.Workbooks.Open($path,0,$false);$sheet=$book.Worksheets.Item(1)
        $decoy=$excel.Workbooks.Add();$decoy.Activate()
        [void](Probe 'OpenDesigner' @($book.Name))
        $ready=[bool](Probe 'ReadPrepare' @($canary));Check 'DesignRead.RealReleasedFixtureReady' $ready
        if(-not $ready){
            $stage=[string](Probe 'ReadSetupStage')
            if($stage -cnotin @('ReleasedProcess','RecipeDraft','RecipeSave','RecipeRelease','ReleasedStatus','Ready')){$stage='Unavailable'}
            throw ('Released fixture setup failed at '+$stage+'; not product RED.')
        }
        foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File|Where-Object{$_.Extension -in '.xlsb','.xlsm' -and $_.Name -notlike '~$*'}){$pins[$file.FullName]=Hash $file.FullName}
        foreach($file in Files){$recordPins[$file]=Hash $file}
        $old=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(18))).Split([char]10)|Where-Object{$_})
        $new=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(19))).Split([char]10)|Where-Object{$_})
        Check 'DesignRead.Catalog19Extends18' ($old.Count -eq 98 -and $new.Count -eq 103 -and @($new|Sort-Object -Unique).Count -eq 103 -and @($old|Where-Object{$_ -cnotin $new}).Count -eq 0)
        foreach($id in $old){$before=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,18));Check ('DesignRead.Catalog.Preserve.'+$id) ($before -ne '' -and $before -ceq [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,19)))}
        foreach($action in $actions){
            $label='DesignRead.'+$action;$positive=Positive $action
            $control='PRODUCTION_'+$action
            $json=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($control,19));$definition=if($json){$json|ConvertFrom-Json}else{$null}
            $caption=switch($action){'PROCESS_LOAD'{'View Process'} 'PROCESS_REUSE'{'Edit as New Version'} 'RECIPE_LOAD'{'Load'} default{'Refresh'}}
            $surface='Operations > Production > '+$(if($action.StartsWith('PROCESS')){'Process Designer'}else{'Recipe Designer'})
            Check ($label+'.FixedMetadata') ($null -ne $definition -and $definition.Caption -ceq $caption -and $definition.Surface -ceq $surface -and $definition.Role -ceq 'Production' -and $definition.Class -ceq 'Command' -and $definition.Capability -ceq 'PROD_POST')
            $absent=$true;foreach($version in 1..18){$absent=$absent -and [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($control,$version)) -ceq ''}
            Check ($label+'.AbsentFromOlderCatalogs') $absent
            foreach($unsupported in @('CONFIRMED','APPLIED','VALIDATED','COMPLETED','REFRESHED','PRESENTED','STAGED')|Where-Object{$_ -cne $positive}){Check ($label+'.RejectsUnsupported.'+$unsupported) ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Outcome' @($control,$unsupported)) -ceq '')}
            foreach($mode in @('Normal',$(if($action.EndsWith('REFRESH')){'EmptyLists'}else{'EmptyArray'}))){
                $before=@(Files);Stage $mode
                Check ($label+'.'+$mode+'.SetupNotUserAction') (@(Files).Count -eq $before.Count)
                $notice=[string](Probe 'ReadAct' @($action))
                Check ($label+'.'+$mode+'.ExistingLocalBehavior') ([bool](Probe 'ReadPreserved' @($action,$mode)))
                Check ($label+'.'+$mode+'.NoUnhandledError') (-not $notice.StartsWith('HANDLER_ERROR|'))
                Pair $before $action $positive ($label+'.'+$mode)
            }
            foreach($guard in @('Loading','Busy')){
                Stage;$state=State;$before=@(Files);[void](Probe 'ReadGuard' @($action,$guard))
                Check ($label+'.'+$guard+'.NoReadsOrMutationOrActivity') ((State) -ceq $state -and [int](Probe 'ReadCalls') -eq 0 -and @(Files).Count -eq $before.Count)
            }
            Stage 'PartialFailure';$state=State;$before=@(Files);$notice=[string](Probe 'ReadAct' @($action))
            Check ($label+'.PartialFailure.LocalEffectsRemain') ((State) -cne $state)
            Check ($label+'.PartialFailure.NoUnhandledError') (-not $notice.StartsWith('HANDLER_ERROR|'))
            Check ($label+'.PartialFailure.GuardsRestored') ([bool](Probe 'ReadGuardsRestored'))
            Pair $before $action 'FAILED' ($label+'.PartialFailure')
            Stage;$before=@(Files);[void](Probe 'ReadGuard' @($action,'Nested'))
            Check ($label+'.Nested.Entered') ([bool](Probe 'ReadNestedEntered'))
            Check ($label+'.Nested.NoSecondRead') ([int](Probe 'ReadCalls') -eq $(if($action.EndsWith('REFRESH')){4}else{1}))
            Check ($label+'.Nested.LocalResultPreserved') ([bool](Probe 'ReadPreserved' @($action,'Normal')))
            Pair $before $action $positive ($label+'.Nested')
            foreach($mode in $(if($action.EndsWith('REFRESH')){@('ReadError')}else{@('NoSelection','Malformed','Unavailable','ReadError')})){
                Stage $mode;$state=State;$before=@(Files);$notice=[string](Probe 'ReadAct' @($action))
                Check ($label+'.'+$mode+'.PriorDraftPreserved') ((State) -ceq $state)
                Check ($label+'.'+$mode+'.NoUnhandledError') (-not $notice.StartsWith('HANDLER_ERROR|'))
                $outcome=if($mode -ceq 'NoSelection'){'REJECTED'}else{'FAILED'}
                Pair $before $action $outcome ($label+'.'+$mode)
            }
        }
        SelectTarget $Fixture 'config-reader';[void](Probe 'ReadReopen' @($book.Name));$decoy.Activate()
        foreach($action in $actions){
            Stage;$state=State;$before=@(Files);[void](Probe 'ReadAct' @($action))
            Check ('DesignRead.Denied.'+$action+'.NoReadsOrLocalMutation') ((State) -ceq $state -and [int](Probe 'ReadCalls') -eq 0)
            Pair $before $action 'DENIED' ('DesignRead.Denied.'+$action) 'config-reader'
        }
        foreach($enabled in @($false,$true)){
            SelectTarget $Fixture
            if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadPolicyForTest' @($enabled))){throw 'Authorized read policy update unavailable; not product RED.'}
            $pins[$Fixture.Config]=Hash $Fixture.Config
            if(-not $enabled){
                SelectTarget $Fixture 'config-producer';[void](Probe 'ReadReopen' @($book.Name));$decoy.Activate()
                foreach($action in $actions){Stage;$before=@(Files);[void](Probe 'ReadAct' @($action));Check ('DesignRead.DisabledTracking.'+$action) ([bool](Probe 'ReadPreserved' @($action,'Normal')) -and @(Files).Count -eq $before.Count)}
            }
        }
        $parent=Join-Path $Fixture.Root 'Training\Activity';$blocked=Join-Path $parent $Fixture.Warehouse;$held=$blocked+'-design-read-held'
        foreach($item in @($blocked,$held)){if(-not [IO.Path]::GetFullPath($item).StartsWith([IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Tracking fault escaped disposable root.'}}
        if(Test-Path -LiteralPath $held){throw 'Preserve existing held fixture.'}
        SelectTarget $Fixture 'config-producer';[void](Probe 'ReadReopen' @($book.Name));$decoy.Activate();$moved=$false
        try{
            if(Test-Path -LiteralPath $blocked){Move-Item -LiteralPath $blocked -Destination $held;$moved=$true}
            [IO.File]::WriteAllText($blocked,'Blocked disposable design-read activity path')
            foreach($action in $actions){Stage;$before=@(Files);$notice=[string](Probe 'ReadAct' @($action));Check ('DesignRead.UnavailableTracking.'+$action+'.LocalActionContinues') ([bool](Probe 'ReadPreserved' @($action,'Normal')));Check ('DesignRead.UnavailableTracking.'+$action+'.VisibleNoFallback') ($notice.Contains('Tracking unavailable') -and @(Files).Count -eq $before.Count)}
        }finally{if(Test-Path -LiteralPath $blocked -PathType Leaf){Remove-Item -LiteralPath $blocked};if($moved){Move-Item -LiteralPath $held -Destination $blocked}}
        Check 'DesignRead.UnknownColumnsInMemory' ($sheet.Cells.Item(1,1).Value2 -ceq 'Operator Extra' -and $sheet.Cells.Item(2,1).Value2 -ceq $canary)
        foreach($guard in @('Target','Session','SignedOut','ClosedWorkbook')){
            SelectTarget $Fixture 'config-producer';[void](Probe 'ReadReopen' @($book.Name));$decoy.Activate()
            if($guard -ceq 'Target'){SelectTarget $Other 'config-producer'}
            if($guard -ceq 'Session'){SelectTarget $Fixture 'config-producer'}
            if($guard -ceq 'SignedOut'){[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
            if($guard -ceq 'ClosedWorkbook'){$book.Close($false);$book=$null}
            $before=@(Files);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            foreach($action in $actions){
                Stage;$state=State;$notice=[string](Probe 'ReadAct' @($action))
                Check ('DesignRead.Guard.'+$guard+'.'+$action+'.NoReadsOrLocalMutation') ((State) -ceq $state -and [int](Probe 'ReadCalls') -eq 0)
                Check ('DesignRead.Guard.'+$guard+'.'+$action+'.RefusalVisible') ($notice.StartsWith('Session, warehouse, or captured workbook changed.') -and $notice.Contains('Reopen Production'))
            }
            Check ('DesignRead.Guard.'+$guard+'.NoRedirectedActivity') (@(Files).Count -eq $before.Count -and @(Get-Slice4beActivityFiles $Other).Count -eq $otherBefore.Count)
        }
        [void](Probe 'CloseDesigner')
        Check 'DesignRead.OperatorBytesPreserved' ((Hash $path) -ceq $bookPin)
        $same=$true;foreach($file in $pins.Keys){$same=$same -and (Hash $file) -ceq $pins[$file]};Check 'DesignRead.SavedAuthorityPreserved' $same
        $same=$true;foreach($file in $recordPins.Keys){$same=$same -and (Hash $file) -ceq $recordPins[$file]};Check 'DesignRead.OlderRecordsImmutable' $same
    } finally {[void](Probe 'CloseDesigner');if($null -ne $book){$book.Close($false)};if($null -ne $decoy){$decoy.Close($false)};SelectTarget $Fixture}
}
