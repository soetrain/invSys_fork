# Synthetic draft values stay in memory; reports contain check names/Booleans only.
function Test-ProductionInstructionActivity($Fixture,$Other) {
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    function Rows([string[]]$Values){$lines=@(for($i=0;$i -lt $Values.Count;$i++){[string]($i+1)+"`t"+$Values[$i]});$lines -join "`n"}
    function State {[string](Probe 'InstructionRows')}
    function Stage([int]$Selected=1,[string]$Value=(' '+$canary+'4 ')){[void](Probe 'InstructionStage' @($canary,$Selected,$Value))}
    function Pair([string[]]$Before,[string]$Action,[string]$Outcome,[string]$Case,[string]$Actor='config-producer'){
        $raw=@(Files|Where-Object {$_ -cnotin $Before}|ForEach-Object {[IO.File]::ReadAllText($_)})
        $records=@($raw|ForEach-Object {$_|ConvertFrom-Json})
        $first=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($records|Where-Object OutcomeCode -CEQ $Outcome)
        $paired=$records.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
        Check ($Case+'.AttemptAndOutcome') $paired
        $id='PRODUCTION_PROCESS_INSTRUCTION_'+$Action
        $context=$paired;$redacted=$paired;$integrity=$paired;$linked=$false;$effect=$false
        foreach($r in $records){
            $context=$context -and $r.ControlId -ceq $id -and $r.OwnerId -ceq 'PRODUCTION_DESIGNER' -and $r.UserId -ceq $Actor -and $r.WarehouseId -ceq $Fixture.Warehouse -and $r.StationId -ceq 'S1' -and $r.CatalogVersion -eq 18
            $redacted=$redacted -and @($r.SourceEventRefs).Count -eq 0
        }
        foreach($value in $raw){
            foreach($secret in @($canary,$Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'mBtn','PinHash')){
                $encoded=ConvertTo-Json -InputObject $secret -Compress
                if($value.Contains($secret) -or $value.Contains($encoded.Substring(1,$encoded.Length-2))){$redacted=$false}
            }
            $m=[regex]::Match($value,',"ContentSha256":"([a-f0-9]{64})"\}$')
            if(-not $m.Success){$integrity=$false;continue}
            $sha=[Security.Cryptography.SHA256]::Create()
            try{$hash=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($value.Substring(0,$m.Index)+'}'))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
            $integrity=$integrity -and $hash -ceq $m.Groups[1].Value
        }
        if($paired){
            $linked=$first[0].ActivityId -cne '' -and $first[0].ActivityId -ceq $last[0].ActivityId -and $first[0].RecordId -cne $last[0].RecordId
            $severity=switch($Outcome){'REJECTED'{'Warning'} 'DENIED'{'Blocked'} 'FAILED'{'Error'} default{'Info'}}
            $dataEffect=if($Outcome -ceq 'FAILED'){'Unknown'}else{'Unchanged'}
            $effect=$first[0].DataEffect -ceq 'Unknown' -and $first[0].Severity -ceq 'Info' -and $last[0].Severity -ceq $severity -and $last[0].DataEffect -ceq $dataEffect -and $last[0].EventCode -ceq ($id+'_'+$Outcome)
        }
        Check ($Case+'.ExactContext') $context
        Check ($Case+'.NoEnteredDataOrSources') $redacted
        Check ($Case+'.ContentIntegrity') $integrity
        Check ($Case+'.LinkedDistinctRecords') $linked
        Check ($Case+'.OwnerFact') $effect
        $requested=$false;$terminal=$false
        if($paired){
            $requested=-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.InstructionTerminalForTest' @(($first[0]|ConvertTo-Json -Depth 20 -Compress)))
            $completed=[bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.InstructionTerminalForTest' @(($last[0]|ConvertTo-Json -Depth 20 -Compress)))
            $terminal=$completed -eq ($Outcome -ceq 'STAGED')
        }
        Check ($Case+'.RequestedIsNotCompletion') $requested
        Check ($Case+'.ExactCommandTerminal') $terminal
    }
    $actions=@('ADD','UPDATE','REMOVE','UP','DOWN');$canary='INSTRUCTION'+[guid]::NewGuid().ToString('N')
    $book=$null;$decoy=$null;$pins=@{}
    SelectTarget $Other
    if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.InstructionPolicyForTest' @($true))){throw 'Other target fixture tracking policy unavailable; not product RED.'}
    SelectTarget $Fixture
    if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.InstructionPolicyForTest' @($true))){throw 'Authorized fixture tracking policy unavailable; not product RED.'}
    foreach($target in @($Fixture,$Other)){foreach($file in Get-ChildItem -LiteralPath $target.Root -Recurse -File|Where-Object {$_.Extension -in '.xlsb','.xlsm' -and -not $_.Name.StartsWith('~$')}){$pins[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash}}
    $old=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(13))).Split("`n")|Where-Object {$_})
    $new=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(14))).Split("`n")|Where-Object {$_})
    Check 'Instructions.Catalog14Extends13' ($old.Count -eq 74 -and $new.Count -eq 79 -and @($old|Where-Object {$_ -cnotin $new}).Count -eq 0 -and @($new|Sort-Object -Unique).Count -eq 79)
    foreach($id in $old){
        $before=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,13))
        $after=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,14))
        Check ('Instructions.Catalog.Preserve.'+$id) ($before -ne '' -and $before -ceq $after)
    }
    try {
        SelectTarget $Fixture 'config-producer'
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$canary
        $path=Join-Path $runRoot 'instruction-operator.xlsb';$book.SaveAs($path,50);$book.Close($false)
        $workbookPin=(Get-FileHash -LiteralPath $path).Hash;$book=$excel.Workbooks.Open($path,0,$false)
        $sheet=$book.Worksheets.Item(1);$decoy=$excel.Workbooks.Add()
        $before=@(Files);[void](Probe 'OpenDesigner' @($book.Name));Stage
        Check 'Instructions.InitializationIsNotUserAction' (@(Files).Count -eq $before.Count)
        foreach($action in $actions){
            $id='PRODUCTION_PROCESS_INSTRUCTION_'+$action
            $json=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,14));$definition=if($json){$json|ConvertFrom-Json}else{$null}
            Check ('Instructions.'+$action+'.FixedMetadata') ($null -ne $definition -and $definition.Caption -ceq [string](Probe 'InstructionCaption' @($action)) -and $definition.Surface -ceq 'Operations > Production > Process Designer > Instructions' -and $definition.Role -ceq 'Production' -and $definition.Class -ceq 'Command' -and $definition.Capability -ceq 'PROD_POST')
            foreach($unsupported in @('COMPLETED','CONFIRMED','VALIDATED','APPLIED')){
                Check ('Instructions.'+$action+'.RejectsUnsupported.'+$unsupported) ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Outcome' @($id,$unsupported)) -ceq '')
            }
            Stage;$before=@(Files);$decoy.Activate();$notice=[string](Probe 'InstructionAct' @($action))
            $expected=switch($action){
                'ADD'{Rows @(($canary+'1'),($canary+'2'),($canary+'3'),($canary+'4'))}
                'UPDATE'{Rows @(($canary+'1'),($canary+'4'),($canary+'3'))}
                'REMOVE'{Rows @(($canary+'1'),($canary+'3'))}
                'UP'{Rows @(($canary+'2'),($canary+'1'),($canary+'3'))}
                'DOWN'{Rows @(($canary+'1'),($canary+'3'),($canary+'2'))}
            }
            Check ('Instructions.'+$action+'.ActualLocalResult') ((State) -ceq $expected)
            if($action -in @('ADD','UPDATE')){
                Check ('Instructions.'+$action+'.EditorNotCleared') ([string](Probe 'InstructionEditor') -ceq (' '+$canary+'4 '))
                Check ('Instructions.'+$action+'.SelectionPreserved') ([long](Probe 'InstructionSelection') -eq 1)
            }
            if($action -in @('UP','DOWN')){Check ('Instructions.'+$action+'.SelectionFollowsMove') ([long](Probe 'InstructionSelection') -eq $(if($action -ceq 'UP'){0}else{2}))}
            Pair $before $action 'STAGED' ('Instructions.'+$action+'.Success')
            Stage;$before=@(Files);$prior=State;$notice=[string](Probe 'InstructionBusy' @($action))
            Check ('Instructions.'+$action+'.BusyGuard') ((State) -ceq $prior -and @(Files).Count -eq $before.Count)
            Stage;$before=@(Files);$prior=State;$notice=[string](Probe 'InstructionLoading' @($action))
            Check ('Instructions.'+$action+'.LoadingGuard') ((State) -ceq $prior -and @(Files).Count -eq $before.Count)
            if($action -in @('UP','DOWN')){
                Stage;$before=@(Files)
                Check ('Instructions.'+$action+'.NestedHandlerEntered') ([bool](Probe 'InstructionNested' @($action)))
                Check ('Instructions.'+$action+'.NestedEditSuppressed') ((State) -ceq $expected)
                Pair $before $action 'STAGED' ('Instructions.'+$action+'.Nested')
            }
            Stage;$before=@(Files);$notice=[string](Probe 'InstructionFailure' @($action))
            Pair $before $action 'FAILED' ('Instructions.'+$action+'.Failure')
        }
        foreach($action in $actions){
            $selected=switch($action){'UP'{0} 'DOWN'{2} default{-1}}
            Stage $selected '  ';$prior=State;$before=@(Files);$notice=[string](Probe 'InstructionAct' @($action))
            Check ('Instructions.'+$action+'.RejectedPreservesList') ((State) -ceq $prior)
            Pair $before $action 'REJECTED' ('Instructions.'+$action+'.Rejected')
        }
        Stage 1 '   ';$before=@(Files);$notice=[string](Probe 'InstructionAct' @('UPDATE'))
        Check 'Instructions.EmptyUpdateRetainsExistingSemantics' ((State) -ceq (Rows @(($canary+'1'),'',($canary+'3'))))
        Pair $before 'UPDATE' 'STAGED' 'Instructions.EmptyUpdate'
        $parent=Join-Path $Fixture.Root 'Training\Activity'
        $blocked=Join-Path $parent $Fixture.Warehouse;$held=$blocked+'-instruction-fixture-held'
        foreach($item in @($blocked,$held)){
            if(-not [IO.Path]::GetFullPath($item).StartsWith([IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Instruction tracking fault escaped the disposable fixture.'}
        }
        if(Test-Path -LiteralPath $held){throw 'Preserve existing held fixture.'}
        $moved=$false
        try {
            if(Test-Path -LiteralPath $blocked){Move-Item -LiteralPath $blocked -Destination $held;$moved=$true}
            [IO.File]::WriteAllText($blocked,'Blocked disposable instruction activity path')
            foreach($action in $actions){
                Stage;$prior=State;$before=@(Files);$notice=[string](Probe 'InstructionAct' @($action))
                Check ('Instructions.StoreUnavailable.'+$action+'.EditContinues') ((State) -cne $prior)
                Check ('Instructions.StoreUnavailable.'+$action+'.VisibleNotice') ($notice.Contains('Tracking unavailable'))
                Check ('Instructions.StoreUnavailable.'+$action+'.NoActivityFallback') (@(Files).Count -eq $before.Count)
            }
        }finally{
            if(Test-Path -LiteralPath $blocked -PathType Leaf){Remove-Item -LiteralPath $blocked}
            if($moved){Move-Item -LiteralPath $held -Destination $blocked}
        }
        SelectTarget $Fixture 'config-reader';[void](Probe 'OpenDesigner' @($book.Name))
        foreach($action in $actions){
            Stage;$prior=State;$before=@(Files);$notice=[string](Probe 'InstructionAct' @($action))
            Check ('Instructions.Denied.'+$action+'.ListUnchanged') ((State) -ceq $prior)
            Pair $before $action 'DENIED' ('Instructions.Denied.'+$action) 'config-reader'
        }
        Check 'Instructions.UnknownColumnInMemory' ($sheet.Cells.Item(1,1).Value2 -ceq 'Operator Extra' -and $sheet.Cells.Item(2,1).Value2 -ceq $canary)
        $same=$true;foreach($file in $pins.Keys){$same=$same -and (Get-FileHash -LiteralPath $file).Hash -ceq $pins[$file]}
        Check 'Instructions.SavedAuthorityBeforePolicyChange' $same
        SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.InstructionPolicyForTest' @($false))){throw 'Authorized disabled-policy fixture unavailable; not product RED.'}
        $pins[$Fixture.Config]=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        SelectTarget $Fixture 'config-producer';[void](Probe 'OpenDesigner' @($book.Name))
        foreach($action in $actions){
            Stage;$prior=State;$before=@(Files);$notice=[string](Probe 'InstructionAct' @($action))
            Check ('Instructions.DisabledTracking.'+$action) ((State) -cne $prior -and @(Files).Count -eq $before.Count)
        }
        SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.InstructionPolicyForTest' @($true))){throw 'Tracking must be enabled for stale-context evidence.'}
        $pins[$Fixture.Config]=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        foreach($guard in @('Target','Session','SignedOut','ClosedWorkbook')){
            SelectTarget $Fixture 'config-producer';[void](Probe 'OpenDesigner' @($book.Name))
            if($guard -ceq 'Target'){SelectTarget $Other 'config-producer'}
            if($guard -ceq 'Session'){SelectTarget $Fixture 'config-producer'}
            if($guard -ceq 'SignedOut'){[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
            if($guard -ceq 'ClosedWorkbook'){$book.Close($false);$book=$null}
            $before=@(Files);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            foreach($action in $actions){
                Stage;$prior=State;$notice=[string](Probe 'InstructionAct' @($action))
                Check ('Instructions.Guard.'+$guard+'.'+$action) ((State) -ceq $prior)
            }
            Check ('Instructions.Guard.'+$guard+'.NoRedirectedActivity') (@(Files).Count -eq $before.Count -and @(Get-Slice4beActivityFiles $Other).Count -eq $otherBefore.Count)
        }
        Check 'Instructions.UnknownColumnAndSavedBytesPreserved' ((Get-FileHash -LiteralPath $path).Hash -ceq $workbookPin)
        $same=$true;foreach($file in $pins.Keys){$same=$same -and (Get-FileHash -LiteralPath $file).Hash -ceq $pins[$file]}
        Check 'Instructions.SavedAuthorityPreserved' $same
    } finally {
        [void](Probe 'CloseDesigner')
        if($null -ne $book){$book.Close($false)}
        if($null -ne $decoy){$decoy.Close($false)}
        SelectTarget $Fixture
    }
}
