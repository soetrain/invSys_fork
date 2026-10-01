# Initial actual-handler baseline. Submission-fault and paired-path gates remain
# separate required evidence; a GREEN here alone cannot accept the new contract.
function Test-ProductionAssignmentActivity($Fixture,$Other) {
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    function Hash([string]$Path){$s=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $s).Hash}finally{$s.Dispose()}}
    function Stage([string]$Mode='Normal'){[void](Probe 'AssignmentStage' @($Mode))}
    function State {[string](Probe 'AssignmentState')}
    function Positive([string]$Action){switch($Action){'REFRESH'{'REFRESHED'} 'PROCESS'{'PRESENTED'} 'PROCESS_SELECT'{'PRESENTED'} 'REQUIREMENT'{'SELECTED'} 'REQUIREMENT_SELECT'{'SELECTED'} 'SAVE'{'CONFIRMED'} default{'STAGED'}}}
    function Pair([string[]]$Before,[string]$Action,[string]$Outcome,[string]$Case){
        $raw=@(Files|Where-Object{$_ -cnotin $Before}|ForEach-Object{[IO.File]::ReadAllText($_)})
        $records=@($raw|ForEach-Object{$_|ConvertFrom-Json})
        $first=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($records|Where-Object OutcomeCode -CEQ $Outcome)
        $paired=$records.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
        $context=$paired;$redacted=$paired;$integrity=$paired;$linked=$false;$facts=$false;$terminal=$false;$sources=$paired
        $control='PRODUCTION_ASSIGNMENT_'+$Action
        foreach($r in $records){
            $context=$context -and $r.ControlId -ceq $control -and $r.OwnerId -ceq 'PRODUCTION_ASSIGNMENT' -and $r.UserId -ceq 'config-producer' -and $r.WarehouseId -ceq $Fixture.Warehouse -and $r.StationId -ceq 'S1' -and $r.CatalogVersion -eq 24
            if($Action -ceq 'SAVE' -and $r.OutcomeCode -ceq 'CONFIRMED'){
                $sources=$sources -and @($r.SourceEventRefs).Count -eq 1
                foreach($ref in $r.SourceEventRefs){$sources=$sources -and $ref.SourceKind -ceq 'Designs' -and $ref.WarehouseId -ceq $Fixture.Warehouse -and $ref.SubmissionState -ceq 'Submitted' -and $ref.EventId -cne ''}
            }else{$sources=$sources -and @($r.SourceEventRefs).Count -eq 0}
        }
        foreach($value in $raw){
            foreach($secret in @($canary,$Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'mBtn','PinHash','ASSIGN-OLD','ASSIGN-NEW','ASSIGN-EXACT-KEY')){
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
            $effect=if($Outcome -in @('FAILED','CONFIRMED')){'Unknown'}else{'Unchanged'}
            $facts=$first[0].DataEffect -ceq 'Unknown' -and $first[0].Severity -ceq 'Info' -and $last[0].Severity -ceq $severity -and $last[0].DataEffect -ceq $effect -and $last[0].EventCode -ceq ($control+'_'+$Outcome)
            $attempt=[bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @(($first[0]|ConvertTo-Json -Depth 20 -Compress)))
            $completed=[bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @(($last[0]|ConvertTo-Json -Depth 20 -Compress)))
            $terminal=-not $attempt -and $completed -eq ($Outcome -ceq (Positive $Action))
        }
        Check ($Case+'.AttemptAndOutcome') $paired
        Check ($Case+'.ExactContext') $context
        Check ($Case+'.NoEnteredData') $redacted
        Check ($Case+'.ContentIntegrity') $integrity
        Check ($Case+'.LinkedDistinctRecords') $linked
        Check ($Case+'.OwnerFact') $facts
        Check ($Case+'.DeclaredSourceEnvelope') $sources
        Check ($Case+'.ExactCommandTerminal') $terminal
    }
    $canary='ASSIGN'+[guid]::NewGuid().ToString('N');$book=$null;$decoy=$null;$pins=@{};$recordPins=@{}
    $actions=@('REFRESH','PROCESS','REQUIREMENT','ADD','REMOVE','CLEAR','PROCESS_SELECT','REQUIREMENT_SELECT','SAVE')
    SelectTarget $Fixture
    if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.AssignmentPolicyForTest' @($true,$true))){throw 'Authorized Assignment policy fixture unavailable; not product RED.'}
    try{
        SelectTarget $Fixture 'config-producer'
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$canary
        $sheet.Cells.Item(2,2).Formula='=1+2'
        $path=Join-Path $runRoot 'assignment-operator.xlsb';$book.SaveAs($path,50);$book.Close($false)
        $bookPin=Hash $path;$book=$excel.Workbooks.Open($path,0,$false);$sheet=$book.Worksheets.Item(1)
        $decoy=$excel.Workbooks.Add();$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name))
        $ready=[bool](Probe 'AssignmentPrepare' @($canary));Check 'Assignment.RealReleasedFixtureReady' $ready
        if(-not $ready){
            $stage=[string](Probe 'ReadSetupStage');if($stage -cnotin @('ReleasedProcess','NEW','SAVE','RELEASE','Ready')){$stage='Unavailable'}
            throw ('Released Assignment fixture setup failed at '+$stage+'; not product RED.')
        }
        foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File|Where-Object{$_.Extension -in '.xlsb','.xlsm' -and $_.Name -notlike '~$*'}){$pins[$file.FullName]=Hash $file.FullName}
        foreach($file in Files){$recordPins[$file]=Hash $file}
        $old=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(22))).Split([char]10)|Where-Object{$_})
        $new=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(23))).Split([char]10)|Where-Object{$_})
        Check 'Assignment.Catalog23Extends22InOrder' ($old.Count -eq 109 -and $new.Count -eq 118 -and @($new|Sort-Object -Unique).Count -eq 118 -and (($new|Select-Object -First 109) -join '|') -ceq ($old -join '|'))
        $preserved=$true
        foreach($id in $old){$prior=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,22));$preserved=$preserved -and $prior -cne '' -and $prior -ceq [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,23))}
        Check 'Assignment.Catalog22DefinitionsPreserved' $preserved
        foreach($action in $actions){
            $label='Assignment.'+$action;$control='PRODUCTION_ASSIGNMENT_'+$action
            $json=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($control,23));$definition=if($json){$json|ConvertFrom-Json}else{$null}
            $caption=switch($action){'REFRESH'{'Refresh'} 'PROCESS'{'Select Process'} 'REQUIREMENT'{'Select Requirement'} 'ADD'{'Add Acceptable'} 'REMOVE'{'Remove Row'} 'CLEAR'{'Clear'} 'SAVE'{'Save Alternatives'} 'PROCESS_SELECT'{'Processes'} 'REQUIREMENT_SELECT'{'Ingredient Requirements'}}
            $class=if($action.EndsWith('_SELECT')){'Navigation'}else{'Command'}
            Check ($label+'.FixedMetadata') ($null -ne $definition -and $definition.Caption -ceq $caption -and $definition.Surface -ceq 'Operations > Production > Ingredients Assignment' -and $definition.OwnerId -ceq 'PRODUCTION_ASSIGNMENT' -and $definition.Role -ceq 'Production' -and $definition.Class -ceq $class -and $definition.Capability -ceq 'PROD_POST')
            $absent=$true;foreach($v in 1..22){$absent=$absent -and [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($control,$v)) -ceq ''};Check ($label+'.AbsentFromOlderCatalogs') $absent
            Stage;$before=@(Files)
            if($action -ceq 'SAVE'){
                $same=$true;foreach($file in $pins.Keys){$same=$same -and (Hash $file) -ceq $pins[$file]};Check 'Assignment.LocalActionsPreserveSavedAuthority' $same
            }
            $notice=[string](Probe 'AssignmentAct' @($action))
            Check ($label+'.ExistingOwnerBehavior') ([bool](Probe 'AssignmentPreserved' @($action,'Normal')))
            Check ($label+'.NoUnhandledError') (-not $notice.StartsWith('HANDLER_ERROR|'))
            Pair $before $action (Positive $action) ($label+'.Normal')
            if($action -ceq 'SAVE'){continue}
            foreach($guard in @('Loading','Busy')){
                Stage;$state=State;$before=@(Files);[void](Probe 'AssignmentGuard' @($action,$guard))
                Check ($label+'.'+$guard+'.NoReadsOrMutationOrActivity') ((State) -ceq $state -and [int](Probe 'ReadCalls') -eq 0 -and @(Files).Count -eq $before.Count)
            }
        }
        # Save was last: local actions must not have changed authority before it.
        # Assert the saved version above, then pin the intended save for refusal tests.
        foreach($file in @($pins.Keys)){$pins[$file]=Hash $file}
        foreach($entry in @(
            @{Action='PROCESS';Mode='NoSelection';Outcome='REJECTED'},
            @{Action='PROCESS_SELECT';Mode='Malformed';Outcome='FAILED'},
            @{Action='REQUIREMENT';Mode='NoSelection';Outcome='REJECTED'},
            @{Action='REQUIREMENT_SELECT';Mode='NoSelection';Outcome='REJECTED'},
            @{Action='ADD';Mode='Duplicate';Outcome='REJECTED'},
            @{Action='ADD';Mode='NoSelection';Outcome='REJECTED'},
            @{Action='REMOVE';Mode='NoSelection';Outcome='REJECTED'},
            @{Action='REMOVE';Mode='NoMatch';Outcome='REJECTED'},
            @{Action='SAVE';Mode='NoSelection';Outcome='REJECTED'})){
            Stage $entry.Mode;$state=State;$before=@(Files);$notice=[string](Probe 'AssignmentAct' @($entry.Action))
            $label='Assignment.'+$entry.Action+'.'+$entry.Mode
            Check ($label+'.PreservesLocalDraft') ((State) -ceq $state)
            Pair $before $entry.Action $entry.Outcome $label
        }
        Stage 'Empty';$before=@(Files);[void](Probe 'AssignmentAct' @('CLEAR'));Pair $before 'CLEAR' 'STAGED' 'Assignment.EmptyClear'
        foreach($action in @('PROCESS','PROCESS_SELECT')){
            Stage 'EmptyArray';$before=@(Files);[void](Probe 'AssignmentAct' @($action))
            Check ('Assignment.'+$action+'.EmptyArrayBehavior') ([bool](Probe 'AssignmentPreserved' @($action,'EmptyArray')))
            Pair $before $action 'PRESENTED' ('Assignment.'+$action+'.EmptyArray')
        }
        $same=$true;foreach($file in $pins.Keys){$same=$same -and (Hash $file) -ceq $pins[$file]};Check 'Assignment.RefusalsPreserveSavedAuthority' $same
        Check 'Assignment.UnknownValuesAndFormula' ($sheet.Cells.Item(1,1).Value2 -ceq 'Operator Extra' -and $sheet.Cells.Item(2,1).Value2 -ceq $canary -and $sheet.Cells.Item(2,2).Formula -ceq '=1+2')
        foreach($guard in @('Target','Session','SignedOut','ClosedWorkbook')){
            SelectTarget $Fixture 'config-producer';[void](Probe 'ReadReopen' @($book.Name));$decoy.Activate()
            if($guard -ceq 'Target'){SelectTarget $Other 'config-producer'}
            if($guard -ceq 'Session'){SelectTarget $Fixture 'config-producer'}
            if($guard -ceq 'SignedOut'){[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
            if($guard -ceq 'ClosedWorkbook'){$book.Close($false);$book=$null}
            $before=@(Files);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            foreach($action in $actions|Where-Object{$_ -cne 'SAVE'}){
                Stage;$state=State;$notice=[string](Probe 'AssignmentAct' @($action))
                Check ('Assignment.Guard.'+$guard+'.'+$action+'.NoReadsOrLocalMutation') ((State) -ceq $state -and [int](Probe 'ReadCalls') -eq 0)
                Check ('Assignment.Guard.'+$guard+'.'+$action+'.RefusalVisible') ($notice.StartsWith('Session, warehouse, or captured workbook changed.') -and $notice.Contains('Reopen Production'))
            }
            Check ('Assignment.Guard.'+$guard+'.NoRedirectedActivity') (@(Files).Count -eq $before.Count -and @(Get-Slice4beActivityFiles $Other).Count -eq $otherBefore.Count)
        }
        [void](Probe 'CloseDesigner');Check 'Assignment.OperatorBytesPreserved' ((Hash $path) -ceq $bookPin)
        $same=$true;foreach($file in $recordPins.Keys){$same=$same -and (Hash $file) -ceq $recordPins[$file]};Check 'Assignment.OlderRecordsImmutable' $same
    }finally{[void](Probe 'CloseDesigner');if($null -ne $book){$book.Close($false)};if($null -ne $decoy){$decoy.Close($false)};SelectTarget $Fixture}
}
