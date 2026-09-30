# D18 component commands: fixed check names and Booleans, no entered values in reports.
function Test-ProductionComponentActivity($Fixture,$Other) {
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    function Stage([string]$Kind,[string]$Mode='Valid'){[void](Probe 'ComponentStage' @($Kind,$canary,$Mode))}
    function State {$value=[string](Probe 'ComponentSnapshot');if($value.Contains('PROBE_ERROR|')){throw 'Component snapshot unavailable; not product RED.'};$value}
    function Rows([string]$Kind){$value=[string](Probe 'ComponentRows' @($Kind));if($value.Contains('PROBE_ERROR|')){throw 'Component row snapshot unavailable; not product RED.'};$value.Split("`n")}
    function Pair([string[]]$Before,[string]$Kind,[string]$Action,[string]$Outcome,[string]$Case,[string]$Actor='config-producer'){
        $raw=@(Files|Where-Object{$_ -cnotin $Before}|ForEach-Object{[IO.File]::ReadAllText($_)})
        $records=@($raw|ForEach-Object{$_|ConvertFrom-Json})
        $first=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($records|Where-Object OutcomeCode -CEQ $Outcome)
        $paired=$records.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
        $id='PRODUCTION_PROCESS_'+$Kind+'_'+$Action
        $context=$paired;$redacted=$paired;$integrity=$paired;$linked=$false;$facts=$false;$terminal=$false
        foreach($r in $records){
            $context=$context -and $r.ControlId -ceq $id -and $r.OwnerId -ceq 'PRODUCTION_DESIGNER' -and $r.UserId -ceq $Actor -and $r.WarehouseId -ceq $Fixture.Warehouse -and $r.StationId -ceq 'S1' -and $r.CatalogVersion -eq 22
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
            $effect=if($Outcome -ceq 'FAILED'){'Unknown'}else{'Unchanged'}
            $facts=$first[0].DataEffect -ceq 'Unknown' -and $first[0].Severity -ceq 'Info' -and $last[0].Severity -ceq $severity -and $last[0].DataEffect -ceq $effect -and $last[0].EventCode -ceq ($id+'_'+$Outcome)
            $attemptCompleted=[bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ComponentTerminalForTest' @(($first[0]|ConvertTo-Json -Depth 20 -Compress)))
            $completed=[bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ComponentTerminalForTest' @(($last[0]|ConvertTo-Json -Depth 20 -Compress)))
            $terminal=-not $attemptCompleted -and $completed -eq ($Outcome -ceq 'STAGED')
        }
        Check ($Case+'.AttemptAndOutcome') $paired
        Check ($Case+'.ExactContext') $context
        Check ($Case+'.NoEnteredDataOrSources') $redacted
        Check ($Case+'.ContentIntegrity') $integrity
        Check ($Case+'.LinkedDistinctRecords') $linked
        Check ($Case+'.OwnerFact') $facts
        Check ($Case+'.ExactCommandTerminal') $terminal
    }
    $kinds=@('REQUIREMENT','OUTPUT');$actions=@('ADD','UPDATE','REMOVE','UP','DOWN')
    $canary='COMPONENT'+[guid]::NewGuid().ToString('N');$book=$null;$decoy=$null;$pins=@{}
    foreach($target in @($Other,$Fixture)){
        SelectTarget $target
        if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ComponentPolicyForTest' @($true))){throw 'Authorized policy fixture unavailable; not product RED.'}
        foreach($file in Get-ChildItem -LiteralPath $target.Root -Recurse -File|Where-Object{$_.Extension -in '.xlsb','.xlsm' -and -not $_.Name.StartsWith('~$')}){$pins[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash}
    }
    $old=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(15))).Split("`n")|Where-Object{$_})
    $new=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(16))).Split("`n")|Where-Object{$_})
    Check 'Components.Catalog16Extends15' ($old.Count -eq 80 -and $new.Count -eq 90 -and @($new|Sort-Object -Unique).Count -eq 90 -and @($old|Where-Object{$_ -cnotin $new}).Count -eq 0)
    foreach($id in $old){
        $before=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,15))
        Check ('Components.Catalog.Preserve.'+$id) ($before -ne '' -and $before -ceq [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,16)))
    }
    try {
        SelectTarget $Fixture 'config-producer'
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$canary
        $path=Join-Path $runRoot 'component-operator.xlsb';$book.SaveAs($path,50);$book.Close($false)
        $workbookPin=(Get-FileHash -LiteralPath $path).Hash;$book=$excel.Workbooks.Open($path,0,$false)
        $sheet=$book.Worksheets.Item(1);$decoy=$excel.Workbooks.Add();$decoy.Activate()
        $before=@(Files);[void](Probe 'OpenDesigner' @($book.Name));Stage 'REQUIREMENT'
        Check 'Components.InitializationIsNotUserAction' (@(Files).Count -eq $before.Count)
        foreach($kind in $kinds){foreach($action in $actions){
            $label='Components.'+$kind+'.'+$action;$id='PRODUCTION_PROCESS_'+$kind+'_'+$action
            $json=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,16));$definition=if($json){$json|ConvertFrom-Json}else{$null}
            $surface='Operations > Production > Process Designer > '+$(if($kind -ceq 'REQUIREMENT'){'Requirements'}else{'Outputs'})
            Check ($label+'.FixedMetadata') ($null -ne $definition -and $definition.Caption -ceq [string](Probe 'ComponentCaption' @($kind,$action)) -and $definition.Surface -ceq $surface -and $definition.Role -ceq 'Production' -and $definition.Class -ceq 'Command' -and $definition.Capability -ceq 'PROD_POST')
            $excluded=$true;foreach($version in 1..15){$excluded=$excluded -and [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,$version)) -ceq ''}
            Check ($label+'.AbsentFromPriorCatalogs') $excluded
            foreach($unsupported in @('COMPLETED','CONFIRMED','VALIDATED','APPLIED')){Check ($label+'.RejectsUnsupported.'+$unsupported) ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Outcome' @($id,$unsupported)) -ceq '')}
            Stage $kind;$oldRows=@(Rows $kind);$before=@(Files);$decoy.Activate();$notice=[string](Probe 'ComponentAct' @($kind,$action));$actual=@(Rows $kind)
            $local=switch($action){
                'ADD'{$actual.Count -eq 4 -and ($actual[0..2] -join "`n") -ceq ($oldRows -join "`n") -and $actual[3].Split("`t")[1] -ceq ($canary+'EDIT') -and $actual[3].Split("`t")[0] -cnotin @($oldRows|ForEach-Object{$_.Split("`t")[0]})}
                'UPDATE'{$actual.Count -eq 3 -and $actual[0] -ceq $oldRows[0] -and $actual[2] -ceq $oldRows[2] -and $actual[1].Split("`t")[0] -ceq $oldRows[1].Split("`t")[0] -and $actual[1].Split("`t")[1] -ceq ($canary+'EDIT')}
                'REMOVE'{($actual -join "`n") -ceq (@($oldRows[0],$oldRows[2]) -join "`n")}
                'UP'{($actual -join "`n") -ceq (@($oldRows[1],$oldRows[0],$oldRows[2]) -join "`n")}
                'DOWN'{($actual -join "`n") -ceq (@($oldRows[0],$oldRows[2],$oldRows[1]) -join "`n")}
            }
            Check ($label+'.ActualLocalRows') $local
            Check ($label+'.NoHandlerError') (-not $notice.StartsWith('HANDLER_ERROR|'))
            if($notice -cmatch '^HANDLER_ERROR\|(-?[0-9]+)$'){
                [pscustomobject]@{Case=$label;ErrorNumber=[long]$Matches[1]}|ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'component-handler-errors.jsonl')
            }
            $editor=([string](Probe 'ComponentEditor' @($kind))).Split("`t");$aux=([string](Probe 'ComponentAux' @($kind))).Split('|')
            if($action -in @('ADD','REMOVE')){Check ($label+'.ExistingEditorReset') ($editor[1] -ceq '' -and $editor[0] -ne '')}
            if($action -ceq 'UPDATE'){Check ($label+'.ExistingEditorRetained') ($editor[1] -ceq (' '+$canary+'EDIT '))}
            if($action -in @('UP','DOWN')){Check ($label+'.SelectionAndOrdinalSemantics') ([int]$aux[0] -eq $(if($action -ceq 'UP'){0}else{2}) -and $aux[3] -ceq '1,2,3,')}
            Check ($label+'.OutputRegulationPreservedOrPreciselyRemoved') ($aux[2] -ceq $(if($kind -ceq 'OUTPUT' -and $action -ceq 'REMOVE'){'B1,B3'}else{'B1,B2,B3'}))
            Pair $before $kind $action 'STAGED' ($label+'.Success')
            if($CaptureEvidence -and $action -ceq 'UPDATE'){
                [void](Probe 'ComponentShow')
                CaptureOwnedFormByCaptionEvidence 'Production' ('component-'+$kind.ToLowerInvariant()+'.png')
            }
            foreach($guard in @('Busy','Loading')){Stage $kind;$prior=State;$before=@(Files);[void](Probe 'ComponentGuard' @($kind,$action,$guard));Check ($label+'.'+$guard+'Guard') ((State) -ceq $prior -and @(Files).Count -eq $before.Count)}
            Stage $kind;$before=@(Files);[void](Probe 'ComponentGuard' @($kind,$action,'Failure'));Pair $before $kind $action 'FAILED' ($label+'.Failure')
            $mode=switch($action){'UP'{'Top'} 'DOWN'{'Bottom'} 'REMOVE'{'NoSelection'} default{'Invalid'}}
            Stage $kind $mode;$prior=@(Rows $kind) -join "`n";$before=@(Files);[void](Probe 'ComponentAct' @($kind,$action))
            Check ($label+'.RejectedRowsPreserved') ((@(Rows $kind) -join "`n") -ceq $prior)
            if($action -ceq 'REMOVE'){Check ($label+'.NoSelectionStillResetsEditor') (([string](Probe 'ComponentEditor' @($kind))).Split("`t")[1] -ceq '')}
            Pair $before $kind $action 'REJECTED' ($label+'.Rejected')
        }}
        foreach($kind in $kinds){
            Stage $kind;$oldRows=@(Rows $kind);$before=@(Files);$notice=[string](Probe 'ComponentGuard' @($kind,'ADD','PartialFailure'));$partial=@(Rows $kind)
            Check ('Components.'+$kind+'.PartialFailure.ExistingRowsRetained') ($partial.Count -eq 4 -and ($partial[0..2] -join "`n") -ceq ($oldRows -join "`n") -and $partial[3].Split("`t")[0] -ne '')
            Check ('Components.'+$kind+'.PartialFailure.GuardsRestored') ([bool](Probe 'ComponentFailureRestored'))
            Pair $before $kind 'ADD' 'FAILED' ('Components.'+$kind+'.PartialFailure')
            foreach($mode in @('Append','Fallback','IdentityLookup','Actual')){
                Stage $kind $mode;$before=@(Files)
                $stagedEditor=([string](Probe 'ComponentEditor' @($kind))).Split("`t")
                if($mode -ceq 'Actual' -and ($stagedEditor[-1] -cne 'ACTUAL' -or $stagedEditor[2] -cne '' -or $stagedEditor[3] -cne '' -or $stagedEditor[4] -cne '')){throw 'ACTUAL component editor setup unavailable; not product RED.'}
                $notice=[string](Probe 'ComponentAct' @($kind,'UPDATE'));$rows=@(Rows $kind)
                $index=if($mode -ceq 'Append'){3}else{1}
                $valid=$rows.Count -eq $(if($mode -ceq 'Append'){4}else{3}) -and $rows[$index].Split("`t")[1] -ceq ($canary+'EDIT')
                if($mode -ceq 'Actual'){$fields=$rows[$index].Split("`t");$start=if($kind -ceq 'REQUIREMENT'){2}else{5};$valid=$valid -and $fields[$start] -ceq '' -and $fields[$start+1] -ceq '' -and $fields[$start+2] -ceq '' -and $fields[-1] -ceq 'ACTUAL'}
                if($mode -ceq 'Actual'){
                    [pscustomobject]@{Kind=$kind;EditorActual=($stagedEditor[-1] -ceq 'ACTUAL');EditorQuantityEmpty=($stagedEditor[2] -ceq '');EditorPercentEmpty=($stagedEditor[3] -ceq '');EditorBasisEmpty=($stagedEditor[4] -ceq '');RowCountThree=($rows.Count -eq 3);RowNameMatches=($fields[1] -ceq ($canary+'EDIT'));RowQuantityEmpty=($fields[$start] -ceq '');RowPercentEmpty=($fields[$start+1] -ceq '');RowBasisEmpty=($fields[$start+2] -ceq '');RowActual=($fields[-1] -ceq 'ACTUAL');NoHandlerError=(-not $notice.StartsWith('HANDLER_ERROR|'))}|ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'component-actual-calibration.jsonl')
                }
                Check ('Components.'+$kind+'.Update.'+$mode+'.ExistingSemantics') $valid
                Pair $before $kind 'UPDATE' 'STAGED' ('Components.'+$kind+'.Update.'+$mode)
            }
            Stage $kind 'Fraction';$prior=@(Rows $kind) -join "`n";$before=@(Files);[void](Probe 'ComponentAct' @($kind,'ADD'))
            Check ('Components.'+$kind+'.WholeUomValidation') ((@(Rows $kind) -join "`n") -ceq $prior)
            Pair $before $kind 'ADD' 'REJECTED' ('Components.'+$kind+'.WholeUom')
            Stage $kind;$before=@(Files);$entered=[string](Probe 'ComponentGuard' @($kind,'UP','Nested'))
            Check ('Components.'+$kind+'.NestedEntered') ($entered -ceq 'True')
            Check ('Components.'+$kind+'.NestedAddSuppressed') (@(Rows $kind).Count -eq 3)
            Pair $before $kind 'UP' 'STAGED' ('Components.'+$kind+'.Nested')
        }
        foreach($mode in @('YieldDefault','YieldChange')){
            Stage 'OUTPUT' $mode;$before=@(Files);[void](Probe 'ComponentAct' @('OUTPUT','UPDATE'));$fields=(@(Rows 'OUTPUT')[1]).Split("`t")
            Check ('Components.Output.'+$mode+'.ExistingDefaults') ($fields[5] -ceq '4' -and $fields[6] -ceq '100' -and $fields[7] -ceq '4')
            Pair $before 'OUTPUT' 'UPDATE' 'STAGED' ('Components.Output.'+$mode)
        }
        SelectTarget $Fixture 'config-reader';$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name))
        foreach($kind in $kinds){foreach($action in $actions){Stage $kind;$prior=State;$before=@(Files);[void](Probe 'ComponentAct' @($kind,$action));Check ('Components.Denied.'+$kind+'.'+$action+'.NoLocalMutation') ((State) -ceq $prior);Pair $before $kind $action 'DENIED' ('Components.Denied.'+$kind+'.'+$action) 'config-reader'}}
        Check 'Components.UnknownColumnsInMemory' ($sheet.Cells.Item(1,1).Value2 -ceq 'Operator Extra' -and $sheet.Cells.Item(2,1).Value2 -ceq $canary)
        foreach($enabled in @($false,$true)){
            SelectTarget $Fixture
            if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ComponentPolicyForTest' @($enabled))){throw 'Authorized policy update unavailable; not product RED.'}
            $pins[$Fixture.Config]=(Get-FileHash -LiteralPath $Fixture.Config).Hash
            if(-not $enabled){
                SelectTarget $Fixture 'config-producer';$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name))
                foreach($kind in $kinds){foreach($action in $actions){Stage $kind;$prior=State;$before=@(Files);[void](Probe 'ComponentAct' @($kind,$action));Check ('Components.DisabledTracking.'+$kind+'.'+$action) ((State) -cne $prior -and @(Files).Count -eq $before.Count)}}
            }
        }
        $parent=Join-Path $Fixture.Root 'Training\Activity';$blocked=Join-Path $parent $Fixture.Warehouse;$held=$blocked+'-component-held'
        foreach($item in @($blocked,$held)){if(-not [IO.Path]::GetFullPath($item).StartsWith([IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Tracking fault escaped the disposable fixture.'}}
        if(Test-Path -LiteralPath $held){throw 'Preserve existing held fixture.'}
        SelectTarget $Fixture 'config-producer';$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name));$moved=$false
        try{
            if(Test-Path -LiteralPath $blocked){Move-Item -LiteralPath $blocked -Destination $held;$moved=$true};[IO.File]::WriteAllText($blocked,'Blocked disposable component activity path')
            foreach($kind in $kinds){foreach($action in $actions){Stage $kind;$prior=State;$before=@(Files);$notice=[string](Probe 'ComponentAct' @($kind,$action));Check ('Components.Unavailable.'+$kind+'.'+$action+'.EditContinues') ((State) -cne $prior);Check ('Components.Unavailable.'+$kind+'.'+$action+'.VisibleNoFallback') ($notice.Contains('Tracking unavailable') -and @(Files).Count -eq $before.Count)}}
        }finally{if(Test-Path -LiteralPath $blocked -PathType Leaf){Remove-Item -LiteralPath $blocked};if($moved){Move-Item -LiteralPath $held -Destination $blocked}}
        foreach($guard in @('Target','Session','SignedOut','ClosedWorkbook')){
            SelectTarget $Fixture 'config-producer';$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name))
            if($guard -ceq 'Target'){SelectTarget $Other 'config-producer'}
            if($guard -ceq 'Session'){SelectTarget $Fixture 'config-producer'}
            if($guard -ceq 'SignedOut'){[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
            if($guard -ceq 'ClosedWorkbook'){$book.Close($false);$book=$null}
            $before=@(Files);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            foreach($kind in $kinds){foreach($action in $actions){Stage $kind;$prior=State;$notice=[string](Probe 'ComponentAct' @($kind,$action));Check ('Components.Guard.'+$guard+'.'+$kind+'.'+$action+'.NoLocalMutation') ((State) -ceq $prior);Check ('Components.Guard.'+$guard+'.'+$kind+'.'+$action+'.RefusalVisible') ($notice.StartsWith('Session, warehouse, or captured workbook changed.') -and $notice.Contains('Reopen Production'))}}
            Check ('Components.Guard.'+$guard+'.NoRedirectedActivity') (@(Files).Count -eq $before.Count -and @(Get-Slice4beActivityFiles $Other).Count -eq $otherBefore.Count)
        }
        Check 'Components.WorkbookBytesAndUnknownColumnsPreserved' ((Get-FileHash -LiteralPath $path).Hash -ceq $workbookPin)
        $same=$true;foreach($file in $pins.Keys){$same=$same -and (Get-FileHash -LiteralPath $file).Hash -ceq $pins[$file]}
        Check 'Components.SavedAuthorityPreserved' $same
    }finally{[void](Probe 'CloseDesigner');if($null -ne $book){$book.Close($false)};if($null -ne $decoy){$decoy.Close($false)};SelectTarget $Fixture}
}
