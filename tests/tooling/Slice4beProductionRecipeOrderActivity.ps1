# Draft values stay in memory; emitted checks and diagnostics contain fixed metadata.
function Test-ProductionRecipeOrderActivity($Fixture,$Other) {
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    function Stage([string]$Mode='Unordered'){[void](Probe 'OrderStage' @($canary,$Mode))}
    function State {[string](Probe 'OrderState')}
    function Rows([string]$Kind='Nodes'){$value=[string](Probe 'OrderRows' @($Kind));if($value -ne ''){$value.Split([char]10)}}
    function Id([string]$Action){if($Action -ceq 'AUTO'){'PRODUCTION_RECIPE_AUTO_ORDER'}else{'PRODUCTION_RECIPE_MOVE_'+$Action}}
    function Ids([string[]]$Values){@($Values|ForEach-Object{$_.Split([char]9)[0]}) -join ','}
    function StableFields([string[]]$Before,[string[]]$After){
        if($Before.Count -ne $After.Count){return $false}
        $a=@($Before|ForEach-Object{($_.Split([char]9)[0..3]) -join [char]9}|Sort-Object)
        $b=@($After|ForEach-Object{($_.Split([char]9)[0..3]) -join [char]9}|Sort-Object)
        return ($a -join [char]10) -ceq ($b -join [char]10)
    }
    function Pair([string[]]$Before,[string]$Action,[string]$Outcome,[string]$Case,[string]$Actor='config-producer'){
        $raw=@(Files|Where-Object{$_ -cnotin $Before}|ForEach-Object{[IO.File]::ReadAllText($_)})
        $records=@($raw|ForEach-Object{$_|ConvertFrom-Json})
        $first=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($records|Where-Object OutcomeCode -CEQ $Outcome)
        $paired=$records.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
        $control=Id $Action;$context=$paired;$redacted=$paired;$integrity=$paired;$linked=$false;$facts=$false;$terminal=$false
        foreach($r in $records){
            $context=$context -and $r.ControlId -ceq $control -and $r.OwnerId -ceq 'PRODUCTION_DESIGNER' -and $r.UserId -ceq $Actor -and $r.WarehouseId -ceq $Fixture.Warehouse -and $r.StationId -ceq 'S1' -and $r.CatalogVersion -eq 23
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
            $facts=$first[0].DataEffect -ceq 'Unknown' -and $first[0].Severity -ceq 'Info' -and $last[0].Severity -ceq $severity -and $last[0].DataEffect -ceq $effect -and $last[0].EventCode -ceq ($control+'_'+$Outcome)
            $attempt=[bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.OrderTerminalForTest' @(($first[0]|ConvertTo-Json -Depth 20 -Compress)))
            $completed=[bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.OrderTerminalForTest' @(($last[0]|ConvertTo-Json -Depth 20 -Compress)))
            $terminal=-not $attempt -and $completed -eq ($Outcome -ceq 'STAGED')
        }
        Check ($Case+'.AttemptAndOutcome') $paired
        Check ($Case+'.ExactContext') $context
        Check ($Case+'.NoEnteredDataOrSources') $redacted
        Check ($Case+'.ContentIntegrity') $integrity
        Check ($Case+'.LinkedDistinctRecords') $linked
        Check ($Case+'.OwnerFact') $facts
        Check ($Case+'.ExactCommandTerminal') $terminal
    }
    $actions=@('UP','DOWN','AUTO');$canary='ORDER'+[guid]::NewGuid().ToString('N')
    $book=$null;$decoy=$null;$pins=@{};$recordPins=@{}
    foreach($target in @($Other,$Fixture)){
        SelectTarget $target
        if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.OrderPolicyForTest' @($true))){throw 'Authorized ordering policy fixture unavailable; not product RED.'}
        foreach($file in Get-ChildItem -LiteralPath $target.Root -Recurse -File|Where-Object{$_.Extension -in '.xlsb','.xlsm' -and -not $_.Name.StartsWith('~$')}){$pins[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash}
    }
    foreach($file in Files){$recordPins[$file]=(Get-FileHash -LiteralPath $file).Hash}
    $old=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(16))).Split([char]10)|Where-Object{$_})
    $new=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(17))).Split([char]10)|Where-Object{$_})
    Check 'RecipeOrder.Catalog17Extends16' ($old.Count -eq 90 -and $new.Count -eq 93 -and @($new|Sort-Object -Unique).Count -eq 93 -and @($old|Where-Object{$_ -cnotin $new}).Count -eq 0)
    foreach($control in $old){
        $before=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($control,16))
        Check ('RecipeOrder.Catalog.Preserve.'+$control) ($before -ne '' -and $before -ceq [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($control,17)))
    }
    try{
        SelectTarget $Fixture 'config-producer'
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$canary
        $path=Join-Path $runRoot 'recipe-order-operator.xlsb';$book.SaveAs($path,50);$book.Close($false)
        $workbookPin=(Get-FileHash -LiteralPath $path).Hash;$book=$excel.Workbooks.Open($path,0,$false)
        $sheet=$book.Worksheets.Item(1);$decoy=$excel.Workbooks.Add();$decoy.Activate()
        $before=@(Files);[void](Probe 'OpenDesigner' @($book.Name));Stage
        Check 'RecipeOrder.InitializationIsNotUserAction' (@(Files).Count -eq $before.Count)
        foreach($action in $actions){
            $label='RecipeOrder.'+$action;$control=Id $action
            $json=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($control,17));$definition=if($json){$json|ConvertFrom-Json}else{$null}
            Check ($label+'.FixedMetadata') ($null -ne $definition -and $definition.Caption -ceq [string](Probe 'OrderCaption' @($action)) -and $definition.Surface -ceq 'Operations > Production > Recipe Designer' -and $definition.Role -ceq 'Production' -and $definition.Class -ceq 'Command' -and $definition.Capability -ceq 'PROD_POST')
            $absent=$true;foreach($version in 1..16){$absent=$absent -and [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($control,$version)) -ceq ''}
            Check ($label+'.AbsentFromPriorCatalogs') $absent
            foreach($unsupported in @('COMPLETED','CONFIRMED','VALIDATED','APPLIED')){Check ($label+'.RejectsUnsupported.'+$unsupported) ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Outcome' @($control,$unsupported)) -ceq '')}
            Stage;$priorRows=@(Rows);$connections=@(Rows 'Connections') -join [char]10;$before=@(Files)
            $notice=[string](Probe 'OrderAct' @($action));$actual=@(Rows)
            $expected=switch($action){'UP'{'N2,N1,N3'} 'DOWN'{'N1,N3,N2'} 'AUTO'{'N3,N1,N2'}}
            Check ($label+'.ExistingRowOrder') ((Ids $actual) -ceq $expected)
            Check ($label+'.AllNonOrderFieldsFollowIdentity') (StableFields $priorRows $actual)
            Check ($label+'.ExecutionOrderRenumbered') ((@($actual|ForEach-Object{$_.Split([char]9)[4]}) -join ',') -ceq '1,2,3')
            Check ($label+'.ConnectionsPreserved') ((@(Rows 'Connections') -join [char]10) -ceq $connections)
            Check ($label+'.InstructionNormalizationPreserved') ((@((Rows 'Instructions')|ForEach-Object{$_.Split([char]9)[0]}) -join ',') -ceq '1,2,3')
            $aux=([string](Probe 'OrderAux')).Split('|')
            Check ($label+'.SelectionFollowsMovedNode') ([int]$aux[0] -eq $(if($action -ceq 'DOWN'){2}else{0}))
            Check ($label+'.ConnectionChoicesRefreshed') ([int]$aux[1] -eq 3)
            Check ($label+'.NoHandlerError') (-not $notice.StartsWith('HANDLER_ERROR|'))
            Pair $before $action 'STAGED' ($label+'.Success')
            if($CaptureEvidence -and $action -in @('DOWN','AUTO')){[void](Probe 'OrderShow');CaptureOwnedFormByCaptionEvidence 'Production' ('recipe-order-'+$action.ToLowerInvariant()+'.png')}
            foreach($guard in @('Busy','Loading')){Stage;$state=State;$before=@(Files);[void](Probe 'OrderGuard' @($action,$guard));Check ($label+'.'+$guard+'Guard') ((State) -ceq $state -and @(Files).Count -eq $before.Count)}
            Stage;$before=@(Files);[void](Probe 'OrderGuard' @($action,'Failure'));Pair $before $action 'FAILED' ($label+'.Failure')
            Stage;$before=@(Files);$notice=[string](Probe 'OrderGuard' @($action,'PartialFailure'))
            Check ($label+'.PartialFailure.RemainsLocalAndNotRolledBack') ((@(Rows)[0]).Split([char]9)[4] -ceq '1' -and (@(Rows)[1]).Split([char]9)[4] -cne '2')
            Check ($label+'.PartialFailure.GuardsRestored') ([bool](Probe 'OrderRestored'))
            Check ($label+'.PartialFailure.NoUnhandledError') (-not $notice.StartsWith('HANDLER_ERROR|'))
            Pair $before $action 'FAILED' ($label+'.PartialFailure')
            Stage;$before=@(Files);$entered=[string](Probe 'OrderGuard' @($action,'Nested'))
            Check ($label+'.NestedEntered') ($entered -ceq 'True')
            Check ($label+'.NestedMutationSuppressed') ((Ids @(Rows)) -ceq $expected)
            Pair $before $action 'STAGED' ($label+'.Nested')
        }
        foreach($action in @('UP','DOWN')){
            foreach($mode in @('NoSelection',$(if($action -ceq 'UP'){'Top'}else{'Bottom'}))){
                Stage $mode;$priorRows=@(Rows);$before=@(Files);[void](Probe 'OrderAct' @($action));$actual=@(Rows)
                Check ('RecipeOrder.'+$action+'.'+$mode+'.RowsDoNotMove') ((Ids $actual) -ceq 'N1,N2,N3' -and (StableFields $priorRows $actual))
                Check ('RecipeOrder.'+$action+'.'+$mode+'.ExistingRenumberingStillOccurs') ((@($actual|ForEach-Object{$_.Split([char]9)[4]}) -join ',') -ceq '1,2,3')
                Check ('RecipeOrder.'+$action+'.'+$mode+'.NoMovementNoInstructionNormalization') ((@((Rows 'Instructions')|ForEach-Object{$_.Split([char]9)[0]}) -join ',') -ceq '9,10,11')
                Pair $before $action 'REJECTED' ('RecipeOrder.'+$action+'.'+$mode)
            }
        }
        foreach($mode in @('Empty','Ordered','Cycle','Self','MissingSource','MissingTarget','CaseInsensitive')){
            Stage $mode;$priorRows=@(Rows);$connections=@(Rows 'Connections') -join [char]10;$before=@(Files)
            $notice=[string](Probe 'OrderAct' @('AUTO'));$actual=@(Rows)
            $expected=switch($mode){'Empty'{''} 'Cycle'{'N1,N3,N2'} 'Self'{'N3,N1,N2'} 'CaseInsensitive'{'N3,N1,N2'} default{'N1,N2,N3'}}
            $outcome=if($mode -in @('Cycle','Self','MissingSource')){'REJECTED'}else{'STAGED'}
            Check ('RecipeOrder.Auto.'+$mode+'.ExistingOrder') ((Ids $actual) -ceq $expected)
            Check ('RecipeOrder.Auto.'+$mode+'.IdentitiesAndOtherFields') (StableFields $priorRows $actual)
            Check ('RecipeOrder.Auto.'+$mode+'.ConnectionsPreserved') ((@(Rows 'Connections') -join [char]10) -ceq $connections)
            Check ('RecipeOrder.Auto.'+$mode+'.NoHandlerError') (-not $notice.StartsWith('HANDLER_ERROR|'))
            Pair $before 'AUTO' $outcome ('RecipeOrder.Auto.'+$mode)
        }
        SelectTarget $Fixture 'config-reader';$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name))
        foreach($action in $actions){Stage;$state=State;$before=@(Files);[void](Probe 'OrderAct' @($action));Check ('RecipeOrder.Denied.'+$action+'.NoLocalMutation') ((State) -ceq $state);Pair $before $action 'DENIED' ('RecipeOrder.Denied.'+$action) 'config-reader'}
        Check 'RecipeOrder.UnknownColumnsInMemory' ($sheet.Cells.Item(1,1).Value2 -ceq 'Operator Extra' -and $sheet.Cells.Item(2,1).Value2 -ceq $canary)
        foreach($enabled in @($false,$true)){
            SelectTarget $Fixture
            if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.OrderPolicyForTest' @($enabled))){throw 'Authorized ordering policy update unavailable; not product RED.'}
            $pins[$Fixture.Config]=(Get-FileHash -LiteralPath $Fixture.Config).Hash
            if(-not $enabled){
                SelectTarget $Fixture 'config-producer';$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name))
                foreach($action in $actions){Stage;$state=State;$before=@(Files);[void](Probe 'OrderAct' @($action));Check ('RecipeOrder.DisabledTracking.'+$action) ((State) -cne $state -and @(Files).Count -eq $before.Count)}
            }
        }
        $parent=Join-Path $Fixture.Root 'Training\Activity';$blocked=Join-Path $parent $Fixture.Warehouse;$held=$blocked+'-ordering-held'
        foreach($item in @($blocked,$held)){if(-not [IO.Path]::GetFullPath($item).StartsWith([IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Tracking fault escaped the disposable fixture.'}}
        if(Test-Path -LiteralPath $held){throw 'Preserve existing held fixture.'}
        SelectTarget $Fixture 'config-producer';$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name));$moved=$false
        try{
            if(Test-Path -LiteralPath $blocked){Move-Item -LiteralPath $blocked -Destination $held;$moved=$true}
            [IO.File]::WriteAllText($blocked,'Blocked disposable ordering activity path')
            foreach($action in $actions){Stage;$state=State;$before=@(Files);$notice=[string](Probe 'OrderAct' @($action));Check ('RecipeOrder.Unavailable.'+$action+'.EditContinues') ((State) -cne $state);Check ('RecipeOrder.Unavailable.'+$action+'.VisibleNoFallback') ($notice.Contains('Tracking unavailable') -and @(Files).Count -eq $before.Count)}
        }finally{if(Test-Path -LiteralPath $blocked -PathType Leaf){Remove-Item -LiteralPath $blocked};if($moved){Move-Item -LiteralPath $held -Destination $blocked}}
        foreach($guard in @('Target','Session','SignedOut','ClosedWorkbook')){
            SelectTarget $Fixture 'config-producer';$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name))
            if($guard -ceq 'Target'){SelectTarget $Other 'config-producer'}
            if($guard -ceq 'Session'){SelectTarget $Fixture 'config-producer'}
            if($guard -ceq 'SignedOut'){[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
            if($guard -ceq 'ClosedWorkbook'){$book.Close($false);$book=$null}
            $before=@(Files);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            foreach($action in $actions){Stage;$state=State;$notice=[string](Probe 'OrderAct' @($action));Check ('RecipeOrder.Guard.'+$guard+'.'+$action+'.NoLocalMutation') ((State) -ceq $state);Check ('RecipeOrder.Guard.'+$guard+'.'+$action+'.RefusalVisible') ($notice.StartsWith('Session, warehouse, or captured workbook changed.') -and $notice.Contains('Reopen Production'))}
            Check ('RecipeOrder.Guard.'+$guard+'.NoRedirectedActivity') (@(Files).Count -eq $before.Count -and @(Get-Slice4beActivityFiles $Other).Count -eq $otherBefore.Count)
        }
        Check 'RecipeOrder.WorkbookBytesAndUnknownColumnsPreserved' ((Get-FileHash -LiteralPath $path).Hash -ceq $workbookPin)
        $same=$true;foreach($file in $pins.Keys){$same=$same -and (Get-FileHash -LiteralPath $file).Hash -ceq $pins[$file]}
        Check 'RecipeOrder.SavedAuthorityPreserved' $same
        $same=$true;foreach($file in $recordPins.Keys){$same=$same -and (Get-FileHash -LiteralPath $file).Hash -ceq $recordPins[$file]}
        Check 'RecipeOrder.ExistingActivityImmutable' $same
    }finally{[void](Probe 'CloseDesigner');if($null -ne $book){$book.Close($false)};if($null -ne $decoy){$decoy.Close($false)};SelectTarget $Fixture}
}
