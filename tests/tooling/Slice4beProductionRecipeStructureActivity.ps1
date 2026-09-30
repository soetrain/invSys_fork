# Draft values stay in memory; emitted checks and diagnostics contain fixed metadata.
function Test-ProductionRecipeStructureActivity($Fixture,$Other) {
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    function Stage([string]$Action,[string]$Mode='Valid'){[void](Probe 'StructureStage' @($canary,$Action,$Mode))}
    function State {[string](Probe 'StructureState')}
    function Rows([string]$Kind='Nodes'){$value=[string](Probe 'StructureRows' @($Kind));if($value -ne ''){$value.Split([char]10)}}
    function Id([string]$Action){'PRODUCTION_RECIPE_'+$Action}
    function Ids([string[]]$Values){@($Values|ForEach-Object{$_.Split([char]9)[0]}) -join ','}
    function Pair([string[]]$Before,[string]$Action,[string]$Outcome,[string]$Case,[string]$Actor='config-producer'){
        $raw=@(Files|Where-Object{$_ -cnotin $Before}|ForEach-Object{[IO.File]::ReadAllText($_)})
        $records=@($raw|ForEach-Object{$_|ConvertFrom-Json})
        $first=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($records|Where-Object OutcomeCode -CEQ $Outcome)
        $paired=$records.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
        $control=Id $Action;$context=$paired;$redacted=$paired;$integrity=$paired;$linked=$false;$facts=$false;$terminal=$false
        foreach($r in $records){
            $context=$context -and $r.ControlId -ceq $control -and $r.OwnerId -ceq 'PRODUCTION_DESIGNER' -and $r.UserId -ceq $Actor -and $r.WarehouseId -ceq $Fixture.Warehouse -and $r.StationId -ceq 'S1' -and $r.CatalogVersion -eq 19
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
            $attempt=[bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.StructureTerminalForTest' @(($first[0]|ConvertTo-Json -Depth 20 -Compress)))
            $completed=[bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.StructureTerminalForTest' @(($last[0]|ConvertTo-Json -Depth 20 -Compress)))
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
    $actions=@('ADD_PROCESS','REMOVE_PROCESS','CONNECT','UPDATE_CONNECTION','DISCONNECT')
    $canary='STRUCTURE'+[guid]::NewGuid().ToString('N')
    $book=$null;$decoy=$null;$pins=@{};$recordPins=@{}
    foreach($target in @($Other,$Fixture)){
        SelectTarget $target
        if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.StructurePolicyForTest' @($true))){throw 'Authorized structure policy fixture unavailable; not product RED.'}
        foreach($file in Get-ChildItem -LiteralPath $target.Root -Recurse -File|Where-Object{$_.Extension -in '.xlsb','.xlsm' -and -not $_.Name.StartsWith('~$')}){$pins[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash}
    }
    foreach($file in Files){$recordPins[$file]=(Get-FileHash -LiteralPath $file).Hash}
    $old=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(17))).Split([char]10)|Where-Object{$_})
    $new=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(18))).Split([char]10)|Where-Object{$_})
    Check 'RecipeStructure.Catalog18Extends17' ($old.Count -eq 93 -and $new.Count -eq 98 -and @($new|Sort-Object -Unique).Count -eq 98 -and @($old|Where-Object{$_ -cnotin $new}).Count -eq 0)
    foreach($control in $old){
        $before=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($control,17))
        Check ('RecipeStructure.Catalog.Preserve.'+$control) ($before -ne '' -and $before -ceq [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($control,18)))
    }
    try{
        SelectTarget $Fixture 'config-producer'
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$canary
        $path=Join-Path $runRoot 'recipe-structure-operator.xlsb';$book.SaveAs($path,50);$book.Close($false)
        $workbookPin=(Get-FileHash -LiteralPath $path).Hash;$book=$excel.Workbooks.Open($path,0,$false)
        $sheet=$book.Worksheets.Item(1);$decoy=$excel.Workbooks.Add();$decoy.Activate()
        $before=@(Files);[void](Probe 'OpenDesigner' @($book.Name));Stage 'CONNECT'
        Check 'RecipeStructure.InitializationIsNotUserAction' (@(Files).Count -eq $before.Count)
        foreach($action in $actions){
            $label='RecipeStructure.'+$action;$control=Id $action
            $json=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($control,18));$definition=if($json){$json|ConvertFrom-Json}else{$null}
            Check ($label+'.FixedMetadata') ($null -ne $definition -and $definition.Caption -ceq [string](Probe 'StructureCaption' @($action)) -and $definition.Surface -ceq 'Operations > Production > Recipe Designer' -and $definition.Role -ceq 'Production' -and $definition.Class -ceq 'Command' -and $definition.Capability -ceq 'PROD_POST')
            $absent=$true;foreach($version in 1..17){$absent=$absent -and [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($control,$version)) -ceq ''}
            Check ($label+'.AbsentFromPriorCatalogs') $absent
            foreach($unsupported in @('COMPLETED','CONFIRMED','VALIDATED','APPLIED')){Check ($label+'.RejectsUnsupported.'+$unsupported) ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Outcome' @($control,$unsupported)) -ceq '')}
            Stage $action;$nodes=@(Rows);$edges=@(Rows 'Connections');$instructions=@(Rows 'Instructions') -join [char]10;$before=@(Files)
            if($action -ceq 'UPDATE_CONNECTION'){
                $editor=[string](Probe 'StructureEditor')
                $expected=@('N1','OEDIT','N2','R1','4','50','EA')
                Write-Output ('Structure Update staged editor matches: '+(@(0..6|ForEach-Object{$editor.Split([char]9)[$_] -ceq $expected[$_]}) -join ','))
            }
            $notice=[string](Probe 'StructureAct' @($action));$actualNodes=@(Rows);$actualEdges=@(Rows 'Connections');$expectedState=State
            $aux=([string](Probe 'StructureAux')).Split('|')
            switch($action){
                'ADD_PROCESS'{
                    Check ($label+'.AppendsIdentityVersionNameAndOrdinal') ($actualNodes.Count -eq 4 -and $actualNodes[3] -ceq (@('N4','P4','2',($canary+'RELEASED'),'4') -join [char]9))
                    Check ($label+'.ExistingNodesPreserved') (($actualNodes[0..2] -join [char]10) -ceq ($nodes -join [char]10))
                    Check ($label+'.ConnectionsPreserved') (($actualEdges -join [char]10) -ceq ($edges -join [char]10))
                    Check ($label+'.SelectionAndChoices') ($aux[0] -ceq '3' -and $aux[2] -ceq '4')
                }
                'REMOVE_PROCESS'{
                    Check ($label+'.SelectedNodeAndIncidentConnectionsRemoved') ((Ids $actualNodes) -ceq 'N1,N3' -and $actualEdges.Count -eq 1 -and $actualEdges[0] -ceq $edges[2])
                    Check ($label+'.SurvivorFieldsAndOrder') ($actualNodes.Count -eq 2 -and $actualNodes[0] -ceq (($nodes[0].Split([char]9)[0..3]+@('1')) -join [char]9) -and $actualNodes[1] -ceq (($nodes[2].Split([char]9)[0..3]+@('2')) -join [char]9))
                    Check ($label+'.ExistingStatusPreserved') ($notice -ceq 'Prior local status')
                }
                'CONNECT'{
                    Check ($label+'.AppendsSevenTrimmedFields') ($actualEdges.Count -eq 4 -and $actualEdges[3] -ceq (@('N1','OEDIT','N3','RNEW','4','50','EA') -join [char]9))
                    Check ($label+'.ExistingConnectionsPreserved') (($actualEdges[0..2] -join [char]10) -ceq ($edges -join [char]10))
                    Check ($label+'.SelectedConnection') ($aux[1] -ceq '3')
                }
                'UPDATE_CONNECTION'{
                    # Fixed booleans only: identify fixture/editor side effects without emitting draft values.
                    $expected=@('N1','OEDIT','N2','R1','4','50','EA');$observed=$actualEdges[0].Split([char]9)
                    Write-Output ('Structure Update field matches: '+(@(0..6|ForEach-Object{$observed[$_] -ceq $expected[$_]}) -join ','))
                    Write-Output ('Structure Update original field matches: '+(@(0..6|ForEach-Object{$observed[$_] -ceq $edges[0].Split([char]9)[$_]}) -join ','))
                    Write-Output ('Structure Update empty fields: '+(@(0..6|ForEach-Object{$observed[$_] -ceq ''}) -join ','))
                    Write-Output ('Structure Update nested selection callbacks: '+[string](Probe 'StructureConnectionClicks'))
                    Check ($label+'.ReplacesSevenTrimmedFields') ($actualEdges.Count -eq 3 -and $actualEdges[0] -ceq (@('N1','OEDIT','N2','R1','4','50','EA') -join [char]9))
                    Check ($label+'.OtherConnectionsPreserved') (($actualEdges[1..2] -join [char]10) -ceq ($edges[1..2] -join [char]10))
                    Check ($label+'.SelectedConnection') ($aux[1] -ceq '0')
                }
                'DISCONNECT'{
                    Check ($label+'.VisibleMappingPrecedesHiddenSelection') ($actualEdges.Count -eq 2 -and ($actualEdges -join [char]10) -ceq ($edges[0,2] -join [char]10))
                    Check ($label+'.EditorCleared') ($aux[3] -ceq '' -and $aux[4] -ceq '')
                }
            }
            if($action -in @('CONNECT','UPDATE_CONNECTION','DISCONNECT')){Check ($label+'.NodeFieldsPreserved') (($actualNodes -join [char]10) -ceq ($nodes -join [char]10))}
            Check ($label+'.InstructionsPreserved') ((@(Rows 'Instructions') -join [char]10) -ceq $instructions)
            Check ($label+'.NoHandlerError') (-not $notice.StartsWith('HANDLER_ERROR|'))
            Pair $before $action 'STAGED' ($label+'.Success')
            if($CaptureEvidence -and $action -in @('ADD_PROCESS','CONNECT')){[void](Probe 'StructureShow');CaptureOwnedFormByCaptionEvidence 'Production' ('recipe-structure-'+$action.ToLowerInvariant()+'.png')}
            foreach($guard in @('Busy','Loading')){Stage $action;$state=State;$before=@(Files);[void](Probe 'StructureGuard' @($action,$guard));Check ($label+'.'+$guard+'Guard') ((State) -ceq $state -and @(Files).Count -eq $before.Count)}
            Stage $action;$before=@(Files);$notice=[string](Probe 'StructureGuard' @($action,'Failure'))
            Check ($label+'.Failure.NoUnhandledError') (-not $notice.StartsWith('HANDLER_ERROR|'));Pair $before $action 'FAILED' ($label+'.Failure')
            Stage $action;$state=State;$before=@(Files);$notice=[string](Probe 'StructureGuard' @($action,'PartialFailure'))
            Check ($label+'.PartialFailure.EditRemains') ((State) -cne $state)
            Check ($label+'.PartialFailure.GuardsRestored') ([bool](Probe 'StructureRestored'))
            Check ($label+'.PartialFailure.NoUnhandledError') (-not $notice.StartsWith('HANDLER_ERROR|'))
            Pair $before $action 'FAILED' ($label+'.PartialFailure')
            Stage $action;$before=@(Files);$entered=[string](Probe 'StructureGuard' @($action,'Nested'))
            Check ($label+'.NestedEntered') ($entered -ceq 'True')
            Check ($label+'.NestedMutationSuppressed') ((State) -ceq $expectedState)
            Pair $before $action 'STAGED' ($label+'.Nested')
        }
        foreach($action in @('ADD_PROCESS','REMOVE_PROCESS','DISCONNECT')){
            Stage $action 'NoSelection';$state=State;$before=@(Files);[void](Probe 'StructureAct' @($action))
            Check ('RecipeStructure.'+$action+'.NoSelection.NoEdit') ((State) -ceq $state)
            Pair $before $action 'REJECTED' ('RecipeStructure.'+$action+'.NoSelection')
        }
        foreach($action in @('CONNECT','UPDATE_CONNECTION')){
            foreach($mode in @('NoSource','NoOutput','NoTarget','NoRequirement','Self','NonPositive','NoUom','FractionalEA','Duplicate','DuplicateCase')){
                Stage $action $mode;$state=State;$before=@(Files);$notice=[string](Probe 'StructureAct' @($action))
                Check ('RecipeStructure.'+$action+'.'+$mode+'.NoEdit') ((State) -ceq $state)
                Check ('RecipeStructure.'+$action+'.'+$mode+'.ValidationVisible') ($notice -cne 'Prior local status' -and -not $notice.StartsWith('HANDLER_ERROR|'))
                Pair $before $action 'REJECTED' ('RecipeStructure.'+$action+'.'+$mode)
            }
            foreach($mode in @('PercentOnly','OtherText','NegativeOther')){
                Stage $action $mode;$before=@(Files);$notice=[string](Probe 'StructureAct' @($action));$edges=@(Rows 'Connections')
                $idx=if($action -ceq 'CONNECT'){3}else{0};$expected=switch($mode){'PercentOnly'{''} 'OtherText'{'not numeric'} 'NegativeOther'{'-2'}}
                Check ('RecipeStructure.'+$action+'.'+$mode+'.ExistingPermissiveOtherField') ($edges.Count -gt $idx -and $edges[$idx].Split([char]9)[4] -ceq $expected -and -not $notice.StartsWith('HANDLER_ERROR|'))
                Pair $before $action 'STAGED' ('RecipeStructure.'+$action+'.'+$mode)
            }
        }
        foreach($mode in @('Collision','CaseInsensitive','AppendFallback','Unchanged','HiddenFallback','NoDisplaySelection','InvalidDisplay')){
            $action=switch($mode){'Collision'{'ADD_PROCESS'} 'CaseInsensitive'{'REMOVE_PROCESS'} 'AppendFallback'{'UPDATE_CONNECTION'} 'Unchanged'{'UPDATE_CONNECTION'} default{'DISCONNECT'}}
            Stage $action $mode;$state=State;$before=@(Files);$edges=@(Rows 'Connections');$notice=[string](Probe 'StructureAct' @($action));$after=@(Rows 'Connections')
            $preserved=switch($mode){
                'Collision'{(Ids @(Rows)) -ceq 'N1,N2,n4,N5'}
                'CaseInsensitive'{(Ids @(Rows)) -ceq 'N1,N3' -and $after.Count -eq 1 -and $after[0] -ceq $edges[2]}
                'AppendFallback'{$after.Count -eq 4 -and $after[3] -ceq (@('N1','OEDIT','N2','RNEW','4','50','EA') -join [char]9)}
                'Unchanged'{($after -join [char]10) -ceq ($edges -join [char]10)}
                'InvalidDisplay'{(State) -ceq $state}
                default{$after.Count -eq 2 -and ($after -join [char]10) -ceq ($edges[1..2] -join [char]10)}
            }
            Check ('RecipeStructure.'+$mode+'.ExistingBehavior') ($preserved -and -not $notice.StartsWith('HANDLER_ERROR|'))
            $outcome=if($mode -ceq 'InvalidDisplay'){'REJECTED'}else{'STAGED'}
            Pair $before $action $outcome ('RecipeStructure.'+$mode)
        }
        SelectTarget $Fixture 'config-reader';$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name))
        foreach($action in $actions){Stage $action;$state=State;$before=@(Files);[void](Probe 'StructureAct' @($action));Check ('RecipeStructure.Denied.'+$action+'.NoLocalMutation') ((State) -ceq $state);Pair $before $action 'DENIED' ('RecipeStructure.Denied.'+$action) 'config-reader'}
        Check 'RecipeStructure.UnknownColumnsInMemory' ($sheet.Cells.Item(1,1).Value2 -ceq 'Operator Extra' -and $sheet.Cells.Item(2,1).Value2 -ceq $canary)
        foreach($enabled in @($false,$true)){
            SelectTarget $Fixture
            if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.StructurePolicyForTest' @($enabled))){throw 'Authorized structure policy update unavailable; not product RED.'}
            $pins[$Fixture.Config]=(Get-FileHash -LiteralPath $Fixture.Config).Hash
            if(-not $enabled){
                SelectTarget $Fixture 'config-producer';$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name))
                foreach($action in $actions){Stage $action;$state=State;$before=@(Files);[void](Probe 'StructureAct' @($action));Check ('RecipeStructure.DisabledTracking.'+$action) ((State) -cne $state -and @(Files).Count -eq $before.Count)}
            }
        }
        $parent=Join-Path $Fixture.Root 'Training\Activity';$blocked=Join-Path $parent $Fixture.Warehouse;$held=$blocked+'-structure-held'
        foreach($item in @($blocked,$held)){if(-not [IO.Path]::GetFullPath($item).StartsWith([IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Tracking fault escaped the disposable fixture.'}}
        if(Test-Path -LiteralPath $held){throw 'Preserve existing held fixture.'}
        SelectTarget $Fixture 'config-producer';$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name));$moved=$false
        try{
            if(Test-Path -LiteralPath $blocked){Move-Item -LiteralPath $blocked -Destination $held;$moved=$true}
            [IO.File]::WriteAllText($blocked,'Blocked disposable structure activity path')
            foreach($action in $actions){Stage $action;$state=State;$before=@(Files);$notice=[string](Probe 'StructureAct' @($action));Check ('RecipeStructure.Unavailable.'+$action+'.EditContinues') ((State) -cne $state);Check ('RecipeStructure.Unavailable.'+$action+'.VisibleNoFallback') ($notice.Contains('Tracking unavailable') -and @(Files).Count -eq $before.Count)}
        }finally{if(Test-Path -LiteralPath $blocked -PathType Leaf){Remove-Item -LiteralPath $blocked};if($moved){Move-Item -LiteralPath $held -Destination $blocked}}
        foreach($guard in @('Target','Session','SignedOut','ClosedWorkbook')){
            SelectTarget $Fixture 'config-producer';$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name))
            if($guard -ceq 'Target'){SelectTarget $Other 'config-producer'}
            if($guard -ceq 'Session'){SelectTarget $Fixture 'config-producer'}
            if($guard -ceq 'SignedOut'){[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
            if($guard -ceq 'ClosedWorkbook'){$book.Close($false);$book=$null}
            $before=@(Files);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            foreach($action in $actions){Stage $action;$state=State;$notice=[string](Probe 'StructureAct' @($action));Check ('RecipeStructure.Guard.'+$guard+'.'+$action+'.NoLocalMutation') ((State) -ceq $state);Check ('RecipeStructure.Guard.'+$guard+'.'+$action+'.RefusalVisible') ($notice.StartsWith('Session, warehouse, or captured workbook changed.') -and $notice.Contains('Reopen Production'))}
            Check ('RecipeStructure.Guard.'+$guard+'.NoRedirectedActivity') (@(Files).Count -eq $before.Count -and @(Get-Slice4beActivityFiles $Other).Count -eq $otherBefore.Count)
        }
        Check 'RecipeStructure.WorkbookBytesAndUnknownColumnsPreserved' ((Get-FileHash -LiteralPath $path).Hash -ceq $workbookPin)
        $same=$true;foreach($file in $pins.Keys){$same=$same -and (Get-FileHash -LiteralPath $file).Hash -ceq $pins[$file]}
        Check 'RecipeStructure.SavedAuthorityPreserved' $same
        $same=$true;foreach($file in $recordPins.Keys){$same=$same -and (Get-FileHash -LiteralPath $file).Hash -ceq $recordPins[$file]}
        Check 'RecipeStructure.ExistingActivityImmutable' $same
    }finally{[void](Probe 'CloseDesigner');if($null -ne $book){$book.Close($false)};if($null -ne $decoy){$decoy.Close($false)};SelectTarget $Fixture}
}
