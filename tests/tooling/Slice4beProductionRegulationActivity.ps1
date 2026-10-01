# Actual operator handlers; fixture values remain in memory and checks are fixed metadata.
function Test-RegulationRecordRedacted([string]$Raw,[string[]]$Sensitive,[string[]]$Identities) {
    $record=$Raw|ConvertFrom-Json
    $decoded=$record|ConvertTo-Json -Depth 30 -Compress
    foreach($value in $Sensitive){
        $encoded=ConvertTo-Json -InputObject $value -Compress
        if($decoded.Contains($value) -or $decoded.Contains($encoded.Substring(1,$encoded.Length-2))){return $false}
    }
    function SafeValue($Value,[string]$Name=''){
        if($null -eq $Value){return $true}
        if($Value -is [string]){
            # Three-character business identities can occur inside unrelated
            # generated identifiers, hashes or fractional timestamps. Exempt
            # only these named fields with their complete technical syntax.
            if($Name -cin @('RecordId','ActivityId','SequenceId') -and $Value -cmatch '^[a-f0-9]{8}(-[a-f0-9]{4}){3}-[a-f0-9]{12}$'){return $true}
            if($Name -ceq 'ContentSha256' -and $Value -cmatch '^[a-f0-9]{64}$'){return $true}
            if($Name -ceq 'OccurredAtUTC' -and $Value -cmatch '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?Z$'){return $true}
            foreach($identity in $Identities){if($Value.Contains($identity)){return $false}}
        }elseif($Value -is [Collections.IEnumerable]){
            foreach($item in $Value){if(-not (SafeValue $item)){return $false}}
        }elseif($Value -is [pscustomobject]){
            foreach($property in $Value.PSObject.Properties){if(-not (SafeValue $property.Value $property.Name)){return $false}}
        }
        return $true
    }
    return (SafeValue $record)
}
function Test-ProductionRegulationActivity($Fixture,$Other) {
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    function Hash([string]$Path){$stream=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}}
    function Stage([string]$Scope='Process',[string]$Mode='Normal'){[void](Probe 'RegulationStage' @($Scope,$Mode,$canary,$processId,$processVersion))}
    function State {[string](Probe 'RegulationState')}
    function Value([string]$Scope,[string]$Node='NODE1',[string]$Output='A01'){[string](Probe 'RegulationValue' @($Scope,$Node,$Output))}
    function Pair([string[]]$Before,[string]$Action,[string]$Outcome,[string]$Case,[string]$Actor='config-producer'){
        $raw=@(Files|Where-Object{$_ -cnotin $Before}|ForEach-Object{[IO.File]::ReadAllText($_)})
        $records=@($raw|ForEach-Object{$_|ConvertFrom-Json})
        $first=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($records|Where-Object OutcomeCode -CEQ $Outcome)
        $paired=$records.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
        $context=$paired;$redacted=$paired;$integrity=$paired;$linked=$false;$facts=$false;$terminal=$false
        $control='PRODUCTION_OUTPUT_REGULATION_'+$Action
        foreach($r in $records){
            $context=$context -and $r.ControlId -ceq $control -and $r.OwnerId -ceq 'PRODUCTION_DESIGNER' -and $r.UserId -ceq $Actor -and $r.WarehouseId -ceq $Fixture.Warehouse -and $r.StationId -ceq 'S1' -and $r.CatalogVersion -eq 23
            $redacted=$redacted -and @($r.SourceEventRefs).Count -eq 0
        }
        foreach($value in $raw){
            $redacted=$redacted -and (Test-RegulationRecordRedacted $value @($canary,$Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'mBtn','PinHash') @($processId,'NODE1','A01'))
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
            $attempt=[bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.RegulationTerminalForTest' @(($first[0]|ConvertTo-Json -Depth 20 -Compress)))
            $completed=[bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.RegulationTerminalForTest' @(($last[0]|ConvertTo-Json -Depth 20 -Compress)))
            $terminal=-not $attempt -and $completed -eq ($Outcome -ceq 'STAGED')
        }
        Check ($Case+'.AttemptAndOutcome') $paired
        Check ($Case+'.ExactContext') $context
        Check ($Case+'.NoEnteredDataOrSources') $redacted
        Check ($Case+'.ContentIntegrity') $integrity
        Check ($Case+'.LinkedDistinctRecords') $linked
        Check ($Case+'.OwnerFact') $facts
        Check ($Case+'.ExactLocalCommandTerminal') $terminal
    }
    $canary='REGULATION'+[guid]::NewGuid().ToString('N');$book=$null;$decoy=$null;$pins=@{};$recordPins=@{}
    $actions=@('APPLY','CLEAR');$scopes=@('Process','Recipe')
    SelectTarget $Fixture
    if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.RegulationPolicyForTest' @($true))){throw 'Authorized regulation policy fixture unavailable; not product RED.'}
    try {
        SelectTarget $Fixture 'config-producer'
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$canary
        $path=Join-Path $runRoot 'regulation-operator.xlsb';$book.SaveAs($path,50);$book.Close($false)
        $bookPin=Hash $path;$book=$excel.Workbooks.Open($path,0,$false);$sheet=$book.Worksheets.Item(1)
        $decoy=$excel.Workbooks.Add();$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name))
        $ready=([string](Probe 'ReleasedProcess' @($canary))).Split('|')
        Check 'Regulation.RealReleasedProcessReady' ($ready.Count -eq 3 -and $ready[0] -ceq 'READY')
        if($ready.Count -ne 3 -or $ready[0] -cne 'READY'){throw 'Released regulation fixture unavailable; not product RED.'}
        $processId=$ready[1];$processVersion=$ready[2]
        foreach($target in @($Fixture,$Other)){
            foreach($file in Get-ChildItem -LiteralPath $target.Root -Recurse -File|Where-Object{$_.Extension -in '.xlsb','.xlsm' -and $_.Name -notlike '~$*'}){$pins[$file.FullName]=Hash $file.FullName}
        }
        foreach($file in Files){$recordPins[$file]=Hash $file}
        $old=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(19))).Split([char]10)|Where-Object{$_})
        $new=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(20))).Split([char]10)|Where-Object{$_})
        Check 'Regulation.Catalog20Extends19' ($old.Count -eq 103 -and $new.Count -eq 105 -and @($new|Sort-Object -Unique).Count -eq 105 -and @($old|Where-Object{$_ -cnotin $new}).Count -eq 0)
        foreach($id in $old){$before=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,19));Check ('Regulation.Catalog.Preserve.'+$id) ($before -ne '' -and $before -ceq [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,20)))}
        foreach($action in $actions){
            $control='PRODUCTION_OUTPUT_REGULATION_'+$action;$label='Regulation.'+$action
            $json=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($control,20));$definition=if($json){$json|ConvertFrom-Json}else{$null}
            Check ($label+'.FixedMetadata') ($null -ne $definition -and $definition.Caption -ceq [string](Probe 'RegulationCaption' @($action)) -and $definition.Surface -ceq 'Operations > Production > Production Settings' -and $definition.Role -ceq 'Production' -and $definition.Class -ceq 'Command' -and $definition.Capability -ceq 'PROD_POST')
            $absent=$true;foreach($version in 1..19){$absent=$absent -and [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($control,$version)) -ceq ''}
            Check ($label+'.AbsentFromOlderCatalogs') $absent
            foreach($unsupported in @('CONFIRMED','APPLIED','VALIDATED','COMPLETED','REFRESHED','PRESENTED')){Check ($label+'.RejectsUnsupported.'+$unsupported) ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Outcome' @($control,$unsupported)) -ceq '')}
            foreach($scope in $scopes){
                $prefix=$label+'.'+$scope
                $before=@(Files);Stage $scope
                Check ($prefix+'.SetupNotUserAction') (@(Files).Count -eq $before.Count)
                $notice=[string](Probe 'RegulationAct' @($action))
                $expected=if($action -ceq 'APPLY'){'True|2|8'}elseif($scope -ceq 'Recipe'){'NONE'}else{'False||'}
                Check ($prefix+'.SelectedLocalEffect') ((Value $scope) -ceq $expected)
                Check ($prefix+'.OtherOutputAndNodePreserved') ((Value 'Process' '' 'B02') -ceq 'True|5|9' -and (Value 'Recipe' 'NODE2') -ceq 'True|5|9' -and (Value 'Recipe' 'NODE1' 'B02') -ceq 'True|6|10')
                Check ($prefix+'.OtherScopePreserved') ((Value $(if($scope -ceq 'Recipe'){'Process'}else{'Recipe'})) -ceq 'True|1|4')
                Check ($prefix+'.NoUnhandledError') (-not $notice.StartsWith('HANDLER_ERROR|'))
                $aux=([string](Probe 'RegulationAux')).Split('|')
                Check ($prefix+'.ExistingRefreshAndEditorReset') ([int]$aux[0] -eq $(if($scope -ceq 'Recipe'){1}else{2}) -and $aux[2] -ceq 'False' -and $aux[3] -ceq '' -and $aux[4] -ceq '')
                if($action -ceq 'APPLY' -and $scope -ceq 'Recipe'){Check ($prefix+'.ExactReleasedProcessOutputBinding') ([string](Probe 'RegulationBinding') -ceq ($processId+'|'+$processVersion+'|NODE1|A01'))}
                Pair $before $action 'STAGED' ($prefix+'.Success')
                if($CaptureEvidence -and (($action -ceq 'APPLY' -and $scope -ceq 'Process') -or ($action -ceq 'CLEAR' -and $scope -ceq 'Recipe'))){[void](Probe 'RegulationShow');CaptureOwnedFormByCaptionEvidence 'Production' ('regulation-'+$scope.ToLowerInvariant()+'-'+$action.ToLowerInvariant()+'.png')}
                foreach($guard in @('Loading','Busy')){Stage $scope;$state=State;$before=@(Files);[void](Probe 'RegulationGuard' @($action,$guard));Check ($prefix+'.'+$guard+'Guard') ((State) -ceq $state -and @(Files).Count -eq $before.Count -and ([string](Probe 'RegulationAux')).Split('|')[1] -ceq '0')}
                Stage $scope;$before=@(Files);[void](Probe 'RegulationGuard' @($action,'Failure'));Pair $before $action 'FAILED' ($prefix+'.Failure')
                Stage $scope;$before=@(Files);$notice=[string](Probe 'RegulationGuard' @($action,'PartialFailure'))
                Check ($prefix+'.PartialEffectNotRolledBack') ((Value $scope) -ceq $expected)
                Check ($prefix+'.PartialFailureGuardsRestored') ([bool](Probe 'RegulationRestored'))
                Check ($prefix+'.PartialFailureHandled') (-not $notice.StartsWith('HANDLER_ERROR|'))
                Pair $before $action 'FAILED' ($prefix+'.PartialFailure')
                Stage $scope;$before=@(Files);$entered=[string](Probe 'RegulationGuard' @($action,'Nested'))
                Check ($prefix+'.NestedEntered') ($entered -ceq 'True')
                Check ($prefix+'.NestedMutationSuppressed') ((Value $scope) -ceq $expected)
                Pair $before $action 'STAGED' ($prefix+'.Nested')
                Stage $scope 'NoSelection';$state=State;$before=@(Files);$notice=[string](Probe 'RegulationAct' @($action))
                Check ($prefix+'.NoSelectionPreservesState') ((State) -ceq $state)
                Check ($prefix+'.NoSelectionExistingFeedback') ($notice -ceq $(if($action -ceq 'APPLY'){'Select an output to regulate.'}else{'Fixture prepared'}))
                Pair $before $action 'REJECTED' ($prefix+'.NoSelection')
            }
        }
        foreach($scope in $scopes){
            foreach($mode in @('Negative','Zero','Reversed','Equal','FractionEA','LowercaseEA','FractionLB','Blank','NonNumeric','DisabledBlank','DisabledText','DisabledNegative')){
                Stage $scope $mode;$state=State;$before=@(Files);$notice=[string](Probe 'RegulationAct' @('APPLY'))
                $outcome=if($mode -in @('Blank','NonNumeric')){'FAILED'}elseif($mode -in @('Negative','Zero','Reversed','FractionEA','LowercaseEA')){'REJECTED'}else{'STAGED'}
                if($outcome -cne 'STAGED'){$preserved=(State) -ceq $state}else{
                    $expected=switch($mode){'Equal'{'True|8|8'} 'FractionLB'{'True|1.5|2.5'} 'DisabledNegative'{'False|-2|-1'} 'DisabledText'{if($scope -ceq 'Recipe'){'False||'}else{'False|'+$canary+'|'+$canary}} default{'False||'}}
                    $preserved=(Value $scope) -ceq $expected
                }
                Check ('Regulation.Validation.'+$scope+'.'+$mode+'.PreservedLocalSemantics') $preserved
                Check ('Regulation.Validation.'+$scope+'.'+$mode+'.Handled') (-not $notice.StartsWith('HANDLER_ERROR|'))
                Pair $before 'APPLY' $outcome ('Regulation.Validation.'+$scope+'.'+$mode)
            }
            Stage $scope 'NoExisting';$before=@(Files);[void](Probe 'RegulationAct' @('CLEAR'))
            Check ('Regulation.ClearNoExisting.'+$scope+'.LocalEffect') ((Value $scope) -ceq $(if($scope -ceq 'Recipe'){'NONE'}else{'False||'}))
            Pair $before 'CLEAR' 'STAGED' ('Regulation.ClearNoExisting.'+$scope)
        }
        foreach($action in $actions){
            Stage 'Recipe' 'ReadUnavailable';$before=@(Files);[void](Probe 'RegulationAct' @($action))
            Check ('Regulation.UnavailableRead.'+$action+'.NoInventedSourceValidity') (([string](Probe 'RegulationAux')).StartsWith('0|1|'))
            Pair $before $action 'STAGED' ('Regulation.UnavailableRead.'+$action)
        }
        SelectTarget $Fixture 'config-reader';$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name))
        foreach($scope in $scopes){foreach($action in $actions){Stage $scope;$state=State;$before=@(Files);[void](Probe 'RegulationAct' @($action));Check ('Regulation.Denied.'+$scope+'.'+$action+'.NoEditOrRead') ((State) -ceq $state -and ([string](Probe 'RegulationAux')).Split('|')[1] -ceq '0');Pair $before $action 'DENIED' ('Regulation.Denied.'+$scope+'.'+$action) 'config-reader'}}
        Check 'Regulation.UnknownColumnsInMemory' ($sheet.Cells.Item(1,1).Value2 -ceq 'Operator Extra' -and $sheet.Cells.Item(2,1).Value2 -ceq $canary)
        foreach($enabled in @($false,$true)){
            SelectTarget $Fixture
            if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.RegulationPolicyForTest' @($enabled))){throw 'Authorized regulation policy update unavailable; not product RED.'}
            $pins[$Fixture.Config]=Hash $Fixture.Config
            if(-not $enabled){
                SelectTarget $Fixture 'config-producer';$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name))
                foreach($scope in $scopes){foreach($action in $actions){Stage $scope;$state=State;$before=@(Files);[void](Probe 'RegulationAct' @($action));Check ('Regulation.DisabledTracking.'+$scope+'.'+$action) ((State) -cne $state -and @(Files).Count -eq $before.Count)}}
            }
        }
        $parent=Join-Path $Fixture.Root 'Training\Activity';$blocked=Join-Path $parent $Fixture.Warehouse;$held=$blocked+'-regulation-held'
        foreach($item in @($blocked,$held)){if(-not [IO.Path]::GetFullPath($item).StartsWith([IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Tracking fault escaped disposable fixture.'}}
        if(Test-Path -LiteralPath $held){throw 'Preserve existing held fixture.'}
        SelectTarget $Fixture 'config-producer';$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name));$moved=$false
        try{
            if(Test-Path -LiteralPath $blocked){Move-Item -LiteralPath $blocked -Destination $held;$moved=$true}
            [IO.File]::WriteAllText($blocked,'Blocked disposable regulation activity path')
            foreach($scope in $scopes){foreach($action in $actions){Stage $scope;$state=State;$before=@(Files);$notice=[string](Probe 'RegulationAct' @($action));Check ('Regulation.Unavailable.'+$scope+'.'+$action+'.EditContinues') ((State) -cne $state);Check ('Regulation.Unavailable.'+$scope+'.'+$action+'.VisibleNoFallback') ($notice.Contains('Tracking unavailable') -and @(Files).Count -eq $before.Count)}}
        }finally{if(Test-Path -LiteralPath $blocked -PathType Leaf){Remove-Item -LiteralPath $blocked};if($moved){Move-Item -LiteralPath $held -Destination $blocked}}
        foreach($guard in @('Target','Session','SignedOut','ClosedWorkbook')){
            SelectTarget $Fixture 'config-producer';$decoy.Activate();[void](Probe 'OpenDesigner' @($book.Name))
            if($guard -ceq 'Target'){SelectTarget $Other 'config-producer'}
            if($guard -ceq 'Session'){SelectTarget $Fixture 'config-producer'}
            if($guard -ceq 'SignedOut'){[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
            if($guard -ceq 'ClosedWorkbook'){$book.Close($false);$book=$null}
            $before=@(Files);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            foreach($scope in $scopes){foreach($action in $actions){Stage $scope;$state=State;$notice=[string](Probe 'RegulationAct' @($action));Check ('Regulation.Guard.'+$guard+'.'+$scope+'.'+$action+'.NoEditOrRead') ((State) -ceq $state -and ([string](Probe 'RegulationAux')).Split('|')[1] -ceq '0');Check ('Regulation.Guard.'+$guard+'.'+$scope+'.'+$action+'.RefusalVisible') ($notice.StartsWith('Session, warehouse, or captured workbook changed.') -and $notice.Contains('Reopen Production'))}}
            Check ('Regulation.Guard.'+$guard+'.NoRedirectedActivity') (@(Files).Count -eq $before.Count -and @(Get-Slice4beActivityFiles $Other).Count -eq $otherBefore.Count)
        }
        Check 'Regulation.WorkbookBytesAndUnknownColumnsPreserved' ((Hash $path) -ceq $bookPin)
        $same=$true;foreach($file in $pins.Keys){$same=$same -and (Hash $file) -ceq $pins[$file]}
        Check 'Regulation.SavedAuthorityPreserved' $same
        $same=$true;foreach($file in $recordPins.Keys){$same=$same -and (Hash $file) -ceq $recordPins[$file]}
        Check 'Regulation.ExistingActivityImmutable' $same
    }finally{[void](Probe 'CloseDesigner');if($null -ne $book){$book.Close($false)};if($null -ne $decoy){$decoy.Close($false)};SelectTarget $Fixture}
}
