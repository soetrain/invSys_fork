# D18 designer-draft observations. Only synthetic fixture values stay in memory.
function Test-ProductionDesignerActivity($Fixture,$Other) {
    function Probe([string]$Action,[object[]]$Values=@()) {
        $method=switch($Action){'Open'{'OpenDesigner'} 'Close'{'CloseDesigner'} default{$Action}}
        Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$method) $Values
    }
    function Files { @(Get-Slice4beActivityFiles $Fixture) }
    function Pair([string[]]$Before,[string]$Id,[string]$Outcome,[string]$Case,[string]$Actor='config-producer',$ObservedFixture=$Fixture) {
        $raw=@(Get-Slice4beActivityFiles $ObservedFixture|Where-Object {$_ -cnotin $Before}|ForEach-Object {[IO.File]::ReadAllText($_)})
        $rows=@($raw|ForEach-Object {$_|ConvertFrom-Json})
        $first=@($rows|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($rows|Where-Object OutcomeCode -CEQ $Outcome)
        $paired=$rows.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
        Check ($Case+'.AttemptAndOutcome') $paired
        $context=$paired;$safe=$paired;$integrity=$paired
        foreach($r in $rows){
            $context=$context -and $r.ControlId -ceq $Id -and $r.OwnerId -ceq 'PRODUCTION_DESIGNER' -and $r.UserId -ceq $Actor -and $r.WarehouseId -ceq $ObservedFixture.Warehouse -and $r.StationId -ceq 'S1'
            $safe=$safe -and @($r.SourceEventRefs).Count -eq 0
        }
        foreach($text in $raw){
            foreach($forbidden in @($canary,$ObservedFixture.Secret,(CredentialHash $ObservedFixture.Secret),$ObservedFixture.Root,'mBtn','PinHash')){if($text.Contains($forbidden)){$safe=$false}}
            $match=[regex]::Match($text,',"ContentSha256":"([a-f0-9]{64})"\}$')
            if(-not $match.Success){$integrity=$false;continue}
            $sha=[Security.Cryptography.SHA256]::Create()
            try{$digest=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($text.Substring(0,$match.Index)+'}'))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
            $integrity=$integrity -and $digest -ceq $match.Groups[1].Value
        }
        Check ($Case+'.ExactContext') $context
        Check ($Case+'.NoEnteredDataOrSources') $safe
        Check ($Case+'.Integrity') $integrity
        $linked=$false;$effect=$false
        if($paired){
            $linked=$first[0].ActivityId -cne '' -and $first[0].ActivityId -ceq $last[0].ActivityId -and $first[0].RecordId -cne $last[0].RecordId
            $severity=switch($Outcome){'REJECTED'{'Warning'} 'DENIED'{'Blocked'} default{'Info'}}
            $effect=$first[0].DataEffect -ceq 'Unknown' -and $last[0].DataEffect -ceq 'Unchanged' -and $last[0].EventCode -ceq ($Id+'_'+$Outcome) -and $last[0].Severity -ceq $severity -and $last[0].CatalogVersion -eq 13
        }
        Check ($Case+'.LinkedDistinctRecords') $linked
        Check ($Case+'.LocalOnlyOutcome') $effect
    }
    $canary='DESIGNER'+[guid]::NewGuid().ToString('N')
    $book=$null;$decoy=$null
    Write-Output 'Production designer: saved-authority snapshot.'
    $pins=@{};foreach($f in @(Get-ChildItem -LiteralPath $Fixture.Root -File -Recurse|Where-Object {$_.Extension -in '.xlsb','.xlsm' -and -not $_.Name.StartsWith('~$')})){$pins[$f.FullName]=(Get-FileHash -LiteralPath $f.FullName).Hash}
    Write-Output 'Production designer: captured-form fixture.'
    try {
        SelectTarget $Fixture 'config-producer'
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$canary
        $path=Join-Path $runRoot 'designer-operator.xlsb';$book.SaveAs($path,50)
        $book.Close($false)
        $workbookPin=(Get-FileHash -LiteralPath $path).Hash
        $book=$excel.Workbooks.Open($path,0,$false)
        $sheet=$book.Worksheets.Item(1)
        $decoy=$excel.Workbooks.Add()
        $before=@(Files);[void](Probe 'Open' @($book.Name))
        Check 'ProductionDesigner.Initialize.NotUserAction' (@(Files).Count -eq $before.Count)
        foreach($designer in @('Process','Recipe')){
            foreach($action in @('New','Clear','Validate')){
                $id='PRODUCTION_'+$designer.ToUpperInvariant()+'_'+$action.ToUpperInvariant()
                $case='ProductionDesigner.'+$designer+'.'+$action
                $definition=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,12))
                $metadata=if($definition){$definition|ConvertFrom-Json}else{$null}
                $caption=[string](Probe 'Caption' @($designer,$action))
                Check ($case+'.ExactVisibleMetadata') ($null -ne $metadata -and $metadata.Caption -ceq $caption -and $metadata.Surface -ceq ('Operations > Production > '+$designer+' Designer') -and $metadata.Class -ceq 'Command' -and $metadata.Role -ceq 'Production' -and $metadata.Capability -ceq 'PROD_POST')
                [void](Probe 'Stage' @($designer,$canary));$before=@(Files);$decoy.Activate()
                $report=[string](Probe 'Act' @($designer,$action))
                $state=[string](Probe 'State' @($designer))
                $outcome=if($action -ceq 'Validate'){'REJECTED'}else{'STAGED'}
                Check ($case+'.ActualLocalResult') $(if($action -ceq 'Validate'){$state.Contains($canary) -and $report -match 'must declare|must select'}else{-not $state.Contains($canary) -and $report -match 'started|cleared'})
                Pair $before $id $outcome $case
            }
        }
        [void](Probe 'ValidProcess');$before=@(Files)
        $report=[string](Probe 'Act' @('Process','Validate'))
        Check 'ProductionDesigner.Process.Valid.ActualValidator' ($report -like 'Process draft is valid*')
        Pair $before 'PRODUCTION_PROCESS_VALIDATE' 'VALIDATED' 'ProductionDesigner.Process.Valid'

        $old=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(11))).Split("`n")|Where-Object {$_})
        $new=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(12))).Split("`n")|Where-Object {$_})
        Check 'ProductionDesigner.Catalog.TwelveExtendsEleven' ($old.Count -eq 62 -and $new.Count -eq 68 -and @($new|Select-Object -Unique).Count -eq 68 -and @($old|Where-Object {$_ -cnotin $new}).Count -eq 0)
        foreach($id in $old){Check ('ProductionDesigner.Catalog.Preserve.'+$id) ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,11)) -ceq [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,12)))}

        SelectTarget $Fixture 'config-reader';[void](Probe 'Open' @($book.Name))
        foreach($designer in @('Process','Recipe')){foreach($action in @('New','Clear','Validate')){
            [void](Probe 'Stage' @($designer,$canary));$state=[string](Probe 'State' @($designer));$before=@(Files)
            $report=[string](Probe 'Act' @($designer,$action))
            $case='ProductionDesigner.Denied.'+$designer+'.'+$action
            Check ($case+'.DraftUnchanged') ($state -ceq [string](Probe 'State' @($designer)) -and $report -match 'permission|authorized|PROD_POST')
            Pair $before ('PRODUCTION_'+$designer.ToUpperInvariant()+'_'+$action.ToUpperInvariant()) 'DENIED' $case 'config-reader'
        }}
        Check 'ProductionDesigner.UnknownWorkbookColumnRemainsInMemory' ($sheet.Cells.Item(1,1).Value2 -ceq 'Operator Extra' -and $sheet.Cells.Item(2,1).Value2 -ceq $canary)
        $same=$true;foreach($file in $pins.Keys){$same=$same -and (Get-FileHash -LiteralPath $file).Hash -ceq $pins[$file]}
        Check 'ProductionDesigner.SavedAuthorityBeforePolicyCommand' $same
        SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.DesignerDisableTrackingForTest')){throw 'Explicit disabled-tracking fixture policy unavailable; not product RED.'}
        # Only the deliberate fixture policy command may change this Config pin.
        $pins[$Fixture.Config]=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        SelectTarget $Fixture 'config-producer';[void](Probe 'Open' @($book.Name))
        foreach($designer in @('Process','Recipe')){foreach($action in @('New','Clear','Validate')){
            [void](Probe 'Stage' @($designer,$canary));$before=@(Files)
            $report=[string](Probe 'Act' @($designer,$action));$state=[string](Probe 'State' @($designer))
            $worked=if($action -ceq 'Validate'){$state.Contains($canary) -and $report -match 'must declare|must select'}else{-not $state.Contains($canary) -and $report -match 'started|cleared'}
            Check ('ProductionDesigner.DisabledTracking.'+$designer+'.'+$action) ($worked -and @(Files).Count -eq $before.Count)
        }}

        foreach($change in @('Target','Session','ClosedWorkbook')){
            SelectTarget $Fixture 'config-producer';[void](Probe 'Open' @($book.Name))
            if($change -ceq 'Target'){SelectTarget $Other 'config-producer'}
            if($change -ceq 'Session'){SelectTarget $Fixture 'config-producer'}
            if($change -ceq 'ClosedWorkbook'){$book.Close($false);$book=$null}
            $before=@(Files);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            foreach($designer in @('Process','Recipe')){foreach($action in @('New','Clear','Validate')){
                [void](Probe 'Stage' @($designer,$canary));$state=[string](Probe 'State' @($designer))
                $report=[string](Probe 'Act' @($designer,$action))
                Check ('ProductionDesigner.Guard.'+$change+'.'+$designer+'.'+$action) ($state -ceq [string](Probe 'State' @($designer)) -and $report -match 'Reopen|reopen')
            }}
            Check ('ProductionDesigner.Guard.'+$change+'.NoRedirectedActivity') (@(Files).Count -eq $before.Count -and @(Get-Slice4beActivityFiles $Other).Count -eq $otherBefore.Count)
        }
        Check 'ProductionDesigner.UnknownWorkbookColumnAndSavedBytesPreserved' ((Get-FileHash -LiteralPath $path).Hash -ceq $workbookPin)
        $same=$true;foreach($file in $pins.Keys){$same=$same -and (Get-FileHash -LiteralPath $file).Hash -ceq $pins[$file]}
        Check 'ProductionDesigner.SavedAuthorityPreserved' $same
        . (Join-Path $PSScriptRoot 'Slice4beProductionRecipeActivity.ps1')
        Test-ProductionRecipeActivity $Other $canary
    } finally {
        [void](Probe 'Close')
        if($null -ne $book){$book.Close($false)}
        if($null -ne $decoy){$decoy.Close($false)}
        SelectTarget $Fixture
    }
}
