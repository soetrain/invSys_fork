# Worksheet-specific actions supplement the shared independent-guide UI gate.
# Values are confined to disposable fixtures; output contains Boolean evidence.
function Get-WorksheetPathInstruction([string]$Action) {
    switch($Action){
        'SEND' {'In Process Designer, prepare a new draft and click Send Process to Sheet. Inspect the new local table; Designs has not been saved.'}
        'ADD_ITEM' {'Select the first Process table and click Add Acceptable Item. Inspect the additional numbered item and SKU columns.'}
        'RETRIEVE' {'Fill both Process tables with valid definitions, Ctrl+select both and click Retrieve Selected Process. Verify both DRAFT imports and selected-table removal; inspect published source evidence separately.'}
    }
}
function Get-WorksheetPathTables($Book) {
    foreach($sheet in $Book.Worksheets){foreach($table in $sheet.ListObjects){if($table.Name -like 'invSys_Process_*'){$table}}}
}
function Invoke-WorksheetPathAction([string]$Action,$Book,$Decoy,[string]$Canary) {
    if($Action -ceq 'SEND'){
        [void](Probe 'WorksheetActivityStage' @($Canary));$Decoy.Activate()
    }else{
        $last=if($Action -ceq 'RETRIEVE'){2}else{1}
        if(-not [bool](Probe 'WorksheetActivitySelect' @($Book.Name,1,$last,($Action -ceq 'RETRIEVE')))){throw 'Worksheet path selection fixture unavailable.'}
    }
    [void](Probe 'WorksheetActivityAct' @($Action))
}
function Test-WorksheetPathAction([string]$Action,$Book,$Decoy,$Terminal,$Attempt,$Fixture,[string]$Label,[int]$Ordinal) {
    $prefix='WorksheetPaths.'+$Label+'.'+$Ordinal+'.'+$Action
    $tables=@(Get-WorksheetPathTables $Book)
    $expected=if($Action -ceq 'RETRIEVE'){0}elseif($Ordinal -eq 3){2}else{1}
    Check ($prefix+'.CapturedTablesSaved') ($tables.Count -eq $expected -and $Book.Saved -and $Decoy.Worksheets.Count -eq 1 -and $Decoy.Worksheets.Item(1).ListObjects.Count -eq 0)
    if($Action -ceq 'ADD_ITEM'){
        $headers=@($tables[0].ListColumns|ForEach-Object Name)
        Check ($prefix+'.NumberedPairAdded') ('Acceptable Managed Item 5' -cin $headers -and 'Accepted SKU 5' -cin $headers)
    }
    $facts=@(([string](Probe 'WorksheetSubmissionFactsForTest')).Split([char]10)|Where-Object{$_})
    $refs=@($Terminal.SourceEventRefs);$count=if($Action -ceq 'RETRIEVE'){2}else{0}
    $exact=$facts.Count -eq $count -and $refs.Count -eq $count -and @($Attempt.SourceEventRefs).Count -eq 0
    $exact=$exact -and ($facts -join "`n") -ceq (@($refs|ForEach-Object{$_.EventId+'|'+$_.SubmissionState}) -join "`n")
    foreach($ref in $refs){$exact=$exact -and $ref.SourceKind -ceq 'Designs' -and $ref.WarehouseId -ceq $Fixture.Warehouse -and $ref.SubmissionState -ceq 'Submitted' -and @($ref.PSObject.Properties).Count -eq 4}
    Check ($prefix+'.ExactOwningSubmissionReferences') $exact
    Check ($prefix+'.ExactRequestedOccurrence') ($Attempt.SequenceId -ceq $Terminal.SequenceId -and $Attempt.Ordinal -eq $Ordinal -and $Attempt.ControlId -ceq $Terminal.ControlId)
}
function Get-WorksheetPathBusinessState($Fixture) {
    $pins=@(foreach($file in Get-ChildItem $Fixture.Root -File|Where-Object{$_.Name -like '*.Auth.xlsb' -or $_.Name -like '*.Config.xlsb'}){[pscustomobject]@{Name=$file.Name;Hash=(SavedHash $file.FullName)}})
    if($pins.Count -ne 2){throw 'Auth and Config fixture files required.'}
    $source=$null;$owned=$false;$inventoryPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    foreach($open in $excel.Workbooks){if([string]$open.FullName -ieq $inventoryPath){$source=$open;break}}
    if($null -eq $source){$source=$excel.Workbooks.Open($inventoryPath,0,$true);$owned=$true}
    try{
        $values=@(foreach($sheet in $source.Worksheets){foreach($table in $sheet.ListObjects){
            if([string]$table.Name -cin @('tblInventoryLog','tblAppliedEvents','tblInventoryEntities','tblSkuBalance','tblLocationBalance','tblSkuCatalog')){[pscustomobject]@{Name=[string]$table.Name;Values=$table.Range.Value2}}
        }})
        if($values.Count -ne 6){throw 'Six Inventory business tables required.'}
        # In-memory comparison only: never write row values or credential material.
        [ordered]@{Pins=@($pins|Sort-Object Name);Inventory=@($values|Sort-Object Name)}|ConvertTo-Json -Depth 10 -Compress
    }finally{if($owned){$source.Close($false)}}
}
function Test-WorksheetPathPublishedSources($Source,$Observed,$Publication,$Fixture) {
    foreach($run in @(@{Name='GuideSource';Value=$Source},@{Name='ObservedRun';Value=$Observed})){
        $refs=@($run.Value.Terminals[-1].SourceEventRefs)
        $exact=$refs.Count -eq 2
        foreach($ref in $refs){
            $group=@($Publication.Groups|Where-Object{$_.Source -ceq 'Designs' -and $_.SourceId -ceq $ref.EventId})
            $exact=$exact -and $group.Count -eq 1
            if($group.Count -eq 1){
                $exact=$exact -and @($group[0].Lines).Count -gt 0 -and @($group[0].Outcomes).Count -eq @($group[0].Lines).Count
                foreach($line in $group[0].Lines){$exact=$exact -and $line.EventID -ceq $ref.EventId -and $line.WarehouseId -ceq $Fixture.Warehouse -and $line.AppliedAtUTC -cne '' -and $line.AppliedSeq -cmatch '^[1-9][0-9]*$'}
            }
        }
        Check ('WorksheetPaths.'+$run.Name+'.BothExactPublishedDesignsEvents') $exact
    }
    Check 'WorksheetPaths.IndependentOwningEventIdentities' (@($Source.Terminals[-1].SourceEventRefs|Where-Object{$_.EventId -cin $Observed.Terminals[-1].SourceEventRefs.EventId}).Count -eq 0)
}
function Test-WorksheetPathLocalTerminals($Observed) {
    foreach($index in @(0,1)){foreach($kind in @('CommandCompleted','SourceEventsApplied')){
        $record=$Observed.Terminals[$index];$label='WorksheetPaths.LocalTerminal.'+$record.ControlId+'.'+$kind
        $steps=@(,@($record.ControlId,'STAGED','True'))
        $ready=Set-EvaluationDraft $steps 0 $kind -StopAtMissingChoice
        Check ($label+'.ActualEditor') $ready
        if(-not $ready){continue}
        $before=@(EvaluationFiles|ForEach-Object FullName);Delivered (ExpectationControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths')
        $fresh=@(EvaluationFiles|Where-Object{$_.FullName -cnotin $before});if($fresh.Count -ne 1){throw 'Local worksheet evaluation unavailable.'}
        $result=Get-Content $fresh[0].FullName -Raw|ConvertFrom-Json
        $valid=if($kind -ceq 'CommandCompleted'){$result.ResultState -ceq 'Concluded' -and 'COMMAND_COMPLETED' -cin @($result.ReasonCodes)}else{$result.ResultState -ceq 'Incomplete' -and 'SOURCE_UNAVAILABLE' -cin @($result.ReasonCodes)}
        Check ($label+'.ExactLocalOnlyConclusion') ($valid -and $result.JournalRecordId -ceq $Observed.Journal.RecordId -and $result.Matches[-1].ActivityId -ceq $record.ActivityId -and @($result.TerminalSources).Count -eq 0)
    }}
}
function Test-WorksheetPathTerminalSources($Result,$Terminal,$Publication) {
    . (Join-Path $PSScriptRoot 'Slice4beEvaluationEvidence.ps1')
    $refs=@($Result.TerminalSources);$expected=@($Terminal.SourceEventRefs)
    $exact=$refs.Count -eq 2 -and ($refs.EventId -join '|') -ceq ($expected.EventId -join '|')
    foreach($ref in $refs){
        $group=@($Publication.Groups|Where-Object{$_.Source -ceq 'Designs' -and $_.SourceId -ceq $ref.EventId})
        $exact=$exact -and $group.Count -eq 1 -and $ref.SourceKind -ceq 'Designs' -and $ref.WarehouseId -ceq $Terminal.WarehouseId -and $ref.SubmissionState -ceq 'Submitted' -and $ref.OwnerStatus -ceq 'Applied' -and @($ref.SystemKeys).Count -eq 0
        if($group.Count -eq 1){
            $body=[ordered]@{Lines=@($group[0].Lines)}|ConvertTo-Json -Depth 16 -Compress
            $exact=$exact -and $ref.LineCount -eq @($group[0].Lines).Count -and $ref.LinesSha256 -ceq (EvaluationSha $body)
        }
    }
    Check 'WorksheetPaths.BothExactAppliedSourceProofs' $exact
}
