# D18: observe actual Print handlers; never infer a printed page from preview return.
function Test-ProductionPrintActivity($Fixture,$Book,$Sheet,$Decoy){
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    function Fingerprint($Worksheet){ConvertTo-Json -Compress -Depth 10 -InputObject @($Worksheet.UsedRange.Formula)}
    function Pins {
        $pins=@{}
        foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File|Where-Object Name -match '\.invSys\.Data\.(Inventory|Designs)\.'){
            $stream=[IO.File]::Open($file.FullName,'Open','Read','ReadWrite')
            try{$pins[$file.FullName]=(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}
        }
        if(@($pins.Keys|Where-Object {$_ -match '\.invSys\.Data\.Inventory\.'}).Count -ne 1){throw 'Canonical Inventory fixture missing; not product RED.'}
        return $pins
    }
    function Preserved($Before){$after=Pins;if($after.Count -ne $Before.Count){return $false};foreach($path in $Before.Keys){if($after[$path] -cne $Before[$path]){return $false}};return $true}
    Test-ProductionPrintCatalog
    $id='PRODUCTION_RUN_PRINT';$owner='PRODUCTION_RECALL_REPORT'
    $catalog=[int](Run 'invSys.Core.xlam' 'TestShippingCatalog.DeclaredCatalogVersionForTest')
    $activities=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $recall=$Sheet.ListObjects.Item('ProductionOutput').ListColumns.Item('RECALL CODE').DataBodyRange.Cells.Item(1,1)
    $originalRecall=[string]$recall.Value2;$foreign=Fingerprint $Decoy.Worksheets.Item('Production')
    SelectTarget $Fixture
    if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadPolicyForTest' @($true))){throw 'Recording policy unavailable; not product RED.'}
    try{
        foreach($case in @('Normal','RefusalAfterPreview','PreviewFailure','Recovery','Nested','Loading','Busy','Reader','Admin','MissingSheet','SignedOut','ClosedForm','Disabled')){
            if($case -ceq 'Disabled'){
                SelectTarget $Fixture
                if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadPolicyForTest' @($false))){throw 'Disabled policy unavailable; not product RED.'}
            }
            $actor=switch($case){'Reader'{'config-reader'} 'Admin'{'config-admin'} default{'config-producer'}}
            SelectTarget $Fixture $actor
            $recall.Value2=if($case -ceq 'RefusalAfterPreview'){''}else{$originalRecall}
            [void](Probe 'OpenDesigner' @($Book.Name));[void](Probe 'RunLocalShowAndCapture' @($Book.Name,'PRINT'))
            [void](Probe 'ResetPrintPreviewForTest');[void](Probe 'PrintPreviewFailureForTest' @($case -ceq 'PreviewFailure'))
            if($case -ceq 'MissingSheet'){$Sheet.Name='PrintMissing'}
            if($case -ceq 'SignedOut'){[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
            [void](Probe 'PrintYieldArm' @('',''));[void](Probe 'PrintNativeArm' @($false,($case -ceq 'ClosedForm')))
            $source=Fingerprint $Sheet;$pins=Pins;$before=Files;$Decoy.Activate();$label='PrintActivity.'+$case
            $stop=Join-Path $runRoot ('print-activity-'+$case)
            $observer=Start-DialogCaptureAndDismiss -ExcelProcessId (@(Get-Process EXCEL)[0].Id) -TimeoutSeconds 30 -StopPath $stop
            try{
                $returned=if($case -ceq 'ClosedForm'){[bool](Probe 'PrintYieldAct')}else{[bool](Probe 'PrintEntryAct' @($case))}
                Check ($label+'.ActualHandlerReturned') $returned
            }finally{
                [IO.File]::WriteAllText($stop,'');Wait-Job $observer -Timeout 35|Out-Null
                Receive-Job $observer -ErrorAction SilentlyContinue|Out-Null
                if($observer.State -ne 'Completed'){Stop-Job $observer;Remove-Job $observer;throw 'Native observer failed; not product RED.'}
                Remove-Job $observer
            }
            if($case -ceq 'ClosedForm'){
                Check ($label+'.ActualCloseHandlerDismissedForm') ([bool](Probe 'PrintNativeFact' @('Closed')))
                Check ($label+'.NoReinitialization') ([bool](Probe 'PrintYieldFact' @('NoReinitialization')))
                Check ($label+'.RemainsUnloaded') (-not [bool](Probe 'PrintYieldFact' @('Loaded')))
            }else{Check ($label+'.GuardsRestored') ([bool](Probe 'PrintEntryFact' @('GuardsRestored')))}
            $suppressed=$case -cin @('Loading','Busy','Reader','MissingSheet','SignedOut')
            $entries=if($suppressed){0}else{1}
            Check ($label+'.OwnerEntries') ([int](Probe 'PrintOwnerEntries') -eq $entries)
            Check ($label+'.ReportReads') ([int](Probe 'PrintReportReads') -eq $entries)
            $previews=if($suppressed -or $case -ceq 'RefusalAfterPreview'){0}else{1}
            Check ($label+'.PreviewEntries') ([int](Probe 'PrintPreviewCountForTest') -eq $previews)
            $outcome=switch($case){'Reader'{'DENIED'} 'MissingSheet'{'FAILED'} 'ClosedForm'{'FAILED'} 'PreviewFailure'{'FAILED'} 'RefusalAfterPreview'{'REJECTED'} default{'PREVIEW_RETURNED'}}
            $expected=switch($case){
                'Reader'{'Production permission changed. Reopen Production before continuing.'}
                'MissingSheet'{'Production sheet not found.'}
                'SignedOut'{'Session, warehouse, or captured workbook changed. Reopen Production before editing the draft.'}
                'Loading'{'PRINT-ACTIVE'} 'Busy'{'PRINT-ACTIVE'}
                'PreviewFailure'{'BTN_PRINT_CODES failed: Print preview unavailable (fixture).'}
                'RefusalAfterPreview'{'No recall-coded ProductionOutput rows found. Generate recall codes from checked output rows before printing.'}
                default{'Print preview closed.'}
            }
            if($case -cne 'ClosedForm'){Check ($label+'.IndependentOwnerStatus') ([string](Probe 'PrintStatus') -ceq $expected)}
            $raw=@(Files|Where-Object{$_ -cnotin $before}|ForEach-Object{[IO.File]::ReadAllText($_)})
            # Actual Close may also record its own action; retain only Print's pair.
            $records=@($raw|ForEach-Object{$_|ConvertFrom-Json}|Where-Object ControlId -CEQ $id)
            if($case -cin @('Loading','Busy','SignedOut','Disabled')){
                Check ($label+'.NoRecords') ($records.Count -eq 0)
            }else{
                $first=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($records|Where-Object OutcomeCode -CEQ $outcome)
                $pair=$records.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
                $context=$pair;$safe=$pair;$integrity=$pair;$linked=$false;$facts=$false;$terminal=$false
                foreach($r in $records){
                    $context=$context -and $r.ControlId -ceq $id -and $r.OwnerId -ceq $owner -and $r.UserId -ceq $actor -and $r.WarehouseId -ceq $Fixture.Warehouse -and $r.StationId -ceq 'S1' -and $r.CatalogVersion -eq $catalog
                    $safe=$safe -and @($r.SourceEventRefs).Count -eq 0
                }
                foreach($value in $raw){
                    foreach($secret in @($originalRecall,$Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'mBtn','PinHash','Print preview unavailable (fixture).')){
                        $encoded=ConvertTo-Json -InputObject $secret -Compress
                        if($value.Contains($secret) -or $value.Contains($encoded.Substring(1,$encoded.Length-2))){$safe=$false}
                    }
                    $match=[regex]::Match($value,',"ContentSha256":"([a-f0-9]{64})"\}$')
                    if(-not $match.Success){$integrity=$false;continue}
                    $sha=[Security.Cryptography.SHA256]::Create()
                    try{$hash=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($value.Substring(0,$match.Index)+'}'))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
                    $integrity=$integrity -and $hash -ceq $match.Groups[1].Value
                }
                if($pair){
                    $linked=$first[0].ActivityId -cne '' -and $first[0].ActivityId -ceq $last[0].ActivityId -and $first[0].RecordId -cne $last[0].RecordId -and $activities.Add([string]$first[0].ActivityId)
                    $severity=switch($outcome){'REJECTED'{'Warning'} 'DENIED'{'Blocked'} 'FAILED'{'Error'} default{'Info'}}
                    $effect=if($outcome -ceq 'FAILED'){'Unknown'}else{'Unchanged'}
                    $facts=$first[0].DataEffect -ceq 'Unknown' -and $last[0].Severity -ceq $severity -and $last[0].DataEffect -ceq $effect -and $last[0].EventCode -ceq ($id+'_'+$outcome)
                    $terminal=-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @(($first[0]|ConvertTo-Json -Depth 20 -Compress))) -and [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @(($last[0]|ConvertTo-Json -Depth 20 -Compress))) -eq ($outcome -ceq 'PREVIEW_RETURNED')
                }
                Check ($label+'.AttemptAndOutcome') $pair
                Check ($label+'.ExactContext') $context
                Check ($label+'.RedactedAndNoSources') $safe
                Check ($label+'.Integrity') $integrity
                Check ($label+'.DistinctLinkedAttempt') $linked
                Check ($label+'.OwnerOutcomeFacts') $facts
                Check ($label+'.ExactCommandTerminal') $terminal
            }
            Check ($label+'.CanonicalBytesPreserved') (Preserved $pins)
            Check ($label+'.SourceKeysAndCustomValuesPreserved') ((Fingerprint $Sheet) -ceq $source)
            Check ($label+'.DecoyPreserved') ((Fingerprint $Decoy.Worksheets.Item('Production')) -ceq $foreign)
            if($case -cin @('Normal','PreviewFailure')){CaptureOwnedFormByCaptionEvidence 'Production' ('print-activity-'+$case.ToLowerInvariant()+'.png')}
            $Sheet.Name='Production';[void](Probe 'PrintNativeArm' @($false,$false));[void](Probe 'RunLocalSafeClose')
        }
    }finally{
        $Sheet.Name='Production';$recall.Value2=$originalRecall
        [void](Probe 'PrintPreviewFailureForTest' @($false));[void](Probe 'PrintNativeArm' @($false,$false));[void](Probe 'RunLocalSafeClose')
        SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadPolicyForTest' @($false))){throw 'Disabled policy cleanup failed.'}
    }
}

function Test-ProductionPrintCatalog {
    $id='PRODUCTION_RUN_PRINT';$owner='PRODUCTION_RECALL_REPORT'
    $old=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(28))).Split("`n")|Where-Object {$_})
    $new=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(29))).Split("`n")|Where-Object {$_})
    Check 'PrintCatalog.Extends28Exactly' ($old.Count -eq 132 -and $new.Count -eq 133 -and @($new|Select-Object -Unique).Count -eq 133 -and ($new[0..131] -join '|') -ceq ($old -join '|') -and $new[-1] -ceq $id)
    $same=$true
    foreach($prior in $old){$before=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($prior,28));$after=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($prior,29));$same=$same -and $before -cne '' -and $after -ceq $before}
    Check 'PrintCatalog.All132PriorDefinitionsPreserved' $same
    $excluded=$true
    foreach($version in 1..28){$excluded=$excluded -and [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,$version)) -ceq ''}
    Check 'PrintCatalog.ExcludedFromHistoricalVersions' $excluded
    $wire=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,29));$record=if($wire){$wire|ConvertFrom-Json}else{$null}
    Check 'PrintCatalog.ExactControlContract' ($null -ne $record -and $record.ControlId -ceq $id -and $record.OwnerId -ceq $owner -and $record.Class -ceq 'Command' -and $record.Role -ceq 'Production' -and $record.Caption -ceq 'Print Recall' -and $record.Surface -ceq 'Operations > Production > Production Run - List' -and $record.Capability -ceq 'PROD_POST')
    foreach($code in @('REQUESTED','DENIED','REJECTED','FAILED','PREVIEW_RETURNED','PRINTED','COMPLETED','APPLIED')){
        $json=@{ControlId=$id;OwnerId=$owner;CatalogVersion=29;OutcomeCode=$code}|ConvertTo-Json -Compress
        Check ('PrintCatalog.Terminal.'+$code) ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @($json)) -eq ($code -ceq 'PREVIEW_RETURNED'))
        $wire=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Outcome' @($id,$code));$record=if($wire){$wire|ConvertFrom-Json}else{$null}
        $supported=$code -cin @('REQUESTED','DENIED','REJECTED','FAILED','PREVIEW_RETURNED');$valid=$null -eq $record
        if($supported){
            $severity=switch($code){'DENIED'{'Blocked'} 'FAILED'{'Error'} 'REJECTED'{'Warning'} default{'Info'}}
            $effect=if($code -cin @('REQUESTED','FAILED')){'Unknown'}else{'Unchanged'}
            $valid=$null -ne $record -and $record.EventCode -ceq ($id+'_'+$code) -and $record.Severity -ceq $severity -and $record.DataEffect -ceq $effect
        }
        Check ('PrintCatalog.Outcome.'+$code) $valid
        Check ('PrintCatalog.EmptyReferences.'+$code) ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,$code,'[]')) -eq $supported)
        $source='[{"WarehouseId":"CATALOG_TEST","SourceKind":"Inventory","EventId":"Source_A","SubmissionState":"Submitted"}]'
        Check ('PrintCatalog.RejectsSource.'+$code) (-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,$code,$source)))
    }
}
