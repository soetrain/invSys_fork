# D13 General Settings observations and captured-context safety.
function Test-GeneralSettings($Fixture,$Other) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beAdminUomProbe.ps1')
    $configBytes=[IO.File]::ReadAllBytes($Fixture.Config)
    $originalHash=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $otherHash=(Get-FileHash -LiteralPath $Other.Config).Hash
    $canary='GENERALTEST'+[guid]::NewGuid().ToString('N').Substring(0,12).ToUpperInvariant()
    $ids=@('ADMIN_SETTINGS_RELOAD','ADMIN_SETTINGS_SELECT_CONFIG','ADMIN_CARRIER_ADD','ADMIN_CARRIER_REMOVE','ADMIN_CARRIER_RESET','ADMIN_CARRIER_SELECT','ADMIN_UOM_SELECT','ADMIN_CONNECTION_SELECT','ADMIN_CONNECTION_SAVE')
    $ordinal=0;$sequence='';$anchor=$null
    function Act([string]$Action,[string]$Value=''){[string](Run 'invSys.Admin.xlam' 'TestD5Commands.GeneralSettingsAction' @($Action,$Value))}
    function Carriers {[string](Run 'invSys.Core.xlam' 'modCarrierSettings.GetConfiguredCarriersText')}
    function Connection {[bool](Run 'invSys.Core.xlam' 'modNasConnection.RequireManualServerCredentials')}
    function OpenGeneral {
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.ShowSettings')
    }
    function Pair([string[]]$Before,[string]$Id,[string]$Outcome,[string]$Name) {
        $files=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $Before})
        $records=@($files|ForEach-Object{[IO.File]::ReadAllText($_)|ConvertFrom-Json})
        $attempt=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$end=@($records|Where-Object OutcomeCode -CEQ $Outcome)
        $pair=$records.Count -eq 2 -and $attempt.Count -eq 1 -and $end.Count -eq 1
        if($pair){$pair=$attempt[0].ActivityId -ceq $end[0].ActivityId -and $attempt[0].RecordId -cne $end[0].RecordId}
        Check ('GeneralSettings.'+$Name+'.ExactlyOneAttemptAndOutcome') $pair
        $owner=if($Id -ceq 'ADMIN_SETTINGS_RELOAD'){'CORE_CONFIGURATION'}elseif($Id -cin @('ADMIN_CARRIER_ADD','ADMIN_CARRIER_REMOVE','ADMIN_CARRIER_RESET','ADMIN_CONNECTION_SAVE')){'CORE_LOCAL_SETTINGS'}else{'ADMIN_SETTINGS_UI'}
        $metadata=$pair;$integrity=$pair;$redacted=$pair
        foreach($record in $records){
            $metadata=$metadata -and $record.ControlId -ceq $Id -and $record.OwnerId -ceq $owner -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.UserId -ceq 'config-admin' -and $record.StationId -ceq 'S1' -and $record.SequenceId -ceq $sequence -and $record.Ordinal -eq $ordinal -and $record.CatalogVersion -eq 31 -and @($record.SourceEventRefs).Count -eq 0 -and $record.EventCode -ceq ($Id+'_'+$record.OutcomeCode)
        }
        if($pair){$metadata=$metadata -and $attempt[0].DataEffect -ceq 'Unknown' -and $end[0].DataEffect -ceq $(if($Outcome -ceq 'COMPLETED'){'Changed'}else{'Unchanged'})}
        Check ('GeneralSettings.'+$Name+'.ExactOwnerContextSequenceAndEffect') $metadata
        foreach($file in $files){
            $raw=[IO.File]::ReadAllText($file)
            foreach($value in @($canary,$Fixture.Secret,$Fixture.Root,'"BatchSize"','"WarehouseId":"'+$Other.Warehouse+'"','mBtn','mLst','PinHash')){if($raw.Contains($value)){$redacted=$false}}
            $match=[regex]::Match($raw,'^(?<body>\{.*),"ContentSha256":"(?<hash>[a-f0-9]{64})"\}$')
            if(-not $match.Success){$integrity=$false;continue}
            $sha=[Security.Cryptography.SHA256]::Create()
            try{$hash=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($match.Groups['body'].Value+'}'))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
            $integrity=$integrity -and $hash -ceq $match.Groups['hash'].Value
        }
        Check ('GeneralSettings.'+$Name+'.RedactedPayload') $redacted
        Check ('GeneralSettings.'+$Name+'.ImmutableIntegrity') $integrity
    }
    function Step([string]$Name,[string]$Id,[string]$Outcome,[string]$Action,[string]$Value,[scriptblock]$Owner) {
        $before=@(Get-Slice4beActivityFiles $Fixture)
        if($Action -ceq 'ResetYes'){$null=Invoke-AdminUomResetChoice 'Yes' ($Name.ToLowerInvariant()+'.png') -Carrier}
        elseif($Action -ceq 'ResetNo'){$null=Invoke-AdminUomResetChoice 'No' ($Name.ToLowerInvariant()+'.png') -Carrier}
        else{$null=Act $Action $Value}
        $valid=[bool](& $Owner)
        Check ('GeneralSettings.'+$Name+'.ActualOwnerResult') $valid
        if(-not $valid){throw 'Existing General Settings owner prerequisite failed; inspect before claiming observation RED.'}
        Pair $before $Id $Outcome $Name
        Check ('GeneralSettings.'+$Name+'.WarehouseConfigPreserved') ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $enabledHash -and (Get-FileHash -LiteralPath $Other.Config).Hash -ceq $otherHash)
    }
    try {
        $excel.Visible=$true;$anchor=$excel.Workbooks.Add();$anchor.Activate()
        CloseRecordingViewer;[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings');SelectTarget $Fixture 'config-admin'
        $old=[string](Run 'invSys.Core.xlam' 'TestGeneralSettings.Definitions' @(30))
        $new=[string](Run 'invSys.Core.xlam' 'TestGeneralSettings.Definitions' @(31))
        $oldRows=@($old -split "`n"|Where-Object {$_});$newRows=@($new -split "`n"|Where-Object {$_})
        Check 'GeneralSettings.Catalog.ExtendsThirtyByNine' ($oldRows.Count -eq 135 -and $newRows.Count -eq 144)
        Check 'GeneralSettings.Catalog.PriorDefinitionsPreserved' ($new.StartsWith($old,[StringComparison]::Ordinal))
        foreach($id in $ids){
            $raw=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,31))
            $definition=if($raw){$raw|ConvertFrom-Json}else{$null}
            Check ('GeneralSettings.Catalog.'+$id) ($null -ne $definition -and $definition.ControlId -ceq $id -and $definition.Role -ceq 'Admin')
            Check ('GeneralSettings.Catalog.OlderExcludes.'+$id) ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,30)) -ceq '')
        }
        $prior=ActivityPins
        if(-not [bool](Run 'invSys.Core.xlam' 'TestSettingsOwnerState.SetupRecordingPolicy')){throw 'Recording policy fixture unavailable.'}
        $enabledHash=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        [void](Run 'invSys.Core.xlam' 'modCarrierSettings.ResetConfiguredCarriers')
        [void](Run 'invSys.Core.xlam' 'modNasConnection.SetRequireManualServerCredentials' @($false))
        OpenRecordingViewer;OpenGeneral
        Check 'GeneralSettings.InitializationAndDirectSetupDoNotInventActions' ((ActivityPins).Count -eq $prior.Count -and (PinsRetained $prior))
        $startsBefore=@(if(Test-Path -LiteralPath $journalRoot){Get-ChildItem -LiteralPath $journalRoot -File|ForEach-Object FullName})
        Check 'GeneralSettings.ActualStartRecording' ((RecordingControl 'Start Recording' 'Click') -ceq 'DELIVERED')
        $starts=@(Get-ChildItem -LiteralPath $journalRoot -File -Filter '*.json'|Where-Object {$_.FullName -cnotin $startsBefore}|ForEach-Object{Get-Content -LiteralPath $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
        if($starts.Count -ne 1){throw 'One actual recording Start required.'};$sequence=$starts[0].SequenceId
        $ordinal++;Step 'SelectConfig' 'ADMIN_SETTINGS_SELECT_CONFIG' 'SELECTED' 'SelectConfig' 'BatchSize' { (Act 'ConfigKey') -ceq 'BatchSize' }
        $ordinal++;Step 'Reload' 'ADMIN_SETTINGS_RELOAD' 'REFRESHED' 'Reload' '' { (Act 'ConfigKey') -ceq '' -and [int](Act 'Rows') -gt 0 }
        $ordinal++;Step 'SelectUom' 'ADMIN_UOM_SELECT' 'SELECTED' 'SelectUom' 'EA' { (Act 'UomDraft') -ceq 'EA' }
        $ordinal++;Step 'CarrierAdd' 'ADMIN_CARRIER_ADD' 'COMPLETED' 'CarrierAdd' $canary { $canary -cin ((Carriers) -split '\r?\n') }
        $ordinal++;Step 'CarrierDuplicate' 'ADMIN_CARRIER_ADD' 'UNCHANGED' 'CarrierAdd' $canary { @((Carriers) -split '\r?\n'|Where-Object {$_ -ceq $canary}).Count -eq 1 }
        $ordinal++;Step 'CarrierEmpty' 'ADMIN_CARRIER_ADD' 'REJECTED' 'CarrierAdd' '' { $canary -cin ((Carriers) -split '\r?\n') }
        $ordinal++;Step 'SelectCarrier' 'ADMIN_CARRIER_SELECT' 'SELECTED' 'SelectCarrier' $canary { (Act 'CarrierDraft') -ceq $canary }
        $ordinal++;Step 'CarrierRemove' 'ADMIN_CARRIER_REMOVE' 'COMPLETED' 'CarrierRemove' $canary { $canary -cnotin ((Carriers) -split '\r?\n') }
        $ordinal++;Step 'CarrierNoSelection' 'ADMIN_CARRIER_REMOVE' 'REJECTED' 'CarrierRemove' '' { $canary -cnotin ((Carriers) -split '\r?\n') }
        $ordinal++;Step 'ConnectionChoice' 'ADMIN_CONNECTION_SELECT' 'STAGED' 'ConnectionChoice' 'True' { (Act 'ConnectionDraft') -ceq 'True' -and -not (Connection) }
        $ordinal++;Step 'ConnectionSave' 'ADMIN_CONNECTION_SAVE' 'COMPLETED' 'ConnectionSave' '' { Connection }
        $ordinal++;Step 'ConnectionUnchanged' 'ADMIN_CONNECTION_SAVE' 'UNCHANGED' 'ConnectionSave' '' { Connection }
        $ordinal++;Step 'CarrierBeforeReset' 'ADMIN_CARRIER_ADD' 'COMPLETED' 'CarrierAdd' $canary { $canary -cin ((Carriers) -split '\r?\n') }
        $ordinal++;Step 'ResetNo' 'ADMIN_CARRIER_RESET' 'CANCELLED' 'ResetNo' '' { $canary -cin ((Carriers) -split '\r?\n') }
        $ordinal++;Step 'ResetYes' 'ADMIN_CARRIER_RESET' 'COMPLETED' 'ResetYes' '' { (Carriers) -ceq "UPS`r`nUSPS`r`nFedEx`r`nDHL" }
        $ordinal++;Step 'ResetUnchanged' 'ADMIN_CARRIER_RESET' 'UNCHANGED' 'ResetYes' '' { (Carriers) -ceq "UPS`r`nUSPS`r`nFedEx`r`nDHL" }
        $before=ActivityPins;$null=Act 'Render'
        Check 'GeneralSettings.ProgrammaticRenderDoesNotInventActions' ((ActivityPins).Count -eq $before.Count -and (PinsRetained $before))
        CaptureOwnedFormByCaptionEvidence 'invSys Settings' 'general-settings.png'
        Check 'GeneralSettings.ActualStopRecording' ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')
        $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close')
        Check 'GeneralSettings.SixteenActionsCompleteJournal' ($closed.Count -eq 1 -and $closed[0].Lifecycle -ceq 'Stopped' -and $closed[0].ActionCount -eq 16 -and @($closed[0].Observations).Count -eq 32 -and (JournalChain $sequence 34))
        foreach($loss in @('SignedOut','Reauthenticated','OtherWarehouse')){
            foreach($action in @('CarrierAdd','ConnectionSave','Reload','SelectConfig')){
                CloseRecordingViewer;SelectTarget $Fixture 'config-admin'
                [void](Run 'invSys.Core.xlam' 'modCarrierSettings.ResetConfiguredCarriers')
                [void](Run 'invSys.Core.xlam' 'modNasConnection.SetRequireManualServerCredentials' @($false))
                OpenGeneral;$null=Act 'SelectConfig' 'BatchSize';$null=Act 'ConnectionChoice' 'True'
                if($loss -ceq 'OtherWarehouse'){SelectTarget $Other 'config-admin'}else{
                    [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
                    if($loss -ceq 'Reauthenticated'){SelectTarget $Fixture 'config-admin'}
                }
                $before=ActivityPins;$carriersBefore=Carriers;$connectionBefore=Connection;$keyBefore=Act 'ConfigKey'
                $value=if($action -ceq 'CarrierAdd'){$canary}else{'WarehouseId'}
                $null=Act $action $value
                Check ('GeneralSettings.'+$loss+'.'+$action+'.CapturedStatePreserved') ((Carriers) -ceq $carriersBefore -and (Connection) -eq $connectionBefore -and (Act 'ConfigKey') -ceq $keyBefore)
                Check ('GeneralSettings.'+$loss+'.'+$action+'.NoRetargetedActivity') ((ActivityPins).Count -eq $before.Count -and (PinsRetained $before))
            }
        }
        Check 'GeneralSettings.PriorActivityImmutable' (PinsRetained $prior)
        Check 'GeneralSettings.OtherWarehouseConfigPreserved' ((Get-FileHash -LiteralPath $Other.Config).Hash -ceq $otherHash)
    } finally {
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings');CloseRecordingViewer
        if($null -ne $anchor){$anchor.Close($false)}
        [IO.File]::WriteAllBytes($Fixture.Config,$configBytes);SelectTarget $Fixture 'config-admin'
    }
    Check 'GeneralSettings.OriginalConfigRestored' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $originalHash)
}
