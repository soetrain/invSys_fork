# Independent actual recordings and authored intent; no synthetic success records.
function Test-GeneralSettingsRecorded($Fixture,$Other) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beEvaluationContracts.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beAdminUomProbe.ps1')
    function Hash([string]$Path){(Get-FileHash -LiteralPath $Path).Hash}
    function Retained($Before,$After){foreach($path in $Before.Keys){if(-not $After.ContainsKey($path) -or $After[$path] -cne $Before[$path]){return $false}};return $true}
    function Act([string]$Action,[string]$Value=''){[string](Run 'invSys.Admin.xlam' 'TestD5Commands.GeneralSettingsAction' @($Action,$Value))}
    function Carriers {[string](Run 'invSys.Core.xlam' 'modCarrierSettings.GetConfiguredCarriersText')}
    function Connection {[bool](Run 'invSys.Core.xlam' 'modNasConnection.RequireManualServerCredentials')}
    function CloseGeneral {[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')}
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    $root=[IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\'
    $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    if(-not $root.StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Generated recording fixture required.'}
    foreach($path in @($Fixture.Config,$authPath,$activityRoot,$journalRoot)){
        if(-not [IO.Path]::GetFullPath($path).StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw 'Recording path escaped generated fixture.'}
    }
    $configBytes=[IO.File]::ReadAllBytes($Fixture.Config);$configHash=Hash $Fixture.Config
    $authBytes=[IO.File]::ReadAllBytes($authPath);$authHash=Hash $authPath;$otherPins=BoundPins $Other.Root
    $hostPath=Join-Path $runRoot 'general-recorded-host.xlsb'
    if(Test-Path -LiteralPath $hostPath){throw 'Preserve existing recording host.'}
    $book=$null;$anchor=$null;$runs=@();$canary='RECORDGENERAL'+[guid]::NewGuid().ToString('N').Substring(0,10)
    $steps=@(
        @('Reload','ADMIN_SETTINGS_RELOAD','REFRESHED',''),
        @('SelectConfig','ADMIN_SETTINGS_SELECT_CONFIG','SELECTED','BatchSize'),
        @('SelectUom','ADMIN_UOM_SELECT','SELECTED','EA'),
        @('CarrierReset','ADMIN_CARRIER_RESET','COMPLETED',''),
        @('CarrierAdd','ADMIN_CARRIER_ADD','COMPLETED',$canary),
        @('SelectCarrier','ADMIN_CARRIER_SELECT','SELECTED',$canary),
        @('CarrierRemove','ADMIN_CARRIER_REMOVE','COMPLETED',$canary),
        @('ConnectionChoice','ADMIN_CONNECTION_SELECT','STAGED','True'),
        @('ConnectionSave','ADMIN_CONNECTION_SAVE','COMPLETED','')
    )
    try {
        $anchor=$excel.Workbooks.Add();$anchor.SaveAs($hostPath,50);$anchor.Close($false);$anchor=$null
        $hostHash=Hash $hostPath;$anchor=$excel.Workbooks.Open($hostPath,0,$false)
        SelectTarget $Fixture 'config-admin'
        if(-not [bool](Run 'invSys.Core.xlam' 'TestSettingsOwnerState.SetupRecordingPolicy')){throw 'Actual recording policy prerequisite unavailable.'}
        $enabledHash=Hash $Fixture.Config
        foreach($kind in @('Source','Observed','Interrupted','Denied')){
            $label='GeneralRecorded.'+$kind
            CloseGeneral;CloseRecordingViewer
            [IO.File]::WriteAllBytes($authPath,$authBytes);SelectTarget $Fixture 'config-admin'
            [void](Run 'invSys.Core.xlam' 'modCarrierSettings.ResetConfiguredCarriers')
            [void](Run 'invSys.Core.xlam' 'modCarrierSettings.AddConfiguredCarrier' @($canary))
            [void](Run 'invSys.Core.xlam' 'modNasConnection.SetRequireManualServerCredentials' @($false))
            OpenRecordingViewer
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.ShowSettings')
            if($kind -ceq 'Denied'){
                $context=[string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
                $book=$excel.Workbooks.Open($authPath,0,$false);$changed=0
                try {
                    $caps=Table $book 'tblCapabilities'
                    foreach($row in $caps.ListRows){if($row.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-admin' -and $row.Range.Cells.Item(1,$caps.ListColumns.Item('Capability').Index).Value2 -ceq 'ADMIN_MAINT'){
                        $row.Range.Cells.Item(1,$caps.ListColumns.Item('Status').Index).Value2='Inactive';$changed++
                    }}
                    $book.Save()
                }finally{$book.Close($false);$book=$null}
                [void](Run 'invSys.Core.xlam' 'modAuth.ReloadAuth' @($Fixture.Warehouse))
                $lost=$changed -eq 1 -and $context -cne '' -and [string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext') -ceq $context -and -not [bool](Run 'invSys.Core.xlam' 'modRoleUiAccess.CanCurrentUserPerformCapabilityCached' @('ADMIN_MAINT'))
                Check ($label+'.SameSessionPermissionLoss') $lost
                if(-not $lost){throw 'Permission-loss fixture unavailable; not product RED.'}
            }
            $prior=BoundPins $journalRoot;$activity=ActivityPins
            Check ($label+'.ActualStart') ((RecordingControl 'Start Recording' 'Click') -ceq 'DELIVERED')
            $starts=@(Get-ChildItem -LiteralPath $journalRoot -Filter '*.json' -File|Where-Object {-not $prior.ContainsKey($_.FullName)}|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
            if($starts.Count -ne 1){throw 'One durable Start required.'}
            $sequence=$starts[0].SequenceId;$ordinal=0;$all=@()
            foreach($step in $steps){
                $ordinal++;$action=$step[0];$id=$step[1];$expected=$step[2]
                $denied=$kind -ceq 'Denied' -and $action -cin @('Reload','CarrierReset','CarrierAdd','CarrierRemove','ConnectionSave')
                if($denied){$expected='DENIED'}
                $before=@(Get-Slice4beActivityFiles $Fixture);$pins=ActivityPins
                $local=Carriers;$connection=Connection;$key=Act 'ConfigKey'
                if($action -ceq 'CarrierReset' -and -not $denied){$status=Invoke-AdminUomResetChoice 'Yes' ('general-recorded-'+$kind.ToLowerInvariant()+'-reset.png') -Carrier}
                else{$status=Act $action $step[3]}
                $owner=if($denied){(Carriers) -ceq $local -and (Connection) -eq $connection -and (Act 'ConfigKey') -ceq $key -and $status.Contains('ADMIN_MAINT')}else{
                    switch($action){
                        Reload {(Act 'ConfigKey') -ceq '' -and [int](Act 'Rows') -gt 0}
                        SelectConfig {(Act 'ConfigKey') -ceq 'BatchSize'}
                        SelectUom {(Act 'UomDraft') -ceq 'EA'}
                        CarrierReset {(Carriers) -ceq "UPS`r`nUSPS`r`nFedEx`r`nDHL"}
                        CarrierAdd {$canary -cin ((Carriers) -split '\r?\n')}
                        SelectCarrier {(Act 'CarrierDraft') -ceq $canary}
                        CarrierRemove {$canary -cnotin ((Carriers) -split '\r?\n')}
                        ConnectionChoice {(Act 'ConnectionDraft') -ceq 'True' -and -not(Connection)}
                        ConnectionSave {Connection}
                    }
                }
                Check ($label+'.'+$action+'.IndependentOwnerState') ([bool]$owner)
                $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json})
                $first=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($records|Where-Object OutcomeCode -CEQ $expected)
                $valid=$records.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
                if($valid){$valid=$first[0].ActivityId -ceq $last[0].ActivityId -and $first[0].RecordId -cne $last[0].RecordId}
                $ownerId=if($action -ceq 'Reload'){'CORE_CONFIGURATION'}elseif($action -cin @('CarrierReset','CarrierAdd','CarrierRemove','ConnectionSave')){'CORE_LOCAL_SETTINGS'}else{'ADMIN_SETTINGS_UI'}
                foreach($record in $records){$valid=$valid -and $record.ControlId -ceq $id -and $record.OwnerId -ceq $ownerId -and $record.SequenceId -ceq $sequence -and $record.Ordinal -eq $ordinal -and $record.CatalogVersion -eq 31 -and @($record.SourceEventRefs).Count -eq 0}
                Check ($label+'.'+$action+'.ExactRecordedPairAndOwner') $valid
                $redacted=$true
                foreach($record in $records){$raw=ConvertTo-Json $record -Depth 20 -Compress;foreach($value in @($canary,$Fixture.Secret,$Fixture.Root,'"BatchSize"','PinHash')){$redacted=$redacted -and -not $raw.Contains($value)}}
                Check ($label+'.'+$action+'.NoInputOrCredentialPayload') $redacted
                Check ($label+'.'+$action+'.PriorEvidenceAndConfigPreserved') ((PinsRetained $pins) -and (Hash $Fixture.Config) -ceq $enabledHash)
                $all+=@($records)
                if($kind -ceq 'Interrupted'){[void](Run 'invSys.Core.xlam' 'modAuth.SignOut');break}
            }
            if($kind -cne 'Interrupted'){Check ($label+'.ActualStop') ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')}
            $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close')
            $count=if($kind -ceq 'Interrupted'){1}else{9}
            $valid=$closed.Count -eq 1
            if($valid){$valid=$closed[0].ActionCount -eq $count -and @($closed[0].Observations).Count -eq 2*$count -and $closed[0].Lifecycle -ceq $(if($kind -ceq 'Interrupted'){'Incomplete'}else{'Stopped'})}
            if($valid -and $kind -ceq 'Interrupted'){$valid=$closed[0].ReasonCode -ceq 'SESSION_CHANGED'}
            Check ($label+'.ExactJournalLifecycle') $valid
            Check ($label+'.ImmutableJournalChain') (JournalChain $sequence (2*$count+2))
            Check ($label+'.PriorActivityImmutable') (PinsRetained $activity)
            if($closed.Count -ne 1){throw 'Closed recording prerequisite missing.'}
            $runs+=@{Kind=$kind;Journal=$closed[0];Records=$all;Success=($kind -cin @('Source','Observed'));Incomplete=($kind -ceq 'Interrupted')}
        }
        CloseGeneral;CloseRecordingViewer
        [IO.File]::WriteAllBytes($authPath,$authBytes);SelectTarget $Fixture 'config-admin'
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Actual publication prerequisite failed.'}
        SelectTarget $Fixture 'config-reader';OpenRecordingViewer
        $training=BoundPins $journalRoot;$observed=@($runs|Where-Object Kind -CEQ 'Observed')[0]
        foreach($step in $steps){
            $id=$step[1]
            if((Select-EvaluationRun $observed.Journal.ActionPathId) -cne 'SELECTED'){throw 'Exact recorded run unavailable.'}
            $ready=Set-EvaluationDraft @(,@($id,$step[2],'True')) 0 'CommandCompleted' -StopAtMissingChoice
            Check ('GeneralRecorded.ExpectedConclusion.'+$id+'.ActualEditorChoice') $ready
            if(-not $ready){continue}
            $before=@(EvaluationFiles|ForEach-Object FullName)
            $clicked=ExpectationControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths'
            $fresh=@(EvaluationFiles|Where-Object {$_.FullName -cnotin $before})
            if($clicked -cne 'DELIVERED' -or $fresh.Count -ne 1){throw 'Evaluate must append exactly one result.'}
            $result=Get-Content $fresh[0].FullName -Raw|ConvertFrom-Json
            $match=@($observed.Records|Where-Object {$_.ControlId -ceq $id -and $_.OutcomeCode -ceq $step[2]})
            Check ('GeneralRecorded.ExpectedConclusion.'+$id+'.IndependentConcludedResult') ($result.ResultState -ceq 'Concluded' -and 'COMMAND_COMPLETED' -cin @($result.ReasonCodes) -and @($result.Matches).Count -eq 1 -and $result.Matches[0].ActivityId -ceq $match[0].ActivityId)
            Check ('GeneralRecorded.ExpectedConclusion.'+$id+'.ExactJournalNoDomainSources') ($result.JournalRecordId -ceq $observed.Journal.RecordId -and $result.JournalSha256 -ceq $observed.Journal.ContentSha256 -and @($result.TerminalSources).Count -eq 0)
        }
        . (Join-Path $PSScriptRoot 'Slice4beGeneralSettingsGuidePaths.ps1')
        Test-GeneralSettingsGuidePaths $Fixture $runs $steps
        Check 'GeneralRecorded.TrainingEvidenceImmutableThroughEvaluation' (Retained $training (BoundPins $journalRoot))
        Check 'GeneralRecorded.OtherWarehousePreserved' (BoundSame $otherPins (BoundPins $Other.Root))
    }finally{
        CloseGeneral;CloseRecordingViewer
        if($null -ne $book){$book.Close($false)}
        if($null -ne $anchor){$anchor.Close($false)}
        [IO.File]::WriteAllBytes($Fixture.Config,$configBytes);[IO.File]::WriteAllBytes($authPath,$authBytes)
        SelectTarget $Fixture 'config-admin'
    }
    Check 'GeneralRecorded.FinalAuthorityAndHostPreserved' ((Hash $Fixture.Config) -ceq $configHash -and (Hash $authPath) -ceq $authHash -and (Hash $hostPath) -ceq $hostHash)
}
