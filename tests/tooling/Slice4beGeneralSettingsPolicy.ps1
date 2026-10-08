# Faults are confined to Admin-generated fixtures; command handlers remain real.
function Test-GeneralSettingsPolicy($Fixture,$Other) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beAdminUomProbe.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beTrackingPolicy.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beUserTrackingPolicy.ps1')
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    $root=[IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\'
    if(-not $root.StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Generated General Settings fixture required.'}
    $leaf=Join-Path $Fixture.Root ('Training/Activity/'+$Fixture.Warehouse)
    $held=$leaf+'-general-policy-held'
    $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    foreach($path in @($leaf,$held,$Fixture.Config,$authPath)){
        if(-not [IO.Path]::GetFullPath($path).StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw 'General policy path escaped generated root.'}
    }
    if(Test-Path -LiteralPath $held){throw 'Preserve existing held activity.'}
    $configBytes=[IO.File]::ReadAllBytes($Fixture.Config);$authBytes=[IO.File]::ReadAllBytes($authPath)
    $originalHash=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $authHash=(Get-FileHash -LiteralPath $authPath).Hash;$otherHash=(Get-FileHash -LiteralPath $Other.Config).Hash
    $ids=@('ADMIN_SETTINGS_RELOAD','ADMIN_SETTINGS_SELECT_CONFIG','ADMIN_CARRIER_ADD','ADMIN_CARRIER_REMOVE','ADMIN_CARRIER_RESET','ADMIN_CARRIER_SELECT','ADMIN_UOM_SELECT','ADMIN_CONNECTION_SELECT','ADMIN_CONNECTION_SAVE')
    $canary='GENERALPOLICY'+[guid]::NewGuid().ToString('N').Substring(0,10).ToUpperInvariant()
    $anchor=$null;$book=$null;$caseStage='CreateSavedHost';$anchorPath=Join-Path $runRoot 'general-policy-host.xlsx';$anchorHash=''
    if(Test-Path -LiteralPath $anchorPath){throw 'Preserve existing General policy host.'}
    function Act([string]$Action,[string]$Value=''){[string](Run 'invSys.Admin.xlam' 'TestD5Commands.GeneralSettingsAction' @($Action,$Value))}
    function Carriers {[string](Run 'invSys.Core.xlam' 'modCarrierSettings.GetConfiguredCarriersText')}
    function Connection {[bool](Run 'invSys.Core.xlam' 'modNasConnection.RequireManualServerCredentials')}
    function CloseGeneral {[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')}
    function OpenGeneral {
        CloseGeneral
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.ShowSettings')
    }
    function ConfigurePolicy([string]$Mode) {
        CloseGeneral;CloseRecordingViewer;SelectTarget $Fixture 'config-admin'
        $previous=Get-TrackingPolicyVersion $Fixture
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
        try {
            if((UserPolicyControl 'btnResetTrackingPolicy' 'Click') -cne 'CLICKED'){throw 'Policy Reset fixture unavailable.'}
            foreach($id in $ids){[void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyControl' @($id,($Mode -cne 'ControlOff')))}
            if($Mode -ceq 'UserOff'){
                if((UserPolicyControl 'lstTrackingUsers' 'Select' 'config-admin') -cne 'SELECTED'){throw 'Generated policy actor unavailable.'}
                if((UserPolicyControl 'chkUserRecord' 'Write' 'False') -cne 'SET'){throw 'User recording choice unavailable.'}
            }
            $before=[int](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyEntries')
            $saved=(UserPolicyControl 'btnSaveTrackingPolicy' 'Click') -ceq 'CLICKED'
            $saved=$saved -and [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyShowsSaved') -and [int](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyEntries') -eq $before+1
            if(-not $saved){throw 'Actual policy save prerequisite failed; not product RED.'}
        }finally{CloseGeneral}
        $version=Get-TrackingPolicyVersion $Fixture
        return [pscustomobject]@{Version=$version;Previous=$previous}
    }
    function Records([string]$Id,[string]$Outcome,[string[]]$Before,[bool]$Expected) {
        $files=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $Before})
        if(-not $Expected){return $files.Count -eq 0}
        $records=@($files|ForEach-Object{[IO.File]::ReadAllText($_)|ConvertFrom-Json})
        $first=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($records|Where-Object OutcomeCode -CEQ $Outcome)
        if($records.Count -ne 2 -or $first.Count -ne 1 -or $last.Count -ne 1){return $false}
        $owner=if($Id -ceq 'ADMIN_SETTINGS_RELOAD'){'CORE_CONFIGURATION'}elseif($Id -cin @('ADMIN_CARRIER_ADD','ADMIN_CARRIER_REMOVE','ADMIN_CARRIER_RESET','ADMIN_CONNECTION_SAVE')){'CORE_LOCAL_SETTINGS'}else{'ADMIN_SETTINGS_UI'}
        $valid=$first[0].ActivityId -ceq $last[0].ActivityId -and $first[0].RecordId -cne $last[0].RecordId
        $valid=$valid -and $first[0].DataEffect -ceq 'Unknown' -and $last[0].DataEffect -ceq $(if($Outcome -ceq 'COMPLETED'){'Changed'}else{'Unchanged'})
        foreach($record in $records){$valid=$valid -and $record.ControlId -ceq $Id -and $record.OwnerId -ceq $owner -and $record.UserId -ceq 'config-admin' -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.CatalogVersion -eq 31 -and @($record.SourceEventRefs).Count -eq 0}
        foreach($file in $files){$raw=[IO.File]::ReadAllText($file);foreach($value in @($canary,$Fixture.Secret,$Fixture.Root,'"BatchSize"','mBtn','mLst','PinHash')){$valid=$valid -and -not $raw.Contains($value)}}
        return $valid
    }
    function Step([string]$Name,[string]$Id,[string]$Outcome,[string]$Action,[string]$Value,[scriptblock]$Owner,[string]$Message) {
        $before=@();$pins=$history
        if(-not $blocked){$before=@(Get-Slice4beActivityFiles $Fixture);$pins=ActivityPins}
        if($Action -cin @('ResetNo','ResetYes')){
            $choice=if($Action -ceq 'ResetNo'){'No'}else{'Yes'}
            $status=Invoke-AdminUomResetChoice $choice ('general-policy-'+$mode.ToLowerInvariant()+'-'+$Name.ToLowerInvariant()+'.png') -Carrier
        }else{$status=Act $Action $Value}
        $prefix='GeneralPolicy.'+$mode+'.'+$Name
        Check ($prefix+'.IndependentOwnerResult') ([bool](& $Owner))
        Check ($prefix+'.OwnerMessagePreserved') ($status.StartsWith($Message,[StringComparison]::Ordinal))
        Check ($prefix+'.TrackingNotice') ($status.Contains('Tracking unavailable') -eq ($mode -cin @('Older','Invalid','StoreUnavailable')))
        $evidence=if($blocked){(Test-Path -LiteralPath $leaf -PathType Leaf) -and [IO.File]::ReadAllText($leaf) -ceq 'Generated General Settings activity unavailable.'}else{Records $Id $Outcome $before ($mode -ceq 'Recovery')}
        Check ($prefix+'.ExactEvidenceOrSuppression') $evidence
        $retained=$true
        if($blocked){
            foreach($name in $pins.Keys){
                $priorFile=Join-Path $held $name
                if(-not (Test-Path -LiteralPath $priorFile -PathType Leaf) -or (Get-FileHash -LiteralPath $priorFile).Hash -cne $pins[$name]){$retained=$false}
            }
        }else{$retained=PinsRetained $pins}
        Check ($prefix+'.PriorEvidencePreserved') $retained
        Check ($prefix+'.ConfigAndOtherWarehousePreserved') ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configuredHash -and (Get-FileHash -LiteralPath $Other.Config).Hash -ceq $otherHash)
    }
    try {
        $anchor=$excel.Workbooks.Add();$anchor.SaveAs($anchorPath,51)
        $anchor.Close($false);$anchor=$null
        $anchorHash=(Get-FileHash -LiteralPath $anchorPath).Hash
        $anchor=$excel.Workbooks.Open($anchorPath,0,$false)
        $modes=@('ControlOff','UserOff','Older','Invalid','StoreUnavailable','Recovery')
        if($GeneralSettingsPolicyOnly){$modes=@('StoreUnavailable','Recovery','Older','Invalid','ControlOff','UserOff')}
        foreach($mode in $modes){
            $blocked=$false;$moved=$false;$book=$null;$caseStage='ConfigurePolicy'
            try {
                $saved=ConfigurePolicy $mode;$version=[int]$saved.Version
                Check ('GeneralPolicy.'+$mode+'.ActualPolicySaveAppendsVersion') ($version -eq $saved.Previous+1)
                if($mode -cin @('Older','Invalid')){
                    $caseStage='OpenGeneratedConfig'
                    $book=$excel.Workbooks.Open($Fixture.Config,0,$false)
                    $caseStage='FindPolicyTables'
                    $headers=Table $book 'tblEventTrackingPolicies';$controls=Table $book 'tblEventTrackingControls';$found=0
                    $caseStage='ChangePolicyHeader'
                    foreach($row in $headers.ListRows){if([int]$row.Range.Cells.Item(1,$headers.ListColumns.Item('PolicyVersion').Index).Value2 -eq $version){
                        $column=if($mode -ceq 'Older'){'CatalogVersion'}else{'SchemaVersion'}
                        $caseStage='ResolvePolicyColumn'
                        $columnIndex=$headers.ListColumns.Item($column).Index
                        $caseStage='ResolvePolicyCell'
                        $policyCell=$row.Range.Cells.Item(1,$columnIndex)
                        $caseStage='WritePolicyCell'
                        if($mode -ceq 'Older'){$policyCell.Value2=30.0}else{$policyCell.Value2=999.0}
                        $found++
                    }}
                    if($found -ne 1){throw 'Unique generated policy required.'}
                    $caseStage='RemoveNewControlRows'
                    if($mode -ceq 'Older'){for($index=$controls.ListRows.Count;$index -ge 1;$index--){
                        $row=$controls.ListRows.Item($index)
                        if([int]$row.Range.Cells.Item(1,$controls.ListColumns.Item('PolicyVersion').Index).Value2 -eq $version -and [string]$row.Range.Cells.Item(1,$controls.ListColumns.Item('ControlId').Index).Value2 -cin $ids){$row.Delete()}
                    }}
                    $caseStage='SaveGeneratedConfig'
                    $book.Save();$book.Close($false);$book=$null
                }
                $caseStage='VerifyEffectivePolicy'
                foreach($id in $ids){
                    $actual=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @($id))
                    $expected=if($mode -cin @('Older','Invalid')){'False|False|False|0'}elseif($mode -cin @('ControlOff','UserOff')){'True|False|'+$(if($mode -ceq 'ControlOff'){'False'}else{'True'})+'|'+$version}else{'True|True|True|'+$version}
                    Check ('GeneralPolicy.'+$mode+'.EffectivePolicy.'+$id) ($actual -ceq $expected)
                    if($actual -cne $expected){
                        if($actual -cmatch '^(True|False)\|(True|False)\|(True|False)\|[0-9]+$'){
                            [pscustomobject]@{Mode=$mode;Control=$id;Expected=$expected;Actual=$actual}|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'general-policy-fixture-facts.json')
                        }
                        throw 'Effective General policy fixture mismatch; not command RED.'
                    }
                }
                $configuredHash=(Get-FileHash -LiteralPath $Fixture.Config).Hash
                $history=ActivityPins
                if($mode -ceq 'StoreUnavailable'){
                    if(Test-Path -LiteralPath $leaf){Move-Item -LiteralPath $leaf -Destination $held;$moved=$true}
                    [IO.File]::WriteAllText($leaf,'Generated General Settings activity unavailable.');$blocked=$true
                }
                [void](Run 'invSys.Core.xlam' 'modCarrierSettings.ResetConfiguredCarriers')
                [void](Run 'invSys.Core.xlam' 'modNasConnection.SetRequireManualServerCredentials' @($false))
                OpenGeneral
                $caseStage='ActualGeneralHandlers'
                Step 'SelectConfig' 'ADMIN_SETTINGS_SELECT_CONFIG' 'SELECTED' 'SelectConfig' 'BatchSize' {(Act 'ConfigKey') -ceq 'BatchSize'} 'Edit the selected value'
                Step 'Reload' 'ADMIN_SETTINGS_RELOAD' 'REFRESHED' 'Reload' '' {(Act 'ConfigKey') -ceq '' -and [int](Act 'Rows') -gt 0} 'Canonical config reloaded.'
                Step 'SelectUom' 'ADMIN_UOM_SELECT' 'SELECTED' 'SelectUom' 'EA' {(Act 'UomDraft') -ceq 'EA'} ''
                Step 'Add' 'ADMIN_CARRIER_ADD' 'COMPLETED' 'CarrierAdd' $canary {$canary -cin ((Carriers) -split '\r?\n')} 'Carrier added.'
                Step 'SelectCarrier' 'ADMIN_CARRIER_SELECT' 'SELECTED' 'SelectCarrier' $canary {(Act 'CarrierDraft') -ceq $canary} ''
                Step 'Remove' 'ADMIN_CARRIER_REMOVE' 'COMPLETED' 'CarrierRemove' $canary {$canary -cnotin ((Carriers) -split '\r?\n')} 'Carrier removed.'
                [void](Run 'invSys.Core.xlam' 'modCarrierSettings.AddConfiguredCarrier' @($canary));$null=Act 'Render'
                Step 'ResetNo' 'ADMIN_CARRIER_RESET' 'CANCELLED' 'ResetNo' '' {$canary -cin ((Carriers) -split '\r?\n')} 'Carrier reset cancelled.'
                Step 'ResetYes' 'ADMIN_CARRIER_RESET' 'COMPLETED' 'ResetYes' '' {(Carriers) -ceq "UPS`r`nUSPS`r`nFedEx`r`nDHL"} 'Defaults restored.'
                Step 'ConnectionChoice' 'ADMIN_CONNECTION_SELECT' 'STAGED' 'ConnectionChoice' 'True' {(Act 'ConnectionDraft') -ceq 'True' -and -not(Connection)} 'Connection option staged.'
                Step 'ConnectionSave' 'ADMIN_CONNECTION_SAVE' 'COMPLETED' 'ConnectionSave' '' {Connection} 'Server connection option saved'
                if($mode -cin @('StoreUnavailable','Recovery')){CaptureOwnedFormByCaptionEvidence 'invSys Settings' ('general-policy-'+$mode.ToLowerInvariant()+'.png')}
                CloseGeneral
            }catch{
                [pscustomobject]@{Mode=$mode;Stage=$caseStage;HResult=$_.Exception.HResult}|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'general-policy-failure-facts.json')
                throw
            }finally{
                CloseGeneral
                if($null -ne $book){$book.Close($false)}
                if($blocked){Remove-Item -LiteralPath $leaf}
                if($moved){Move-Item -LiteralPath $held -Destination $leaf}
                [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
                [void](Run 'invSys.Core.xlam' 'modConfig.Reload')
            }
            Check ('GeneralPolicy.'+$mode+'.PriorActivityRestoredOrRetained') (PinsRetained $history)
            Check ('GeneralPolicy.'+$mode+'.OriginalConfigRestored') ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $originalHash)
        }
        $saved=ConfigurePolicy 'PermissionLoss'
        Check 'GeneralPolicy.PermissionLoss.ActualPolicySaveAppendsVersion' ($saved.Version -eq $saved.Previous+1)
        [void](Run 'invSys.Core.xlam' 'modCarrierSettings.ResetConfiguredCarriers')
        [void](Run 'invSys.Core.xlam' 'modNasConnection.SetRequireManualServerCredentials' @($false))
        OpenGeneral;$null=Act 'SelectConfig' 'BatchSize';$null=Act 'ConnectionChoice' 'True'
        $context=[string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
        $book=$excel.Workbooks.Open($authPath,0,$false)
        try {
            $caps=Table $book 'tblCapabilities';$changed=0
            foreach($row in $caps.ListRows){if($row.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-admin' -and $row.Range.Cells.Item(1,$caps.ListColumns.Item('Capability').Index).Value2 -ceq 'ADMIN_MAINT'){$row.Range.Cells.Item(1,$caps.ListColumns.Item('Status').Index).Value2='Inactive';$changed++}}
            $book.Save()
        }finally{$book.Close($false);$book=$null}
        [void](Run 'invSys.Core.xlam' 'modAuth.ReloadAuth' @($Fixture.Warehouse))
        $lost=$changed -gt 0 -and -not[bool](Run 'invSys.Core.xlam' 'modRoleUiAccess.CanCurrentUserPerformCapabilityCached' @('ADMIN_MAINT')) -and $context -cne '' -and [string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext') -ceq $context
        Check 'GeneralPolicy.PermissionLoss.SameSessionLosesOnlyCapability' $lost
        if(-not $lost){throw 'Permission-loss fixture invalid; not product RED.'}
        $configPin=(Get-FileHash -LiteralPath $Fixture.Config).Hash;$local=Carriers;$key=Act 'ConfigKey'
        foreach($case in @(@('Reload','ADMIN_SETTINGS_RELOAD'),@('CarrierAdd','ADMIN_CARRIER_ADD'),@('CarrierRemove','ADMIN_CARRIER_REMOVE'),@('CarrierReset','ADMIN_CARRIER_RESET'),@('ConnectionSave','ADMIN_CONNECTION_SAVE'))){
            $before=@(Get-Slice4beActivityFiles $Fixture);$pins=ActivityPins
            $status=Act $case[0] 'UPS'
            Check ('GeneralPolicy.PermissionLoss.'+$case[0]+'.IndependentStatePreserved') ((Carriers) -ceq $local -and -not(Connection) -and (Act 'ConfigKey') -ceq $key -and (Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configPin)
            Check ('GeneralPolicy.PermissionLoss.'+$case[0]+'.ExactDeniedPair') (Records $case[1] 'DENIED' $before $true)
            Check ('GeneralPolicy.PermissionLoss.'+$case[0]+'.VisibleDenied') ($status.Contains('ADMIN_MAINT') -and -not $status.Contains('Tracking unavailable'))
            Check ('GeneralPolicy.PermissionLoss.'+$case[0]+'.PriorEvidencePreserved') (PinsRetained $pins)
        }
        CaptureOwnedFormByCaptionEvidence 'invSys Settings' 'general-policy-permission-denied.png'
    }catch{
        [pscustomobject]@{Stage=$caseStage;HResult=$_.Exception.HResult;Line=$_.InvocationInfo.ScriptLineNumber}|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'general-policy-outer-failure.json')
        throw
    }finally{
        CloseGeneral;CloseRecordingViewer
        if($null -ne $book){$book.Close($false)}
        [IO.File]::WriteAllBytes($Fixture.Config,$configBytes);[IO.File]::WriteAllBytes($authPath,$authBytes)
        [void](Run 'invSys.Core.xlam' 'modAuth.ReloadAuth' @($Fixture.Warehouse))
        [void](Run 'invSys.Core.xlam' 'modConfig.Reload')
        if($null -ne $anchor){$anchor.Close($false)}
    }
    Check 'GeneralPolicy.Final.SavedHostPreserved' ((Get-FileHash -LiteralPath $anchorPath).Hash -ceq $anchorHash)
    Check 'GeneralPolicy.Final.ConfigAuthAndOtherWarehousePreserved' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $originalHash -and (Get-FileHash -LiteralPath $authPath).Hash -ceq $authHash -and (Get-FileHash -LiteralPath $Other.Config).Hash -ceq $otherHash)
}
