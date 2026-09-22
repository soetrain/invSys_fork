# D18 Admin UOM coverage through the actual packaged form handlers.
function Test-AdminUomActivity($Fixture,$Other) {
    $entryVisibility=[bool]$excel.Visible
    $configBytes=[IO.File]::ReadAllBytes($Fixture.Config)
    $configPin=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $otherBytes=[IO.File]::ReadAllBytes($Other.Config)
    $otherInitialPin=(Get-FileHash -LiteralPath $Other.Config).Hash
    $original=@{}; foreach($path in @(Get-Slice4beActivityFiles $Fixture)){$original[$path]=(Get-FileHash -LiteralPath $path).Hash}
    $canary='UOMTEST'+[guid]::NewGuid().ToString('N').Substring(0,12).ToUpperInvariant()
    function CurrentUoms { [string](Run 'invSys.Core.xlam' 'modUomSettings.GetConfiguredUomsPackedText') }
    function Act([string]$Action,[string]$Value='') { [string](Run 'invSys.Admin.xlam' 'TestD5Commands.UomActivityAction' @($Action,$Value)) }
    function OpenUoms([string]$Actor='config-admin') {
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
        SelectTarget $Fixture $Actor
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.ShowSettings')
    }
    function Pair([string[]]$Before,[string]$Control,[string]$Outcome,[string]$Case,[string]$Actor='config-admin') {
        $effect=if($Outcome -ceq 'COMPLETED'){'Changed'}elseif($Outcome -ceq 'FAILED'){'Unknown'}else{'Unchanged'}
        $severity=switch($Outcome){'DENIED'{'Blocked'} 'REJECTED'{'Warning'} 'CANCELLED'{'Notice'} 'FAILED'{'Error'} default{'Info'}}
        Test-Slice4beObservedAction $Fixture $Before $Control ($Control+'_REQUESTED') ($Control+'_'+$Outcome) $effect $severity $Actor ('AdminUom.'+$Case)
        $fresh=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $Before})
        $safe=$fresh.Count -eq 2
        foreach($path in $fresh){
            $raw=Get-Content -LiteralPath $path -Raw;$record=$raw|ConvertFrom-Json
            $safe=$safe -and @($record.SourceEventRefs).Count -eq 0 -and -not $raw.Contains($canary) -and -not $raw.Contains('mBtnUom')
        }
        Check ('AdminUom.'+$Case+'.NoEnteredValueOrInventoryReferences') $safe
    }
    try {
        $excel.Visible=$true
        OpenUoms
        $before=@(Get-Slice4beActivityFiles $Fixture)
        [void](Act 'Add' $canary)
        Check 'AdminUom.Add.ActualOwnerChangesCatalog' ($canary -cin (CurrentUoms).Split('|'))
        Pair $before 'ADMIN_UOM_ADD' 'COMPLETED' 'Add'

        $before=@(Get-Slice4beActivityFiles $Fixture);$pin=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        [void](Act 'Add' $canary)
        Check 'AdminUom.Duplicate.SuccessPreservesConfigBytes' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $pin)
        Pair $before 'ADMIN_UOM_ADD' 'UNCHANGED' 'Duplicate'

        $before=@(Get-Slice4beActivityFiles $Fixture)
        [void](Act 'Add' '')
        Check 'AdminUom.InvalidAdd.PreservesConfigBytes' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $pin)
        Pair $before 'ADMIN_UOM_ADD' 'REJECTED' 'InvalidAdd'

        $before=@(Get-Slice4beActivityFiles $Fixture)
        [void](Act 'Remove' $canary)
        Check 'AdminUom.Remove.ActualOwnerChangesCatalog' ($canary -cnotin (CurrentUoms).Split('|'))
        Pair $before 'ADMIN_UOM_REMOVE' 'COMPLETED' 'Remove'

        $before=@(Get-Slice4beActivityFiles $Fixture);$pin=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        [void](Act 'Remove' '')
        Check 'AdminUom.NoSelection.PreservesConfigBytes' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $pin)
        Pair $before 'ADMIN_UOM_REMOVE' 'REJECTED' 'NoSelection'

        OpenUoms 'config-reader'
        foreach($action in @('Add','Remove','Reset')){
            $before=@(Get-Slice4beActivityFiles $Fixture);$pin=(Get-FileHash -LiteralPath $Fixture.Config).Hash
            $value=if($action -ceq 'Remove'){(CurrentUoms).Split('|')[0]}else{$canary}
            [void](Act $action $value)
            Check ('AdminUom.Denied'+$action+'.PreservesConfigBytes') ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $pin)
            Pair $before ('ADMIN_UOM_'+$action.ToUpperInvariant()) 'DENIED' ('Denied'+$action) 'config-reader'
        }

        OpenUoms
        [void](Act 'Add' $canary)
        $before=@(Get-Slice4beActivityFiles $Fixture);$pin=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        [void](Invoke-AdminUomResetChoice 'No' 'admin-uom-reset-no.png')
        Check 'AdminUom.ResetNo.ActualNoPreservesConfigBytes' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $pin -and $canary -cin (CurrentUoms).Split('|'))
        Pair $before 'ADMIN_UOM_RESET' 'CANCELLED' 'ResetNo'

        $before=@(Get-Slice4beActivityFiles $Fixture)
        [void](Invoke-AdminUomResetChoice 'Yes' 'admin-uom-reset-yes.png')
        Check 'AdminUom.ResetYes.ActualYesRestoresDefaultCatalog' ((CurrentUoms) -ceq 'EA|LB|LBS|OZ|KG|G|GAL|QT|PT|L|ML|CS')
        Pair $before 'ADMIN_UOM_RESET' 'COMPLETED' 'ResetYes'

        $before=@(Get-Slice4beActivityFiles $Fixture);$pin=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        [void](Invoke-AdminUomResetChoice 'Yes' 'admin-uom-reset-unchanged.png')
        Check 'AdminUom.ResetUnchanged.ActualYesPreservesConfigBytes' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $pin)
        Pair $before 'ADMIN_UOM_RESET' 'UNCHANGED' 'ResetUnchanged'

        $before=@(Get-Slice4beActivityFiles $Fixture)
        $direct=[bool](Run 'invSys.Core.xlam' 'modUomSettings.AddConfiguredUom' @($canary))
        Check 'AdminUom.DirectService.PreservesObservationCount' ($direct -and @(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count -and $canary -cin (CurrentUoms).Split('|'))

        OpenUoms
        [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
        SelectTarget $Fixture
        $before=@(Get-Slice4beActivityFiles $Fixture);$pin=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        [void](Act 'Add' ($canary+'STALE'))
        Check 'AdminUom.ReplacedSession.CannotReviveCapturedForm' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $pin -and @(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count)
        OpenUoms
        $pin=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        SelectTarget $Other
        $otherPin=(Get-FileHash -LiteralPath $Other.Config).Hash
        [void](Act 'Add' ($canary+'OTHER'))
        Check 'AdminUom.ChangedTarget.CannotRedirectCapturedForm' ((Get-FileHash -LiteralPath $Other.Config).Hash -ceq $otherPin -and (Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $pin)

        $retained=$true;foreach($path in $original.Keys){$retained=$retained -and (Get-FileHash -LiteralPath $path).Hash -ceq $original[$path]}
        Check 'AdminUom.PriorActivityBytesPreserved' $retained
    } finally {
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
        $excel.Visible=$entryVisibility
        # Only this Admin-generated disposable fixture is restored.
        [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
        [IO.File]::WriteAllBytes($Other.Config,$otherBytes)
        SelectTarget $Fixture
        [void](Run 'invSys.Core.xlam' 'modConfig.Reload')
    }
    Check 'AdminUom.GeneratedConfigRestoredExactly' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configPin)
    Check 'AdminUom.OtherGeneratedConfigRestoredExactly' ((Get-FileHash -LiteralPath $Other.Config).Hash -ceq $otherInitialPin)
}
