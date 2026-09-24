# D8-A. Only Admin-generated disposable authority is changed by fixture setup.
# Public packaged Auth callers remain real; probes observe the resolved path and
# expire the in-memory TTL only. No credentials or operational values are logged.
function Install-Slice4beAuthReadProbe {
    $module=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modAuth').CodeModule
    $source=$module.Lines(1,$module.CountOfLines)
    foreach($anchor in @('Public Function LoadAuth(', '    Set preOpen = CaptureOpenWorkbookPathsAuth()', '    mAuthWorkbook = wb.Name')){
        if([regex]::Matches($source,[regex]::Escape($anchor),[Text.RegularExpressions.RegexOptions]::IgnoreCase).Count -ne 1){throw 'Auth observation anchor unavailable.'}
    }
    $source=$source.Replace('Public Function LoadAuth(',"Private AuthObservedPathForTest As String`r`n`r`nPublic Function LoadAuth(")
    $source=$source.Replace('    Set preOpen = CaptureOpenWorkbookPathsAuth()',"    AuthObservedPathForTest = vbNullString`r`n    Set preOpen = CaptureOpenWorkbookPathsAuth()")
    $source=[regex]::Replace($source,[regex]::Escape('    mAuthWorkbook = wb.Name'),"    mAuthWorkbook = wb.Name`r`n    AuthObservedPathForTest = wb.FullName",[Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($source)
    $module.AddFromString(@'
Public Function AuthReadPathForTest() As String
    AuthReadPathForTest = AuthObservedPathForTest
End Function
Public Sub ExpireAuthCacheForTest()
    mLoadedAt = DateSerial(1900, 1, 1)
End Sub
'@)
}

function Get-AuthReadFixtureHash([string]$Path) {
    $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
    $sha=[Security.Cryptography.SHA256]::Create()
    try{[BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-','')}finally{$sha.Dispose();$stream.Dispose()}
}

function Test-Slice4beAuthReadOnly($Fixture,$Other) {
    $ownedRoot=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $otherPath=Join-Path $Other.Root ($Other.Warehouse+'.invSys.Auth.xlsb')
    foreach($path in @($authPath,$otherPath)){
        if(-not [IO.Path]::GetFullPath($path).StartsWith($ownedRoot,[StringComparison]::OrdinalIgnoreCase)){throw 'Auth fixture escaped generated root.'}
    }
    Check 'AuthRead.ExplicitAdminProvisioningCreatedBothAuthorities' ((Test-Path -LiteralPath $authPath) -and (Test-Path -LiteralPath $otherPath))
    $auth=$null;$otherBook=$null;$lock=$null
    $original=[IO.File]::ReadAllBytes($authPath)
    $configOriginal=[IO.File]::ReadAllBytes($Fixture.Config)
    $otherHash=(Get-FileHash -LiteralPath $otherPath).Hash
    $configHash=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    try {
        $auth=$excel.Workbooks.Open($authPath,0,$false)
        foreach($tableName in @('tblUsers','tblCapabilities')){
            $table=Table $auth $tableName
            $extra=$table.ListColumns.Add(1);$extra.Name='OperatorAuthExtra'
            $extra.DataBodyRange.Value2='retained fixture annotation'
        }
        $auth.Save();$auth.Close($false);$auth=$null
        $healthy=[IO.File]::ReadAllBytes($authPath)
        $healthyHash=(Get-FileHash -LiteralPath $authPath).Hash
        foreach($operation in @('LoadAuth','ReloadAuth')){
            $loaded=Run 'invSys.Core.xlam' ('modAuth.'+$operation) @($Fixture.Warehouse)
            Check ('AuthRead.Healthy.'+$operation+'.Loads') ($loaded -is [bool] -and $loaded)
            Check ('AuthRead.Healthy.'+$operation+'.ExactPath') ((Run 'invSys.Core.xlam' 'modAuth.AuthReadPathForTest') -ieq $authPath)
            Check ('AuthRead.Healthy.'+$operation+'.BytesPreserved') ((Get-FileHash -LiteralPath $authPath).Hash -ceq $healthyHash)
        }
        SelectTarget $Fixture 'config-reader'
        Check 'AuthRead.Healthy.SignInAndAllowedCapability' ([bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('RECEIVE_POST','config-reader',$Fixture.Warehouse,'S1')))
        Check 'AuthRead.Healthy.SignInBytesPreserved' ((Get-FileHash -LiteralPath $authPath).Hash -ceq $healthyHash)
        Check 'AuthRead.Healthy.ProcessorCapabilityRetained' ([bool](Run 'invSys.Core.xlam' 'modAuth.HasProvisionedCapabilityForSystem' @('INBOX_PROCESS','svc_processor',$Fixture.Warehouse,'S1')))
        [IO.File]::Delete($authPath)
        Check 'AuthRead.ValidTtlRetainsCurrentCapabilityCache' ([bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('RECEIVE_POST','config-reader',$Fixture.Warehouse,'S1')))
        Check 'AuthRead.ValidTtlDoesNotCreateAuthority' (-not (Test-Path -LiteralPath $authPath))
        [IO.File]::WriteAllBytes($authPath,$healthy)
        $auth=$excel.Workbooks.Open($authPath,0,$false)
        $users=Table $auth 'tblUsers'
        $users.ListColumns.Item('OperatorAuthExtra').DataBodyRange.Cells.Item(1,1).Value2='unsaved local annotation'
        $dirtyBefore=-not [bool]$auth.Saved
        $loaded=Run 'invSys.Core.xlam' 'modAuth.ReloadAuth' @($Fixture.Warehouse)
        Check 'AuthRead.OpenHealthy.LoadsWithoutSavingOrClosing' ($dirtyBefore -and [bool]$loaded -and -not [bool]$auth.Saved -and $auth.FullName -ieq $authPath)
        Check 'AuthRead.OpenHealthy.UnknownValueAndDiskPreserved' ($users.ListColumns.Item('OperatorAuthExtra').DataBodyRange.Cells.Item(1,1).Value2 -ceq 'unsaved local annotation' -and (Get-AuthReadFixtureHash $authPath) -ceq $healthyHash)
        $auth.Close($false);$auth=$null

        foreach($operation in @('LoadAuth','ReloadAuth','SignIn','CapabilityRefresh','ProcessorCapability')){
            SelectTarget $Fixture 'config-reader'
            [IO.File]::Delete($authPath)
            switch($operation){
                'SignIn' {$result=Run 'invSys.Core.xlam' 'modAuth.SignInCurrentTargetForAutomation' @('config-reader',$Fixture.Secret,'');$denied=-not ([string]$result).StartsWith('OK|')}
                'CapabilityRefresh' {
                    [void](Run 'invSys.Core.xlam' 'modAuth.ExpireAuthCacheForTest')
                    $result=Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('RECEIVE_POST','config-reader',$Fixture.Warehouse,'S1');$denied=$result -is [bool] -and -not $result
                }
                'ProcessorCapability' {
                    [void](Run 'invSys.Core.xlam' 'modAuth.ExpireAuthCacheForTest')
                    $result=Run 'invSys.Core.xlam' 'modAuth.HasProvisionedCapabilityForSystem' @('INBOX_PROCESS','svc_processor',$Fixture.Warehouse,'S1');$denied=$result -is [bool] -and -not $result
                }
                default {$result=Run 'invSys.Core.xlam' ('modAuth.'+$operation) @($Fixture.Warehouse);$denied=$result -is [bool] -and -not $result}
            }
            Check ('AuthRead.Missing.'+$operation+'.FailsClosed') $denied
            Check ('AuthRead.Missing.'+$operation+'.DoesNotCreateAuthority') (-not (Test-Path -LiteralPath $authPath))
            [IO.File]::WriteAllBytes($authPath,$healthy)
        }

        SelectTarget $Fixture
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
        try {
            [IO.File]::Delete($authPath)
            [void](Run 'invSys.Core.xlam' 'modAuth.ExpireAuthCacheForTest')
            $saved=Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('BatchSize','799')
            Check 'AuthRead.Missing.ActualSettingsSaveDenied' ($saved -is [bool] -and -not $saved)
            Check 'AuthRead.Missing.ActualSettingsSavePreservesConfig' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configHash)
            Check 'AuthRead.Missing.ActualSettingsSaveDoesNotCreateAuth' (-not (Test-Path -LiteralPath $authPath))
        } finally {
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
            [IO.File]::WriteAllBytes($authPath,$healthy)
        }

        SelectTarget $Fixture
        $held=$Fixture.Root+'-auth-read-held';$created=$Fixture.Root+'-auth-read-created'
        foreach($path in @($Fixture.Root,$held,$created)){
            if(-not [IO.Path]::GetFullPath($path).StartsWith($ownedRoot,[StringComparison]::OrdinalIgnoreCase)){throw 'Directory fixture escaped generated root.'}
        }
        if((Test-Path -LiteralPath $held) -or (Test-Path -LiteralPath $created)){throw 'Directory fixture destination already exists.'}
        Move-Item -LiteralPath $Fixture.Root -Destination $held
        try {
            $result=Run 'invSys.Core.xlam' 'modAuth.ReloadAuth' @($Fixture.Warehouse)
            Check 'AuthRead.MissingDirectory.FailsClosed' ($result -is [bool] -and -not $result)
            Check 'AuthRead.MissingDirectory.DoesNotCreateDirectory' (-not (Test-Path -LiteralPath $Fixture.Root))
        } finally {
            if(Test-Path -LiteralPath $Fixture.Root){Move-Item -LiteralPath $Fixture.Root -Destination $created}
            Move-Item -LiteralPath $held -Destination $Fixture.Root
        }

        foreach($tableName in @('tblUsers','tblCapabilities')){
            SelectTarget $Fixture
            $auth=$excel.Workbooks.Open($authPath,0,$false)
            $table=Table $auth $tableName
            $required=if($tableName -ceq 'tblUsers'){'Status'}else{'Capability'}
            $table.ListColumns.Item($required).Name='RemovedRequiredHeader'
            $auth.Save();$auth.Close($false);$auth=$null
            $invalidHash=(Get-FileHash -LiteralPath $authPath).Hash
            $result=Run 'invSys.Core.xlam' 'modAuth.ReloadAuth' @($Fixture.Warehouse)
            Check ('AuthRead.Invalid.'+$tableName+'.FailsClosed') ($result -is [bool] -and -not $result)
            Check ('AuthRead.Invalid.'+$tableName+'.DoesNotRepairOrSave') ((Get-FileHash -LiteralPath $authPath).Hash -ceq $invalidHash)
            [IO.File]::WriteAllBytes($authPath,$healthy)
        }

        SelectTarget $Fixture
        $otherBook=$excel.Workbooks.Open($otherPath,0,$false)
        $otherSaved=[bool]$otherBook.Saved
        $lock=[IO.File]::Open($authPath,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
        try {
            $result=Run 'invSys.Core.xlam' 'modAuth.ReloadAuth' @($Fixture.Warehouse)
            Check 'AuthRead.Unreadable.FailsClosedWithoutOtherWarehouseFallback' ($result -is [bool] -and -not $result)
            Check 'AuthRead.Unreadable.DoesNotReadOtherWarehouse' ((Run 'invSys.Core.xlam' 'modAuth.AuthReadPathForTest') -cne $otherPath)
        } finally {$lock.Dispose();$lock=$null}
        Check 'AuthRead.Unreadable.BothAuthoritiesPreserved' ((Get-FileHash -LiteralPath $authPath).Hash -ceq $healthyHash -and (Get-AuthReadFixtureHash $otherPath) -ceq $otherHash -and [bool]$otherBook.Saved -eq $otherSaved)
        $otherBook.Close($false);$otherBook=$null

        SelectTarget $Fixture
        $filesBefore=@(Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File | ForEach-Object FullName)
        $result=Run 'invSys.Core.xlam' 'modAuth.LoadAuth' @($Other.Warehouse)
        Check 'AuthRead.WrongWarehouseArgument.FailsClosed' ($result -is [bool] -and -not $result)
        Check 'AuthRead.WrongWarehouseArgument.DoesNotCreateAuthority' (@(Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File | Where-Object {$_.FullName -cnotin $filesBefore}).Count -eq 0)
        Check 'AuthRead.WrongWarehouseArgument.DoesNotReadOtherWarehouse' ((Run 'invSys.Core.xlam' 'modAuth.AuthReadPathForTest') -cne $otherPath)
        Check 'AuthRead.WrongWarehouseArgument.AuthoritiesPreserved' ((Get-FileHash -LiteralPath $authPath).Hash -ceq $healthyHash -and (Get-FileHash -LiteralPath $otherPath).Hash -ceq $otherHash)
        Check 'AuthRead.ConfigAuthorityPreserved' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configHash)

        [void](Run 'invSys.Core.xlam' 'modNasConnection.ClearWarehouseTarget')
        $configuredRoot=[string](Run 'invSys.Core.xlam' 'modConfig.GetString' @('PathDataRoot',''))
        if($configuredRoot.TrimEnd('\') -ine $Fixture.Root.TrimEnd('\')){throw 'Unselected-source fixture is not contained in the generated runtime.'}
        $result=Run 'invSys.Core.xlam' 'modAuth.LoadAuth' @($Fixture.Warehouse)
        Check 'AuthRead.Unselected.DoesNotDiscoverAuthority' ($result -is [bool] -and -not $result)
        Check 'AuthRead.Unselected.ExistingBytesPreserved' ((Get-FileHash -LiteralPath $authPath).Hash -ceq $healthyHash)
        [void](Run 'invSys.Core.xlam' 'modRuntimeWorkbooks.SetCoreDataRootOverride' @($Fixture.Root))
        $configured=Run 'invSys.Core.xlam' 'modConfig.LoadConfig' @($Fixture.Warehouse,'S1')
        if($configured -isnot [bool] -or -not $configured){throw 'Explicit headless runtime configuration is unavailable.'}
        $result=Run 'invSys.Core.xlam' 'modAuth.LoadAuth' @($Fixture.Warehouse)
        Check 'AuthRead.ExplicitHeadlessRuntime.Loads' ($result -is [bool] -and $result)
        Check 'AuthRead.ExplicitHeadlessRuntime.ExactPath' ((Run 'invSys.Core.xlam' 'modAuth.AuthReadPathForTest') -ieq $authPath)
        Check 'AuthRead.ExplicitHeadlessRuntime.ProcessorCapabilityRetained' ([bool](Run 'invSys.Core.xlam' 'modAuth.HasProvisionedCapabilityForSystem' @('INBOX_PROCESS','svc_processor',$Fixture.Warehouse,'S1')))
        Check 'AuthRead.ExplicitHeadlessRuntime.BytesPreserved' ((Get-FileHash -LiteralPath $authPath).Hash -ceq $healthyHash)
        $unexpected=Join-Path $Fixture.Root ($Other.Warehouse+'.invSys.Auth.xlsb')
        if(-not [IO.Path]::GetFullPath($unexpected).StartsWith($ownedRoot,[StringComparison]::OrdinalIgnoreCase)){throw 'Mismatched Auth fixture escaped generated root.'}
        if(Test-Path -LiteralPath $unexpected){[IO.File]::Delete($unexpected)}
        $result=Run 'invSys.Core.xlam' 'modAuth.LoadAuth' @($Other.Warehouse)
        Check 'AuthRead.ExplicitHeadlessRuntime.WrongWarehouseFailsClosed' ($result -is [bool] -and -not $result)
        Check 'AuthRead.ExplicitHeadlessRuntime.WrongWarehouseDoesNotCreate' (-not (Test-Path -LiteralPath $unexpected))

        SelectTarget $Fixture
        [IO.File]::Delete($authPath)
        $setupArgs=@($Fixture.Warehouse,'S1','d8-provisioned','Provisioning fixture','RECEIVE',$authPath,'svc_processor') -join '|'
        $setup=Run 'invSys.Core.xlam' 'modAuth.EnsureStationRoleAuthPackedForAutomation' @($setupArgs)
        Check 'AuthRead.ExplicitStationProvisioningCreatesAuthority' (([string]$setup).StartsWith('OK|') -and (Test-Path -LiteralPath $authPath))
        $auth=$excel.Workbooks.Open($authPath,0,$true)
        $caps=Table $auth 'tblCapabilities'
        $grants=@($caps.ListRows | Where-Object {$_.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'd8-provisioned' -and $_.Range.Cells.Item(1,$caps.ListColumns.Item('Capability').Index).Value2 -ceq 'RECEIVE_POST' -and $_.Range.Cells.Item(1,$caps.ListColumns.Item('WarehouseId').Index).Value2 -ceq $Fixture.Warehouse -and $_.Range.Cells.Item(1,$caps.ListColumns.Item('StationId').Index).Value2 -ceq 'S1'})
        Check 'AuthRead.ExplicitStationProvisioningOwnsExactGrant' ($grants.Count -eq 1)
        $auth.Close($false);$auth=$null
        [IO.File]::WriteAllBytes($authPath,$healthy)

        # D2 rule8 is an existing authorized post-credential command, distinct
        # from D8-A's ordinary reads. Protect it instead of silently removing it.
        $station=[string](Run 'invSys.Core.xlam' 'modStationIdentity.CurrentComputerStationId')
        if([string]::IsNullOrWhiteSpace($station) -or $station -ieq 'S1'){throw 'Current-computer transition fixture is unavailable.'}
        $selected=Run 'invSys.Core.xlam' 'modNasConnection.SelectWarehouseTargetForAutomation' @($Fixture.Root,$Fixture.Root,$station,$false)
        if(-not ([string]$selected).StartsWith('OK|')){throw 'Current-computer fixture selection failed.'}
        foreach($case in @('WrongCredential','MissingCapability')){
            $secret=if($case -ceq 'WrongCredential'){[guid]::NewGuid().ToString('N')}else{$Fixture.Secret}
            $capability=if($case -ceq 'WrongCredential'){'RECEIVE_POST'}else{'ADMIN_MAINT'}
            $result=Run 'invSys.Core.xlam' 'modAuth.SignInCurrentTargetForAutomation' @('config-reader',$secret,$capability)
            Check ('AuthRead.StationTransition.'+$case+'.Denied') (-not ([string]$result).StartsWith('OK|'))
            Check ('AuthRead.StationTransition.'+$case+'.AuthorityPreserved') ((Get-FileHash -LiteralPath $authPath).Hash -ceq $healthyHash)
        }
        $result=Run 'invSys.Core.xlam' 'modAuth.SignInCurrentTargetForAutomation' @('config-reader',$Fixture.Secret,'RECEIVE_POST')
        Check 'AuthRead.StationTransition.ValidCredentialStillAuthorized' (([string]$result).StartsWith('OK|'))
        $auth=$excel.Workbooks.Open($authPath,0,$false)
        $caps=Table $auth 'tblCapabilities'
        $grants=@($caps.ListRows | Where-Object {$_.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-reader' -and $_.Range.Cells.Item(1,$caps.ListColumns.Item('Capability').Index).Value2 -ceq 'RECEIVE_POST' -and $_.Range.Cells.Item(1,$caps.ListColumns.Item('WarehouseId').Index).Value2 -ceq $Fixture.Warehouse})
        $stationRows=@($grants | Where-Object {$_.Range.Cells.Item(1,$caps.ListColumns.Item('StationId').Index).Value2 -ceq $station})
        $legacyRows=@($grants | Where-Object {$_.Range.Cells.Item(1,$caps.ListColumns.Item('StationId').Index).Value2 -ceq 'S1'})
        Check 'AuthRead.StationTransition.ExactGrantAndOriginalPreserved' ($stationRows.Count -eq 1 -and $legacyRows.Count -eq 1)
        if($stationRows.Count -ne 1){throw 'Authorized transition fixture did not produce exactly one station grant.'}
        $stationRows[0].Range.Cells.Item(1,$caps.ListColumns.Item('Status').Index).Value2='DENY'
        $auth.Save();$auth.Close($false);$auth=$null
        $denyHash=(Get-FileHash -LiteralPath $authPath).Hash
        $result=Run 'invSys.Core.xlam' 'modAuth.SignInCurrentTargetForAutomation' @('config-reader',$Fixture.Secret,'RECEIVE_POST')
        Check 'AuthRead.StationTransition.ExplicitDenyRetained' (-not ([string]$result).StartsWith('OK|') -and (Get-FileHash -LiteralPath $authPath).Hash -ceq $denyHash)
    } finally {
        if($null -ne $lock){$lock.Dispose()}
        if($null -ne $auth){$auth.Close($false)}
        if($null -ne $otherBook){$otherBook.Close($false)}
        [IO.File]::WriteAllBytes($authPath,$original)
        [IO.File]::WriteAllBytes($Fixture.Config,$configOriginal)
    }
    SelectTarget $Fixture
    Check 'AuthRead.FixtureRestoredForPackagedCommandRegressions' ((Get-FileHash -LiteralPath $otherPath).Hash -ceq $otherHash -and (Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configHash)
}
