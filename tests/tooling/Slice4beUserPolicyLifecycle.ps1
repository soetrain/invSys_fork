# Membership changes are confined to generated Auth fixtures; only UserId cells
# are inspected. Policy changes use the real packaged Settings handlers.
function Test-UserPolicyLifecycle($Fixture) {
    $auth=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $authBytes=[IO.File]::ReadAllBytes($auth);$authHash=(Get-FileHash -LiteralPath $auth).Hash
    $configBytes=[IO.File]::ReadAllBytes($Fixture.Config);$configHash=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $newId='00041'
    try {
        $book=$excel.Workbooks.Open($auth,0,$false)
        try {
            $users=Table $book 'tblUsers';$removed=0;$added=0
            for($index=$users.ListRows.Count;$index -ge 1;$index--){
                $row=$users.ListRows.Item($index)
                $cell=$row.Range.Cells.Item(1,$users.ListColumns.Item('UserId').Index)
                if($cell.Value2 -ceq $newId){throw 'New-user fixture identity already exists.'}
                if($cell.Value2 -ceq 'config-reader'){$row.Delete();$removed++}
                elseif($cell.Value2 -ceq 'config-producer'){$cell.NumberFormat='@';$cell.Value2=$newId;$added++}
            }
            if($removed -ne 1 -or $added -ne 1){throw 'Roster membership fixture was not calibrated.'}
            $book.Save()
        } finally {$book.Close($false)}
        $changedAuth=(Get-FileHash -LiteralPath $auth).Hash
        [void](UserPolicyControl 'btnReloadTrackingPolicy' 'Click')
        $selected=(UserPolicyControl 'lstTrackingUsers' 'Select' 'config-reader') -ceq 'SELECTED'
        $request=Get-TrackingPolicyRequest
        Check 'UserPolicy.Lifecycle.RemovedOverrideShownUnavailable' ($selected -and (UserPolicyControl 'lstTrackingUsers' 'Availability') -ceq 'Unavailable' -and (UserPolicyControl 'chkUserRecord' 'Enabled') -ceq 'False' -and (UserPolicyDisabled $request 'config-reader'))
        [void](UserPolicyControl 'chkUserRecord' 'Write' 'True')
        Check 'UserPolicy.Lifecycle.UnavailableToggleCannotChangeOverride' ((Get-TrackingPolicyRequest) -ceq $request -and $configHash -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
        $version=Get-TrackingPolicyVersion $Fixture
        [void](UserPolicyControl 'btnSaveTrackingPolicy' 'Click')
        [void](UserPolicyControl 'btnReloadTrackingPolicy' 'Click')
        Check 'UserPolicy.Lifecycle.SaveRetainsUnchangedRemovedOverride' ((Get-TrackingPolicyVersion $Fixture) -eq $version+1 -and (UserPolicyDisabled (Get-TrackingPolicyRequest) 'config-reader'))
        $version=Get-TrackingPolicyVersion $Fixture
        $saved=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        $bad=Get-TrackingPolicyRequest|ConvertFrom-Json
        foreach($row in $bad.Users){if($row.UserId -ieq 'config-reader'){$row.Record=$true}}
        $context=[string](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyContext')
        $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicySaveDirect' @($context,$version,($bad|ConvertTo-Json -Depth 8 -Compress)))
        Check 'UserPolicy.Lifecycle.ChangedRemovedOverrideRejectedAtSave' (-not $ok -and $saved -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
        $selected=(UserPolicyControl 'lstTrackingUsers' 'Select' $newId) -ceq 'SELECTED'
        Check 'UserPolicy.Lifecycle.NewUserDefaultsEnabled' ($selected -and (UserPolicyControl 'lstTrackingUsers' 'Availability') -ceq 'Available' -and (UserPolicyControl 'chkUserRecord' 'Read') -ceq 'True')
        [void](UserPolicyControl 'chkUserRecord' 'Write' 'False')
        Check 'UserPolicy.Lifecycle.NewUserToggleStagesOnly' ((UserPolicyDisabled (Get-TrackingPolicyRequest) $newId) -and $saved -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
        [void](UserPolicyControl 'btnSaveTrackingPolicy' 'Click')
        [void](UserPolicyControl 'btnReloadTrackingPolicy' 'Click')
        $model=Get-TrackingPolicyRequest|ConvertFrom-Json
        $exact=@($model.Users|Where-Object {$_.UserId -is [string] -and $_.UserId -ceq $newId -and $_.Record -is [bool] -and -not $_.Record})
        Check 'UserPolicy.Lifecycle.NewUserExactTextIdentityPersists' ($exact.Count -eq 1 -and (Get-TrackingPolicyVersion $Fixture) -eq $version+1)
        Check 'UserPolicy.Lifecycle.PolicyCommandsDoNotChangeAuth' ($changedAuth -ceq (Get-FileHash -LiteralPath $auth).Hash)
        # Restore membership while retaining saved overrides: one user returns,
        # the newly appearing identity is now unavailable.
        [IO.File]::WriteAllBytes($auth,$authBytes)
        [void](UserPolicyControl 'btnReloadTrackingPolicy' 'Click')
        [void](UserPolicyControl 'lstTrackingUsers' 'Select' 'config-reader')
        Check 'UserPolicy.Lifecycle.ReturnedUserKeepsSavedFlag' ((UserPolicyControl 'lstTrackingUsers' 'Availability') -ceq 'Available' -and (UserPolicyControl 'chkUserRecord' 'Enabled') -ceq 'True' -and (UserPolicyControl 'chkUserRecord' 'Read') -ceq 'False')
        [void](UserPolicyControl 'lstTrackingUsers' 'Select' $newId)
        Check 'UserPolicy.Lifecycle.NewlyMissingOverrideRetained' ((UserPolicyControl 'lstTrackingUsers' 'Availability') -ceq 'Unavailable' -and (UserPolicyDisabled (Get-TrackingPolicyRequest) $newId))
        $saved=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        $version=Get-TrackingPolicyVersion $Fixture
        [void](UserPolicyControl 'btnResetTrackingPolicy' 'Click')
        $model=Get-TrackingPolicyRequest|ConvertFrom-Json
        Check 'UserPolicy.Lifecycle.ResetStagesEmptyOverridesOnly' (@($model.Users).Count -eq 0 -and $saved -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
        [void](UserPolicyControl 'btnSaveTrackingPolicy' 'Click')
        [void](UserPolicyControl 'btnReloadTrackingPolicy' 'Click')
        $model=Get-TrackingPolicyRequest|ConvertFrom-Json
        Check 'UserPolicy.Lifecycle.ExplicitResetSaveClearsOverrides' (@($model.Users).Count -eq 0 -and (Get-TrackingPolicyVersion $Fixture) -eq $version+1)
    } finally {
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCloseAction')
        foreach($book in @($excel.Workbooks)){
            if($book.FullName -ieq $auth -or $book.FullName -ieq $Fixture.Config){$book.Close($false)}
        }
        [IO.File]::WriteAllBytes($auth,$authBytes)
        [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
    }
    Check 'UserPolicy.Lifecycle.GeneratedFixturesRestoredExactly' ($authHash -ceq (Get-FileHash -LiteralPath $auth).Hash -and $configHash -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
}
