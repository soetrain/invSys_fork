# Uses the actual Ribbon and Create handler. All destinations are disposable.
function Install-InitialAdminFailureProbe {
    $core=$packages['invSys.Core.xlam'].VBProject
    $probe=$core.VBComponents.Add(1)
    $probe.Name='TestInitialAdmin'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mArmed As Boolean, mHit As Boolean
Public Sub Arm()
    mArmed = True: mHit = False
End Sub
Public Function RefuseAccountWrite() As Boolean
    If mArmed Then mArmed = False: mHit = True: RefuseAccountWrite = True
End Function
Public Function Hit() As Boolean
    Hit = mHit
End Function
Public Function ExistingPublicationEligible(ByVal root As String) As Boolean
    Dim spec As WarehouseSpec
    ExistingPublicationEligible = modWarehousePublishSource.ReadExisting(root, spec)
End Function
'@)
    $auth=$core.VBComponents.Item('modAuth').CodeModule
    $start=$auth.ProcStartLine('EnsureStationRoleAuth',0)
    $count=$auth.ProcCountLines('EnsureStationRoleAuth',0)
    $body=$auth.Lines($start,$count)
    $anchor='(?im)^([ \t]*)EnsureUserRow loUsers, resolvedUser, resolvedDisplay[ \t]*\r?$'
    if([regex]::Matches($body,$anchor).Count -ne 1){throw 'Destination account-write boundary unavailable.'}
    $body=[regex]::Replace($body,$anchor,[Text.RegularExpressions.MatchEvaluator]{param($match)
        '    If TestInitialAdmin.RefuseAccountWrite() Then report = "Destination account write failed.": GoTo CleanExit'+"`r`n"+$match.Value
    })
    $auth.DeleteLines($start,$count)
    $auth.InsertLines($start,$body)
}

function Test-InitialAdminOnboarding {
    $empty=New-CreateCase 'initial-admin-no-entry'
    [void](Invoke-Create 'Create')
    $facts=[string](Invoke-Create 'Facts')
    Check 'Onboarding.EmptyPin.RefusesBeforeBootstrap' ($facts -ceq 'False|False|Create|False|0|0')
    Check 'Onboarding.EmptyPin.NoDestination' (-not (Test-Path -LiteralPath $empty.Root))
    [void](Invoke-Create 'CloseForm')

    # Characterize the existing bootstrap recovery boundary independently of
    # the future credential writer. Never manufacture an Auth row or a grant.
    foreach($existing in @($false,$true)){
        $kind=if($existing){'ExistingHub'}else{'NewHub'}
        $case=New-CreateCase ('initial-admin-'+$kind)
        if($existing){
            [void][IO.Directory]::CreateDirectory($case.Root)
            $sentinel=Join-Path $case.Root 'unrelated.txt'
            [IO.File]::WriteAllText($sentinel,'Preserve this fixture file.')
            $before=(Get-FileHash -LiteralPath $sentinel).Hash
        }
        [void](Run 'invSys.Core.xlam' 'TestInitialAdmin.Arm')
        [void](Invoke-Create 'Create')
        $hit=[bool](Run 'invSys.Core.xlam' 'TestInitialAdmin.Hit')
        if(-not $hit){throw 'Destination account-write fault was not reached; recovery result unavailable.'}
        $facts=[string](Invoke-Create 'Facts')
        Check ("Onboarding.$kind.ReportsCreationFailure") ($facts.StartsWith('False|False|') -and $facts.EndsWith('|1|0'))
        $eligible=[bool](Run 'invSys.Core.xlam' 'TestInitialAdmin.ExistingPublicationEligible' @($case.Root))
        Check ("Onboarding.$kind.IncompleteDestinationNotPublishable") (-not $eligible)
        if($existing){
            Check 'Onboarding.ExistingHub.PreservesUnrelatedFile' ((Get-FileHash -LiteralPath $sentinel).Hash -ceq $before)
            [void](Invoke-Create 'Create')
            $retry=[string](Invoke-Create 'Facts')
            Write-Host ('Onboarding existing hub after account-write fault: AuthFilePresent={0}, ConfigFilePresent={1}, PublicationEligible={2}, RetryReachedBootstrap={3}' -f
                (Test-Path -LiteralPath (Join-Path $case.Root ($case.Id+'.invSys.Auth.xlsb'))),(Test-Path -LiteralPath $case.Config),$eligible,($retry -like '*|2|0'))
        }else{
            Check 'Onboarding.NewHub.RemovesOwnedRoot' (-not (Test-Path -LiteralPath $case.Root))
        }
        [void](Invoke-Create 'CloseForm')
    }
}
