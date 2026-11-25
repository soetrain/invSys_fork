# Uses the actual Ribbon and Create handler. All destinations are disposable.
function Install-InitialAdminFailureProbe {
    $core=$packages['invSys.Core.xlam'].VBProject
    $probe=$core.VBComponents.Add(1)
    $probe.Name='TestInitialAdmin'
    $admin=$packages['invSys.Admin.xlam'].VBProject
    $ownerModule=$admin.VBComponents.Item('TestCreateOwner').CodeModule
    $ownerModule.InsertLines($ownerModule.CountOfDeclarationLines+1,'Private mClosedWithoutEntry As Boolean')
    $messages=$admin.VBComponents.Item('modAdminUsersAction').CodeModule
    $line=$messages.ProcBodyLine('ShowNotice',0)
    $presentation='MsgBox TRACKING_NOTICE, vbExclamation, "Users & Roles"'
    while(!$messages.Lines($line,1).Contains($presentation)){
        if($messages.Lines($line,1).Trim() -ceq 'End Sub'){throw 'Users notice presentation changed.'}
        $line++
    }
    $messages.ReplaceLine($line,$messages.Lines($line,1).Replace($presentation,'TestCreateOwner.ObserveRoleNotice TRACKING_NOTICE'))
    $ownerModule.AddFromString(@'
Public Sub ObserveRoleNotice(ByVal notice As String)
    If notice <> "Tracking unavailable: the Users & Roles action could not be fully recorded." Then Err.Raise 5
    ObserveTrackingNotice notice
End Sub
Public Sub ObservePrivateClose(ByVal view As Object)
    Dim ctl As Object, cleared As Long
    For Each ctl In view.Controls
        If ctl.Name = "txtInitialPin" Or ctl.Name = "txtConfirmPin" Then
            If CStr(ctl.Value) = "" Then cleared = cleared + 1
        End If
    Next ctl
    mClosedWithoutEntry = (cleared = 2)
End Sub
Public Function ClosedWithoutEntry() As Boolean
    ClosedWithoutEntry = mClosedWithoutEntry
End Function
Public Sub EnterPrivate(ByVal entry As String, ByVal confirmation As String)
    Dim ctl As Object
    For Each ctl In frmCreateWarehouse.Controls
        If ctl.Name = "txtInitialPin" Then
            ctl.Value = entry
            frmCreateWarehouse.Controls("txtConfirmPin").Value = confirmation
        End If
    Next ctl
End Sub
Public Function PrivateCleared() As Boolean
    PrivateCleared = (frmCreateWarehouse.Controls("txtInitialPin").Value = "" And frmCreateWarehouse.Controls("txtConfirmPin").Value = "")
End Function
Public Function PrivateFieldsReady() As Boolean
    Dim ctl As Object, count As Long
    For Each ctl In frmCreateWarehouse.Controls
        If ctl.Name = "txtInitialPin" Or ctl.Name = "txtConfirmPin" Then
            If ctl.Visible And ctl.Enabled And ctl.PasswordChar <> "" And ctl.TabStop And _
                ctl.Left >= 0 And ctl.Top >= 0 And ctl.Left + ctl.Width <= frmCreateWarehouse.InsideWidth And _
                ctl.Top + ctl.Height <= frmCreateWarehouse.InsideHeight Then count = count + 1
        End If
    Next ctl
    If count = 2 Then PrivateFieldsReady = (frmCreateWarehouse.Controls("txtConfirmPin").TabIndex = frmCreateWarehouse.Controls("txtInitialPin").TabIndex + 1)
End Function
Public Sub GuideChoice(ByVal user As String, ByVal grant As Boolean)
    Dim view As New frmCreateDeleteUser
    view.GuideChoiceForTest user, grant
    Unload view
End Sub
Public Sub PendingChoice(ByVal user As String, ByVal scopeTest As String)
    frmCreateDeleteUser.GuideChoiceForTest user, True, True, scopeTest
End Sub
Public Sub SubmitPendingChoice()
    frmCreateDeleteUser.SubmitPendingChoiceForTest
    Unload frmCreateDeleteUser
End Sub
Public Function GuideAfterScopeChange(ByVal user As String) As Boolean
    Dim view As New frmCreateDeleteUser
    GuideAfterScopeChange = view.GuideAfterScopeChangeForTest(user)
    Unload view
End Function
'@)
    $createForm=$admin.VBComponents.Item('frmCreateWarehouse').CodeModule
    $line=$createForm.ProcBodyLine('UserForm_QueryClose',0)
    while($line -le $createForm.CountOfLines -and $createForm.Lines($line,1).Trim() -cne 'End Sub'){$line++}
    if($line -gt $createForm.CountOfLines){throw 'Create Close boundary unavailable.'}
    $createForm.InsertLines($line,'    TestCreateOwner.ObservePrivateClose Me')
    $admin.VBComponents.Item('frmCreateDeleteUser').CodeModule.AddFromString(@'
Public Sub GuideChoiceForTest(ByVal user As String, ByVal grant As Boolean, Optional ByVal hold As Boolean = False, Optional ByVal scopeTest As String = "")
    Dim target As WarehouseTarget, ctl As Object
    Set target = modNasConnection.GetCurrentTarget()
    mTxtAuthPath.Value = target.AuthPath
    mTxtWarehouseId.Value = target.WarehouseId
    mChkRoamingStation.Value = False
    mTxtStationId.Value = target.StationId
    If scopeTest = "WrongStation" Then mTxtStationId.Value = "S2"
    If scopeTest = "WrongWarehouse" Then mTxtWarehouseId.Value = "different-training-warehouse"
    LoadSelectedUserForm user
    For Each ctl In Me.Controls
        If ctl.Name = "chkGuideMaint" Then ctl.Value = grant
    Next ctl
    If Not hold Then mBtnSave_Click
End Sub
Public Sub SubmitPendingChoiceForTest()
    mBtnSave_Click
End Sub
Public Function GuideAfterScopeChangeForTest(ByVal user As String) As Boolean
    Dim target As WarehouseTarget, ctl As Object
    Set target = modNasConnection.GetCurrentTarget()
    mTxtAuthPath.Value = target.AuthPath
    mTxtWarehouseId.Value = target.WarehouseId
    mChkRoamingStation.Value = True
    LoadSelectedUserForm user
    mChkRoamingStation.Value = False
    mTxtStationId.Value = target.StationId
    For Each ctl In Me.Controls
        If ctl.Name = "chkGuideMaint" Then GuideAfterScopeChangeForTest = CBool(ctl.Value)
    Next ctl
End Function
'@)
    $core.VBComponents.Item('frmSignIn').CodeModule.AddFromString(@'
Public Sub EnterForTest(ByVal user As String, ByVal entry As String)
    mTxtUser.Value = user
    mTxtSecret.Value = entry
    mBtnOK_Click
End Sub
'@)
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mArmed As Boolean, mHit As Boolean
Private mFailSave As Boolean, mSaveHit As Boolean
Private mLateWrite As Boolean, mLateHit As Boolean
Public Sub Arm()
    mArmed = True: mHit = False
End Sub
Public Sub ArmSave()
    mFailSave = True: mSaveHit = False
End Sub
Public Sub BeforePersist(ByVal staging As String)
    Dim stream As Object
    If Not mFailSave Then Exit Sub
    mFailSave = False
    Set stream = CreateObject("Scripting.FileSystemObject").CreateTextFile(staging & "\unrelated-arrival.txt", False)
    stream.Write "Independent fixture content"
    stream.Close
    mSaveHit = True
    Err.Raise 75
End Sub
Public Function SaveWasRefused() As Boolean
    SaveWasRefused = mSaveHit
End Function
Public Sub ArmLateWrite()
    mLateWrite = True: mLateHit = False
End Sub
Public Function RefuseLateWrite() As Boolean
    If mLateWrite Then mLateWrite = False: mLateHit = True: RefuseLateWrite = True
End Function
Public Function LateWriteSeen() As Boolean
    LateWriteSeen = mLateHit
    mLateWrite = False
End Function
Public Function RefuseAccountWrite() As Boolean
    If mArmed Then mArmed = False: mHit = True: RefuseAccountWrite = True
End Function
Public Function Hit() As Boolean
    Hit = mHit
End Function
Public Function SignIn(ByVal user As String, ByVal entry As String) As Boolean
    Dim view As New frmSignIn
    view.InitializeSignIn modNasConnection.GetCurrentTarget()
    view.EnterForTest user, entry
    SignIn = (view.ResultStatus = AUTH_OK And modAuth.IsSignedIn())
    Unload view
End Function
Public Function HasGuide(ByVal user As String) As Boolean
    Dim target As WarehouseTarget
    Set target = modNasConnection.GetCurrentTarget()
    If Not modAuth.LoadAuth(target.WarehouseId) Then Err.Raise 5
    HasGuide = modAuth.HasProvisionedCapabilityForSystem("ACTION_PATH_MAINT", user, target.WarehouseId, target.StationId)
End Function
Public Function LegacySave(ByVal user As String, ByVal mask As Long) As Boolean
    Dim target As WarehouseTarget, report As String
    Set target = modNasConnection.GetCurrentTarget()
    LegacySave = modAuthUserCommands.SaveUser(modAuthUserCommands.CaptureContext(), target.AuthPath, user, user, "", target.WarehouseId, target.StationId, mask, report)
End Function
Public Function TransferAllowed() As Boolean
    Dim notice As String
    TransferAllowed = modGuideTransfer.CanImport(modActivity.CaptureContext(), notice)
End Function
Public Function ExistingPublicationEligible(ByVal root As String) As Boolean
    Dim spec As WarehouseSpec
    ExistingPublicationEligible = modWarehousePublishSource.ReadExisting(root, spec)
End Function
'@)
    $auth=$core.VBComponents.Item('modAuth').CodeModule
    if($auth.Lines(1,$auth.CountOfLines).Contains('Public Sub WriteUserCredential(')){
        $probe.CodeModule.AddFromString(@'
Public Function TypedRejects(ByVal id As String, ByVal root As String, ByVal entry As String, ByVal confirmation As String) As Boolean
    Dim report As String, path As String, outcome As String
    If modWarehouseCreateCommands.CreateWarehouse(modAuthUserCommands.CaptureContext(), id, "Training", "S1", "create-fixture-admin", root, "", "Training", report, path, outcome, entry, confirmation) Then Exit Function
    TypedRejects = (outcome = "REJECTED" And path = "" And Dir$(root, vbDirectory) = "")
End Function
'@)
        $start=$auth.ProcBodyLine('WriteUserCredential',0)
        $auth.InsertLines($start+1,'    If TestInitialAdmin.RefuseAccountWrite() Then Err.Raise 5')
        $bootstrap=$core.VBComponents.Item('modWarehouseBootstrap').CodeModule
        $start=$bootstrap.ProcBodyLine('PersistInitialAdmin',0)
        $end=$start+$bootstrap.ProcCountLines('PersistInitialAdmin',0)
        $anchors=@(for($line=$start;$line -lt $end;$line++){if($bootstrap.Lines($line,1) -match '^\s*wb\.SaveAs\s+Filename\s*:=\s*stagedFile\s*,'){$line}})
        if($anchors.Count -ne 1){throw 'Isolated credential persistence boundary unavailable.'}
        $bootstrap.InsertLines($anchors[0],'    TestInitialAdmin.BeforePersist staging')
        $start=$auth.ProcBodyLine('EnsureStationRoleAuth',0)
        $end=$start+$auth.ProcCountLines('EnsureStationRoleAuth',0)
        $anchors=@(for($line=$start;$line -lt $end;$line++){if($auth.Lines($line,1).Trim() -eq 'EnsureUserRow loUsers, resolvedUser, resolvedDisplay'){$line}})
        if($anchors.Count -ne 1){throw 'Original account-provisioning boundary unavailable.'}
        $auth.InsertLines($anchors[0],'    If TestInitialAdmin.RefuseLateWrite() Then GoTo CleanExit')
        return
    }
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
    function Test-GuideChoice($user,[int]$priorMask,[string]$scope){
        function HasGuide { [bool](Run 'invSys.Core.xlam' 'TestInitialAdmin.HasGuide' @($user)) }
        [void](Invoke-Create 'GuideChoice' @($user,$true))
        Check "Onboarding.$scope.ExplicitGrant" (HasGuide)
        Check "Onboarding.$scope.SwitchToExactScopeShowsSavedGrant" ([bool](Invoke-Create 'GuideAfterScopeChange' @($user)))
        Check "Onboarding.$scope.LegacySaveWithGrant" ([bool](Run 'invSys.Core.xlam' 'TestInitialAdmin.LegacySave' @($user,$priorMask)) -and (HasGuide))
        [void](Invoke-Create 'GuideChoice' @($user,$false))
        Check "Onboarding.$scope.ExplicitRevoke" (-not (HasGuide))
        Check "Onboarding.$scope.LegacySaveWithoutGrant" ([bool](Run 'invSys.Core.xlam' 'TestInitialAdmin.LegacySave' @($user,$priorMask)) -and -not (HasGuide))
    }
    function Test-CreatedDestination($case){
        Check 'Onboarding.CompletedClearsBoth' ([bool](Invoke-Create 'PrivateCleared'))
        Check 'Onboarding.Create.SourceSessionPreserved' (([string](Run 'invSys.Core.xlam' 'TestCreateCalls.ContextFacts')) -ceq 'SignedIn=True|VersionChanged=False|TargetMissing=False|RootChanged=False|HubChanged=False|StationChanged=False')
        Check 'Onboarding.Create.SourceAuthPreserved' ($sourceHash -ceq (Get-FileHash -LiteralPath $sourceAuth).Hash)
        [void](Invoke-Create 'CloseForm')
        $destinationAuth=Join-Path $case.Root ($case.Id+'.invSys.Auth.xlsb')
        $authority=$excel.Workbooks.Open($destinationAuth,0,$true)
        try{
            $roles=Table $authority 'tblCapabilities'
            $initialRoles=@(foreach($row in $roles.ListRows){
                if($row.Range.Cells.Item(1,$roles.ListColumns.Item('UserId').Index).Value2 -eq 'create-fixture-admin' -and
                    $row.Range.Cells.Item(1,$roles.ListColumns.Item('Status').Index).Value2 -eq 'ACTIVE'){
                    $scopeOk=$row.Range.Cells.Item(1,$roles.ListColumns.Item('WarehouseId').Index).Value2 -ceq $case.Id -and
                        $row.Range.Cells.Item(1,$roles.ListColumns.Item('StationId').Index).Value2 -ceq 'S1'
                    if($scopeOk){[string]$row.Range.Cells.Item(1,$roles.ListColumns.Item('Capability').Index).Value2}else{'UnexpectedScope'}
                }
            })
            Check 'Onboarding.Destination.ExactTwoDefaultRoles' ((($initialRoles|Sort-Object) -join ',') -ceq 'ADMIN_MAINT,RECEIVE_POST')
        }finally{$authority.Close($false)}
        [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
        [void](Run 'invSys.Core.xlam' 'modRuntimeWorkbooks.SetCoreDataRootOverride' @($case.Root))
        $selection=[string](Run 'invSys.Core.xlam' 'modNasConnection.SelectWarehouseTargetForAutomation' @($case.Root,$case.Root,'S1',$false))
        if(-not $selection.StartsWith('OK|')){throw 'Generated Training destination could not be selected.'}
        [void](Run 'invSys.Core.xlam' 'modNasConnection.SetCurrentTargetPathsForTest' @('\\fixture-host\config-command',$case.Root))
        $user='create-fixture-admin'
        $signed=[bool](Run 'invSys.Core.xlam' 'TestInitialAdmin.SignIn' @($user,$secret))
        Check 'Onboarding.Destination.OrdinarySignIn' $signed
        if($signed){
            Check 'Onboarding.Destination.NoImplicitGuideGrant' (-not [bool](Run 'invSys.Core.xlam' 'TestInitialAdmin.HasGuide' @($user)))
            Test-GuideChoice $user 3 'Destination'
            [void](Invoke-Create 'GuideChoice' @($user,$true))
            if($GuideDraftOnly){
                . (Join-Path $PSScriptRoot 'Slice4beOnboardingGuides.ps1')
                $guide=$null
                try {New-OnboardingGuide $case ([ref]$guide)} catch {
                    $where=$_.InvocationInfo
                    Write-Host ('Onboarding guide fixture stopped at '+[IO.Path]::GetFileName($where.ScriptName)+' line '+$where.ScriptLineNumber)
                    throw
                }
            }
            . (Join-Path $PSScriptRoot 'Slice4beOnboardingPeer.ps1')
            Test-OnboardingPeer $case $user $secret
            $textLeak=$false
            $storedRepresentation=[string](Run 'invSys.Core.xlam' 'modAuth.HashUserCredential' @($secret))
            foreach($file in Get-ChildItem -LiteralPath $runRoot,$reportRoot -Recurse -File -Include '*.json','*.jsonl','*.log','*.txt','*.md'){
                $content=[IO.File]::ReadAllText($file.FullName)
                if($content.Contains($secret) -or $content.Contains($differentEntry) -or $content.Contains($storedRepresentation)){$textLeak=$true;break}
            }
            Check 'Onboarding.Privacy.NoPrivateInputInTextEvidence' (-not $textLeak)
        }
    }
    $secret=[guid]::NewGuid().ToString('N')
    $differentEntry=[guid]::NewGuid().ToString('N')
    $sourceAuth=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $sourceHash=(Get-FileHash -LiteralPath $sourceAuth).Hash
    if($GuideDraftOnly){
        $case=New-CreateCase 'onboard-guide-author'
        [void](Invoke-Create 'EnterPrivate' @($secret,$secret))
        [void](Invoke-Create 'Create')
        if(([string](Invoke-Create 'Facts')) -cne 'True|True|Close|True|1|0'){throw 'Ordinary Training creation failed.'}
        Test-CreatedDestination $case
        return
    }
    $empty=New-CreateCase 'initial-admin-no-entry'
    $privateFields=[bool](Invoke-Create 'PrivateFieldsReady')
    Check 'Onboarding.Create.MaskedPrivateFields' $privateFields
    [void](Invoke-Create 'EnterPrivate' @('',''))
    [void](Invoke-Create 'Create')
    $facts=[string](Invoke-Create 'Facts')
    Check 'Onboarding.EmptyPin.RefusesBeforeBootstrap' ($facts -ceq 'False|False|Create|False|0|0')
    Check 'Onboarding.EmptyPin.NoDestination' (-not (Test-Path -LiteralPath $empty.Root))
    [void](Invoke-Create 'CloseForm')
    $mismatch=New-CreateCase 'initial-admin-mismatch'
    [void](Invoke-Create 'EnterPrivate' @($secret,$differentEntry))
    [void](Invoke-Create 'Create')
    Check 'Onboarding.Mismatch.NoDestination' (-not (Test-Path -LiteralPath $mismatch.Root))
    [void](Invoke-Create 'CloseForm')
    $cancel=New-CreateCase 'private-entry-cancel'
    [void](Invoke-Create 'EnterPrivate' @($secret,$secret))
    [void](Invoke-Create 'CloseForm')
    Check 'Onboarding.Cancel.NoDestination' (-not (Test-Path -LiteralPath $cancel.Root))
    if($privateFields){
        Check 'Onboarding.Cancel.ClearsBothBeforeUnload' ([bool](Invoke-Create 'ClosedWithoutEntry'))
        Check 'Onboarding.Core.EmptyRefused' ([bool](Run 'invSys.Core.xlam' 'TestInitialAdmin.TypedRejects' @($cancel.Id,$cancel.Root,'','')))
        Check 'Onboarding.Core.MismatchRefused' ([bool](Run 'invSys.Core.xlam' 'TestInitialAdmin.TypedRejects' @($cancel.Id,$cancel.Root,$secret,[guid]::NewGuid().ToString('N'))))
    }
    SelectTarget $Other
    Test-GuideChoice 'config-reader' 4 'IndependentUser'
    if($privateFields){
        $interrupted=New-CreateCase 'interrupted-credential-save'
        [void](Run 'invSys.Core.xlam' 'TestInitialAdmin.ArmSave')
        [void](Invoke-Create 'Create')
        if(-not [bool](Run 'invSys.Core.xlam' 'TestInitialAdmin.SaveWasRefused')){throw 'Credential persistence fault was not reached.'}
        $remaining=@(if(Test-Path -LiteralPath $interrupted.Root){Get-ChildItem -LiteralPath $interrupted.Root -File -Recurse})
        Check 'Onboarding.UncertainCleanup.UnrelatedArrivalPreserved' ($remaining.Count -eq 1 -and $remaining[0].Name -ceq 'unrelated-arrival.txt' -and [IO.File]::ReadAllText($remaining[0].FullName) -ceq 'Independent fixture content')
        Check 'Onboarding.UncertainCleanup.NoCompletionOrPublication' (([string](Invoke-Create 'Facts')).StartsWith('False|False|') -and -not [bool](Run 'invSys.Core.xlam' 'TestInitialAdmin.ExistingPublicationEligible' @($interrupted.Root)))
        [void](Invoke-Create 'CloseForm')
        $late=New-CreateCase 'account-persistence-in-existing-directory'
        [void][IO.Directory]::CreateDirectory($late.Root)
        [void](Run 'invSys.Core.xlam' 'TestInitialAdmin.ArmLateWrite')
        [void](Invoke-Create 'Create')
        $hit=[bool](Run 'invSys.Core.xlam' 'TestInitialAdmin.LateWriteSeen')
        $complete=([string](Invoke-Create 'Facts')).StartsWith('True|True|')
        $eligible=[bool](Run 'invSys.Core.xlam' 'TestInitialAdmin.ExistingPublicationEligible' @($late.Root))
        Check 'Onboarding.AccountProvisioning.NoPublishableFailureAfterCredentialPersistence' (($hit -and -not $complete -and -not $eligible) -or (-not $hit -and $complete -and $eligible))
        [void](Invoke-Create 'CloseForm')
        $otherAuth=Join-Path $Other.Root ($Other.Warehouse+'.invSys.Auth.xlsb')
        [void](Invoke-Create 'ResetTrackingNotices')
        foreach($change in @('SignedOut','Replaced','NewWarehouse','Reader','ReaderAtOpen','WrongStation','WrongWarehouse')){
            if($change -eq 'ReaderAtOpen'){SelectTarget $Other 'config-reader'}else{SelectTarget $Other}
            $origin=(Get-FileHash -LiteralPath $otherAuth).Hash
            [void](Invoke-Create 'PendingChoice' @('config-reader',$change))
            switch($change){
                'SignedOut' {[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
                'Replaced' {SelectTarget $Other}
                'NewWarehouse' {SelectTarget $Fixture}
                'Reader' {SelectTarget $Other 'config-reader'}
            }
            [void](Invoke-Create 'SubmitPendingChoice')
            Check "Onboarding.Users.${change}PreservesStoredRoles" ($origin -ceq (Get-FileHash -LiteralPath $otherAuth).Hash)
        }
        Check 'Onboarding.Users.StaleSessionTrackingNoticePresented' ([int](Invoke-Create 'TrackingNotices') -gt 0)
    }
    # Characterize the existing bootstrap recovery boundary independently of
    # the future credential writer. Never manufacture an Auth row or a grant.
    foreach($existing in @($false,$true)){
        $kind=if($existing){'ExistingHub'}else{'NewHub'}
        $case=New-CreateCase ('initial-admin-'+$kind)
        if($existing){
            [void][IO.Directory]::CreateDirectory($case.Root)
            $inbox=Join-Path $case.Root 'inbox'
            [void][IO.Directory]::CreateDirectory($inbox)
            $sentinel=Join-Path $inbox 'unrelated.txt'
            [IO.File]::WriteAllText($sentinel,'Preserve this fixture file.')
            $before=(Get-FileHash -LiteralPath $sentinel).Hash
        }
        [void](Run 'invSys.Core.xlam' 'TestInitialAdmin.Arm')
        [void](Invoke-Create 'EnterPrivate' @($secret,$secret))
        [void](Invoke-Create 'Create')
        $hit=[bool](Run 'invSys.Core.xlam' 'TestInitialAdmin.Hit')
        if(-not $hit){throw 'Destination account-write fault was not reached; recovery result unavailable.'}
        $facts=[string](Invoke-Create 'Facts')
        Check ("Onboarding.$kind.ReportsCreationFailure") ($facts.StartsWith('False|False|') -and $facts.EndsWith('|1|0'))
        $eligible=[bool](Run 'invSys.Core.xlam' 'TestInitialAdmin.ExistingPublicationEligible' @($case.Root))
        Check ("Onboarding.$kind.IncompleteDestinationNotPublishable") (-not $eligible)
        if($existing){
            Check 'Onboarding.ExistingHub.PreservesUnrelatedFile' ((Get-FileHash -LiteralPath $sentinel).Hash -ceq $before)
            $remaining=@(Get-ChildItem -LiteralPath $case.Root -Force -Recurse)
            Check 'Onboarding.ExistingHub.OnlyOriginalEntriesRemain' ($remaining.Count -eq 2 -and (Test-Path -LiteralPath $inbox))
            [void](Invoke-Create 'EnterPrivate' @($secret,$secret))
            [void](Invoke-Create 'Create')
            $retry=[string](Invoke-Create 'Facts')
            Check 'Onboarding.ExistingHub.OrdinaryRetryCreatesWarehouse' ($retry -ceq 'True|True|Close|True|2|0')
            if($retry -ceq 'True|True|Close|True|2|0'){
                Test-CreatedDestination $case
                continue
            }
        }else{
            Check 'Onboarding.NewHub.RemovesOwnedRoot' (-not (Test-Path -LiteralPath $case.Root))
        }
        [void](Invoke-Create 'CloseForm')
    }
}
