# A1 user policy: drive live packaged controls; never fabricate saved policy.
# Only fixed check names/Booleans leave the generated fixture in reports.
function Install-UserTrackingPolicyProbe($TestModule) {
    $TestModule.CodeModule.AddFromString(@'
Public Function UserPolicyControl(ByVal name As String, ByVal command As String, Optional ByVal value As String = "") As String
    Dim control As Object, kind As String, index As Long
    Select Case name
        Case "lstTrackingUsers": kind = "ListBox"
        Case "chkUserRecord": kind = "CheckBox"
        Case Else: kind = "CommandButton"
    End Select
    Set control = TrackingControl(mForm, kind, "", name)
    If control Is Nothing Then UserPolicyControl = "MISSING": Exit Function
    Select Case command
        Case "Exists": UserPolicyControl = "PRESENT"
        Case "Select"
            control.ListIndex = -1
            For index = 0 To control.ListCount - 1
                If StrComp(CStr(control.List(index, 0)), value, vbTextCompare) = 0 Then
                    control.ListIndex = index: UserPolicyControl = "SELECTED": Exit Function
                End If
            Next index
            UserPolicyControl = "ABSENT"
        Case "Read": UserPolicyControl = CStr(control.Value)
        Case "Write": control.Value = CBool(value): UserPolicyControl = "SET"
        Case "Click": control.Value = True: UserPolicyControl = "CLICKED"
    End Select
End Function
Public Function UserPolicyRecording(ByVal command As String) As String
    Dim notice As String, active As Boolean, canStart As Boolean, context As String
    context = modActivity.CaptureContext()
    If command = "Status" Then
        notice = modActionRecording.Status(context, active, canStart)
        UserPolicyRecording = CStr(active) & "|" & CStr(canStart)
    Else
        UserPolicyRecording = CStr(modActionRecording.Control(context, command, notice))
    End If
End Function
Public Function UserPolicyCanRead(ByVal id As String) As Boolean
    UserPolicyCanRead = (Left$(modActivity.ReadActivityRecord(id), 3) = "OK|")
End Function
'@)
}

function UserPolicyControl([string]$Name,[string]$Command,[string]$Value='') {
    [string](Run 'invSys.Admin.xlam' 'TestD5Commands.UserPolicyControl' @($Name,$Command,$Value))
}
function UserPolicyDisabled([string]$Request,[string]$User) {
    $model=$Request|ConvertFrom-Json
    if(-not $model.PSObject.Properties['Users']){return $false}
    $rows=@($model.Users|Where-Object UserId -IEQ $User)
    return ($rows.Count -eq 1 -and $rows[0].Record -is [bool] -and -not $rows[0].Record)
}

function Test-UserTrackingPolicy($Fixture,$Other) {
    $before=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $auth=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $authBefore=(Get-FileHash -LiteralPath $auth).Hash
    $otherBefore=(Get-FileHash -LiteralPath $Other.Config).Hash
    $request=Get-TrackingPolicyRequest;$model=$request|ConvertFrom-Json
    $v2=$model.SchemaVersion -eq 2 -and $null -ne $model.PSObject.Properties['Users']
    Check 'UserPolicy.EditorProjectsV2' $v2
    Check 'UserPolicy.DefaultUsersEnabled' ($v2 -and @($model.Users).Count -eq 0)
    $list=(UserPolicyControl 'lstTrackingUsers' 'Exists') -ceq 'PRESENT'
    $flag=(UserPolicyControl 'chkUserRecord' 'Exists') -ceq 'PRESENT'
    Check 'UserPolicy.UserSelectorPresent' $list
    Check 'UserPolicy.RecordUserPresent' $flag
    $selected=(UserPolicyControl 'lstTrackingUsers' 'Select' 'config-reader') -ceq 'SELECTED'
    Check 'UserPolicy.SelectsWarehouseUser' $selected
    Check 'UserPolicy.SelectedUserDefaultsEnabled' ((UserPolicyControl 'chkUserRecord' 'Read') -ceq 'True')
    [void](UserPolicyControl 'chkUserRecord' 'Write' 'False')
    $staged=Get-TrackingPolicyRequest
    $disabled=UserPolicyDisabled $staged 'config-reader'
    Check 'UserPolicy.ToggleStagesDisabledUser' $disabled
    Check 'UserPolicy.StagingDoesNotWriteAuthorities' ($before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash -and $authBefore -ceq (Get-FileHash -LiteralPath $auth).Hash)
    $context=[string](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyContext')
    foreach($case in @('UnknownField','DuplicateUser','InvalidRecord','BlankUser','UntrimmedUser','UnknownUser')) {
        $bad=$request|ConvertFrom-Json
        $bad.SchemaVersion=2
        $bad|Add-Member -Force NoteProperty Users @([pscustomobject]@{UserId='config-reader';Record=$false})
        switch($case) {
            UnknownField {$bad.Users[0]|Add-Member NoteProperty Unexpected $true}
            DuplicateUser {$bad.Users+=,[pscustomobject]@{UserId='CONFIG-READER';Record=$true}}
            InvalidRecord {$bad.Users[0].Record='False'}
            BlankUser {$bad.Users[0].UserId=''}
            UntrimmedUser {$bad.Users[0].UserId=' config-reader '}
            UnknownUser {$bad.Users[0].UserId='not-a-warehouse-user'}
        }
        $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicySaveDirect' @($context,0,($bad|ConvertTo-Json -Depth 8 -Compress)))
        Check ('UserPolicy.Reject'+$case) (-not $ok -and $before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    }
    $entries=[int](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyEntries')
    if($disabled){[void](UserPolicyControl 'btnSaveTrackingPolicy' 'Click')}
    $version=Get-TrackingPolicyVersion $Fixture
    Check 'UserPolicy.RealSaveHandler' ($disabled -and [int](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyEntries') -eq $entries+1)
    Check 'UserPolicy.SavesOnePolicyVersion' ($disabled -and $version -eq 1)
    [void](UserPolicyControl 'btnReloadTrackingPolicy' 'Click')
    $saved=Get-TrackingPolicyRequest
    Check 'UserPolicy.DisabledFlagReloads' (UserPolicyDisabled $saved 'config-reader')
    [void](UserPolicyControl 'lstTrackingUsers' 'Select' 'config-reader')
    Check 'UserPolicy.SavedFlagRendered' ((UserPolicyControl 'chkUserRecord' 'Read') -ceq 'False')
    if($version -eq 1) {
        $legacy=$saved|ConvertFrom-Json;$legacy.SchemaVersion=1;$legacy.PSObject.Properties.Remove('Users')
        $pin=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicySaveDirect' @($context,$version,($legacy|ConvertTo-Json -Depth 8 -Compress)))
        Check 'UserPolicy.DownlevelWriteCannotEraseOverrides' (-not $ok -and $pin -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    } else { Check 'UserPolicy.DownlevelWriteCannotEraseOverrides' $false }
    $pin=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    [void](UserPolicyControl 'btnResetTrackingPolicy' 'Click')
    $reset=Get-TrackingPolicyRequest|ConvertFrom-Json
    $empty=$null -ne $reset.PSObject.Properties['Users']
    if($empty){$empty=@($reset.Users).Count -eq 0}
    Check 'UserPolicy.ResetStagesDefaultsOnly' ($empty -and $pin -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCloseAction')
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
    Check 'UserPolicy.CloseDiscardsUserEdits' ((UserPolicyDisabled (Get-TrackingPolicyRequest) 'config-reader') -and $pin -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    Check 'UserPolicy.AuthAndOtherWarehouseUnchanged' ($authBefore -ceq (Get-FileHash -LiteralPath $auth).Hash -and $otherBefore -ceq (Get-FileHash -LiteralPath $Other.Config).Hash)
    Test-UserPolicyRecording $Fixture
}

function Test-UserPolicyRecording($Fixture) {
    . (Join-Path $PSScriptRoot 'Slice4beActivityAssertions.ps1')
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCapture' @($true))
    [void](UserPolicyControl 'btnSaveTrackingPolicy' 'Click')
    Check 'UserPolicy.EnabledActorCanRecord' ((Run 'invSys.Admin.xlam' 'TestD5Commands.UserPolicyRecording' @('Status')) -ceq 'False|True')
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCloseAction')
    SelectTarget $Fixture 'config-reader'
    Check 'UserPolicy.DisabledActorCannotStartCapture' ((Run 'invSys.Admin.xlam' 'TestD5Commands.UserPolicyRecording' @('Status')) -ceq 'False|False')
    $started=(Run 'invSys.Admin.xlam' 'TestD5Commands.UserPolicyRecording' @('Start')) -ceq 'True'
    Check 'UserPolicy.DisabledActorStartRefused' (-not $started)
    if($started){[void](Run 'invSys.Admin.xlam' 'TestD5Commands.UserPolicyRecording' @('Stop'))}
    SelectTarget $Fixture
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
    $before=@(Get-Slice4beActivityFiles $Fixture)
    $saved=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('BatchSize','730'))
    $fresh=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -notin $before}|ForEach-Object {Get-Content -LiteralPath $_ -Raw|ConvertFrom-Json})
    $history=@($fresh|Where-Object ControlId -CEQ 'ADMIN_SETTINGS_SAVE_VALUE')
    Check 'UserPolicy.EnabledActorOrdinaryOwnerRecords' ($saved -and $history.Count -eq 2)
    $started=(Run 'invSys.Admin.xlam' 'TestD5Commands.UserPolicyRecording' @('Start')) -ceq 'True'
    Check 'UserPolicy.RecordingFixtureStarted' $started
    [void](UserPolicyControl 'lstTrackingUsers' 'Select' 'config-admin')
    [void](UserPolicyControl 'chkUserRecord' 'Write' 'False')
    [void](UserPolicyControl 'btnSaveTrackingPolicy' 'Click')
    $journal=Join-Path $Fixture.Root ('Training/ActionPaths/'+$Fixture.Warehouse)
    $closed=@(Get-ChildItem -LiteralPath $journal -File -Filter '*.json'|ForEach-Object {Get-Content -LiteralPath $_.FullName -Raw|ConvertFrom-Json}|Where-Object {$_.RecordType -ceq 'Close' -and $_.CreatedByUserId -ceq 'config-admin' -and $_.Lifecycle -ceq 'Incomplete' -and $_.ReasonCode -ceq 'POLICY_CHANGED'})
    Check 'UserPolicy.PolicyChangeEndsRecordingPartial' ($started -and $closed.Count -eq 1)
    Check 'UserPolicy.DisabledAdminCaptureUnavailable' ((Run 'invSys.Admin.xlam' 'TestD5Commands.UserPolicyRecording' @('Status')) -ceq 'False|False')
    $before=@(Get-Slice4beActivityFiles $Fixture)
    $saved=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('BatchSize','731'))
    Check 'UserPolicy.DisabledActorOrdinaryWorkContinues' ($saved -and [long](Run 'invSys.Core.xlam' 'modConfig.GetLong' @('BatchSize',0)) -eq 731)
    Check 'UserPolicy.DisabledActorHasNoOptionalActivity' (@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -notin $before}).Count -eq 0)
    $readable=$history.Count -eq 2
    foreach($row in $history){$one=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.UserPolicyCanRead' @($row.RecordId));$readable=$readable -and $one}
    Check 'UserPolicy.DisabledRecordingKeepsHistoryVisible' $readable
    [void](UserPolicyControl 'lstTrackingUsers' 'Select' 'config-admin')
    [void](UserPolicyControl 'chkUserRecord' 'Write' 'True')
    [void](UserPolicyControl 'btnSaveTrackingPolicy' 'Click')
    Check 'UserPolicy.AdminCanReenableOwnRecording' ((Run 'invSys.Admin.xlam' 'TestD5Commands.UserPolicyRecording' @('Status')) -ceq 'False|True')
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCapture' @($false))
    [void](UserPolicyControl 'btnSaveTrackingPolicy' 'Click')
}
