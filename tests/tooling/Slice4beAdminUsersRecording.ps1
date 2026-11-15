# Actual packaged form handlers and real Auth persistence. No target-account
# payload, credential or activity body is emitted to diagnostic reports.
. (Join-Path $PSScriptRoot 'Slice4beActivityAssertions.ps1')
. (Join-Path $PSScriptRoot 'Slice4beAdminUsersRecordingStorage.ps1')
. (Join-Path $PSScriptRoot 'Slice4beAdminUsersRecordingContext.ps1')
function Install-AdminUsersRecordingProbe {
    $probe=$packages['invSys.Core.xlam'].VBProject.VBComponents.Add(1);$probe.Name='TestUsersRecording'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Public Function CurrentCatalogVersion() As Long
    CurrentCatalogVersion = modActivityCatalog.CATALOG_VERSION
End Function
Public Function Policy(Optional ByVal mode As String = "Enabled") As Boolean
    Dim context As String, version As Long, wire As String, report As String
    Dim model As Object, row As Variant, user As Object
    context = modActivity.CaptureContext()
    If context = "" Then Exit Function
    If Not modTrackingPolicySettings.ReadEditor(context, version, wire, report) Then Exit Function
    Set model = modTrackingPolicyModel.Defaults(mode <> "Disabled")
    If mode = "Capture" Then model("ViewerActionPathCaptureEnabled") = True
    If mode = "ControlOff" Then
        For Each row In model("Controls")
            If Left$(row("ControlId"), 11) = "ADMIN_USER_" Then row("Collect") = False
        Next row
    ElseIf mode = "UserOff" Then
        Set user = CreateObject("Scripting.Dictionary")
        user.Add "UserId", modAuth.GetCurrentUserId(): user.Add "Record", False
        model("Users").Add user
    End If
    Policy = modTrackingPolicySettings.SavePolicy(context, version, modTrainingJson.EncodeObject(model), report)
End Function
Public Function Compatible() As Boolean
    Dim id As Variant, before As Object, after As Object, ids As Variant
    ids = modActivityCatalog.ControlIds(37)
    If UBound(ids) <> 156 Or UBound(modActivityCatalog.ControlIds()) <> 159 Then Exit Function
    For Each id In ids
        Set before = modActivityCatalog.Control(CStr(id), 37)
        Set after = modActivityCatalog.Control(CStr(id))
        If before Is Nothing Or after Is Nothing Then Exit Function
        If modTrainingJson.EncodeObject(before) <> modTrainingJson.EncodeObject(after) Then Exit Function
    Next id
    For Each id In Array("ADMIN_USER_SAVE", "ADMIN_USER_DEACTIVATE", "ADMIN_USER_DELETE")
        If Not modActivityCatalog.Control(CStr(id), 37) Is Nothing Then Exit Function
    Next id
    Compatible = True
End Function
Public Function Concludes(ByVal wire As String) As Boolean
    Concludes = modEvaluationMatches.CommandCompleted(modTrainingJson.DecodeObject(wire))
End Function
'@)
    Install-AdminUsersRecordingStorageProbe
    Install-AdminUsersRecordingContextProbe
}

function Test-AdminUsersRecording($Fixture,$Other) {
    $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    if(-not [IO.Path]::GetFullPath($authPath).StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Recording requires generated Auth fixture.'}
    $original=[IO.File]::ReadAllBytes($authPath);$originalHash=Get-AuthReadFixtureHash $authPath
    $auth=$null
    SelectTarget $Fixture
    if(-not [bool](Run 'invSys.Core.xlam' 'TestUsersRecording.Policy')){throw 'Eligible recording policy fixture unavailable; not product RED.'}
    Check 'UsersRecording.Catalog.Prior157DefinitionsUnchanged' ([bool](Run 'invSys.Core.xlam' 'TestUsersRecording.Compatible'))
    try {
        foreach($case in @('Saved','Dirty','SaveCancelled','ReaderDenied','Cancelled')){
            foreach($action in @('Save','Deactivate','Delete')){
                if($case -ceq 'Cancelled' -and $action -cne 'Delete'){continue}
                [IO.File]::WriteAllBytes($authPath,$original)
                SelectTarget $Fixture
                if($case -ceq 'ReaderDenied'){SelectTarget $Fixture 'config-reader'}
                [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.SetCancelDelete' @($case -ceq 'Cancelled'))
                if(-not [bool](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.LoadForTest' @($authPath))){throw 'Generated form selection unavailable.'}
                if($case -in @('Dirty','SaveCancelled')){
                    $auth=$excel.Workbooks.Open($authPath,0,$false)
                    if($case -eq 'Dirty'){
                        if(-not [bool](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.BorrowedCanaryForTest' @($auth.Name,$true))){throw 'Dirty fixture unavailable.'}
                    }else{[void](Run 'invSys.Core.xlam' 'TestUsersAuthority.CancelSave' @($authPath))}
                }
                $before=@(Get-Slice4beActivityFiles $Fixture)
                $wasEnabled=$excel.EnableEvents
                try {
                    if($case -eq 'SaveCancelled'){$excel.EnableEvents=$true}
                    [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.WriteReaderForTest' @($action))
                }finally{$excel.EnableEvents=$wasEnabled}
                $changed=(Get-AuthReadFixtureHash $authPath) -cne $originalHash
                $owner=if($case -ceq 'Saved'){$changed}elseif($case -in @('Dirty','SaveCancelled')){-not $changed -and $null -ne $auth -and -not $auth.Saved}else{-not $changed}
                if($case -eq 'SaveCancelled'){$owner=$owner -and [bool](Run 'invSys.Core.xlam' 'TestUsersAuthority.CancelSeen')}
                Check ('UsersRecording.'+$case+'.'+$action+'.OwnerPersistence') $owner
                if(-not $owner){throw 'Owner fixture failed; not missing-record RED.'}
                $id='ADMIN_USER_'+$action.ToUpperInvariant()
                $outcome=switch($case){'Saved'{'SAVED'}'Dirty'{'REJECTED'}'SaveCancelled'{'FAILED'}'ReaderDenied'{'DENIED'}'Cancelled'{'CANCELLED'}}
                $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content -LiteralPath $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -CEQ $id)
                $requests=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$terminals=@($records|Where-Object OutcomeCode -CEQ $outcome)
                $pair=$records.Count -eq 2 -and $requests.Count -eq 1 -and $terminals.Count -eq 1
                Check ('UsersRecording.'+$case+'.'+$action+'.CorrelatedPair') ($pair -and $requests[0].ActivityId -ceq $terminals[0].ActivityId -and $requests[0].RecordId -cne $terminals[0].RecordId)
                $safe=$pair;$bound=$pair;$concludes=$pair
                foreach($record in $records){
                    $json=$record|ConvertTo-Json -Depth 20 -Compress
                    # Only the required actor field may name the signed-in Reader.
                    $privacyWire=if($case -ceq 'ReaderDenied'){$record|Select-Object * -ExcludeProperty UserId|ConvertTo-Json -Depth 20 -Compress}else{$json}
                    foreach($value in @($Fixture.Root,$Fixture.Secret,(CredentialHash $Fixture.Secret),'config-reader','Updated fixture reader','PinHash','SaveUser','OperatorAuthReadCanary')){$safe=$safe -and -not $privacyWire.Contains($value)}
                    $effect=if($record.OutcomeCode -cin @('REJECTED','CANCELLED')){'Unchanged'}else{'Unknown'}
                    $actor=if($case -ceq 'ReaderDenied'){'config-reader'}else{'config-admin'}
                    $bound=$bound -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.StationId -ceq 'S1' -and $record.UserId -ceq $actor -and $record.OwnerId -ceq 'ADMIN_USERS_ROLES' -and $record.DataEffect -ceq $effect -and @($record.SourceEventRefs).Count -eq 0
                    $concludes=$concludes -and ([bool](Run 'invSys.Core.xlam' 'TestUsersRecording.Concludes' @($json)) -eq ($record.OutcomeCode -ceq 'SAVED'))
                }
                Check ('UsersRecording.'+$case+'.'+$action+'.NoAccountOrCredentialPayload') $safe
                Check ('UsersRecording.'+$case+'.'+$action+'.ActorEffectAndNoSourceClaim') $bound
                Check ('UsersRecording.'+$case+'.'+$action+'.OnlySavedConcludes') $concludes
                [void](Run 'invSys.Core.xlam' 'TestUsersAuthority.ReleaseSaveGate')
                [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.CloseForTest')
                if($null -ne $auth){$auth.Close($false);$auth=$null}
            }
        }
        [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.SetCancelDelete' @($false))
        foreach($mode in @('Disabled','ControlOff','UserOff')){
            foreach($action in @('Save','Deactivate','Delete')){
                [IO.File]::WriteAllBytes($authPath,$original);SelectTarget $Fixture
                if(-not [bool](Run 'invSys.Core.xlam' 'TestUsersRecording.Policy' @($mode))){throw 'Suppression policy fixture unavailable.'}
                if(-not [bool](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.LoadForTest' @($authPath))){throw 'Suppression form unavailable.'}
                $before=@(Get-Slice4beActivityFiles $Fixture)
                [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.WriteReaderForTest' @($action))
                Check ('UsersRecording.'+$mode+'.'+$action+'.BusinessStillPersists') ((Get-AuthReadFixtureHash $authPath) -cne $originalHash)
                Check ('UsersRecording.'+$mode+'.'+$action+'.NoOptionalRecords') (@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}).Count -eq 0)
                Check ('UsersRecording.'+$mode+'.'+$action+'.NoTrackingWarning') ([int](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.TrackingNotices') -eq 0)
                [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.CloseForTest')
            }
        }
    }finally{
        [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.SetCancelDelete' @($false))
        [void](Run 'invSys.Core.xlam' 'TestUsersAuthority.ReleaseSaveGate')
        [void](Run 'invSys.Admin.xlam' 'TestAdminAuthRefresh.CloseForTest')
        if($null -ne $auth){$auth.Close($false)}
        [IO.File]::WriteAllBytes($authPath,$original)
        SelectTarget $Fixture
    }
    Check 'UsersRecording.AuthFixtureRestored' ((Get-AuthReadFixtureHash $authPath) -ceq $originalHash)
    Test-AdminUsersRecordingStorage $Fixture $Other
    Test-AdminUsersRecordingContext $Fixture $Other
}
