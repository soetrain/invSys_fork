function Install-UserPolicyAccessProbe($TestModule) {
    $TestModule.CodeModule.AddFromString(@'
Public Function UserPolicyRoster(ByVal request As String) As String
    Dim projection As String, report As String
    UserPolicyRoster = "DENIED"
    If modTrackingUserSettings.ReadUsers(modActivity.CaptureContext(), request, projection, report) Then UserPolicyRoster = projection
End Function
Public Function UserPolicyPersonalProjection() As String
    Dim choice As String, effective As String, evidence As String, report As String, version As Long, request As String
    UserPolicyPersonalProjection = "DENIED"
    If modActionPathPreference.ReadPreference(modActivity.CaptureContext(), choice, effective, evidence, report, version, request) Then UserPolicyPersonalProjection = request
End Function
'@)
}

function Test-UserPolicyAccess($Fixture) {
    $request=Get-TrackingPolicyRequest
    $pin=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $version=Get-TrackingPolicyVersion $Fixture
    $roster=[string](Run 'invSys.Admin.xlam' 'TestD5Commands.UserPolicyRoster' @($request))
    $model=$roster|ConvertFrom-Json
    $fields=@($model.PSObject.Properties).Count -eq 1 -and $null -ne $model.PSObject.Properties['Users']
    foreach($row in $model.Users){$fields=$fields -and @($row.PSObject.Properties).Count -eq 2 -and $null -ne $row.PSObject.Properties['UserId'] -and $null -ne $row.PSObject.Properties['Available']}
    Check 'UserPolicy.Access.AdminRosterHasOnlyIdentityAndAvailability' ($fields -and -not $roster.Contains($Fixture.Secret))
    $personal=([string](Run 'invSys.Admin.xlam' 'TestD5Commands.UserPolicyPersonalProjection'))|ConvertFrom-Json
    Check 'UserPolicy.Access.AdminPersonalViewExcludesOtherUserFlags' (@($personal.Users).Count -eq 0)
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCloseAction')
    SelectTarget $Fixture 'config-reader'
    $context=[string](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyContext')
    Check 'UserPolicy.Access.ReaderCannotReadRoster' ((Run 'invSys.Admin.xlam' 'TestD5Commands.UserPolicyRoster' @($request)) -ceq 'DENIED')
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicySaveDirect' @($context,$version,$request))
    Check 'UserPolicy.Access.ReaderCannotWritePolicy' (-not $ok -and $pin -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    $personal=([string](Run 'invSys.Admin.xlam' 'TestD5Commands.UserPolicyPersonalProjection'))|ConvertFrom-Json
    Check 'UserPolicy.Access.ReaderSeesOnlyOwnFlag' (@($personal.Users).Count -eq 1 -and $personal.Users[0].UserId -ieq 'config-reader' -and -not $personal.Users[0].Record)
    SelectTarget $Fixture
    $auth=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $authHash=(Get-FileHash -LiteralPath $auth).Hash
    $book=$excel.Workbooks.Open($auth,0,$false)
    try {
        $cell=$book.Worksheets.Item(1).Cells.Item(1,20);$cell.Value2='unsaved roster fixture'
        $denied=(Run 'invSys.Admin.xlam' 'TestD5Commands.UserPolicyRoster' @($request)) -ceq 'DENIED'
        $memoryPreserved=$denied -and -not $book.Saved -and $cell.Value2 -ceq 'unsaved roster fixture'
    } finally {$book.Close($false)}
    Check 'UserPolicy.Access.DirtyAuthIsUnavailableAndPreserved' ($memoryPreserved -and $authHash -ceq (Get-FileHash -LiteralPath $auth).Hash)
    $hidden=$auth+'.unavailable'
    $fixtureRoot=[IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\'
    if(-not [IO.Path]::GetFullPath($auth).StartsWith($fixtureRoot,[StringComparison]::OrdinalIgnoreCase) -or -not [IO.Path]::GetFullPath($hidden).StartsWith($fixtureRoot,[StringComparison]::OrdinalIgnoreCase)){throw 'Auth fixture paths escaped their generated root.'}
    Move-Item -LiteralPath $auth -Destination $hidden
    try {
        Check 'UserPolicy.Access.MissingAuthIsUnavailableWithoutCreation' ((Run 'invSys.Admin.xlam' 'TestD5Commands.UserPolicyRoster' @($request)) -ceq 'DENIED' -and -not (Test-Path -LiteralPath $auth))
    } finally {Move-Item -LiteralPath $hidden -Destination $auth}
    Check 'UserPolicy.Access.AuthAndConfigBytesPreserved' ($authHash -ceq (Get-FileHash -LiteralPath $auth).Hash -and $pin -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
}
