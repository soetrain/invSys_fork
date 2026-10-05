# Corruption is confined to generated fixture Config files and restored byte-for-byte.
function Test-UserPolicyStorage($Fixture,$Other) {
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCapture' @($true))
    [void](UserPolicyControl 'btnSaveTrackingPolicy' 'Click')
    $baseline=Join-Path $runRoot 'user-policy-config-baseline.xlsb'
    Copy-Item -LiteralPath $Fixture.Config -Destination $baseline
    $baselineHash=(Get-FileHash -LiteralPath $baseline).Hash
    $version=Get-TrackingPolicyVersion $Fixture
    foreach($case in @('MissingUsers','MissingCount','CountMismatch','InvalidFlag','DuplicateUser','OrphanVersion','UnknownSchema')) {
        $book=$excel.Workbooks.Open($Fixture.Config,0,$false)
        try {
            $headers=Table $book 'tblEventTrackingPolicies';$users=Table $book 'tblEventTrackingUsers'
            $selected=0
            for($index=1;$index -le $headers.ListRows.Count;$index++){if($headers.DataBodyRange.Cells.Item($index,$headers.ListColumns.Item('PolicyVersion').Index).Value2 -eq $version){$selected=$index}}
            if(-not $selected){throw 'Current user policy fixture missing.'}
            $userRow=0
            for($index=1;$index -le $users.ListRows.Count;$index++){if($users.DataBodyRange.Cells.Item($index,$users.ListColumns.Item('PolicyVersion').Index).Value2 -eq $version){$userRow=$index;break}}
            if(-not $userRow){throw 'Disabled-user storage fixture missing.'}
            switch($case) {
                MissingUsers {
                    $users.Delete()
                    $remaining=0
                    foreach($sheet in $book.Worksheets){foreach($table in $sheet.ListObjects){if($table.Name -ceq 'tblEventTrackingUsers'){$remaining++}}}
                    if($remaining){throw 'Missing-user-table fixture failed to remove its table.'}
                }
                MissingCount {$headers.ListColumns.Item('UserCount').Delete()}
                CountMismatch {$headers.DataBodyRange.Cells.Item($selected,$headers.ListColumns.Item('UserCount').Index).Value2=99.0}
                InvalidFlag {$cell=$users.DataBodyRange.Cells.Item($userRow,$users.ListColumns.Item('Record').Index);$cell.NumberFormat='@';$cell.Value2='False'}
                DuplicateUser {
                    $row=$users.ListRows.Add()
                    $row.Range.Cells.Item(1,$users.ListColumns.Item('PolicyVersion').Index).Value2=[double]$version
                    $row.Range.Cells.Item(1,$users.ListColumns.Item('UserId').Index).Value2='CONFIG-READER'
                    $row.Range.Cells.Item(1,$users.ListColumns.Item('Record').Index).Value2=$false
                }
                OrphanVersion {$users.DataBodyRange.Cells.Item($userRow,$users.ListColumns.Item('PolicyVersion').Index).Value2=2147483646.0}
                UnknownSchema {$headers.DataBodyRange.Cells.Item($selected,$headers.ListColumns.Item('SchemaVersion').Index).Value2=999.0}
            }
            $book.Save()
        } finally {$book.Close($false)}
        try {
            $broken=(Get-FileHash -LiteralPath $Fixture.Config).Hash
            [void](UserPolicyControl 'btnReloadTrackingPolicy' 'Click')
            $unavailable=(Get-TrackingPolicyRequest) -ceq ''
            $captureOff=(Run 'invSys.Admin.xlam' 'TestD5Commands.UserPolicyRecording' @('Status')) -ceq 'False|False'
            $preserved=$broken -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash
            [pscustomobject]@{Case=$case;EditorUnavailable=$unavailable;CaptureUnavailable=$captureOff;BytesPreserved=$preserved}|ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'user-policy-storage-facts.jsonl')
            Check ('UserPolicy.Storage.'+$case+'.FailsClosedWithoutRepair') ($unavailable -and $captureOff -and $preserved)
        } finally {
            Copy-Item -LiteralPath $baseline -Destination $Fixture.Config -Force
            [void](UserPolicyControl 'btnReloadTrackingPolicy' 'Click')
        }
    }
    Check 'UserPolicy.Storage.ValidFixtureRestored' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $baselineHash -and (UserPolicyDisabled (Get-TrackingPolicyRequest) 'config-reader'))

    [void](UserPolicyControl 'lstTrackingUsers' 'Select' 'config-admin')
    [void](UserPolicyControl 'chkUserRecord' 'Write' 'False')
    $staged=Get-TrackingPolicyRequest;$events=$excel.EnableEvents
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCancelSave' @($Fixture.Config))
    try {
        $excel.EnableEvents=$true
        [void](UserPolicyControl 'btnSaveTrackingPolicy' 'Click')
        $cancelled=[int](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCancelledCount')
    } finally {
        $excel.EnableEvents=$events
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyStopCancelling')
    }
    Check 'UserPolicy.Storage.CancelledSaveKeepsBytesAndStaging' ($cancelled -eq 1 -and $baselineHash -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash -and $staged -ceq (Get-TrackingPolicyRequest) -and (Get-TrackingPolicyVersion $Fixture) -eq $version)
    [void](UserPolicyControl 'btnReloadTrackingPolicy' 'Click')

    $book=$excel.Workbooks.Open($Fixture.Config,0,$false)
    try {
        $users=Table $book 'tblEventTrackingUsers';$column=$users.ListColumns.Add(1);$column.Name='User Extra';$column.DataBodyRange.Value2='preserved user note'
        $prior=$users.DataBodyRange.Value2;$priorCount=$users.ListRows.Count
        $book.Save()
    } finally {$book.Close($false)}
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCapture' @($false))
    [void](UserPolicyControl 'btnSaveTrackingPolicy' 'Click')
    $book=$excel.Workbooks.Open($Fixture.Config,0,$true)
    try {
        $users=Table $book 'tblEventTrackingUsers';$same=$users.ListRows.Count -gt $priorCount
        for($r=1;$r -le $priorCount;$r++){for($c=1;$c -le $users.ListColumns.Count;$c++){$same=$same -and $users.DataBodyRange.Cells.Item($r,$c).Value2 -ceq $prior.GetValue($r,$c)}}
        Check 'UserPolicy.Storage.AppendsPreservingHistoryAndUnknownColumns' $same
    } finally {$book.Close($false)}
    Test-UserPolicyLegacy $Fixture $Other
}

function Test-UserPolicyLegacy($Fixture,$Other) {
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCloseAction')
    SelectTarget $Other
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
    $legacy=Get-TrackingPolicyRequest|ConvertFrom-Json;$legacy.SchemaVersion=1;$legacy.PSObject.Properties.Remove('Users')
    $context=[string](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyContext')
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicySaveDirect' @($context,0,($legacy|ConvertTo-Json -Depth 8 -Compress)))
    Check 'UserPolicy.Legacy.V1RequestAcceptedBeforeUpgrade' ($ok -and (Get-TrackingPolicyVersion $Other) -eq 1)
    if(-not $ok){throw 'Legacy policy fixture save failed.'}
    $pin=(Get-FileHash -LiteralPath $Other.Config).Hash
    [void](UserPolicyControl 'btnReloadTrackingPolicy' 'Click')
    $model=Get-TrackingPolicyRequest|ConvertFrom-Json
    Check 'UserPolicy.Legacy.V1ProjectsV2WithoutWriting' ($model.SchemaVersion -eq 2 -and @($model.Users).Count -eq 0 -and $pin -ceq (Get-FileHash -LiteralPath $Other.Config).Hash)
    [void](UserPolicyControl 'lstTrackingUsers' 'Select' 'config-reader')
    [void](UserPolicyControl 'chkUserRecord' 'Write' 'False')
    [void](UserPolicyControl 'btnSaveTrackingPolicy' 'Click')
    $book=$excel.Workbooks.Open($Other.Config,0,$true)
    try {
        $headers=Table $book 'tblEventTrackingPolicies'
        Check 'UserPolicy.Legacy.UpgradeAppendsWithoutRewritingV1' ($headers.ListRows.Count -eq 2 -and $headers.DataBodyRange.Cells.Item(1,$headers.ListColumns.Item('SchemaVersion').Index).Value2 -eq 1 -and $headers.DataBodyRange.Cells.Item(2,$headers.ListColumns.Item('SchemaVersion').Index).Value2 -eq 2)
    } finally {$book.Close($false)}
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyCloseAction')
    SelectTarget $Fixture
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
}
