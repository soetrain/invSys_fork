# Unsaved instrumentation at the existing form's real pending-status yield.
# Only generated fixture authorization changes; no business owner is replaced here.
function Install-Slice4beShippingAccessInterruptionProbe($Form,$Module) {
    $source=$Form.Lines(1,$Form.CountOfLines)
    $source=$source.Replace('Option Explicit',"Option Explicit`r`nPrivate ActivityPendingPermissionLoss As Boolean")
    $pattern='(?s)(Private Sub ShowPersistencePending\(ByVal messageText As String\).*?    DoEvents)'
    if([regex]::Matches($source,$pattern).Count -ne 1){throw 'Shipping permission-yield anchor unavailable.'}
    $source=[regex]::Replace($source,$pattern,'$1'+"`r`n    If ActivityPendingPermissionLoss Then`r`n        ActivityPendingPermissionLoss = False`r`n        modTS_Shipments.ActivityShippingRevokeAtPending`r`n        ActivityPendingInterruptions = ActivityPendingInterruptions + 1`r`n    End If")
    $Form.DeleteLines(1,$Form.CountOfLines);$Form.AddFromString($source)
    $Form.AddFromString(@'
Public Sub ActivityShippingArmPermissionLoss()
    ActivityPendingPermissionLoss = True
    ActivityPendingInterruptions = 0
End Sub
'@)
    $Module.AddFromString(@'
Private ActivityFixtureAuthPath As String
Public Sub ActivityShippingArmPermissionLoss(ByVal authPath As String)
    ActivityFixtureAuthPath = authPath
    mShipmentsLauncherForm.ActivityShippingArmPermissionLoss
End Sub
Public Sub ActivityShippingRevokeAtPending()
    Dim wb As Workbook, ws As Worksheet, caps As ListObject, candidate As ListObject
    Dim row As ListRow, revoked As Long
    On Error GoTo Failed
    Set wb = Application.Workbooks.Open(ActivityFixtureAuthPath, 0, False)
    For Each ws In wb.Worksheets
        For Each candidate In ws.ListObjects
            If candidate.Name = "tblCapabilities" Then Set caps = candidate
        Next candidate
    Next ws
    If caps Is Nothing Then GoTo Failed
    For Each row In caps.ListRows
        If CStr(row.Range.Cells(1, caps.ListColumns("UserId").Index).Value2) = "config-reader" And _
           CStr(row.Range.Cells(1, caps.ListColumns("Capability").Index).Value2) = "SHIP_POST" Then
            row.Range.Cells(1, caps.ListColumns("Status").Index).Value2 = "Inactive"
            revoked = revoked + 1
        End If
    Next row
    If revoked <> 1 Then GoTo Failed
    wb.Save
    wb.Close False
    Exit Sub
Failed:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0
    Err.Raise vbObjectError + 261, , "Shipping permission interruption fixture unavailable."
End Sub
'@)
}

function Test-Slice4beShippingAccessInterruptions($Fixture,$Operator,$Other,$Ship,$Hold) {
    $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $authBytes=[IO.File]::ReadAllBytes($authPath)
    $authHash=Get-ShippingActivityHash $authPath
    $configHash=Get-ShippingActivityHash $Fixture.Config
    $inventory=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $authorityHash=Get-ShippingActivityHash $inventory
    $otherHash=Get-ShippingActivityHash $Other.FullName
    $root=[IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\'
    $notice='Shipping permission could not be verified. Review Shipping access before continuing.'
    foreach($path in @($authPath,$Fixture.Config)) {
        if(-not [IO.Path]::GetFullPath($path).StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw 'Access fixture path escaped its generated root.'}
    }
    [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingSetProbeMode' @($true))
    $completed=$false
    try {
        foreach($case in @('DuringYield','AuthUnavailable','ConfigUnavailable')) {
            $actions=if($case -eq 'DuringYield'){@('Add','Stage','Send')}else{@('Add','Update','Remove','Hold','Return','Stage','Send')}
            foreach($action in $actions) {
                SelectTarget $Fixture 'config-reader'
                [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingRelaunchReuses')
                $key=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingPrepare' @($action))
                if(-not $key){throw 'Shipping access-interruption fixture lacks selected staging.'}
                $session=[long](Run 'invSys.Core.xlam' 'TestShippingSession.SessionVersion')
                $allowed=[bool](Run 'invSys.Core.xlam' 'TestShippingSession.CanShip')
                if(-not $allowed){throw 'Shipping access fixture lacks initial permission.'}
                $rows=@(Get-ShippingActivityRows $Ship)|ConvertTo-Json -Depth 5 -Compress
                $held=@(Get-ShippingActivityRows $Hold)|ConvertTo-Json -Depth 5 -Compress
                $before=[long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingProbeCount')
                $path='';$saved='';$moved=$false
                try {
                    if($case -eq 'DuringYield') {
                        [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingArmPermissionLoss' @($authPath))
                    } else {
                        $path=if($case -eq 'AuthUnavailable'){$authPath}else{$Fixture.Config}
                        $saved=$path+'.fixture-held'
                        if(-not [IO.Path]::GetFullPath($saved).StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $saved)){throw 'Unavailable-access fixture destination is invalid.'}
                        Move-Item -LiteralPath $path -Destination $saved
                        $moved=$true
                    }
                    $Other.Activate()
                    [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingClick' @($action))
                    $prefix='Shipping.Access.'+$case+'.'+$action
                    if($case -eq 'DuringYield') {
                        Check ($prefix+'.RealYieldRevokedOnce') ([long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingInterruptionCount') -eq 1)
                    } else {
                        Check ($prefix+'.MissingFileNotRecreated') (-not (Test-Path -LiteralPath $path))
                    }
                    Check ($prefix+'.CorePermissionUnavailable') (-not [bool](Run 'invSys.Core.xlam' 'TestShippingSession.CanShip'))
                    Check ($prefix+'.SameSignedInSession') ([bool](Run 'invSys.Core.xlam' 'modAuth.IsSignedIn') -and $session -eq [long](Run 'invSys.Core.xlam' 'TestShippingSession.SessionVersion'))
                    Check ($prefix+'.StopsBeforeMutationOwner') ($before -eq [long](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingProbeCount'))
                    Check ($prefix+'.FixedAccessNotice') ([string](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingStatus') -ceq $notice)
                    Check ($prefix+'.StagingAndUnknownValuesPreserved') ($rows -ceq (@(Get-ShippingActivityRows $Ship)|ConvertTo-Json -Depth 5 -Compress) -and $held -ceq (@(Get-ShippingActivityRows $Hold)|ConvertTo-Json -Depth 5 -Compress))
                    Check ($prefix+'.CapturedWorkbookRetained') ([bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingBound' @($Operator.Name)))
                } finally {
                    if($moved) {
                        if(Test-Path -LiteralPath $path){
                            $unexpected=$path+'.unexpected-'+$case+'-'+$action
                            if(-not [IO.Path]::GetFullPath($unexpected).StartsWith($root,[StringComparison]::OrdinalIgnoreCase) -or (Test-Path -LiteralPath $unexpected)){throw 'Unexpected authority fixture destination is invalid.'}
                            Move-Item -LiteralPath $path -Destination $unexpected
                        }
                        Move-Item -LiteralPath $saved -Destination $path
                    }
                    if($case -eq 'DuringYield'){[IO.File]::WriteAllBytes($authPath,$authBytes)}
                }
            }
        }
        Check 'Shipping.Access.AuthBytesRestored' ($authHash -ceq (Get-ShippingActivityHash $authPath))
        Check 'Shipping.Access.ConfigBytesRestored' ($configHash -ceq (Get-ShippingActivityHash $Fixture.Config))
        Check 'Shipping.Access.AuthorityBytesPreserved' ($authorityHash -ceq (Get-ShippingActivityHash $inventory))
        Check 'Shipping.Access.UnrelatedWorkbookPreserved' ($otherHash -ceq (Get-ShippingActivityHash $Other.FullName) -and $Other.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'shipping unrelated sentinel')
        $completed=$true
    } finally {
        [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.ActivityShippingSetProbeMode' @($false))
        if($completed){SelectTarget $Fixture 'config-reader'}
    }
}
