# D18 reachability discovery through the actual packaged Receiving Ribbon route.
# All workbook preparation is confined to disposable Admin-generated fixtures.
function Test-ReceivingWorksheetRoute($Button) {
    $route=[string]$Button.OnAction
    $split=$route.LastIndexOf('!')
    # Excel may normalize a loaded add-in assignment to its unqualified macro.
    # This check proves the assigned handler only; native invocation is separate.
    if($route -ceq 'modTS_Received.ConfirmWrites'){return $true}
    if($split -lt 1){return $false}
    $package=$route.Substring(0,$split).Trim("'").Replace("''", "'")
    $macro=$route.Substring($split+1)
    $expected=$packages['invSys.Operations.xlam']
    $packageMatches=if([IO.Path]::IsPathRooted($package)){$package -ieq $expected.FullName}else{$package -ieq $expected.Name}
    return $packageMatches -and $macro -ceq 'modTS_Received.ConfirmWrites'
}

function Test-ReceivingSurfaceCoverage($Fixture) {
    if($CheckReceivingNativeSurface){Install-ReceivingNativeSurfaceSeam}
    SelectTarget $Fixture 'config-reader'
    $authority=Get-ReceivingAuthorityHashes $Fixture
    $other=$excel.Workbooks.Add()
    $other.Worksheets.Item(1).Cells.Item(1,1).Value2='surface unrelated sentinel'
    $other.SaveAs((Join-Path $runRoot 'surface-other.xlsm'),52)
    $otherHash=Get-ReceivingFixtureHash $other.FullName
    $operator=$null
    $observations=[Collections.Generic.List[object]]::new()
    try {
        [void](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestRibbonOpen')
        $state=[string](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherState')
        if(-not $state.EndsWith('|True')){throw 'Surface fixture Ribbon failed to establish a current visible form.'}
        $operator=$excel.Workbooks.Item($state.Split('|')[0])
        $sheet=$operator.Worksheets.Item('ReceivedTally')
        $button=$sheet.Shapes.Item('btnConfirmWrites')
        Check 'Surface.Provisioned.FormCaptured' ([bool](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherBoundTo' @($operator.Name)))
        Check 'Surface.Provisioned.ConfirmSupportSheetHidden' ($sheet.Visible -eq 2)
        Check 'Surface.Provisioned.ConfirmRouteRetained' (Test-ReceivingWorksheetRoute $button)
        $observations.Add([pscustomobject]@{Case='Provisioned';SheetVisibility=[int]$sheet.Visible;ButtonVisible=($button.Visible -ne 0);AssignedHandlerValid=(Test-ReceivingWorksheetRoute $button);NativeInvocationVerified=$false})
        [void](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherDismiss' @('Button'))
        $extra=(Table $operator 'ReceivedTally').ListColumns.Add(); $extra.Name='Surface Extra'
        $operator.Save()
        foreach($mode in @('Reused','OnlyStagingVisible','SavedReopen')) {
            if($mode -eq 'OnlyStagingVisible') {
                # A supported existing role workbook may have only this sheet visible.
                # Do not delete a sheet, change business data or create a new owner.
                $sheet.Visible=-1
                foreach($ws in $operator.Worksheets){if($ws.Name -ne 'ReceivedTally'){$ws.Visible=2}}
                $operator.Save()
            }
            if($mode -eq 'SavedReopen') {
                $operatorPath=$operator.FullName
                $operator.Close($false)
                $operator=$excel.Workbooks.Open($operatorPath)
                $sheet=$operator.Worksheets.Item('ReceivedTally')
            }
            $operatorHash=Get-ReceivingFixtureHash $operator.FullName
            $operator.Activate()
            [void](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestRibbonOpen')
            Check ('Surface.'+$mode+'.Captured') ([bool](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherBoundTo' @($operator.Name)))
            $button=$sheet.Shapes.Item('btnConfirmWrites')
            $visible=$sheet.Visible -eq -1 -and $button.Visible -ne 0
            Write-Output ('Surface observation '+$mode+': worksheet confirm visible='+$visible)
            Check ('Surface.'+$mode+'.ExpectedSupportVisibility') ($visible -eq ($mode -ne 'Reused'))
            $observation=[pscustomobject]@{Case=$mode;SheetVisibility=[int]$sheet.Visible;ButtonVisible=($button.Visible -ne 0);AssignedHandlerValid=(Test-ReceivingWorksheetRoute $button);NativeInvocationVerified=$false}
            $observations.Add($observation)
            Check ('Surface.'+$mode+'.AssignedPublicRoute') (Test-ReceivingWorksheetRoute $button)
            Check ('Surface.'+$mode+'.UnknownHeaderPreserved') ((Table $operator 'ReceivedTally').ListColumns.Item('Surface Extra').Name -ceq 'Surface Extra')
            [void](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherDismiss' @('Button'))
            Check ('Surface.'+$mode+'.SameVisibilityAfterFormClose') (($sheet.Visible -eq -1 -and $button.Visible -ne 0) -eq $visible)
            Check ('Surface.'+$mode+'.SavedBytesPreserved') ($operatorHash -ceq (Get-ReceivingFixtureHash $operator.FullName))
            if($visible -and $CheckReceivingNativeSurface){
                $native=@(Invoke-ReceivingNativeSurface $operator $sheet $button $mode)
                $native | Where-Object {$_ -is [string]} | Write-Output
                $proof=$native[-1]
                $observation.NativeInvocationVerified=$proof.Entered -and $proof.ShapeCaller
                Check ('Surface.'+$mode+'.SavedBytesPreservedAfterNativeInput') ($operatorHash -ceq (Get-ReceivingFixtureHash $operator.FullName))
            }
        }
        Check 'Surface.AuthorityBytesPreserved' (Test-ReceivingAuthorityHashes $Fixture $authority)
        Check 'Surface.UnrelatedWorkbookPreserved' ($otherHash -ceq (Get-ReceivingFixtureHash $other.FullName) -and $other.Worksheets.Item(1).ListObjects.Count -eq 0)
    } finally {
        $observations | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $reportRoot 'surface-reachability.json') -Encoding UTF8
        if([string](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherState') -ne ''){[void](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestLauncherDismiss' @('Internal'))}
        if($null -ne $operator){$operator.Close($false)}
        $other.Close($false)
    }
}
