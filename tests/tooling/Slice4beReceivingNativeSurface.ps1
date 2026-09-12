# D18 discovery only. These observation seams are never saved into the XLAM.
. (Join-Path $PSScriptRoot 'Slice4beWorksheetInput.ps1')
function Install-ReceivingNativeSurfaceSeam {
    . (Join-Path $repo 'tools/plan022-dialog-observer.ps1')
    if(-not ('Plan022NativeDialogs' -as [type])) {
        Invoke-Plan022NativeDialogObservation -ProcessId ([WorksheetInput]::Owner([IntPtr]$excel.Hwnd)) -TimeoutSeconds 0
    }
    $module=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modTS_Received').CodeModule
    $module.AddFromString(@'
Public SurfaceEntries As Long
Public SurfaceNativeCaller As Boolean
Public SurfaceProbeEntries As Long
Public Sub ActivityTestSurfaceInput()
    SurfaceProbeEntries = SurfaceProbeEntries + 1
End Sub
Public Function ActivityTestSurfaceInputCount() As Long
    ActivityTestSurfaceInputCount = SurfaceProbeEntries
End Function
Public Function ActivityTestSurfaceEntries() As Long
    ActivityTestSurfaceEntries = SurfaceEntries
End Function
Public Function ActivityTestSurfaceNativeCaller() As Boolean
    ActivityTestSurfaceNativeCaller = SurfaceNativeCaller
End Function
'@)
    $start=$module.ProcStartLine('ConfirmWrites',0)
    $count=$module.ProcCountLines('ConfirmWrites',0)
    $source=$module.Lines($start,$count)
    $pattern='(?im)^    mLastConfirmSucceeded = False\r?$'
    if([regex]::Matches($source,$pattern).Count -ne 1){throw 'Unique worksheet owner-entry observation boundary unavailable.'}
    $instrumentation=@'
    SurfaceEntries = SurfaceEntries + 1
    SurfaceNativeCaller = False
    If VarType(Application.Caller) = vbString Then SurfaceNativeCaller = (Application.Caller = "btnConfirmWrites")
'@
    $changed=[regex]::Replace($source,$pattern,[Text.RegularExpressions.MatchEvaluator]{param($match) $instrumentation+"`r`n"+$match.Value})
    $module.DeleteLines($start,$count);$module.InsertLines($start,$changed)
}

function Invoke-ReceivingNativeSurface($Operator,$Sheet,$Button,[string]$Mode) {
    $excel.Visible=$true
    Write-Output ('Native fixture window visible before preparation='+[bool]$Operator.Windows.Item(1).Visible)
    $Operator.Windows.Item(1).Visible=$true
    $Operator.Windows.Item(1).Activate();$Operator.Activate();$Sheet.Activate()
    $window=[IntPtr]$Operator.Windows.Item(1).Hwnd
    $excel.ActiveWindow.ScrollRow=1;$excel.ActiveWindow.ScrollColumn=1
    if(-not [WorksheetInput]::Activate($window)){throw 'Owned Receiving workbook foreground unavailable.'}
    Start-Sleep -Milliseconds 350
    # Calibrate in this workbook, using the same package, before the real control.
    $probe=$Sheet.Shapes.AddFormControl(0,$Button.Left,$Button.Top+28,$Button.Width,$Button.Height)
    try {
        $probe.OnAction="'"+$packages['invSys.Operations.xlam'].FullName.Replace("'","''")+"'!modTS_Received.ActivityTestSurfaceInput"
        $before=[int](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestSurfaceInputCount')
        $point=Get-WorksheetButtonPoint $excel $probe $Operator $Sheet
        Save-WorksheetInputCapture $excel $point (Join-Path $reportRoot ('probe-'+$Mode+'.png'))
        [WorksheetInput]::Click($window,$point.X,$point.Y)
        Start-Sleep -Milliseconds 500
        if([int](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestSurfaceInputCount') -ne $before+1){throw 'Packaged worksheet input calibration failed.'}
        Check ('Surface.'+$Mode+'.NativeInputCalibrated') $true
    } finally {$probe.Delete()}
    $before=[int](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestSurfaceEntries')
    $point=Get-WorksheetButtonPoint $excel $Button $Operator $Sheet
    Save-WorksheetInputCapture $excel $point (Join-Path $reportRoot ('confirm-'+$Mode+'.png'))
    [WorksheetInput]::Click($window,$point.X,$point.Y)
    $macroBlocked=$false
    for($poll=0;$poll -lt 5;$poll++) {
        Start-Sleep -Milliseconds 250
        # Arbitrary dialog text stays only in memory; dismiss informational OK only.
        $notices=@([Plan022NativeDialogs]::Poll([WorksheetInput]::Owner([IntPtr]$excel.Hwnd)))
        if(@($notices | Where-Object {$_ -match 'Cannot run the macro|macro may not be available'}).Count -gt 0){$macroBlocked=$true}
    }
    $entered=[int](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestSurfaceEntries') -eq $before+1
    $shapeCaller=$entered -and [bool](Run 'invSys.Operations.xlam' 'modTS_Received.ActivityTestSurfaceNativeCaller')
    Check ('Surface.'+$Mode+'.NativeOwnerEntry') $entered
    Check ('Surface.'+$Mode+'.NativeShapeCaller') $shapeCaller
    Check ('Surface.'+$Mode+'.NoMacroUnavailableNotice') (-not $macroBlocked)
    return [pscustomobject]@{Entered=$entered;ShapeCaller=$shapeCaller}
}
