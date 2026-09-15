# Diagnostic instrumentation of disposable copies only; actual Viewer entry is retained.
function Install-Slice4beViewerStartupProbe {
    $manager=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modInventoryViewer').CodeModule
    $form=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmInventoryViewer').CodeModule
    $manager.InsertLines($manager.CountOfDeclarationLines+1,"Public ViewerStartupTraceForTest As String`r`nPrivate ViewerStartupHistoryForTest As String")
    function TraceBefore($Module,[string]$Procedure,[string]$Statement,[string]$Stage) {
        $first=$Module.ProcStartLine($Procedure,0)
        $last=$first+$Module.ProcCountLines($Procedure,0)-1
        $traceLines=@(for($index=$first;$index -le $last;$index++){
            if(([string]$Module.Lines($index,1)).Trim() -ieq $Statement){$index}
        })
        if($traceLines.Count -ne 1){throw ('Viewer startup trace boundary is not unique: '+$Stage+'; count='+$traceLines.Count)}
        $Module.InsertLines($traceLines[0],('    modInventoryViewer.RecordViewerStartupStageForTest "'+$Stage+'"'))
    }
    TraceBefore $manager 'OpenInventoryViewer' 'Set mInventoryViewer = New frmInventoryViewer' 'Launcher.Create'
    TraceBefore $manager 'OpenInventoryViewer' 'mInventoryViewer.SetGeneration mViewerGeneration' 'Launcher.SetGeneration'
    TraceBefore $manager 'OpenInventoryViewer' 'mInventoryViewer.SetWarehouse warehouseId' 'Launcher.SetWarehouse'
    TraceBefore $manager 'OpenInventoryViewer' 'mInventoryViewer.RefreshInventory' 'Launcher.RefreshInventory'
    TraceBefore $manager 'OpenInventoryViewer' 'If Not mInventoryViewer.Visible Then mInventoryViewer.Show vbModeless' 'Launcher.Show'
    TraceBefore $form 'SetWarehouse' 'If mSettingsContext <> modActivity.CaptureContext() Then ClosePathLibrary' 'Warehouse.ClosePaths'
    TraceBefore $form 'SetWarehouse' 'If mSettingsContext <> modActivity.CaptureContext() Then ClearViewerContent' 'Warehouse.ClearContent'
    TraceBefore $form 'SetWarehouse' 'mWarehouseId = Trim$(warehouseId)' 'Warehouse.AssignIdentity'
    TraceBefore $form 'SetWarehouse' 'mSettingsContext = modActivity.CaptureContext()' 'Warehouse.CaptureContext'
    TraceBefore $form 'SetWarehouse' 'If Not mRecording Is Nothing Then mRecording.BindContext mSettingsContext' 'Warehouse.BindRecording'
    TraceBefore $form 'SetWarehouse' 'Me.Caption = "Viewer - " & mWarehouseId' 'Warehouse.Caption'
    TraceBefore $form 'UserForm_Initialize' 'BuildLayout' 'Initialize.BuildLayout'
    TraceBefore $form 'UserForm_Activate' 'If Not mRecording Is Nothing Then mRecording.Render' 'Activate.Recording'
    TraceBefore $form 'UserForm_Activate' 'modUserFormResizeWin.EnableResizableUserForm Me, True, True' 'Activate.Resize'
    TraceBefore $form 'UserForm_Activate' 'If Not mLayout Is Nothing Then mLayout.ApplyAnchoredLayout' 'Activate.Anchors'
    TraceBefore $form 'UserForm_Activate' 'ConfigureViewerHeaderGeometry' 'Activate.Headers'
    TraceBefore $form 'UserForm_Layout' 'If Not mLayout Is Nothing Then mLayout.ApplyAnchoredLayout' 'Layout.Anchors'
    TraceBefore $form 'UserForm_Layout' 'ConfigureViewerHeaderGeometry' 'Layout.Headers'
    foreach($procedure in @('UserForm_Initialize','UserForm_Activate','UserForm_Layout','SetWarehouse')){
        TraceBefore $form $procedure 'End Sub' (($procedure -replace '_','')+'.Return')
    }
    $manager.AddFromString(@'
Public Sub RecordViewerStartupStageForTest(ByVal stage As String)
    ViewerStartupTraceForTest = stage
    ViewerStartupHistoryForTest = ViewerStartupHistoryForTest & stage & "|"
    If Len(ViewerStartupHistoryForTest) > 8192 Then ViewerStartupHistoryForTest = Mid$(ViewerStartupHistoryForTest, InStr(4096, ViewerStartupHistoryForTest, "|") + 1)
End Sub
Public Function ReadViewerStartupHistoryForTest() As String
    ReadViewerStartupHistoryForTest = ViewerStartupHistoryForTest
End Function
Public Function ViewerStartupDiagnosticForTest() As String
    On Error GoTo Failed
    ViewerStartupHistoryForTest = ""
    RecordViewerStartupStageForTest "Launcher.Enter"
    OpenInventoryViewer
    If mInventoryViewer Is Nothing Then Err.Raise 5
    If Not mInventoryViewer.Visible Then Err.Raise 5
    ViewerStartupDiagnosticForTest = "OK|" & CStr(mViewerGeneration)
    Exit Function
Failed:
    ViewerStartupDiagnosticForTest = "ERROR|" & CStr(Err.Number) & "|" & ViewerStartupTraceForTest
End Function
Public Function CloseViewerStartupDiagnosticForTest() As String
    On Error GoTo Failed
    CloseInventoryViewerForTest
    CloseViewerStartupDiagnosticForTest = "OK"
    Exit Function
Failed:
    CloseViewerStartupDiagnosticForTest = "ERROR|" & CStr(Err.Number)
End Function
'@)
}

function Initialize-Slice4beViewerStartupWorkbook {
    $file=Join-Path $runRoot 'viewer-startup-operator.xlsm'
    if(Test-Path -LiteralPath $file){throw 'Preserve an existing startup workbook fixture.'}
    $book=$excel.Workbooks.Add()
    try {$book.SaveAs($file,52)}
    finally {
        $book.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    }
    $script:ViewerStartupWorkbookPath=$file
    $script:ViewerStartupWorkbookHash=(Get-FileHash -LiteralPath $file).Hash
    $script:ViewerStartupWorkbook=$excel.Workbooks.Open($file,0,$false)
    if($null -eq $script:ViewerStartupWorkbook -or $script:ViewerStartupWorkbook.Saved -isnot [bool] -or -not $script:ViewerStartupWorkbook.Saved){throw 'Saved/reopened startup workbook is not verified.'}
}

function Close-Slice4beViewerStartupWorkbook {
    if($null -eq $script:ViewerStartupWorkbook -or $script:ViewerStartupWorkbook.FullName -cne $script:ViewerStartupWorkbookPath){throw 'Captured startup workbook identity changed.'}
    $script:ViewerStartupWorkbook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:ViewerStartupWorkbook)
    $script:ViewerStartupWorkbook=$null
    Check 'ViewerStartup.SavedWorkbookBytesPreservedAfterClose' ((Get-FileHash -LiteralPath $script:ViewerStartupWorkbookPath).Hash -ceq $script:ViewerStartupWorkbookHash)
}

function Test-Slice4beViewerStartup($Fixture,[hashtable]$Pins,[long]$AuthorityBefore,[long]$PublicationBefore) {
    $context=Run 'invSys.Core.xlam' 'modActivity.CaptureContext'
    if($context -isnot [string] -or $context -ceq ''){throw 'Viewer startup fixture has no captured reader context.'}
    $visible=$excel.Visible
    if($visible -isnot [bool]){throw 'Viewer entry visibility is not a Boolean observation.'}
    $bookCount=$excel.Workbooks.Count
    if($bookCount -isnot [int]){throw 'Viewer entry workbook count is not an integer observation.'}
    $ordinary=0
    foreach($book in $excel.Workbooks){
        $isAddin=$book.IsAddin
        if($isAddin -isnot [bool]){throw 'Workbook classification is unavailable.'}
        if(-not $isAddin){$ordinary++}
    }
    [pscustomobject]@{EarlyVisibleRequested=[bool]$GuideCaptureVisibleExcelForTest;SavedWorkbookRequested=[bool]$ViewerStartupSavedWorkbookForTest;ExcelVisible=$visible;VisibilityType=$visible.GetType().FullName;WorkbookCount=$bookCount;OrdinaryWorkbookCount=$ordinary;ActiveWorkbookPresent=($null -ne $excel.ActiveWorkbook)}|
        ConvertTo-Json|Set-Content -LiteralPath (Join-Path $reportRoot 'viewer-startup-window-state.json')
    Check 'ViewerStartup.EntryVisibilityObserved' $true
    if($GuideCaptureVisibleExcelForTest -and -not $visible){throw 'The requested early-visible fixture no longer has visible Excel.'}
    $report=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.ViewerStartupDiagnosticForTest')
    if($report -notmatch '^(OK\|[1-9][0-9]*|ERROR\|-?[0-9]+\|[A-Za-z.]+)$'){throw 'Viewer startup diagnostic envelope is invalid.'}
    # Fixed stage identifiers and numeric error/generation only; no raw error text.
    $report|Set-Content -LiteralPath (Join-Path $reportRoot 'viewer-startup-observation.txt')
    $history=Run 'invSys.Operations.xlam' 'modInventoryViewer.ReadViewerStartupHistoryForTest'
    if($history -isnot [string] -or $history.Length -gt 8192 -or $history -cnotmatch '^(?:[A-Za-z.]+\|)+$'){throw 'Startup event trace is not a bounded fixed-stage envelope.'}
    $history|Set-Content -LiteralPath (Join-Path $reportRoot 'viewer-startup-event-trace.txt')
    Write-Output ('Viewer startup observation: '+$report)
    $opened=$report.StartsWith('OK|',[StringComparison]::Ordinal)
    Check 'ViewerStartup.ActualPublicCallbackOpens' $opened
    if($opened){
        $again=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.ViewerStartupDiagnosticForTest')
        if($again -notmatch '^(OK\|[1-9][0-9]*|ERROR\|-?[0-9]+\|[A-Za-z.]+)$'){throw 'Repeated entry diagnostic envelope is invalid.'}
        Check 'ViewerStartup.RepeatedEntryReusesGeneration' ($again -ceq $report)
    } else {Check 'ViewerStartup.RepeatedEntryReusesGeneration' $false}
    Check 'ViewerStartup.CapturedContextRetained' ((Run 'invSys.Core.xlam' 'modActivity.CaptureContext') -ceq $context)
    $unchanged=$true
    foreach($path in $Pins.Keys){if((Get-FileHash -LiteralPath $path).Hash -cne $Pins[$path]){$unchanged=$false}}
    $unchanged=$unchanged -and @(Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File).Count -eq $Pins.Count
    Check 'ViewerStartup.WarehouseBytesPreserved' $unchanged
    $authority=Run 'invSys.Operations.xlam' 'modTS_Shipments.PublishedReadAuthorityCallsForTest'
    $publication=Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest'
    if($null -eq $authority -or $null -eq $publication){throw 'Startup read/publication counters are unavailable.'}
    Check 'ViewerStartup.NoShippingAuthorityOrPublication' ($authority -eq $AuthorityBefore -and $publication -eq $PublicationBefore)
    if($ViewerStartupSavedWorkbookForTest){
        $active=$excel.ActiveWorkbook
        $saved=$script:ViewerStartupWorkbook.Saved
        Check 'ViewerStartup.SavedWorkbookActiveAndSaved' ($null -ne $active -and $active.FullName -ceq $script:ViewerStartupWorkbook.FullName -and
            $saved -is [bool] -and $saved)
    }
}
