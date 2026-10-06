# Entry RED uses a real authored guide and the same public form-action probe.
# This gate is intentionally separate from round-trip/hostile-file acceptance.
function Write-GuideTransferHostState([string]$Stage) {
    if(-not $GuideCaptureSavedWorkbookForTest){return}
    $book=$script:ViewerStartupWorkbook
    [pscustomobject]@{Stage=$Stage;UTC=[DateTimeOffset]::UtcNow.ToString('o');IdentityPreserved=($null -ne $book -and $book.FullName -ceq $script:ViewerStartupWorkbookPath);Saved=$book.Saved;Sheets=$book.Sheets.Count;Names=$book.Names.Count}|
        ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'transfer-host-state.jsonl')
}

function Install-GuideTransferDialogProbe {
    # The ordinary publisher appends Admin audit sheets. Supply its explicit
    # disposable Admin workbook instead of consuming the saved capture host.
    $admin=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('modAdminConsole').CodeModule
    $first=$admin.ProcStartLine('PublishReadFixtureForTest',0)
    $lines=$admin.ProcCountLines('PublishReadFixtureForTest',0)
    $admin.DeleteLines($first,$lines)
    $admin.InsertLines($first,@'
Public Function PublishReadFixtureForTest() As Boolean
    Dim report As String, auditBook As Workbook
    On Error GoTo Failed
    Set auditBook = Application.Workbooks.Add(xlWBATWorksheet)
    PublishReadFixtureForTest = GenerateInventorySnapshot("config-admin", "", Nothing, "", auditBook, report)
Cleanup:
    On Error GoTo 0
    If Not auditBook Is Nothing Then auditBook.Close SaveChanges:=False
    Exit Function
Failed:
    PublishReadFixtureForTest = False
    Resume Cleanup
End Function
'@)
    $script:guideTransferDialogProbeInstalled=$false
    $component=@($packages['invSys.Operations.xlam'].VBProject.VBComponents|Where-Object Name -CEQ 'modGuideTransferUi')
    if($component.Count -eq 0){return}
    if($component.Count -ne 1){throw 'Ambiguous transfer dialog owner; not product RED.'}
    $module=$component[0].CodeModule
    $start=$module.ProcStartLine('SelectFileForTransfer',0)
    $count=$module.ProcCountLines('SelectFileForTransfer',0)
    $body=[string]$module.Lines($start,$count)
    if($body -notmatch 'Private Function SelectFileForTransfer\(ByVal exporting As Boolean\) As String'){throw 'Transfer dialog seam changed; not product RED.'}
    $module.DeleteLines($start,$count)
    $module.InsertLines($start,@'
Private Function SelectFileForTransfer(ByVal exporting As Boolean) As String
    If TransferFileForTest = "TRANSFER_PICKER_FAULT" Then Err.Raise 5, , "Transfer fixture diagnostic content"
    If TransferSignOutForTest Then Application.Run "'invSys.Core.xlam'!modAuth.SignOut"
    SelectFileForTransfer = TransferFileForTest
End Function
'@)
    $module.InsertLines($module.CountOfDeclarationLines+1,"Private TransferFileForTest As String`r`nPrivate TransferSignOutForTest As Boolean")
    $module.AddFromString(@'
Public Sub SetTransferFileForTest(ByVal path As String, Optional ByVal signOut As Boolean = False)
    TransferFileForTest = path: TransferSignOutForTest = signOut
End Sub
'@)
    $script:guideTransferDialogProbeInstalled=$true
}

function Test-GuideTransferEntry($Fixture,$Other,$Guide) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    Write-GuideTransferHostState 'BeforeEntry'
    SelectTarget $Fixture 'config-admin';OpenRecordingViewer
    if((BoundLibrary 'Open') -cne 'DELIVERED'){throw 'Existing library unavailable; not transfer RED.'}
    BoundOpen;BoundSelect $Guide
    $training=BoundPins $journalRoot;$activity=ActivityPins
    $config=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $otherConfig=(Get-FileHash -LiteralPath $Other.Config).Hash
    $instructions=BoundControl 'txtPublishedInstructions' 'Text'
    $source=BoundControl 'lblPublishedGuideSource' 'Label'
    if(-not $instructions.Contains([string]$Guide.Name) -or -not $source.Contains([string]$Guide.ContentSha256)){throw 'Exact published-guide prerequisite unavailable.'}
    Check 'GuideTransfer.ExactPublishedVersionLoaded' $true
    try {
        foreach($control in @('btnExportGuide','btnImportGuide')){
            $state=BoundControl $control 'State'
            Check ('GuideTransfer.AdminEnabled.'+$control) ($state -ceq 'True|True')
            # A present implementation needs the compiled file-dialog cancellation
            # seam before this gate may invoke it; never open an unattended dialog.
            if($state -cne 'MISSING' -and -not $guideTransferDialogProbeInstalled){throw 'Transfer cancellation probe unavailable; not product RED.'}
            Check ('GuideTransfer.ActualCancelledHandler.'+$control) ((BoundControl $control 'Click') -ceq 'DELIVERED')
        }
        Check 'GuideTransfer.EntryAndCancellationPreserveEvidence' ((BoundSame $training (BoundPins $journalRoot)) -and (PinsRetained $activity) -and (Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $config)
        if($CaptureEvidence){CaptureOwnedFormByCaptionEvidence 'Published guides' 'guide-transfer-admin-entry.png'}
        Write-GuideTransferHostState 'AfterAdminEntry'
        CloseRecordingViewer;SelectTarget $Fixture 'config-reader';OpenRecordingViewer
        if((BoundLibrary 'Open') -cne 'DELIVERED'){throw 'Existing reader library unavailable.'}
        BoundOpen;BoundSelect $Guide
        foreach($control in @('btnExportGuide','btnImportGuide')){
            Check ('GuideTransfer.ReaderDenied.'+$control) ((BoundControl $control 'State') -ceq 'True|False' -and (BoundControl $control 'Click') -ceq 'DISABLED')
        }
        Check 'GuideTransfer.ReaderAndOtherWarehousePreserved' ((BoundSame $training (BoundPins $journalRoot)) -and (Get-FileHash -LiteralPath $Other.Config).Hash -ceq $otherConfig)
    } finally {CloseRecordingViewer;SelectTarget $Fixture 'config-admin';Write-GuideTransferHostState 'AfterReaderEntry'}
}
