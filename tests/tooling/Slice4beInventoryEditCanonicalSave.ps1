# Actual Upload, owned generated warehouses, and failures at the real catalog
# save boundary. Persisted values stay in memory/ignored owned copies; publish
# fixed assertion labels and booleans only. No product source or package is saved.
function Install-InventoryEditCanonicalSaveProbe {
    $domainOwner=Test-Path (Join-Path $repo 'src/InventoryDomain/Modules/modInventoryEditReceipts.bas')
    $script:InventoryEditSaveProbePackage=if($domainOwner){'invSys.Inventory.Domain.xlam'}else{'invSys.Admin.xlam'}
    $project=$packages[$script:InventoryEditSaveProbePackage].VBProject
    $probe=$project.VBComponents.Add(1);$probe.Name='TestInventoryEditSave'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mMode As String, mSource As String, mHost As String, mCopy As String
Private mReached As Boolean, mCopied As Boolean, mSaved As Boolean, mSaves As Long
Private mSourceHash As String, mHostHash As String
Public Sub Arm(ByVal mode As String, ByVal source As String, ByVal host As String, ByVal copyPrefix As String)
    mMode = mode: mSource = source: mHost = host: mCopy = copyPrefix
    mReached = False: mCopied = False: mSaved = False: mSaves = 0
    mSourceHash = "": mHostHash = ""
End Sub
Private Sub Observe()
    Dim fso As Object
    mReached = True
    mSourceHash = modTrainingWire.FileSha256(mSource)
    mHostHash = modTrainingWire.FileSha256(mHost)
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile mHost, mCopy & "-save.xlsx", False
    mCopied = (mSourceHash <> "" And mHostHash <> "" And _
        mSourceHash = modTrainingWire.FileSha256(mSource) And _
        mHostHash = modTrainingWire.FileSha256(mHost))
End Sub
Public Sub BeforeSave()
    If mMode = "BeforeCatalogSaveAuthorityLost" Then
        Observe
        modAuth.SignOut
    End If
End Sub
Public Sub SaveReturn(ByVal book As Workbook)
    If mMode = "" Then book.Save: Exit Sub
    If StrComp(book.FullName, mSource, vbTextCompare) <> 0 Then _
        Err.Raise vbObjectError + 2211, , "Owned catalog save source changed."
    mSaves = mSaves + 1
    If mMode = "BeforeCatalogSaveThrows" Then
        Observe
        Err.Raise vbObjectError + 2212, , "Owned catalog save failure."
    End If
    book.Save
    mSaved = book.Saved
    If Not mReached Then Observe
    If mMode = "AfterCatalogSaveThrows" Then _
        Err.Raise vbObjectError + 2213, , "Owned save-return failure."
    If mMode = "AfterCatalogSaveAuthorityLost" Then modAuth.SignOut
End Sub
Public Function Facts() As String
    Dim sourceSame As Boolean, hostSame As Boolean
    If mSourceHash <> "" Then sourceSame = (mSourceHash = modTrainingWire.FileSha256(mSource))
    If mHostHash <> "" Then hostSame = (mHostHash = modTrainingWire.FileSha256(mHost))
    Facts = CStr(mReached) & "|" & CStr(mCopied) & "|" & CStr(mSaved) & "|" & _
        CStr(mSaves) & "|" & CStr(sourceSame) & "|" & CStr(hostSame)
End Function
'@)
    $packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('TestInventoryCommands').CodeModule.AddFromString(@'
Public Function CanonicalSaveUnconfirmedForTest() As Boolean
    CanonicalSaveUnconfirmedForTest = (InStr(1, mLastNotice, "partial", vbTextCompare) > 0 Or _
        InStr(1, mLastNotice, "unconfirmed", vbTextCompare) > 0)
End Function
'@)
    $ownerName=if($domainOwner){'modInventoryEditReceipts'}else{'modAdmin'}
    $procedure=if($domainOwner){'SaveCatalog'}else{'UpdateInventoryItemCatalogForWarehouse'}
    $owner=$project.VBComponents.Item($ownerName).CodeModule
    $start=$owner.ProcStartLine($procedure,0)
    $count=$owner.ProcCountLines($procedure,0)
    $body=$owner.Lines($start,$count)
    $boundary=if($domainOwner){'(?im)^    If Not modInventoryCommands\.BeginCatalogSave\(context, report\) Then GoTo Failed\r?$'}else{'(?im)^    WriteInventoryCatalogFieldsAdmin lo, rowIndex, fields\r?$'}
    $save=if($domainOwner){'(?im)^    book\.Save\r?$'}else{'(?im)^    wb\.Save\r?$'}
    if([regex]::Matches($body,$boundary).Count -ne 1 -or [regex]::Matches($body,$save).Count -ne 1){throw 'Canonical catalog save seam unavailable; not RED.'}
    $body=[regex]::Replace($body,$boundary,[Text.RegularExpressions.MatchEvaluator]{
        param($m)
        if($domainOwner){'    TestInventoryEditSave.BeforeSave'+"`r`n"+$m.Value.TrimEnd("`r")}
        else{$m.Value.TrimEnd("`r")+"`r`n"+'    TestInventoryEditSave.BeforeSave'}
    })
    $replacement=if($domainOwner){'    TestInventoryEditSave.SaveReturn book'}else{'    TestInventoryEditSave.SaveReturn wb'}
    $body=[regex]::Replace($body,$save,$replacement)
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$body)
}
function Test-InventoryEditCanonicalSaveOutcome($Fixture,[ref]$HostRef,[string]$HostPath,$Table,$Created,$After,$Snapshot,$Saved,[string]$Sku,[string]$Key,[string]$Item,[string]$Reason,[string]$Case,[string]$Prefix) {
    $id='InventoryEditReceipt.'+$Case
    $sourcePath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $boundary=([string](Run $script:InventoryEditSaveProbePackage 'TestInventoryEditSave.Facts')).Split('|')
    $observed=$boundary.Count -eq 6 -and $boundary[0] -ceq 'True' -and $boundary[1] -ceq 'True'
    Check ($id+'.ActualCanonicalSaveBoundaryObserved') $observed
    if(-not $observed){throw 'Canonical save observer unavailable; not RED.'}
    $wasSaved=$Case -like 'After*'
    if($wasSaved -and $boundary[2] -cne 'True'){throw 'Owned save did not return before after-save injection; not RED.'}
    $expectedSaves=1;if($Case -eq 'BeforeCatalogSaveAuthorityLost'){$expectedSaves=0}
    Check ($id+'.OnlyAllowedCatalogSaveAttempted') ([int]$boundary[3] -eq $expectedSaves)
    Check ($id+'.SourceBytesPreservedAfterBoundary') ($boundary[4] -ceq 'True')
    if($Case -like '*AuthorityLost'){
        Check ($id+'.OperatorBytesPreservedAfterAuthorityLoss') ($boundary[5] -ceq 'True')
        Check ($id+'.AuthorityLossObserved') (-not [bool](Run 'invSys.Core.xlam' 'modAuth.IsSignedIn'))
    }
    $prepared=Get-InventoryEditSavedResult ($Prefix+'-save.xlsx')
    Check ($id+'.CommandReferenceSavedBeforeCanonicalSave') ($prepared.Command -ne '' -and $prepared.Preserved)
    Check ($id+'.FailureRetainsOriginalCommandReference') ($prepared.Command -ne '' -and $Saved.Command -ceq $prepared.Command)
    Check ($id+'.UnconfirmedResultNotImported') ($Saved.Status -cne 'IMPORTED')
    Check ($id+'.ReadOnlyInspectionPreservedBytes') ($Snapshot.Preserved -and $Saved.Preserved)
    $queue=([string](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.Facts')).Split('|')
    Check ($id+'.NoQuantityQueuedAfterSaveFailureOrAuthorityLoss') ($queue[0] -ceq 'False')
    Check ($id+'.NoQuantityAppliedAfterBoundary') ($After.LogRows -eq $Created.LogRows -and $After.AppliedRows -eq $Created.AppliedRows -and $After.EntityQuantities.Count -eq $Created.EntityQuantities.Count -and $After.EntityQuantities[$Key] -eq 0)
    $planned=''
    if($wasSaved){
        Check ($id+'.SavedCatalogAndReasonIndependentlyConfirmed') ($Snapshot.CatalogMatches -eq 1 -and $Snapshot.Item -ceq $Item -and $Snapshot.Reason -ceq $Reason -and [regex]::Matches($Snapshot.History,[regex]::Escape($Reason)).Count -eq 1)
        $bound=Test-InventoryEditBinding $Snapshot $prepared $Fixture $Sku
        if($bound){$planned=$Snapshot.Receipts[0].Event}
        Check ($id+'.CatalogReceiptAndEventPlanShareConfirmedSave') ($bound -and $planned -ne '')
    }else{
        Check ($id+'.UnsavedCanonicalFilePreserved') ($After.FileHash -ceq $Created.FileHash)
        Check ($id+'.NoFalseCatalogReceiptPersisted') ($Snapshot.Receipts.Count -eq 0 -and [regex]::Matches($Snapshot.History,[regex]::Escape($Reason)).Count -eq 0)
    }
    if($Case -ne 'BeforeCatalogSaveAuthorityLost'){
        Check ($id+'.SaveUncertaintyDisclosed') ($Saved.Text -match '(?i)partial|unconfirmed' -or [bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.CanonicalSaveUnconfirmedForTest'))
    }
    Check ($id+'.UnknownOperatorColumnPreserved') ([string]$Table.ListColumns.Item('USER NOTE').DataBodyRange.Cells.Item(1,1).Value2 -ceq 'Retained receipt extension')
    # Reopen the same saved operator, under the original source/user's new session.
    # Replace the caller's owned reference so its finally closes the new instance.
    $HostRef.Value.Close($false)
    $retryBook=$excel.Workbooks.Open($HostPath,0,$false)
    $HostRef.Value=$retryBook
    try{
        SelectTarget $Fixture
        $retryTable=@($retryBook.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -like 'invSys_Inventory_*'})[0]
        $retryBook.Activate();$retryTable.DataBodyRange.Cells.Item(1,1).Select()
        [void](Run $script:InventoryEditSaveProbePackage 'TestInventoryEditSave.Arm' @('ObserveRetry',$sourcePath,$HostPath,($Prefix+'-retry')))
        [void](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.Arm' @('Retry',$sourcePath,$HostPath,($Prefix+'-retry')))
        [void](Run 'invSys.Admin.xlam' 'TestInventoryWorksheetResult.Arm' @('CanonicalRetry',$retryBook.Name))
        Invoke-InventoryEditObservedUpload $Fixture $retryBook ($id+'.Retry')
        $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
        $reached=$facts.Count -eq 5 -and $facts[0] -ceq 'True' -and $facts[2] -ceq '1' -and $facts[3] -ceq 'True'
        Check ($id+'.FreshRetryActualCapturedHandlerReached') $reached
        if(-not $reached){throw 'Fresh retry handler unavailable; not RED.'}
        $retry=Get-AdminInventoryEntryFacts $Fixture -Persisted
        $copy=$Prefix+'-result-retry.xlsx';Copy-Item -LiteralPath $HostPath -Destination $copy
        $retryResult=Get-InventoryEditSavedResult $copy
        $retrySnapshot=Read-InventoryEditReceiptCopy $sourcePath $Sku $retryResult.Command
        $retryBoundary=([string](Run $script:InventoryEditSaveProbePackage 'TestInventoryEditSave.Facts')).Split('|')
        $expectedRetrySaves=1;if($wasSaved){$expectedRetrySaves=0}
        Check ($id+'.FreshRetryOnlyCommitsUnsavedCatalog') ([int]$retryBoundary[3] -eq $expectedRetrySaves)
        Check ($id+'.FreshRetryCatalogAppliedExactlyOnce') ($retrySnapshot.Item -ceq $Item -and $retrySnapshot.Reason -ceq $Reason -and [regex]::Matches($retrySnapshot.History,[regex]::Escape($Reason)).Count -eq 1)
        Check ($id+'.FreshRetryKeepsOriginalCommandReference') ($prepared.Command -ne '' -and $retryResult.Command -ceq $prepared.Command)
        Check ($id+'.FreshRetryReciprocalSavedCommand') (Test-InventoryEditBinding $retrySnapshot $retryResult $Fixture $Sku)
        Check ($id+'.FreshRetryImportedResultSaved') ($retryResult.Status -ceq 'IMPORTED')
        $retryEvent=[string](Run 'invSys.Core.xlam' 'TestInventoryEditReceipt.PlannedEventForOwnedTest')
        if($wasSaved){Check ($id+'.FreshRetryKeepsSavedEventPlan') ($planned -ne '' -and $retryEvent -ceq $planned)}
        Check ($id+'.FreshRetryExactEntityQuantity') ($retry.EntityQuantities.Count -eq $Created.EntityQuantities.Count -and $retry.EntityQuantities[$Key] -eq 5)
        Check ($id+'.FreshRetryOneQuantityApplication') ($retry.LogRows -eq $Created.LogRows+1 -and $retry.AppliedRows -eq $Created.AppliedRows+1)
        $exact=@($retrySnapshot.Log|Where-Object {$_.Event -ceq $retryEvent -and $_.Key -ceq $Key -and [double]$_.Delta -eq 5})
        Check ($id+'.FreshRetryPlannedEventHasSavedProof') ($retryEvent -ne '' -and @($retrySnapshot.Applied|Where-Object {$_ -ceq $retryEvent}).Count -eq 1 -and $exact.Count -eq 1)
        Check ($id+'.RetryCustomAndInspectionPreserved') ($retrySnapshot.Custom -ceq 'Retained receipt extension' -and $retrySnapshot.Preserved -and $retryResult.Preserved)
    }finally{
        [void](Run $script:InventoryEditSaveProbePackage 'TestInventoryEditSave.Arm' @('','','',''))
    }
}
