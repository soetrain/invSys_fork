# Real Print handler entry and nested entry at the declared preview seam.
function Install-ProductionPrintEntryProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $form.InsertLines(1,'Private mPrintEntryRestored As Boolean')
    $form.AddFromString(@'
Public Function PrintEntryActForTest(ByVal mode As String) As Boolean
    Dim priorLoading As Boolean, priorBusy As Boolean, entryLoading As Boolean, entryBusy As Boolean
    On Error GoTo Failed
    priorLoading = mLoading: priorBusy = mDesignerActionInProgress
    mPrintEntryRestored = False
    If mode = "Loading" Then mLoading = True
    If mode = "Busy" Then mDesignerActionInProgress = True
    entryLoading = mLoading: entryBusy = mDesignerActionInProgress
    mTxtStatus.Text = "PRINT-ACTIVE"
    mBtnManagerPrint_Click
    mPrintEntryRestored = (mLoading = entryLoading And mDesignerActionInProgress = entryBusy)
    PrintEntryActForTest = True
Failed:
    mLoading = priorLoading: mDesignerActionInProgress = priorBusy
End Function
Public Function PrintEntryNestedForTest() As Boolean
    Dim before As String
    before = mTxtStatus.Text
    mBtnManagerPrint_Click
    PrintEntryNestedForTest = (mTxtStatus.Text = before)
End Function
Public Function PrintEntryRestoredForTest() As Boolean
    PrintEntryRestoredForTest = mPrintEntryRestored
End Function
'@)
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,'Private mPrintNestedArm As Boolean, mPrintNestedReached As Boolean, mPrintNestedStatus As Boolean')
    $line=$adapter.ProcBodyLine('PrintPreviewForTest',0)
    $adapter.InsertLines($line+1,'    PrintEntryPreviewBoundary')
    $adapter.AddFromString(@'
Public Sub PrintEntryPreviewBoundary()
    If Not mPrintNestedArm Then Exit Sub
    mPrintNestedArm = False: mPrintNestedReached = True
    mPrintNestedStatus = mForm.PrintEntryNestedForTest()
End Sub
Public Function PrintEntryAct(ByVal mode As String) As Boolean
    mPrintOwners = 0: mPrintReads = 0: Set mPrintBook = Nothing
    mPrintNestedArm = (mode = "Nested"): mPrintNestedReached = False: mPrintNestedStatus = False
    PrintEntryAct = mForm.PrintEntryActForTest(mode)
End Function
Public Function PrintEntryFact(ByVal fact As String) As Boolean
    Select Case fact
        Case "GuardsRestored": PrintEntryFact = mForm.PrintEntryRestoredForTest()
        Case "NestedReached": PrintEntryFact = mPrintNestedReached
        Case "NestedStatusPreserved": PrintEntryFact = mPrintNestedStatus
        Case "Permitted": PrintEntryFact = modRoleUiAccess.CanCurrentUserPerformCapability("PROD_POST") Or modRoleUiAccess.CanCurrentUserPerformCapability("ADMIN_MAINT")
    End Select
End Function
'@)
}

function Test-ProductionPrintEntry($Fixture,$Book,$Sheet,$Decoy){
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Fingerprint($Worksheet){ConvertTo-Json -Compress -Depth 10 -InputObject @($Worksheet.UsedRange.Formula)}
    function AuthorityPins {
        $pins=@{}
        foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File|Where-Object Name -NotLike '~$*'){
            $stream=[IO.File]::Open($file.FullName,'Open','Read','ReadWrite')
            try{$pins[$file.FullName]=(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}
        }
        return $pins
    }
    function AuthorityPreserved($Before){
        $after=AuthorityPins
        if($after.Count -ne $Before.Count){return $false}
        foreach($path in $Before.Keys){if(-not $after.ContainsKey($path) -or $after[$path] -cne $Before[$path]){return $false}}
        return $true
    }
    $report=$Book.Worksheets.Item('RecallCodesPrint')
    $source=Fingerprint $Sheet
    $decoyBefore=Fingerprint $Decoy.Worksheets.Item('Production')
    foreach($case in @('Loading','Busy','Nested','Reader','Admin','Producer')){
        $actor=switch($case){'Reader'{'config-reader'} 'Admin'{'config-admin'} default{'config-producer'}}
        SelectTarget $Fixture $actor
        [void](Probe 'OpenDesigner' @($Book.Name));[void](Probe 'RunLocalShowAndCapture' @($Book.Name,'PRINT'))
        $allowed=$case -cne 'Reader'
        if([bool](Probe 'PrintEntryFact' @('Permitted')) -ne $allowed){throw 'Print entry permission prerequisite mismatch; not product RED.'}
        $pins=AuthorityPins
        $reportBefore=Fingerprint $report
        $Decoy.Activate();[void](Probe 'ResetPrintPreviewForTest')
        $label='PrintEntry.'+$case
        Check ($label+'.ActualHandlerReturned') ([bool](Probe 'PrintEntryAct' @($case)))
        Check ($label+'.GuardsRestored') ([bool](Probe 'PrintEntryFact' @('GuardsRestored')))
        $suppressed=$case -in @('Loading','Busy','Reader')
        $count=if($suppressed){0}else{1}
        Check ($label+'.OwnerEntries') ([int](Probe 'PrintOwnerEntries') -eq $count)
        Check ($label+'.ReportReads') ([int](Probe 'PrintReportReads') -eq $count)
        Check ($label+'.PreviewEntries') ([int](Probe 'PrintPreviewCountForTest') -eq $count)
        $expected=if($case -in @('Loading','Busy')){'PRINT-ACTIVE'}elseif($case -ceq 'Reader'){'Production permission changed. Reopen Production before continuing.'}else{'Print preview closed.'}
        Check ($label+'.Status') ([string](Probe 'PrintStatus') -ceq $expected)
        if($suppressed){Check ($label+'.ReportPreserved') ((Fingerprint $report) -ceq $reportBefore)}
        if($case -ceq 'Nested'){
            foreach($fact in @('NestedReached','NestedStatusPreserved')){Check ($label+'.'+$fact) ([bool](Probe 'PrintEntryFact' @($fact)))}
        }
        Check ($label+'.SourceAndExactKeyPreserved') ((Fingerprint $Sheet) -ceq $source)
        Check ($label+'.DecoyPreserved') ((Fingerprint $Decoy.Worksheets.Item('Production')) -ceq $decoyBefore)
        Check ($label+'.AuthorityPreserved') (AuthorityPreserved $pins)
        CaptureOwnedFormByCaptionEvidence 'Production' ('print-entry-'+$case.ToLowerInvariant()+'.png')
        [void](Probe 'CloseDesigner')
    }
}
