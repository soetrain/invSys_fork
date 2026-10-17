# Disposable-fixture entry audit. Invoke the packaged Ribbon and modal OK handler.
# Adapters supply fixture fields, expose fixed facts and suppress raw owner notices.
function Install-AdminInventoryEntryProbe {
    $project=$packages['invSys.Admin.xlam'].VBProject
    $control=$project.VBComponents.Add(2);$control.Name='TestInventoryEntryRibbon'
    $control.CodeModule.AddFromString(@'
Option Explicit
Implements Office.IRibbonControl
Private Property Get IRibbonControl_Id() As String
    IRibbonControl_Id = "btnAdminAddInventoryItem"
End Property
Private Property Get IRibbonControl_Tag() As String
    IRibbonControl_Tag = ""
End Property
Private Property Get IRibbonControl_Context() As Object
    Set IRibbonControl_Context = Application.ActiveWindow
End Property
'@)
    $probe=$project.VBComponents.Add(1);$probe.Name='TestInventoryEntry'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mMode As String, mRoot As String, mSecret As String, mOtherBook As String
Private mHostName As String, mCapturedHost As Boolean
Private mSourceRoot As String, mWarehouse As String, mCapturedKind As String
Private mShows As Long, mAccepts As Long, mNotices As Long
Private mAccepted As Boolean, mYieldSucceeded As Boolean
Private mPayloadSeen As Boolean, mLegacyPayload As Boolean, mLegacyDisplay As Boolean
Public Sub Execute(ByVal hostName As String, ByVal mode As String, _
                   ByVal targetRoot As String, ByVal secret As String, ByVal otherBook As String, _
                   ByVal sourceRoot As String, ByVal warehouseId As String)
    Dim control As New TestInventoryEntryRibbon
    mMode = mode: mRoot = targetRoot: mSecret = secret: mOtherBook = otherBook
    mHostName = hostName: mCapturedHost = False
    mSourceRoot = sourceRoot: mWarehouse = warehouseId: mCapturedKind = "Unavailable"
    mShows = 0: mAccepts = 0: mNotices = 0: mAccepted = False: mYieldSucceeded = False
    mPayloadSeen = False: mLegacyPayload = False: mLegacyDisplay = False
    Application.Workbooks(hostName).Activate
    RibbonOnActionAdmin control
    mSecret = ""
End Sub
Public Sub Drive(ByVal form As frmAddInventoryItem)
    Dim selected As String, signed As String
    If TestInventoryCommands.IsDriving Then TestInventoryCommands.Drive form: Exit Sub
    mShows = mShows + 1
    If mShows > 1 Then form.InventoryEntryCancelForTest: Exit Sub
    mLegacyDisplay = form.InventoryEntryLegacyDisplayForTest()
    mCapturedHost = form.InventoryEntryOwnerBoundForTest(mHostName)
    mCapturedKind = form.InventoryEntryCapturedKindForTest(mHostName, mSourceRoot, mWarehouse)
    Select Case mMode
        Case "ActiveWorkbook"
            Application.Workbooks(mOtherBook).Activate
            mYieldSucceeded = True
        Case "SignedOut"
            modAuth.SignOut
            mYieldSucceeded = Not modAuth.IsSignedIn()
        Case "OtherWarehouseAdmin"
            modRuntimeWorkbooks.SetCoreDataRootOverride mRoot
            selected = modNasConnection.SelectWarehouseTargetForAutomation(mRoot, mRoot, "S1", False)
            modNasConnection.SetCurrentTargetPathsForTest "\\fixture-host\config-command", mRoot
            signed = modAuth.SignInCurrentTargetForAutomation("config-admin", mSecret, "")
            mYieldSucceeded = (Left$(selected, 3) = "OK|" And Left$(signed, 3) = "OK|")
        Case "ReplacementSession"
            modAuth.SignOut
            signed = modAuth.SignInCurrentTargetForAutomation("config-admin", mSecret, "")
            mYieldSucceeded = (Left$(signed, 3) = "OK|")
        Case "Baseline"
            mYieldSucceeded = True
    End Select
    If Not mYieldSucceeded Then form.InventoryEntryCancelForTest: Exit Sub
    mAccepts = mAccepts + 1
    mAccepted = form.InventoryEntryAcceptForTest()
    If Not mAccepted Then form.InventoryEntryCancelForTest
End Sub
Public Sub Notice()
    mNotices = mNotices + 1
End Sub
Public Sub RecordPayload(ByVal payload As String)
    mPayloadSeen = True
    mLegacyPayload = (InStr(1, payload, Chr$(34) & "ROW" & Chr$(34), vbTextCompare) > 0)
End Sub
Public Function Facts() As String
    Facts = CStr(mShows) & "|" & CStr(mAccepts) & "|" & CStr(mAccepted) & "|" & _
            CStr(mYieldSucceeded) & "|" & CStr(mNotices) & "|" & CStr(mPayloadSeen) & "|" & _
            CStr(mLegacyPayload) & "|" & CStr(mLegacyDisplay) & "|" & CStr(mCapturedHost) & "|" & mCapturedKind
End Function
'@)
    $form=$project.VBComponents.Item('frmAddInventoryItem').CodeModule
    $start=$form.ProcStartLine('UserForm_Activate',0);$count=$form.ProcCountLines('UserForm_Activate',0)
    $body=$form.Lines($start,$count)
    $boundary='    If Not mAnchors Is Nothing Then mAnchors.ResizeControls'
    if([regex]::Matches($body,[regex]::Escape($boundary)).Count -ne 1){throw 'Inventory activation observation boundary changed; not product RED.'}
    $body=$body.Replace($boundary,($boundary+"`r`n"+'    TestInventoryEntry.Drive Me'))
    $form.DeleteLines($start,$count);$form.InsertLines($start,$body)
    $form.AddFromString(@'
Public Function InventoryEntryAcceptForTest() As Boolean
    mTxtItemName.Value = "Inventory entry fixture"
    mCmbUom.Value = "LB"
    mCmbLocation.Value = "CLEARVIEW"
    mTxtQty.Value = "1"
    mBtnOK_Click
    InventoryEntryAcceptForTest = mAccepted
End Function
Public Sub InventoryEntryCancelForTest()
    mBtnCancel_Click
End Sub
Public Function InventoryEntryLegacyDisplayForTest() As Boolean
    InventoryEntryLegacyDisplayForTest = (InStr(1, mLblGenerated.Caption, "ROW:", vbTextCompare) > 0)
End Function
Public Function InventoryEntryOwnerBoundForTest(ByVal hostName As String) As Boolean
    Dim expected As Workbook
    Set expected = Application.Workbooks(hostName)
    InventoryEntryOwnerBoundForTest = (mOperatorWorkbook Is expected)
End Function
Public Function InventoryEntryCapturedKindForTest(ByVal hostName As String, ByVal root As String, _
                                                 ByVal warehouseId As String) As String
    InventoryEntryCapturedKindForTest = "OtherWorkbook"
    If mOperatorWorkbook Is Nothing Then InventoryEntryCapturedKindForTest = "NoWorkbook": Exit Function
    If InventoryEntryOwnerBoundForTest(hostName) Then InventoryEntryCapturedKindForTest = "ExpectedOperator": Exit Function
    Dim prefix As String
    prefix = root & "\" & warehouseId
    If StrComp(mOperatorWorkbook.FullName, prefix & ".invSys.Config.xlsb", vbTextCompare) = 0 Then InventoryEntryCapturedKindForTest = "SourceConfig"
    If StrComp(mOperatorWorkbook.FullName, prefix & ".invSys.Auth.xlsb", vbTextCompare) = 0 Then InventoryEntryCapturedKindForTest = "SourceAuth"
    If StrComp(mOperatorWorkbook.FullName, prefix & ".invSys.Data.Inventory.xlsb", vbTextCompare) = 0 Then InventoryEntryCapturedKindForTest = "SourceInventory"
End Function
'@)
    $owner=$project.VBComponents.Item('modAdmin').CodeModule
    $start=$owner.ProcStartLine('AddInventoryItemForWarehouse',0);$count=$owner.ProcCountLines('AddInventoryItemForWarehouse',0)
    $body=$owner.Lines($start,$count)
    $boundary='    payloadJson = modRoleEventWriter.BuildPayloadJsonFromCollection(payloadItems)'
    if([regex]::Matches($body,[regex]::Escape($boundary)).Count -eq 1){
        $body=$body.Replace($boundary,($boundary+"`r`n"+'    TestInventoryEntry.RecordPayload payloadJson'))
    }else{
        # The owning Add boundary still observes the completed payload after the
        # shared builder returns; readonly receipt preflight is not an action.
        $pattern='(?im)^[\t ]*payloadJson = BuildInventoryCreatePayloadAdmin\([^\r\n]*\r?\n[^\r\n]*customFields\)[\t ]*\r?$'
        if([regex]::Matches($body,$pattern).Count -ne 1){throw 'Inventory payload observation boundary changed; not product RED.'}
        $body=[regex]::Replace($body,$pattern,[Text.RegularExpressions.MatchEvaluator]{param($m) $m.Value.TrimEnd("`r")+"`r`n"+'    TestInventoryEntry.RecordPayload payloadJson'})
    }
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$body)
    $owner.AddFromString(@'
Private Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
                        Optional ByVal Title As String = "") As VbMsgBoxResult
    TestInventoryEntry.Notice
    TestInventoryCommands.RecordNotice Prompt
    MsgBox = vbOK
End Function
'@)
    Install-InventoryCommandBoundaryProbe
}

function Get-AdminInventoryEntryHash([string]$Path) {
    $stream=[IO.FileStream]::new($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
    try {(Get-FileHash -InputStream $stream).Hash}
    finally {$stream.Dispose()}
}

function Get-AdminInventoryEntryFacts($Fixture,[switch]$Persisted) {
    $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $book=$null;$opened=$false
    foreach($candidate in $excel.Workbooks){if($candidate.FullName -ieq $path){$book=$candidate;break}}
    # Disposable fixture only: discard unsaved memory, never save for the owner.
    if($Persisted -and $null -ne $book){$book.Close($false);$book=$null}
    if($null -eq $book){$book=$excel.Workbooks.Open($path,0,$true);$opened=$true}
    try {
        $log=Table $book 'tblInventoryLog';$applied=Table $book 'tblAppliedEvents'
        $legacy=@($book.Worksheets|ForEach-Object {$_.ListObjects}|ForEach-Object {$_.ListColumns}|Where-Object {$_.Name.Trim().ToUpperInvariant() -ceq 'ROW'}).Count
        $balances=[Collections.Generic.Dictionary[string,double]]::new([StringComparer]::Ordinal)
        foreach($row in $log.ListRows){
            $key=[string]$row.Range.Cells.Item(1,$log.ListColumns.Item('System_Key').Index).Value2
            if($key){
                $delta=[double]$row.Range.Cells.Item(1,$log.ListColumns.Item('QtyDelta').Index).Value2
                if(-not $balances.ContainsKey($key)){$balances.Add($key,0)}
                $balances[$key]+=$delta
            }
        }
        [pscustomobject]@{LogRows=$log.ListRows.Count;AppliedRows=$applied.ListRows.Count;LegacyHeaders=$legacy;DistinctKeys=$balances.Count;EntityQuantities=$balances;Saved=$book.Saved;FileHash=(Get-AdminInventoryEntryHash $path)}
    } finally {if($opened){$book.Close($false)}}
}

function Test-AdminInventoryEntry($Fixture,$Other) {
    $script:InventoryConfigReadChecks=[Collections.Generic.List[object]]::new()
    SelectTarget $Fixture
    $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($Fixture.Warehouse,'S1','config-admin'))
    Check 'InventoryEntry.GeneratedFixtureSeedApplied' $seed.StartsWith('OK|')
    if(-not $seed.StartsWith('OK|')){throw 'Generated inventory fixture unavailable; not product RED.'}
    $hostBook=$excel.Workbooks.Add();$otherBook=$excel.Workbooks.Add()
    $hostBook.SaveAs((Join-Path $runRoot 'inventory-entry-host.xlsx'),51)
    $otherBook.SaveAs((Join-Path $runRoot 'inventory-entry-unrelated.xlsx'),51)
    $otherSheetCount=$otherBook.Worksheets.Count
    $authHash=Get-AdminInventoryEntryHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))
    $configHash=Get-AdminInventoryEntryHash $Fixture.Config
    $otherInventoryHash=Get-AdminInventoryEntryHash (Join-Path $Other.Root ($Other.Warehouse+'.invSys.Data.Inventory.xlsb'))
    $otherAuthHash=Get-AdminInventoryEntryHash (Join-Path $Other.Root ($Other.Warehouse+'.invSys.Auth.xlsb'))
    $otherConfigHash=Get-AdminInventoryEntryHash $Other.Config
    try {
        foreach($mode in @('Baseline','ActiveWorkbook','SignedOut','OtherWarehouseAdmin','ReplacementSession')){
            SelectTarget $Fixture
            $before=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $target=if($mode -ceq 'OtherWarehouseAdmin'){$Other}else{$Fixture}
            $configBefore=Get-InventoryConfigDiagnosticSnapshot $Fixture
            [void](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Reset')
            [void](Run 'invSys.Admin.xlam' 'TestInventoryEntry.Execute' @($hostBook.Name,$mode,$target.Root,$target.Secret,$otherBook.Name,$Fixture.Root,$Fixture.Warehouse))
            $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryEntry.Facts')).Split('|')
            $after=Get-AdminInventoryEntryFacts $Fixture
            $persisted=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $configAfter=Get-InventoryConfigDiagnosticSnapshot $Fixture
            $saveFacts=([string](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Facts')).Split('|')
            if($saveFacts.Count -ne 3){throw 'Config save observation unavailable; not product RED.'}
            $phase=[pscustomobject]@{Calls=[int]$saveFacts[0];ByteChangesDuringSave=[int]$saveFacts[1];UnavailableHashComparisons=[int]$saveFacts[2];Procedure='modConfig.SaveConfigWorkbookIfWritable'}
            Write-InventoryConfigDiagnosticComparison $configBefore $configAfter $mode $phase
            $newKeys=@($persisted.EntityQuantities.Keys|Where-Object {-not $before.EntityQuantities.ContainsKey($_)})
            $exactSaved=($newKeys.Count -eq 1 -and $persisted.EntityQuantities[$newKeys[0]] -eq 1 -and $persisted.FileHash -cne $before.FileHash)
            # Historical assertion ID denotes reaching the operator's OK attempt;
            # an authorized control must accept, while a legal refusal need not.
            $mustAccept=$mode -cin @('Baseline','ActiveWorkbook')
            $actionReached=$facts.Count -eq 10 -and [int]$facts[1] -eq 1 -and $facts[3] -ceq 'True' -and (-not $mustAccept -or $facts[2] -ceq 'True')
            Check ('InventoryEntry.'+$mode+'.ActualModalAcceptReached') $actionReached
            if(-not $actionReached){throw 'Inventory modal action fixture did not reach OK; not product RED.'}
            Check ('InventoryEntry.'+$mode+'.OriginalHostCapturedBeforeYield') ($facts[8] -ceq 'True')
            $wrote=$after.LogRows -gt $before.LogRows -or $after.AppliedRows -gt $before.AppliedRows
            if($mode -cin @('Baseline','ActiveWorkbook')){
                Check ('InventoryEntry.'+$mode+'.OwningCreateApplied') ($after.LogRows -eq $before.LogRows+1 -and $after.AppliedRows -eq $before.AppliedRows+1)
                Check ('InventoryEntry.'+$mode+'.NewDistinctEntity') ($after.DistinctKeys -eq $before.DistinctKeys+1)
                Check ('InventoryEntry.'+$mode+'.ExactStartingQuantitySaved') $exactSaved
                Check ('InventoryEntry.'+$mode+'.CreatePayloadContainsNoROW') ($facts[5] -ceq 'True' -and $facts[6] -ceq 'False')
            }elseif($mode -cne 'ReplacementSession'){
                Check ('InventoryEntry.'+$mode+'.NoInventoryMutation') (-not $wrote)
            }
            Check ('InventoryEntry.'+$mode+'.ManagedHeadersContainNoROW') ($after.LegacyHeaders -eq 0)
            Check ('InventoryEntry.'+$mode+'.DisplayContainsNoROWIdentity') ($facts[7] -ceq 'False')
            Check ('InventoryEntry.'+$mode+'.UnrelatedWorkbookStructurePreserved') ($otherBook.Worksheets.Count -eq $otherSheetCount)
            [pscustomobject]@{Case=$mode;ModalHandlerReached=$actionReached;ModalAcceptReached=($facts[2] -ceq 'True');CapturedHostKind=$facts[9];InventoryWriteObserved=$wrote;OwnerWorkbookSavedFlag=$after.Saved;PersistedInventoryBytesChanged=($persisted.FileHash -cne $before.FileHash);ExactNewEntityQuantityPersisted=$exactSaved;SourceAuthUnchanged=($authHash -ceq (Get-AdminInventoryEntryHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))));SourceConfigUnchanged=($configHash -ceq (Get-AdminInventoryEntryHash $Fixture.Config));OriginalDiagnosticDidNotAssertReplacementRefusal=($mode -ceq 'ReplacementSession')}|
                ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'inventory-entry-observations.jsonl')
        }
        Check 'InventoryEntry.SourceAuthAndConfigPreserved' ($authHash -ceq (Get-AdminInventoryEntryHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))) -and $configHash -ceq (Get-AdminInventoryEntryHash $Fixture.Config))
        Check 'InventoryEntry.OtherWarehouseInventoryPreserved' ($otherInventoryHash -ceq (Get-AdminInventoryEntryHash (Join-Path $Other.Root ($Other.Warehouse+'.invSys.Data.Inventory.xlsb'))))
        Check 'InventoryEntry.OtherWarehouseAuthorityPreserved' ($otherAuthHash -ceq (Get-AdminInventoryEntryHash (Join-Path $Other.Root ($Other.Warehouse+'.invSys.Auth.xlsb'))) -and $otherConfigHash -ceq (Get-AdminInventoryEntryHash $Other.Config))
    } finally {$hostBook.Close($false);$otherBook.Close($false)}
}
