# Actual Create handler, owned operator workbooks and read-only result inspection.
function Install-InventoryCreatePhaseProbe {
    $project=$packages['invSys.Admin.xlam'].VBProject
    $probe=$project.VBComponents.Add(1);$probe.Name='TestInventoryCreatePhase'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mMode As String, mHost As String, mReached As Boolean
Private mBefore As String, mHash As String, mSuccess As Boolean, mPartial As Boolean
Private mSavedPartial As Boolean
Public Sub Arm(ByVal mode As String, ByVal host As String)
    mMode = mode: mHost = host: mReached = False
    mBefore = "": mHash = "": mSuccess = False: mPartial = False
    mSavedPartial = False
End Sub
Private Function Signature() As String
    Dim book As Workbook, sheet As Worksheet, table As ListObject, cell As Range, result As String
    Set book = Application.Workbooks(mHost)
    For Each sheet In book.Worksheets
        result = result & "|" & sheet.Name & "|" & CStr(sheet.Visible) & "|" & sheet.UsedRange.Address
        For Each cell In sheet.UsedRange.Cells
            If IsError(cell.Value2) Then
                result = result & "|Error"
            Else
                result = result & "|" & CStr(Len(CStr(cell.Value2))) & ":" & CStr(cell.Value2)
            End If
        Next cell
        For Each table In sheet.ListObjects
            result = result & "|" & table.Name & "|" & table.Range.Address
        Next table
    Next sheet
    Signature = result
End Function
Public Sub AtPhase(ByVal phase As String)
    Dim sheet As Worksheet, table As ListObject
    If mReached Then Exit Sub
    If mMode = "ChangedNewTable" Or mMode = "RenamedNewTable" Then
        If phase <> "BeforeSave" Then Exit Sub
        For Each sheet In Application.Workbooks(mHost).Worksheets
            For Each table In sheet.ListObjects
                If Left$(table.Name, 17) = "invSys_Inventory_" Then
                    If mMode = "ChangedNewTable" Then
                        table.ListColumns("Quantity").DataBodyRange.Cells(1, 1).Value2 = 1
                    Else
                        table.Name = table.Name & "_changed"
                    End If
                End If
            Next table
        Next sheet
    ElseIf mMode <> phase Then
        Exit Sub
    End If
    mReached = True
    mBefore = Signature()
    mHash = modTrainingWire.FileSha256(Application.Workbooks(mHost).FullName)
    If mBefore = "" Or mHash = "" Then Err.Raise vbObjectError + 2195, , "Create observation unavailable."
    If mMode <> "ChangedNewTable" And mMode <> "RenamedNewTable" Then modAuth.SignOut
End Sub
Public Sub SaveReturn(ByVal book As Workbook)
    If mMode = "SaveThrows" Then
        mReached = True
        mBefore = Signature(): mHash = modTrainingWire.FileSha256(book.FullName)
        Err.Raise vbObjectError + 2196, , "Owned Create save failure."
    End If
    book.Save
    AtPhase "AfterSave"
End Sub
Public Sub ObserveResult(ByVal form As frmAddInventoryItem)
    Dim values As Variant
    values = form.InventoryCreateResultForTest()
    mSuccess = values(0): mPartial = values(1): mSavedPartial = values(2)
End Sub
Public Function Facts() As String
    Dim sameBuffer As Boolean, sameDisk As Boolean
    If mReached Then
        sameBuffer = (mBefore = Signature())
        sameDisk = (mHash <> "" And mHash = modTrainingWire.FileSha256(Application.Workbooks(mHost).FullName))
    End If
    Facts = CStr(mReached) & "|" & CStr(sameBuffer) & "|" & CStr(sameDisk) & "|" & _
            CStr(mSuccess) & "|" & CStr(mPartial) & "|" & CStr(mSavedPartial)
End Function
'@)
    $form=$project.VBComponents.Item('frmAddInventoryItem').CodeModule
    $form.AddFromString(@'
Public Function InventoryCreateResultForTest() As Variant
    InventoryCreateResultForTest = Array(mLastInventoryTableName <> "", _
        InStr(1, mLastInventoryWorksheetReport, "partial", vbTextCompare) > 0, _
        InStr(1, mLastInventoryWorksheetReport, "saved", vbTextCompare) > 0 And _
        InStr(1, mLastInventoryWorksheetReport, "unconfirmed", vbTextCompare) > 0)
End Function
'@)
    $start=$form.ProcStartLine('mBtnCreateInventoryTable_Click',0);$count=$form.ProcCountLines('mBtnCreateInventoryTable_Click',0)
    $body=$form.Lines($start,$count);$line='(?im)^    If Not CurrentContextAllowed\(\) Then (?:Exit Sub|GoTo Refused)\r?$'
    if([regex]::Matches($body,$line).Count -ne 1){throw 'Create form entry seam unavailable; not product RED.'}
    $body=[regex]::Replace($body,$line,[Text.RegularExpressions.MatchEvaluator]{param($m) $m.Value.TrimEnd("`r")+"`r`n"+'    TestInventoryCreatePhase.AtPhase "BeforeHelper"'})
    $form.DeleteLines($start,$count);$form.InsertLines($start,$body)
    $driver=$project.VBComponents.Item('TestInventoryCommands').CodeModule
    $start=$driver.ProcStartLine('Drive',0);$count=$driver.ProcCountLines('Drive',0)
    $body=$driver.Lines($start,$count);$line='        form.InventoryCommandInvokeForTest mAction'
    if([regex]::Matches($body,[regex]::Escape($line)).Count -ne 1){throw 'Actual Create result observation unavailable.'}
    $body=$body.Replace($line,$line+"`r`n"+'        If mAction = "Create" Then TestInventoryCreatePhase.ObserveResult form')
    $driver.DeleteLines($start,$count);$driver.InsertLines($start,$body)
    $owner=$project.VBComponents.Item('modAdminInventoryWorksheet').CodeModule
    $start=$owner.ProcStartLine('CreateInventoryWorksheetTable',0);$count=$owner.ProcCountLines('CreateInventoryWorksheetTable',0)
    $body=$owner.Lines($start,$count)
    foreach($seam in @(
        @('(?im)^    Set ws = EnsureInventoryEditorSheet\(wb(?:, interactiveContext, report)?\)\r?$','AfterSheet'),
        @('(?im)^    Next idx\r?$','AfterHeaders'),
        @('(?im)^    FormatInventoryWorksheet ws, lo(?:, interactiveContext, report)?\r?$','BeforeSave')
    )){
        if([regex]::Matches($body,$seam[0]).Count -ne 1){throw 'Create mutation phase unavailable; not product RED.'}
        $body=[regex]::Replace($body,$seam[0],[Text.RegularExpressions.MatchEvaluator]{param($m) $m.Value.TrimEnd("`r")+"`r`n"+'    TestInventoryCreatePhase.AtPhase "'+$seam[1]+'"'})
    }
    if([regex]::Matches($body,'(?im)^    wb\.Save\r?$').Count -ne 1){throw 'Create save-return seam unavailable.'}
    $body=[regex]::Replace($body,'(?im)^    wb\.Save\r?$','    TestInventoryCreatePhase.SaveReturn wb')
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$body)
}

function Test-InventoryCreatePhases($Fixture,$Other) {
    $script:InventoryConfigReadChecks=[Collections.Generic.List[object]]::new()
    $cases=if($AdminInventoryCreateNameOnly){@('Authorized','RenamedNewTable')}else{@('Authorized','AuthorizedExisting','BeforeHelper','AfterSheet','AfterHeaders','BeforeSave','ChangedNewTable','RenamedNewTable','AfterSave','SaveThrows')}
    foreach($case in $cases){
        SelectTarget $Fixture
        $hostBook=$excel.Workbooks.Add();$path=Join-Path $runRoot ('create-phase-'+$case+'.xlsx')
        $hostBook.Worksheets.Item(1).Range('A1').Value2='Unrelated operator content'
        $hostBook.SaveAs($path,51)
        try{
            $priorTableName=''
            if($case -ceq 'AuthorizedExisting'){
                if(-not [bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.PrepareBoundaryTable' @($hostBook.Name))){throw 'Existing owned Create fixture unavailable.'}
                $priorTable=@($hostBook.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -like 'invSys_Inventory_*'})[0]
                $priorTableName=[string]$priorTable.Name
                $extension=$priorTable.ListColumns.Add();$extension.Name='USER NOTE'
                $extension.DataBodyRange.Cells.Item(1,1).Value2='Preserved existing staging extension'
                $hostBook.Save();$priorTable.DataBodyRange.Cells.Item(1,1).Select()
            }
            $before=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $configBefore=Get-InventoryConfigDiagnosticSnapshot $Fixture
            $authBefore=Get-AdminInventoryEntryHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))
            [void](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Reset')
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCreatePhase.Arm' @($case,$hostBook.Name))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetSelection' @($Other.Root,''))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,'Create','Authorized','ModalReturn',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
            $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
            $id='InventoryCreatePhase.'+$case
            $ready=$facts.Count -eq 5 -and $facts[0] -ceq 'True' -and $facts[1] -ceq 'True' -and $facts[2] -ceq '1'
            Check ($id+'.ActualCreateHandlerReached') $ready
            if(-not $ready){throw 'Actual Create handler unavailable; not product RED.'}
            Check ($id+'.OriginalOperatorCaptured') ($facts[3] -ceq 'True')
            $phaseFacts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCreatePhase.Facts')).Split('|')
            if($phaseFacts.Count -ne 6){throw 'Create phase facts unavailable.'}
            if($case -cnotin @('Authorized','AuthorizedExisting')){
                Check ($id+'.RequiredPhaseReached') ($phaseFacts[0] -ceq 'True')
                if($phaseFacts[0] -cne 'True'){throw 'Create fault phase unavailable; not product RED.'}
                if($case -ceq 'RenamedNewTable'){Check ($id+'.LiveSessionRetained') ([bool](Run 'invSys.Core.xlam' 'modAuth.IsSignedIn'))}
                Check ($id+'.NoFurtherStagingMutation') ($phaseFacts[1] -ceq 'True')
                Check ($id+'.NoFurtherOperatorSave') ($phaseFacts[2] -ceq 'True')
                Check ($id+'.NoFalseSuccess') ($phaseFacts[3] -ceq 'False')
                if($case -cne 'BeforeHelper'){Check ($id+'.HonestPartialReported') ($phaseFacts[4] -ceq 'True')}
                if($case -ceq 'AfterSave'){Check ($id+'.SavedUnconfirmedReported') ($phaseFacts[5] -ceq 'True')}
            }else{
                Check ($id+'.SuccessReported') ($phaseFacts[3] -ceq 'True')
                $hash=Get-AdminInventoryEntryHash $path
                $copy=Join-Path $runRoot ('create-cold-'+[guid]::NewGuid().ToString('N')+'.xlsx')
                Copy-Item -LiteralPath $path -Destination $copy
                $cold=$excel.Workbooks.Open($copy,0,$true)
                try{
                    $tables=@($cold.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -like 'invSys_Inventory_*'})
                    Check ($id+'.IndependentTableSaved') ($tables.Count -eq $(if($priorTableName){2}else{1}) -and @($tables|Where-Object {$_.ListRows.Count -ne 20}).Count -eq 0)
                    if($priorTableName){
                        $prior=@($tables|Where-Object {$_.Name -ceq $priorTableName})
                        Check ($id+'.UnknownExistingColumnSaved') ($prior.Count -eq 1 -and [string]$prior[0].ListColumns.Item('USER NOTE').DataBodyRange.Cells.Item(1,1).Value2 -ceq 'Preserved existing staging extension')
                    }
                    $names=@($tables[0].ListColumns|ForEach-Object {$_.Name})
                    Check ($id+'.NoPhysicalIdentityHeaders') ('System_Key' -cnotin $names -and 'ROW' -cnotin $names)
                    Check ($id+'.UnrelatedContentSaved') ([string]$cold.Worksheets.Item(1).Range('A1').Value2 -ceq 'Unrelated operator content')
                }finally{$cold.Close($false)}
                Check ($id+'.ColdInspectionPreservedBytes') ((Get-AdminInventoryEntryHash $path) -ceq $hash)
            }
            Check ($id+'.UnrelatedContentPreserved') ([string]$hostBook.Worksheets.Item(1).Range('A1').Value2 -ceq 'Unrelated operator content')
            $after=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $same=$after.LogRows -eq $before.LogRows -and $after.AppliedRows -eq $before.AppliedRows -and $after.EntityQuantities.Count -eq $before.EntityQuantities.Count
            foreach($key in $before.EntityQuantities.Keys){if(-not $after.EntityQuantities.ContainsKey($key) -or $after.EntityQuantities[$key] -ne $before.EntityQuantities[$key]){$same=$false}}
            Check ($id+'.CanonicalInventoryPreserved') $same
            $configAfter=Get-InventoryConfigDiagnosticSnapshot $Fixture
            $saveFacts=([string](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Facts')).Split('|')
            if($saveFacts.Count -ne 3){throw 'Create Config save observations unavailable.'}
            $savePhase=[pscustomobject]@{Calls=[int]$saveFacts[0];ByteChangesDuringSave=[int]$saveFacts[1];UnavailableHashComparisons=[int]$saveFacts[2];Procedure='modConfig.SaveConfigWorkbookIfWritable'}
            Write-InventoryConfigDiagnosticComparison $configBefore $configAfter $case $savePhase
            [pscustomobject]@{Case=$case;AuthBytesPreserved=($authBefore -ceq (Get-AdminInventoryEntryHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))))}|
                ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'create-authority-observations.jsonl')
        }finally{$hostBook.Close($false)}
    }
    Test-InventoryConfigDiagnosticPreservation
}
