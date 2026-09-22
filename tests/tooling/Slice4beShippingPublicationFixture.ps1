# D18/D14 fixture: owners create inventory/BOM/staging. Only unknown fixture
# columns are added directly. Values remain in memory; reports contain checks.
function Remove-Slice4beShippingPublicationFiles($Paths) {
    foreach($path in $Paths) {
        # Each absolute file was checked absent before the owning actions ran.
        # Never enumerate/delete another warehouse's profile-local state.
        if(Test-Path -LiteralPath $path){Remove-Item -LiteralPath $path -Force}
    }
}

function New-Slice4beShippingPublicationFixture($Fixture,[ref]$Result) {
    $localRoot=Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'invSys'
    if($Fixture.Warehouse -cnotmatch '^WHD5[A-F0-9]{6}$'){throw 'Shipping publication fixture identity is not disposable.'}
    $localFiles=@('hold','active','sent','projected' | ForEach-Object {
        [IO.Path]::GetFullPath((Join-Path $localRoot ('shipping_'+$_+'_'+$Fixture.Warehouse+'.tsv')))
    })
    foreach($path in $localFiles){if(Test-Path -LiteralPath $path){throw 'Shipping publication fixture would overwrite existing local state.'}}
    # Ignored fixture-only cleanup receipt records exact absent-before paths.
    # No operational rows, identities, credentials or form text are written.
    $localFiles | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $reportRoot 'shipping-fixture-owned-files.json')
    $operator=$null;$dialogJob=$null;$succeeded=$false
    $module=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modTS_Shipments').CodeModule
    $form=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmShipmentsTally').CodeModule
    $formSource=$form.Lines(1,$form.CountOfLines)
    # Reuse the existing proven form-action facades, without its interruption,
    # authority or submission probes. Select only the form here-string via AST.
    $tokens=$null;$errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot 'Slice4beShippingActivity.ps1'),[ref]$tokens,[ref]$errors)
    if($errors.Count){throw 'Shipping form fixture source does not parse.'}
    $assignment=@($ast.FindAll({param($node) $node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -ceq '$formFacade'},$true) | Where-Object {$_.Right.Extent.Text.StartsWith("@'" )})
    if($assignment.Count -ne 1){throw 'Shipping form fixture declaration is ambiguous.'}
    $facadeSource=$assignment[0].Right.Extent.Text
    $facade=''
    foreach($name in @('ActivityShippingWorkbook','ActivityShippingCreateBox','ActivityShippingPrepare','ActivityShippingClick')) {
        $matches=[regex]::Matches($facadeSource,'(?ms)^Public (?:Function|Sub) '+$name+'\(.*?^End (?:Function|Sub)')
        if($matches.Count -ne 1){throw 'Shipping form-action fixture is unavailable.'}
        $facade+=$matches[0].Value+"`r`n"
    }
    # The same Box Designer handler now adds two distinct available components.
    $old='mBtnBoxBuilderAddComponent_Click'+"`r`n"+'    If mLstBoxBuilderComponents.ListCount <> 1 Then Exit Function'
    $new=@'
mBtnBoxBuilderAddComponent_Click
    chosen = False
    For i = mLstBoxBuilderInventory.ListIndex + 1 To mLstBoxBuilderInventory.ListCount - 1
        If ParseNumber(NzText(mLstBoxBuilderInventory.List(i, 6))) >= 10 Then
            mLstBoxBuilderInventory.ListIndex = i: chosen = True: Exit For
        End If
    Next i
    If Not chosen Then Exit Function
    mTxtBoxBuilderComponentQty.Value = "1"
    mBtnBoxBuilderAddComponent_Click
    If mLstBoxBuilderComponents.ListCount <> 2 Then Exit Function
'@
    $facade=$facade -replace "`r?`n","`r`n"
    if(-not $facade.Contains($old)){throw 'Two-component fixture anchor is unavailable.'}
    $facade=$facade.Replace($old,($new -replace "`r?`n","`r`n")).Replace('If mLstBoxMakerComponents.ListCount <> 1','If mLstBoxMakerComponents.ListCount <> 2')
    if($formSource -notmatch '(?im)^Private Function NzText\('){
        $facade=$facade.Replace('NzText(', 'modShippingFormValues.NzText(').Replace('ParseNumber(', 'modShippingFormValues.ParseNumber(')
    }
    $form.AddFromString($facade)
    $module.AddFromString(@'
Public Function PublicationShippingOpenForTest() As String
    BtnOpenShipmentsForm
    If mShipmentsLauncherForm Is Nothing Then Exit Function
    mShipmentsLauncherForm.CancelAutoSync
    PublicationShippingOpenForTest = mShipmentsLauncherForm.ActivityShippingWorkbook()
End Function
Public Function PublicationShippingBoxForTest() As Boolean
    PublicationShippingBoxForTest = mShipmentsLauncherForm.ActivityShippingCreateBox()
End Function
Public Function PublicationShippingActionForTest(ByVal action As String) As String
    PublicationShippingActionForTest = mShipmentsLauncherForm.ActivityShippingPrepare(action)
    If PublicationShippingActionForTest <> "" Then mShipmentsLauncherForm.ActivityShippingClick action
End Function
Public Sub PublicationShippingCloseForTest()
    If Not mShipmentsLauncherForm Is Nothing Then Unload mShipmentsLauncherForm
    Set mShipmentsLauncherForm = Nothing
    Set mShipmentsAutoSyncForm = Nothing
    mShipmentsLauncherWorkbookName = ""
End Sub
Public Function PublicationShippingColumnForTest(ByVal workbookName As String) As Boolean
    Dim wb As Workbook, ws As Worksheet, lo As ListObject, column As ListColumn
    Set wb = Application.Workbooks(workbookName)
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = "tblShippingBOM" Then
                Set column = lo.ListColumns.Add(1)
                column.Name = "Publication fixture extra"
                column.DataBodyRange.Value2 = "DO-NOT-DISPLAY"
                wb.Save
                PublicationShippingColumnForTest = (column.Name = "Publication fixture extra" And wb.Saved)
                Exit Function
            End If
        Next lo
    Next ws
End Function
Public Function PublicationShippingDirtyForTest(ByVal workbookName As String) As Boolean
    Dim wb As Workbook, ws As Worksheet, lo As ListObject
    Set wb = Application.Workbooks(workbookName)
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = "tblShippingBOM" Then
                lo.ListColumns("Publication fixture extra").DataBodyRange.Value2 = "UNSAVED-FIXTURE"
                PublicationShippingDirtyForTest = Not wb.Saved
                Exit Function
            End If
        Next lo
    Next ws
End Function
'@)
    $dialogStop=Join-Path $runRoot 'publication-shipping-dialog.stop'
    try {
        $auth=$excel.Workbooks.Open((Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')),0,$false)
        try {
            $caps=Table $auth 'tblCapabilities';$row=$caps.ListRows.Add()
            foreach($pair in @{UserId='config-reader';Capability='SHIP_POST';WarehouseId=$Fixture.Warehouse;StationId='S1';Status='Active'}.GetEnumerator()){$row.Range.Cells.Item(1,$caps.ListColumns.Item($pair.Key).Index).Value2=$pair.Value}
            $auth.Save()
        } finally {$auth.Close($false)}
        SelectTarget $Fixture 'config-admin'
        $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($Fixture.Warehouse,'S1','config-admin'))
        if(-not $seed.StartsWith('OK|')){throw 'Publication Shipping fixture seed failed.'}
        SelectTarget $Fixture 'config-reader'
        [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride' @((Join-Path $repo 'deploy/current/templates')))
        [void](Run 'invSys.Core.xlam' 'modWarehouseBootstrap.SetLocalOperatorRootOverrideForAutomation' @((Join-Path $runRoot 'publication-shipping-operators')))
        $name=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.PublicationShippingOpenForTest')
        if(-not $name){throw 'Publication Shipping launcher did not establish its form.'}
        $operator=$excel.Workbooks.Item($name)
        $observerPath=Join-Path $repo 'tools/plan022-dialog-observer.ps1'
        . $observerPath
        if(-not ('Plan022NativeDialogs' -as [type])){Invoke-Plan022NativeDialogObservation -ProcessId 0 -TimeoutSeconds 0}
        [uint32]$ownedProcess=0
        [void][Plan022NativeDialogs]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$ownedProcess)
        if($ownedProcess -eq 0){throw 'Publication fixture process unavailable.'}
        $dialogJob=Start-Job -ArgumentList $observerPath,$ownedProcess,$dialogStop -ScriptBlock {
            param($observer,$owned,$stop)
            . $observer
            Invoke-Plan022NativeDialogObservation -ProcessId $owned -TimeoutSeconds 180 -StopPath $stop | Out-Null
        }
        $created=[bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.PublicationShippingBoxForTest')
        if(-not $created){throw 'Publication two-component box fixture was not created by the form handlers.'}
        Check 'ViewerPublication.ShippingFixture.ActualBoxDesignerAndMaker' $created
        $boxMakerPublication=$null
        $ownerPublicationPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')
        if(Test-Path -LiteralPath $ownerPublicationPath){$boxMakerPublication=[IO.File]::ReadAllText($ownerPublicationPath)|ConvertFrom-Json}
        $added=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.PublicationShippingActionForTest' @('Add'))
        $held=[string](Run 'invSys.Operations.xlam' 'modTS_Shipments.PublicationShippingActionForTest' @('Hold'))
        $hold=Table $operator 'NotShipped'
        if(-not $added -or $added -cne $held -or $hold.ListRows.Count -ne 1){throw 'Publication held shipment was not prepared by Add/Hold handlers.'}
        Check 'ViewerPublication.ShippingFixture.ActualAddAndHold' $true
        if($CaptureEvidence){CaptureOwnedFormByCaptionEvidence 'Shipping Shipments' 'publication-shipping-held.png'}
        [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.PublicationShippingCloseForTest')
        $operator.Save();$operator.Close($false);$operator=$null
        $bomPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.ShippingBOM.xlsb')
        $bomBook=$excel.Workbooks.Open($bomPath,0,$false)
        try {
            if(-not [bool](Run 'invSys.Operations.xlam' 'modTS_Shipments.PublicationShippingColumnForTest' @($bomBook.Name))){throw 'Shipping BOM unknown-column fixture was not saved by the calibrated VBA edit.'}
            $bom=Table $bomBook 'tblShippingBOM'
            $bomRows=@(foreach($row in $bom.ListRows){
                $fields=[ordered]@{}
                foreach($column in $bom.ListColumns){if($column.Name -cne 'Publication fixture extra'){$fields[$column.Name]=[string]$row.Range.Cells.Item(1,$column.Index).Value2}}
                [pscustomobject]$fields
            })
        } finally {$bomBook.Close($false)}
        $reopenedBom=$excel.Workbooks.Open($bomPath,0,$true)
        try {
            if(-not $reopenedBom.ReadOnly -or (Table $reopenedBom 'tblShippingBOM').ListRows.Count -ne 2){throw 'Saved Shipping BOM fixture did not reopen with its two owner rows.'}
            Check 'ViewerPublication.ShippingFixture.SavedBomReopensReadOnly' $true
        } finally {$reopenedBom.Close($false)}
        $holdPath=$localFiles[0]
        $heldRows=@(foreach($line in [IO.File]::ReadAllLines($holdPath)){
            if(-not $line){continue}
            $values=$line.Split([char]9)
            if($values.Count -ne 12){throw 'Held fixture field contract is unavailable.'}
            $fields=[ordered]@{};$i=0
            foreach($field in @('Ref','Item','Qty','','UOM','Location','Description','Area','Carrier','ShipmentLineId','ReserveEventId','System_Key')){
                if($field){$fields[$field]=$values[$i]};$i++
            }
            [pscustomobject]$fields
        })
        if($bomRows.Count -ne 2 -or @($bomRows|Where-Object {$_.PackageSystemKey -ceq $added}).Count -ne 2 -or $heldRows.Count -ne 1 -or $heldRows[0].System_Key -cne $added -or -not $heldRows[0].ShipmentLineId){throw 'Owning Shipping source rows failed fixture calibration.'}
        if($bomRows[0].ComponentSystemKey -ceq $bomRows[1].ComponentSystemKey){throw 'BOM fixture lacks two distinct exact component keys.'}
        Check 'ViewerPublication.ShippingFixture.PersistedOwnerRowsCalibrated' $true
        $inventoryPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
        foreach($book in @($excel.Workbooks)){
            if($book.FullName -ieq $inventoryPath){
                # Do not turn unsaved owner/projection state into durable evidence.
                # Close the generated handle without save and verify saved events.
                $book.Close($false)
            }
        }
        $persisted=$excel.Workbooks.Open($inventoryPath,0,$true)
        try {
            $log=Table $persisted 'tblInventoryLog';$builds=0;$buildId=''
            foreach($row in $log.ListRows){
                if([string]$row.Range.Cells.Item(1,$log.ListColumns.Item('System_Key').Index).Value2 -ceq $added -and
                   [string]$row.Range.Cells.Item(1,$log.ListColumns.Item('EventType').Index).Value2 -ceq 'BOX_BUILD' -and
                   [double]$row.Range.Cells.Item(1,$log.ListColumns.Item('QtyDelta').Index).Value2 -eq 10){$builds++;$buildId=[string]$row.Range.Cells.Item(1,$log.ListColumns.Item('EventID').Index).Value2}
            }
            if($builds -ne 1){throw 'Box Maker did not persist its owning inventory event.'}
            Check 'ViewerPublication.ShippingFixture.OwnerBuildWasAlreadyDurable' $true
        } finally {$persisted.Close($false)}
        # Prove ordinary publication through the real Box Maker handler before
        # the explicit Admin publication command or volume-source substitution.
        $publishedBuild=$false
        if($null -ne $boxMakerPublication){
            $ownerGroups=@($boxMakerPublication.Groups|Where-Object {$_.Source -ceq 'Inventory' -and $_.SourceId -ceq $buildId})
            if($ownerGroups.Count -eq 1){$publishedBuild=@($ownerGroups[0].Lines|Where-Object {$_.EventType -ceq 'BOX_BUILD' -and $_.System_Key -ceq $added -and [string]$_.QtyDelta -ceq '10'}).Count -eq 1}
        }
        Check 'ViewerPublication.ShippingFixture.ActualBoxMakerPublishedOwnerEvent' $publishedBuild
        $pins=@{};foreach($path in $localFiles){if(Test-Path -LiteralPath $path){$pins[$path]=PublicationSourceHash $path}}
        $succeeded=$true
        $Result.Value=[pscustomobject]@{BomRows=$bomRows;HoldRows=$heldRows;LocalFiles=$localFiles;Pins=$pins}
    } finally {
        if($null -ne $dialogJob){
            [IO.File]::WriteAllText($dialogStop,'stop')
            [void](Wait-Job $dialogJob -Timeout 5)
            if($dialogJob.State -eq 'Running'){Stop-Job $dialogJob}
            Remove-Job $dialogJob
        }
        [void](Run 'invSys.Operations.xlam' 'modTS_Shipments.PublicationShippingCloseForTest')
        if($null -ne $operator){$operator.Close($false)}
        if(-not $succeeded){Remove-Slice4beShippingPublicationFiles $localFiles}
        SelectTarget $Fixture 'config-admin'
    }
}

function Test-Slice4beShippingPublication($Artifact,$Fixture) {
    $state=@();$sources=@()
    if($null -ne $Artifact){$state=@($Artifact.CurrentState);$sources=@($Artifact.Coverage.Sources)}
    foreach($source in @('ShippingBOM','ShippingHolds')) {
        $entries=@($state|Where-Object {$_.Source -ceq $source})
        $lines=@($entries|ForEach-Object {$_.Lines})
        $expected=@(if($source -ceq 'ShippingBOM'){$Fixture.BomRows}else{$Fixture.HoldRows})
        $preserved=$lines.Count -eq $expected.Count
        foreach($row in $expected){
            $matched=@($lines|Where-Object {
                if($source -ceq 'ShippingBOM'){$_.PackageSystemKey -ceq $row.PackageSystemKey -and $_.ComponentSystemKey -ceq $row.ComponentSystemKey -and [string]$_.BomVersion -ceq $row.BomVersion}
                else {$_.System_Key -ceq $row.System_Key -and $_.ShipmentLineId -ceq $row.ShipmentLineId}
            })
            $preserved=$preserved -and $matched.Count -eq 1
            if($matched.Count -eq 1){
                $fields=if($source -ceq 'ShippingBOM'){@('PackageItem','PackageUOM','PackageLocation','PackageDescription','BomVersionLabel','IsActive','EffectiveFromUTC','EffectiveToUTC','RetiredAtUTC','ComponentItemCode','ComponentItem','ComponentQty','ComponentUOM','ComponentLocation','ComponentDescription','UpdatedAtUTC','UpdatedBy')}else{@('Ref','Item','Qty','UOM','Location','Description','Area','Carrier','ReserveEventId')}
                foreach($field in $fields){
                    $expectedValue=[string]$row.$field
                    if($source -ceq 'ShippingBOM' -and $field.EndsWith('UTC') -and $expectedValue -ne ''){$expectedValue=[DateTime]::FromOADate([double]$expectedValue).ToString('yyyy-MM-ddTHH:mm:ss')}
                    $preserved=$preserved -and [string]$matched[0].$field -ceq $expectedValue
                }
                $preserved=$preserved -and @($matched[0].PSObject.Properties).Count -eq $(if($source -ceq 'ShippingBOM'){20}else{11})
            }
        }
        Check ('ViewerPublication.'+$source+'.EveryExactOwnerLineRetained') $preserved
        $classified=$entries.Count -gt 0
        foreach($entry in $entries){
            $classified=$classified -and $entry.SourceKind -ceq 'Current state'
            foreach($field in @('SourceId','EventID','RecordedAt','Outcomes')){
                $property=$entry.PSObject.Properties[$field]
                if($null -ne $property -and @($property.Value|Where-Object {$_}).Count -gt 0){$classified=$false}
            }
        }
        Check ('ViewerPublication.'+$source+'.CurrentStateWithoutInventedHistory') $classified
        $coverage=@($sources|Where-Object {$_.Source -ceq $source})
        $scope=if($source -ceq 'ShippingBOM'){'Warehouse'}else{'Station profile'}
        $covered=$coverage.Count -eq 1
        if($covered){$covered=$coverage[0].Availability -ceq 'Available' -and $coverage[0].Scope -ceq $scope -and $coverage[0].AvailableLines -eq $expected.Count -and $coverage[0].IncludedLines -eq $expected.Count -and $coverage[0].OmittedLines -eq 0}
        Check ('ViewerPublication.'+$source+'.AvailableCoverageAndScope') $covered
    }
    $unchanged=$true
    foreach($path in $Fixture.Pins.Keys){$unchanged=$unchanged -and (Test-Path -LiteralPath $path) -and (PublicationSourceHash $path) -ceq $Fixture.Pins[$path]}
    Check 'ViewerPublication.ShippingLocalSourceBytesUnchanged' $unchanged
}
