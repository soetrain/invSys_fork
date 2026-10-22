# Gestures use the real Inventory Ribbon/form and native MSForms events.
function Install-InventorySelectionProbe {
    $core=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestInventoryRecordingPolicy').CodeModule
    $core.AddFromString(@'
Public Function Prior190Preserved() As Boolean
    Dim id As Variant, before As Object, after As Object
    For Each id In modActivityCatalog.ControlIds(44)
        Set before = modActivityCatalog.Control(CStr(id), 44)
        Set after = modActivityCatalog.Control(CStr(id))
        If before Is Nothing Or after Is Nothing Then Exit Function
        If modTrainingJson.EncodeObject(before) <> modTrainingJson.EncodeObject(after) Then Exit Function
    Next id
    Prior190Preserved = True
End Function
'@)
    $admin=$packages['invSys.Admin.xlam'].VBProject
    $probe=$admin.VBComponents.Add(1);$probe.Name='TestInventorySelections'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mFacts As String
Public Sub Facts(ByVal value As String)
    mFacts = value
End Sub
Public Function Result() As String
    Result = mFacts
End Function
'@)
    $driver=$admin.VBComponents.Item('TestInventoryRecording').CodeModule
    $first=$driver.ProcStartLine('Drive',0);$count=$driver.ProcCountLines('Drive',0)
    $body=$driver.Lines($first,$count)
    $body=$body.Replace('form.RecordingNavigationForTest mMode','form.InventorySelectionsForTest mMode')
    if($body -notmatch 'InventorySelectionsForTest'){throw 'Inventory selection gesture boundary missing.'}
    $driver.DeleteLines($first,$count);$driver.InsertLines($first,$body)
    $first=$driver.ProcStartLine('Execute',0);$count=$driver.ProcCountLines('Execute',0)
    $body=$driver.Lines($first,$count)
    $body=[regex]::Replace($body,'(?im)^Failed:\s*$',('Failed:'+"`r`n"+'    TestInventorySelections.Facts "ExecuteError:" & CStr(Err.Number)'))
    $driver.DeleteLines($first,$count);$driver.InsertLines($first,$body)
    $form=$admin.VBComponents.Item('frmAddInventoryItem').CodeModule
    $form.AddFromString(@'
Public Sub InventorySelectionsForTest(ByVal mode As String)
    Dim picked As Boolean, localFields As Boolean, options As Boolean, index As Long, selected As Long
    TestInventorySelections.Facts "Entered"
    mBtnEditMode_Click
    TestInventorySelections.Facts "EditMode"
    selected = -1
    For index = 0 To mCmbEditItem.ListCount - 1
        If mCmbEditItem.List(index, 1) = "DEMO-SELECT-01" Then selected = index: Exit For
    Next index
    If selected < 0 Then Err.Raise 5
    mCmbEditItem.ListIndex = selected
    TestInventorySelections.Facts "ItemCombo"
    picked = (mSelectedEditSku = "DEMO-SELECT-01")
    mCmbEditItem.Text = "Selection fixture"
    mLstEditItemResults.ListIndex = 0
    TestInventorySelections.Facts "ItemResults"
    picked = picked And (mSelectedEditSku = "DEMO-SELECT-01")
    mLoading = True
    mCmbLocation.AddItem "SelectionLocation"
    mCmbCategory.AddItem "SelectionCategory"
    mCmbLocation.ListIndex = -1: mCmbCategory.ListIndex = -1
    mCmbUom.ListIndex = 0
    mLoading = False
    TestInventorySelections.Facts "PreparedOptions"
    ' Free typing is not a deliberate dropdown choice.
    mCmbLocation.Text = "TypedLocationSentinel"
    mCmbCategory.Text = "TypedCategorySentinel"
    mTxtQty.Text = "12"
    mCmbUom.ListIndex = 1
    mTxtQty.ListIndex = 1
    mCmbLocation.ListIndex = mCmbLocation.ListCount - 1
    mCmbCategory.ListIndex = mCmbCategory.ListCount - 1
    TestInventorySelections.Facts "SelectedOptions"
    options = (mCmbLocation.Value = "SelectionLocation" And mCmbCategory.Value = "SelectionCategory" And NonCountedItem)
    mTxtCustomName.Value = "": mTxtCustomValue.Value = ""
    mBtnAddField_Click
    localFields = (mLstCustomFields.ListCount = 0 And mLblStatus.Caption <> "")
    mTxtCustomName.Value = "SelectionFieldSentinel": mTxtCustomValue.Value = "SelectionValueSentinel"
    mBtnAddField_Click
    localFields = localFields And (mLstCustomFields.ListCount = 1)
    mTxtCustomName.Value = "SelectionFieldSentinel": mTxtCustomValue.Value = "ReplacementValueSentinel"
    mBtnAddField_Click
    localFields = localFields And (mLstCustomFields.ListCount = 1 And mLstCustomFields.List(0, 1) = "ReplacementValueSentinel")
    mLstCustomFields.ListIndex = 0
    mBtnRemoveField_Click
    If mode = "InterruptBeforeLastRemove" Then modAuth.SignOut
    mBtnRemoveField_Click
    localFields = localFields And (mLstCustomFields.ListCount = 0)
    TestInventorySelections.Facts CStr(picked) & "|" & CStr(options) & "|" & CStr(localFields)
    mBtnCancel_Click
End Sub
'@)
}

function Initialize-InventorySelectionFixture($Fixture) {
    SelectTarget $Fixture
    $csv=Join-Path $runRoot 'selection-fixture.csv'
    @('ITEM_CODE,ITEM,UOM,LOCATION,QTY,CONDITION','DEMO-SELECT-01,Selection fixture,LB,CLEARVIEW,10,GOOD')|Set-Content $csv -Encoding ASCII
    $created=[bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.PrepareExactQuantityFixture' @($Fixture.Warehouse,$csv))
    Check 'InventorySelections.OwningGeneratedFixture' $created
    if(-not $created){throw 'Generated selection fixture unavailable; not product RED.'}
    $memory=Get-AdminInventoryEntryFacts $Fixture
    Write-Output ('Selection fixture in-memory saved: '+$memory.Saved)
    # As in the retained Edit tests, prove persisted seed results and close only
    # the disposable source before exercising a fresh interactive owner capture.
    $persisted=Get-AdminInventoryEntryFacts $Fixture -Persisted
    $book=$excel.Workbooks.Open((Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')),0,$true)
    $keys=@()
    try{
        $table=Table $book 'tblInventoryEntities'
        foreach($row in $table.ListRows){
            if([string]$row.Range.Cells.Item(1,$table.ListColumns.Item('SKU').Index).Value2 -ceq 'DEMO-SELECT-01' -and [string]$row.Range.Cells.Item(1,$table.ListColumns.Item('InventoryState').Index).Value2 -ceq 'ACTIVE'){
                $keys+=[string]$row.Range.Cells.Item(1,$table.ListColumns.Item('System_Key').Index).Value2
            }
        }
    }finally{$book.Close($false)}
    $ready=$keys.Count -eq 1
    if($ready){$ready=$persisted.EntityQuantities.ContainsKey($keys[0]) -and $persisted.EntityQuantities[$keys[0]] -eq 10}
    Check 'InventorySelections.PersistedGeneratedFixture' $ready
    if(-not $ready){throw 'Persisted generated selection fixture unavailable; not product RED.'}
}

function Test-InventorySelections($Fixture) {
    Initialize-InventorySelectionFixture $Fixture
    Check 'InventorySelections.Prior190DefinitionsPreserved' ([bool](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.Prior190Preserved'))
    $ids=@('SELECT_ITEM','SELECT_UOM','SELECT_QUANTITY_MODE','SELECT_LOCATION','SELECT_CATEGORY','SELECT_CUSTOM_FIELD','CUSTOM_FIELD_ADD','CUSTOM_FIELD_REMOVE')
    foreach($id in $ids){
        $wire=[string](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.Definition' @(('ADMIN_INVENTORY_'+$id),45))
        $row=if($wire){$wire|ConvertFrom-Json}else{$null}
        $kind=if($id -like 'SELECT_*'){'Navigation'}else{'Command'}
        Check ('InventorySelections.Catalog.'+$id) ($null -ne $row -and $row.OwnerId -ceq 'ADMIN_INVENTORY' -and $row.Capability -ceq 'ADMIN_MAINT' -and $row.Class -ceq $kind)
        Check ('InventorySelections.Version44Excludes.'+$id) ([string](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.Definition' @(('ADMIN_INVENTORY_'+$id),44)) -ceq '')
    }
    $operatorBook=$excel.Workbooks.Add();$operatorBook.SaveAs((Join-Path $runRoot 'inventory-selection-host.xlsx'),51)
    try {
        foreach($mode in @('Enabled','Disabled','ControlOff','UserOff')){
            SelectTarget $Fixture
            if(-not [bool](Run 'invSys.Core.xlam' 'TestInventoryRecordingPolicy.SetPolicy' @($mode))){throw 'Selection recording policy unavailable.'}
            $pins=@{};foreach($store in @('Config','Auth','Data.Inventory')){$path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.'+$store+'.xlsb');$pins[$path]=Get-AdminInventoryEntryHash $path}
            $before=@(Get-Slice4beActivityFiles $Fixture)
            [void](Run 'invSys.Admin.xlam' 'TestInventoryRecording.Execute' @($operatorBook.Name,$mode))
            $facts=[string](Run 'invSys.Admin.xlam' 'TestInventorySelections.Result')
            Write-Output ('Selection gesture flags: '+$facts)
            Write-Output ('Selection presentation count: '+[long](Run 'invSys.Admin.xlam' 'TestInventoryRecording.Gestures'))
            Write-Output ('Selection entry flags: '+[string](Run 'invSys.Admin.xlam' 'TestInventoryEntry.Facts'))
            Check ('InventorySelections.'+$mode+'.ActualOwnerHandlers') ($facts -ceq 'True|True|True' -and [bool](Run 'invSys.Admin.xlam' 'TestInventoryRecording.Finished'))
            Check ('InventorySelections.'+$mode+'.NoUnexpectedRibbonFailure') ([long](Run 'invSys.Admin.xlam' 'TestInventoryRecording.RibbonFailures') -eq 0)
            if($facts -cne 'True|True|True'){throw 'Selection gesture fixture did not complete; not product RED.'}
            $rows=@(Get-Slice4beActivityFiles $Fixture|Where-Object{$_ -cnotin $before}|ForEach-Object{Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object ControlId -Like 'ADMIN_INVENTORY_*')
            $enabled=$mode -ceq 'Enabled'
            Check ('InventorySelections.'+$mode+'.ExactRecordCount') ($rows.Count -eq $(if($enabled){30}else{0}))
            if($enabled){
                foreach($id in $ids){
                    $part=@($rows|Where-Object ControlId -CEQ ('ADMIN_INVENTORY_'+$id));$attempts=@($part|Where-Object OutcomeCode -CEQ 'REQUESTED');$ends=@($part|Where-Object OutcomeCode -CNE 'REQUESTED')
                    $expected=@(switch($id){'SELECT_ITEM'{@('SELECTED','SELECTED')} 'CUSTOM_FIELD_ADD'{@('REJECTED','STAGED','STAGED')} 'CUSTOM_FIELD_REMOVE'{@('STAGED','UNCHANGED')} default{@('SELECTED')}})
                    $codes=@($ends|ForEach-Object{$_.OutcomeCode}|Sort-Object);$wanted=@($expected|Sort-Object)
                    $valid=$attempts.Count -eq $expected.Count -and $ends.Count -eq $expected.Count -and ($codes -join '|') -ceq ($wanted -join '|')
                    foreach($ending in $ends){$valid=$valid -and @($attempts|Where-Object ActivityId -CEQ $ending.ActivityId).Count -eq 1 -and $ending.DataEffect -ceq 'Unchanged'}
                    Check ('InventorySelections.'+$id+'.ExactCorrelatedOutcomes') $valid
                    Check ('InventorySelections.'+$id+'.DistinctActions') (@($attempts|ForEach-Object{$_.ActivityId}|Sort-Object -Unique).Count -eq $expected.Count)
                }
            }
            $valid=$true;foreach($row in $rows){$valid=$valid -and $row.WarehouseId -ceq $Fixture.Warehouse -and $row.UserId -ceq 'config-admin' -and $row.OwnerId -ceq 'ADMIN_INVENTORY' -and @($row.SourceEventRefs).Count -eq 0}
            Check ('InventorySelections.'+$mode+'.CapturedContextNoCanonicalRefs') $valid
            $json=[string]($rows|ConvertTo-Json -Depth 12 -Compress)
            Check ('InventorySelections.'+$mode+'.Redacted') ($json -notmatch 'Sentinel|DEMO-SELECT-01|Selection fixture|SelectionLocation|SelectionCategory')
            $same=$true;foreach($path in $pins.Keys){$same=$same -and (Get-AdminInventoryEntryHash $path) -ceq $pins[$path]}
            Check ('InventorySelections.'+$mode+'.CanonicalBytesPreserved') $same
        }
    }finally{$operatorBook.Close($false)}
}
